// SPDX-License-Identifier: MIT OR Apache-2.0

// Package oxi is a Gotenberg module that converts .docx with the Oxi engine
// (oxidocs) on the LibreOffice route, and hands every other request to
// LibreOffice unchanged.
//
// It is a middleware on the multipart stack: it runs after Gotenberg has
// parsed the form and before the /forms/libreoffice/convert handler. A request
// it claims never reaches LibreOffice; a request it declines, or one Oxi
// cannot convert, goes on to the handler as if the module were absent. The
// response says which engine produced it:
//
//	Gotenberg-Engine: oxi | libreoffice
//	Gotenberg-Engine-Reason: why LibreOffice was used (absent for oxi)
//
// The module is off unless OXI_ENABLE=true (--oxi-enable). A client can send
// "Gotenberg-Oxi: off" to force LibreOffice for one request.
package oxi

import (
	"context"
	"errors"
	"fmt"
	"net/http"
	"os/exec"
	"slices"
	"sort"
	"strings"
	"time"

	"github.com/labstack/echo/v5"
	flag "github.com/spf13/pflag"

	"github.com/gotenberg/gotenberg/v9/pkg/gotenberg"
	"github.com/gotenberg/gotenberg/v9/pkg/modules/api"
	"github.com/gotenberg/gotenberg/v9/pkg/modules/pdf"
)

func init() {
	gotenberg.MustRegisterModule(new(Oxi))
}

const (
	engineHeader       = "Gotenberg-Engine"
	engineReasonHeader = "Gotenberg-Engine-Reason"
	optOutHeader       = "Gotenberg-Oxi"
	webhookUrlHeader   = "Gotenberg-Webhook-Url"
	convertRoutePath   = "/forms/libreoffice/convert"
)

// libreOfficeFields are the form fields the Gotenberg 9 LibreOffice route
// reads besides "stages". Each one asks LibreOffice for something Oxi does
// not do, so a request that sets one to anything but its default goes to
// LibreOffice. A field outside this list, and outside "stages", also goes
// there, so that the route rejects it with its usual 400.
var libreOfficeFields = []string{
	"addOriginalDocumentAsStream", "allowDuplicateFieldNames", "centerWindow", "convertOooTargetToPdfTarget",
	"displayPDFDocumentTitle", "exportBookmarksToPdfDestination", "exportFormFields", "exportHiddenSlides",
	"exportLinksRelativeFsys", "exportNotes", "exportNotesInMargin", "exportNotesPages", "exportOnlyNotesPages",
	"exportPlaceholders", "hideViewerMenubar", "hideViewerToolbar", "hideViewerWindowControls", "initialPage",
	"initialView", "landscape", "magnification", "nativeBookmarks", "nativeLosslessImageCompression",
	"nativeMaxImageResolution", "nativePages", "nativePdfa", "nativePdfua", "nativeQuality",
	"nativeReduceImageResolution", "nativeTaggedPdf", "nativeTiledWatermarkText", "nativeWatermarkColor",
	"nativeWatermarkFont", "nativeWatermarkFontSize", "nativeWatermarkRotation", "nativeWatermarkText",
	"openBookmarkLevels", "openInFullScreenMode", "pageLayout", "paperHeight", "paperSize", "paperWidth",
	"password", "resizeWindowToInitialPage", "singlePage", "skipEmptyPages", "updateIndexes",
	"useTransitionEffects", "zoom",
}

// passThroughFields are handled the same whichever engine converts:
// "stages" runs on Oxi's PDFs exactly as on LibreOffice's, and the
// downloads of "downloadFrom" are ordinary files by the time this runs.
var passThroughFields = []string{pdf.StagesFormField, "downloadFrom"}

// Oxi is the module.
type Oxi struct {
	enabled  bool
	binPath  string
	threads  int
	timeout  time.Duration
	fallback bool

	processor *pdf.Processor
	server    *server
}

// Descriptor returns the module's descriptor.
func (o *Oxi) Descriptor() gotenberg.ModuleDescriptor {
	return gotenberg.ModuleDescriptor{
		ID: "oxi",
		FlagSet: func() *flag.FlagSet {
			fs := flag.NewFlagSet("oxi", flag.ExitOnError)
			fs.Bool("oxi-enable", false, "Convert .docx on the LibreOffice route with the Oxi engine")
			fs.String("oxi-bin-path", "oxidocs", "Path to the oxidocs binary")
			fs.Int("oxi-threads", 2, "Documents oxidocs converts in parallel")
			fs.Duration("oxi-timeout", 0, "Time limit for one document in oxidocs, retry included; 0 gives it 90% of what is left of the request's --api-timeout")
			fs.Bool("oxi-fallback", true, "Hand documents Oxi cannot convert to LibreOffice; when false, answer 422 (unsupported content) or 500 (failure) instead")
			return fs
		}(),
		New: func() gotenberg.Module { return new(Oxi) },
	}
}

// Provision reads the flags and gets the stage processor Gotenberg runs after
// every conversion.
func (o *Oxi) Provision(ctx *gotenberg.Context) error {
	flags := ctx.ParsedFlags()
	o.enabled = flags.MustBool("oxi-enable")
	o.binPath = flags.MustString("oxi-bin-path")
	o.threads = flags.MustInt("oxi-threads")
	o.timeout = flags.MustDuration("oxi-timeout")
	o.fallback = flags.MustBool("oxi-fallback")
	if !o.enabled {
		return nil
	}

	provider, err := ctx.Module(new(pdf.Provider))
	if err != nil {
		return fmt.Errorf("get PDF provider: %w", err)
	}
	o.processor = provider.(pdf.Provider).Processor()

	o.server = newServer(o.binPath, o.threads, gotenberg.Logger(o))
	return nil
}

// Validate checks the binary when the module is on.
func (o *Oxi) Validate() error {
	if !o.enabled {
		return nil
	}
	if _, err := exec.LookPath(o.binPath); err != nil {
		return fmt.Errorf("oxidocs binary %q: %w; check --oxi-bin-path (OXI_BIN_PATH)", o.binPath, err)
	}
	if o.threads < 1 {
		return errors.New("--oxi-threads (OXI_THREADS) must be at least 1")
	}
	return nil
}

// Start does nothing: oxidocs starts on the first .docx, so a deployment
// that never sees one never pays for it.
func (o *Oxi) Start() error { return nil }

// StartupMessage tells whether .docx goes to Oxi.
func (o *Oxi) StartupMessage() string {
	if !o.enabled {
		return "off; .docx goes to LibreOffice (set OXI_ENABLE=true to change)"
	}
	return fmt.Sprintf(".docx on %s goes to %s (fallback to LibreOffice: %t)", convertRoutePath, o.binPath, o.fallback)
}

// Stop ends the oxidocs process.
func (o *Oxi) Stop(_ context.Context) error {
	if o.server != nil {
		o.server.stop()
	}
	return nil
}

// Middlewares returns the middleware.
func (o *Oxi) Middlewares() ([]api.Middleware, error) {
	if !o.enabled {
		return nil, nil
	}
	return []api.Middleware{{
		Stack:   api.MultipartStack,
		Handler: o.middleware,
	}}, nil
}

func (o *Oxi) middleware(next echo.HandlerFunc) echo.HandlerFunc {
	return func(c *echo.Context) error {
		if !strings.HasSuffix(c.Request().URL.Path, convertRoutePath) {
			return next(c)
		}
		ctx, ok := c.Get("context").(*api.Context)
		if !ok {
			return next(c)
		}
		useLibreOffice := func(reason string) error {
			ctx.SetOutputHeader(engineHeader, "libreoffice")
			ctx.SetOutputHeader(engineReasonHeader, reason)
			ctx.Log().Info(fmt.Sprintf("oxi: LibreOffice converts this request (%s)", reason))
			return next(c)
		}

		inputs, stages, stageFiles, reason := o.claim(c, ctx)
		if reason != "" {
			return useLibreOffice(reason)
		}

		outputs := make([]string, len(inputs))
		for i, input := range inputs {
			// document.docx -> document.docx.pdf, as the LibreOffice route names it.
			outputs[i] = ctx.GeneratePathFromFilename(ctx.OriginalFilename(input) + ".pdf")
			// One time budget for the document, retry included.
			budget := o.budget(ctx)
			deadline := time.Now().Add(budget)
			r, err := o.server.convert(ctx, input, outputs[i], true, budget)
			if left := time.Until(deadline); errors.Is(err, errServerGone) && left > time.Second {
				// The process went down under this request, most often for
				// another one that ran out of time. Once more on a fresh one.
				r, err = o.server.convert(ctx, input, outputs[i], true, left)
			}
			filename := ctx.OriginalFilename(input)
			switch {
			case err != nil && errors.Is(err, context.Canceled):
				return err
			case err != nil:
				ctx.Log().Warn(fmt.Sprintf("oxi: '%s' failed: %s", filename, err))
				if !o.fallback {
					return api.WrapError(fmt.Errorf("oxi: %w", err), api.NewSentinelHttpError(http.StatusInternalServerError,
						fmt.Sprintf("Oxi failed to convert the document '%s'.", filename)))
				}
				if errors.Is(err, errTimeout) {
					return useLibreOffice("oxi-timeout")
				}
				return useLibreOffice("oxi-failed")
			case r.OK:
				ctx.Log().Debug(fmt.Sprintf("oxi: '%s' converted, %d pages in %.0f ms", filename, r.Pages, r.Ms))
			case r.Exit == 3:
				kinds := unsupportedKinds(r.Unsupported)
				if !o.fallback {
					return api.WrapError(fmt.Errorf("oxi: unsupported content: %s", kinds), api.NewSentinelHttpError(http.StatusUnprocessableEntity,
						fmt.Sprintf("The document '%s' holds content Oxi does not draw: %s.", filename, kinds)))
				}
				return useLibreOffice("unsupported-content: " + kinds)
			default:
				ctx.Log().Warn(fmt.Sprintf("oxi: '%s' failed (exit %d): %s", filename, r.Exit, r.Error))
				if !o.fallback {
					return api.WrapError(fmt.Errorf("oxi: exit %d: %s", r.Exit, r.Error), api.NewSentinelHttpError(http.StatusInternalServerError,
						fmt.Sprintf("Oxi failed to convert the document '%s'.", filename)))
				}
				return useLibreOffice("oxi-failed")
			}
		}

		outputs, err := o.processor.Run(ctx, stages, outputs, stageFiles)
		if err != nil {
			return fmt.Errorf("run stages: %w", err)
		}
		ctx.SetOutputHeader(engineHeader, "oxi")
		return ctx.AddOutputPaths(outputs...)
	}
}

// claim decides whether Oxi takes the request. It returns the .docx inputs
// and the stages to run on them, or the reason LibreOffice should take it.
func (o *Oxi) claim(c *echo.Context, ctx *api.Context) (inputs []string, stages pdf.Stages, stageFiles map[string]string, reason string) {
	req := c.Request()
	if strings.EqualFold(req.Header.Get(optOutHeader), "off") {
		return nil, nil, nil, "requested"
	}
	if req.Header.Get(webhookUrlHeader) != "" {
		// The webhook middleware and this one share the lowest priority, so
		// their order is not fixed; claiming here could answer synchronously
		// a request that asked for a callback.
		return nil, nil, nil, "webhook"
	}

	if req.MultipartForm != nil {
		keys := make([]string, 0, len(req.MultipartForm.Value))
		for key := range req.MultipartForm.Value {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			if slices.Contains(passThroughFields, key) {
				continue
			}
			if !slices.Contains(libreOfficeFields, key) {
				return nil, nil, nil, "unknown-field: " + key
			}
			value := strings.TrimSpace(req.MultipartForm.Value[key][0])
			if value != "" && !strings.EqualFold(value, "false") {
				return nil, nil, nil, "option: " + key
			}
		}
	}

	form := ctx.FormData()
	stages, stageFiles, err := pdf.FormDataStages(form)
	if err != nil || form.Validate() != nil {
		// Let the LibreOffice route report the invalid stages as it always has.
		return nil, nil, nil, "invalid-form"
	}
	form.Paths([]string{".docx"}, &inputs)
	if len(inputs) == 0 {
		return nil, nil, nil, "no-docx"
	}
	if ctx.FileCount() != len(inputs)+len(stageFiles) {
		return nil, nil, nil, "other-files"
	}
	return inputs, stages, stageFiles, ""
}

// budget is how long oxidocs may take over one document. By default that is
// most of what the request has left: a document too slow for Oxi -- in
// practice a reference work of ten thousand paragraphs -- is as slow for
// LibreOffice, so stopping Oxi early only adds the wait to LibreOffice's.
// Content Oxi does not draw and failures still fall back at once.
func (o *Oxi) budget(ctx context.Context) time.Duration {
	if o.timeout > 0 {
		return o.timeout
	}
	if deadline, ok := ctx.Deadline(); ok {
		if left := time.Until(deadline) * 9 / 10; left > time.Second {
			return left
		}
		return time.Second
	}
	return time.Minute
}

func unsupportedKinds(list []unsupported) string {
	var kinds []string
	for _, u := range list {
		if !slices.Contains(kinds, u.Kind) {
			kinds = append(kinds, u.Kind)
		}
	}
	if len(kinds) == 0 {
		return "unknown"
	}
	return strings.Join(kinds, ", ")
}

// Interface guards.
var (
	_ gotenberg.Module       = (*Oxi)(nil)
	_ gotenberg.Provisioner  = (*Oxi)(nil)
	_ gotenberg.Validator    = (*Oxi)(nil)
	_ gotenberg.App          = (*Oxi)(nil)
	_ api.MiddlewareProvider = (*Oxi)(nil)
)
