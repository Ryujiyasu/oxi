// SPDX-License-Identifier: MIT OR Apache-2.0

package oxi

import (
	"context"
	"errors"
	"log/slog"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
	"time"
	"unsafe"

	"github.com/labstack/echo/v5"

	"github.com/gotenberg/gotenberg/v9/pkg/gotenberg"
	"github.com/gotenberg/gotenberg/v9/pkg/modules/api"
	"github.com/gotenberg/gotenberg/v9/pkg/modules/pdf"
)

// The tests run the real engine. OXI_TEST_BIN is the oxidocs binary,
// OXI_TEST_DOCX a document it converts, OXI_TEST_UNSUPPORTED_DOCX one that
// --strict refuses.
func fixtures(t *testing.T) (bin, docx, unsupported string) {
	t.Helper()
	bin, docx, unsupported = os.Getenv("OXI_TEST_BIN"), os.Getenv("OXI_TEST_DOCX"), os.Getenv("OXI_TEST_UNSUPPORTED_DOCX")
	if bin == "" || docx == "" || unsupported == "" {
		t.Skip("set OXI_TEST_BIN, OXI_TEST_DOCX and OXI_TEST_UNSUPPORTED_DOCX")
	}
	return bin, docx, unsupported
}

func copyInto(t *testing.T, dir, src, name string) string {
	t.Helper()
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatal(err)
	}
	dst := filepath.Join(dir, name)
	if err := os.WriteFile(dst, data, 0o600); err != nil {
		t.Fatal(err)
	}
	return dst
}

type result struct {
	err         error
	libreOffice bool
	engine      string
	reason      string
	residual    string
	outputs     []string
	names       []string
}

// initFilenames gives the mock the map behind OriginalFilename, which the
// real request context always has; ContextMock has no setter for it.
func initFilenames(ctx *api.ContextMock) {
	m := reflect.ValueOf(ctx.Context).Elem().FieldByName("diskToOriginal")
	reflect.NewAt(m.Type(), unsafe.Pointer(m.UnsafeAddr())).Elem().Set(reflect.ValueOf(map[string]string{}))
}

// run sends one request through the middleware. files maps an uploaded
// filename to its source on disk.
func run(t *testing.T, o *Oxi, files map[string]string, values map[string][]string, headers map[string]string) result {
	t.Helper()
	dir := t.TempDir()
	onDisk := make(map[string]string)
	for name, src := range files {
		onDisk[name] = copyInto(t, dir, src, name)
	}
	ctx := &api.ContextMock{Context: new(api.Context)}
	ctx.SetDirPath(dir)
	ctx.SetFiles(onDisk)
	ctx.SetValues(values)
	ctx.SetLogger(slog.New(slog.DiscardHandler))
	ctx.SetMkdirAll(new(gotenberg.OsMkdirAll))
	ctx.SetPathRename(new(gotenberg.OsPathRename))
	ctx.Context.Context = context.Background()
	initFilenames(ctx)

	req := httptest.NewRequest(http.MethodPost, convertRoutePath, nil)
	req.MultipartForm = &multipart.Form{Value: values}
	for k, v := range headers {
		req.Header.Set(k, v)
	}
	rec := httptest.NewRecorder()
	c := echo.New().NewContext(req, rec)
	ctx.SetEchoContext(c)
	c.Set("context", ctx.Context)

	var r result
	r.err = o.middleware(func(*echo.Context) error {
		r.libreOffice = true
		return nil
	})(c)
	headers = ctx.OutputHeaders()
	r.engine = headers[engineHeader]
	r.reason = headers[engineReasonHeader]
	r.residual = headers[pdf.ResidualHeader]
	r.outputs = ctx.OutputPaths()
	for _, out := range r.outputs {
		r.names = append(r.names, ctx.OriginalFilename(out))
	}
	return r
}

func newOxi(bin string, fallback bool) *Oxi {
	return &Oxi{
		enabled: true, binPath: bin, threads: 2, timeout: 60 * time.Second, fallback: fallback,
		processor: new(pdf.Processor), server: newServer(bin, 2, slog.New(slog.DiscardHandler)),
	}
}

func TestMiddleware(t *testing.T) {
	bin, docx, unsupported := fixtures(t)
	o := newOxi(bin, true)
	defer o.server.stop()

	cases := []struct {
		name       string
		files      map[string]string
		values     map[string][]string
		headers    map[string]string
		wantEngine string
		wantReason string
	}{
		{"docx goes to oxi", map[string]string{"a.docx": docx}, nil, nil, "oxi", ""},
		{"defaults sent explicitly stay on oxi", map[string]string{"a.docx": docx},
			map[string][]string{"landscape": {"false"}, "password": {""}}, nil, "oxi", ""},
		{"an option oxi lacks goes to libreoffice", map[string]string{"a.docx": docx},
			map[string][]string{"landscape": {"true"}}, nil, "libreoffice", "option: landscape"},
		{"native pdf/a goes to libreoffice", map[string]string{"a.docx": docx},
			map[string][]string{"nativePdfa": {"PDF/A-2b"}}, nil, "libreoffice", "option: nativePdfa"},
		{"an unknown field goes to libreoffice, which rejects it", map[string]string{"a.docx": docx},
			map[string][]string{"merge": {"true"}}, nil, "libreoffice", "unknown-field: merge"},
		{"unsupported content goes to libreoffice", map[string]string{"a.docx": unsupported}, nil, nil,
			"libreoffice", "unsupported-content: "},
		{"other formats go to libreoffice", map[string]string{"a.xlsx": docx}, nil, nil, "libreoffice", "no-docx"},
		{"a mixed upload goes to libreoffice", map[string]string{"a.docx": docx, "b.xlsx": docx}, nil, nil,
			"libreoffice", "other-files"},
		{"the client can opt out", map[string]string{"a.docx": docx}, nil,
			map[string]string{optOutHeader: "off"}, "libreoffice", "requested"},
		{"webhook requests go to libreoffice", map[string]string{"a.docx": docx}, nil,
			map[string]string{webhookUrlHeader: "http://example.invalid"}, "libreoffice", "webhook"},
		{"a broken docx falls back", map[string]string{"a.docx": "oxi_test.go"}, nil, nil, "libreoffice", "oxi-failed"},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			r := run(t, o, tc.files, tc.values, tc.headers)
			if r.err != nil {
				t.Fatalf("error: %v", r.err)
			}
			if r.engine != tc.wantEngine || !strings.HasPrefix(r.reason, tc.wantReason) {
				t.Fatalf("engine %q reason %q, want %q %q", r.engine, r.reason, tc.wantEngine, tc.wantReason)
			}
			if r.libreOffice != (tc.wantEngine == "libreoffice") {
				t.Fatalf("LibreOffice handler called: %t", r.libreOffice)
			}
			if tc.wantEngine == "oxi" {
				if len(r.outputs) != 1 {
					t.Fatalf("outputs %v", r.outputs)
				}
				head := make([]byte, 5)
				f, err := os.Open(r.outputs[0])
				if err != nil {
					t.Fatal(err)
				}
				defer f.Close()
				if _, err := f.Read(head); err != nil || string(head) != "%PDF-" {
					t.Fatalf("output is not a PDF: %q %v", head, err)
				}
			}
		})
	}
}

func TestStagesRunOnOxiOutput(t *testing.T) {
	bin, docx, _ := fixtures(t)
	o := newOxi(bin, true)
	defer o.server.stop()
	files := map[string]string{"a.docx": docx, "b.docx": docx}

	r := run(t, o, files, nil, nil)
	if r.err != nil || r.engine != "oxi" || len(r.outputs) != 2 {
		t.Fatalf("two outputs: err %v engine %q outputs %v", r.err, r.engine, r.outputs)
	}
	for _, name := range r.names {
		if name != "a.docx.pdf" && name != "b.docx.pdf" {
			t.Fatalf("output named %q, want <original>.pdf as LibreOffice names it", name)
		}
	}

	r = run(t, o, files, map[string][]string{"stages": {`[{"op":"merge"}]`}}, nil)
	if r.err != nil || r.engine != "oxi" || len(r.outputs) != 1 {
		t.Fatalf("merge: err %v engine %q outputs %v", r.err, r.engine, r.outputs)
	}

	r = run(t, o, map[string]string{"a.docx": docx}, map[string][]string{"stages": {`[{"op":"remediate","profile":"PDF/A-2b"}]`}}, nil)
	if r.err != nil || r.engine != "oxi" || len(r.outputs) != 1 {
		t.Fatalf("remediate: err %v engine %q outputs %v", r.err, r.engine, r.outputs)
	}
	t.Logf("PDF/A-2b remediation of Oxi's output: %s violations left", r.residual)
	if r.residual != "0" {
		t.Errorf("remediated Oxi output still has %q PDF/A-2b violations", r.residual)
	}

	r = run(t, o, map[string]string{"a.docx": docx}, map[string][]string{"stages": {`[{"op":"nope"}]`}}, nil)
	if r.engine != "libreoffice" || r.reason != "invalid-form" {
		t.Fatalf("bad stages: engine %q reason %q", r.engine, r.reason)
	}
}

func TestStrictWithoutFallback(t *testing.T) {
	bin, _, unsupported := fixtures(t)
	o := newOxi(bin, false)
	defer o.server.stop()

	r := run(t, o, map[string]string{"a.docx": unsupported}, nil, nil)
	if r.libreOffice || r.err == nil {
		t.Fatalf("want an error and no LibreOffice, got err %v libreoffice %t", r.err, r.libreOffice)
	}
	var httpErr api.HttpError
	if !errors.As(r.err, &httpErr) {
		t.Fatalf("want an HTTP error, got %v", r.err)
	}
	if code, msg := httpErr.HttpError(); code != http.StatusUnprocessableEntity {
		t.Fatalf("want 422, got %d %q", code, msg)
	}
}

func TestServerRestartsAfterCrash(t *testing.T) {
	bin, docx, _ := fixtures(t)
	s := newServer(bin, 1, slog.New(slog.DiscardHandler))
	defer s.stop()
	out := filepath.Join(t.TempDir(), "a.pdf")

	if r, err := s.convert(context.Background(), docx, out, true, time.Minute); err != nil || !r.OK {
		t.Fatalf("first: %+v %v", r, err)
	}
	s.mu.Lock()
	first := s.proc
	s.mu.Unlock()
	// Kill it behind the server's back, as a crash would.
	_ = first.cmd.Process.Kill()
	<-first.done
	if r, err := s.convert(context.Background(), docx, out, true, time.Minute); err != nil || !r.OK {
		t.Fatalf("after crash: %+v %v", r, err)
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	if s.proc == first {
		t.Fatal("process was not restarted")
	}
}

func TestServerTimeoutKillsProcess(t *testing.T) {
	bin, docx, _ := fixtures(t)
	s := newServer(bin, 1, slog.New(slog.DiscardHandler))
	defer s.stop()
	out := filepath.Join(t.TempDir(), "a.pdf")
	if _, err := s.convert(context.Background(), docx, out, true, time.Nanosecond); err == nil {
		t.Fatal("want a timeout")
	}
	if r, err := s.convert(context.Background(), docx, out, true, time.Minute); err != nil || !r.OK {
		t.Fatalf("after timeout: %+v %v", r, err)
	}
}
