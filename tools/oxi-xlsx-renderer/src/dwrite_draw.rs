// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! Draws a sheet through Direct2D and DirectWrite, beside the GDI path.
//!
//! The geometry is the same one — every rule derived from Excel applies to
//! both. What differs is the engine the glyphs come out of, and that engine
//! is why this is no longer the path that runs: Excel's sheet text is GDI's
//! own ClearType, and Direct2D over a WIC bitmap cannot be asked for it.
//! This is kept reachable with `OXI_XLSX_DWRITE=1`, because a second engine
//! that agrees about geometry is the cheapest check there is on the first.

use oxicells_core::ir::{CellValue, Sheet};
use windows::core::{Interface, Result, HSTRING, PCWSTR};
use windows::Win32::Graphics::Direct2D::Common::*;
use windows::Win32::Graphics::Direct2D::*;
use windows::Win32::Graphics::DirectWrite::*;
use windows::Win32::Graphics::Imaging::*;
use windows::Win32::System::Com::*;

use crate::{
    alignment, cell_text, dressed_by_table, has_filter_button, merges, room_after, room_before,
    rule_for, Align, Broken, Geometry, Merged, FILTER_BUTTON, FILTER_BUTTON_FOOT,
};

fn shade(hex: Option<&str>, fallback: u32) -> D2D1_COLOR_F {
    let value = hex
        .and_then(|hex| u32::from_str_radix(hex.trim_start_matches('#'), 16).ok())
        .unwrap_or(fallback);
    D2D1_COLOR_F {
        r: ((value >> 16) & 0xFF) as f32 / 255.0,
        g: ((value >> 8) & 0xFF) as f32 / 255.0,
        b: (value & 0xFF) as f32 / 255.0,
        a: 1.0,
    }
}

/// How advances are measured. Excel is an old application drawing through a
/// modern engine, so which of these it asks for is worth checking rather than
/// assuming.
fn measuring_mode() -> DWRITE_MEASURING_MODE {
    match std::env::var("OXI_XLSX_MEASURE").as_deref() {
        Ok("natural") => DWRITE_MEASURING_MODE_NATURAL,
        Ok("gdi_natural") => DWRITE_MEASURING_MODE_GDI_NATURAL,
        _ => DWRITE_MEASURING_MODE_GDI_CLASSIC,
    }
}

/// How tall to ask for a font, in pixels.
///
/// Excel asks GDI for its fonts, and GDI takes a whole number of pixels: ten
/// point at 96dpi is thirteen pixels, not thirteen and a third. Asking for the
/// third draws ＭＳ ゴシック a pixel larger in each direction than Excel does,
/// which is visible in every glyph on the sheet.
fn em_pixels(points: f32, scale: f32) -> f32 {
    (points * scale * 96.0 / 72.0).round()
}

fn box_of(left: f32, top: f32, right: f32, bottom: f32) -> D2D_RECT_F {
    D2D_RECT_F {
        left,
        top,
        right,
        bottom,
    }
}

pub fn draw(
    sheet: &Sheet,
    layout: &Geometry,
    width: u32,
    height: u32,
    scale: f32,
) -> Result<image::RgbImage> {
    unsafe {
        let _ = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
        let wic: IWICImagingFactory =
            CoCreateInstance(&CLSID_WICImagingFactory, None, CLSCTX_INPROC_SERVER)?;
        let d2d: ID2D1Factory =
            D2D1CreateFactory::<ID2D1Factory>(D2D1_FACTORY_TYPE_SINGLE_THREADED, None)?;
        let writer: IDWriteFactory = DWriteCreateFactory(DWRITE_FACTORY_TYPE_SHARED)?;

        let bitmap: IWICBitmap = wic.CreateBitmap(
            width,
            height,
            &GUID_WICPixelFormat32bppPBGRA,
            WICBitmapCacheOnLoad,
        )?;
        let props = D2D1_RENDER_TARGET_PROPERTIES {
            r#type: D2D1_RENDER_TARGET_TYPE_DEFAULT,
            pixelFormat: D2D1_PIXEL_FORMAT {
                format: windows::Win32::Graphics::Dxgi::Common::DXGI_FORMAT_B8G8R8A8_UNORM,
                alphaMode: D2D1_ALPHA_MODE_PREMULTIPLIED,
            },
            dpiX: 96.0,
            dpiY: 96.0,
            usage: D2D1_RENDER_TARGET_USAGE_NONE,
            minLevel: D2D1_FEATURE_LEVEL_DEFAULT,
        };
        let target: ID2D1RenderTarget = d2d.CreateWicBitmapRenderTarget(&bitmap, &props)?;

        target.BeginDraw();
        target.Clear(Some(&shade(None, 0x00FF_FFFF)));
        // Excel's own picture of a sheet carries coloured glyph edges, so this
        // asks for ClearType rather than the greyscale the docx side wants
        // against Word's EMF.
        target.SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_CLEARTYPE);
        // Every glyph lands in the same place as Excel's and the edges still
        // differ, because Excel's are GDI's own ClearType bitmaps. Asking for
        // GDI-compatible glyphs here does nothing: a Direct2D
        // target over a WIC bitmap ignores its text rendering parameters
        // altogether — gamma swept from 1.0 to 2.4 leaves the picture
        // byte-identical. Excel's edges are GDI's own ClearType, which only
        // the GDI path can lay down; this is left in as the record of what
        // was tried.
        if let Ok(display) = writer.CreateRenderingParams() {
            if let Ok(alike) = writer.CreateCustomRenderingParams(
                display.GetGamma(),
                display.GetEnhancedContrast(),
                display.GetClearTypeLevel(),
                display.GetPixelGeometry(),
                DWRITE_RENDERING_MODE_GDI_CLASSIC,
            ) {
                target.SetTextRenderingParams(&alike);
            }
        }

        let brush = target.CreateSolidColorBrush(&shade(None, 0), None)?;
        let merged = merges(sheet);

        // A table dresses the whole of its range, and not only the cells a
        // sheet happens to record: doi-list bands eight columns whose rows
        // stop recording cells two thirds of the way across, and Excel bands
        // the rest of them all the same. Painted before the cells, so a cell
        // with a fill of its own still covers it.
        let recorded: std::collections::HashSet<(u32, u32)> = sheet
            .rows
            .iter()
            .flat_map(|row| row.cells.iter().map(move |cell| (row.index, cell.col)))
            .collect();
        for table in &sheet.tables {
            for row in table.start_row..=table.end_row {
                let top_at = row.saturating_sub(layout.first_row) as usize;
                if row < layout.first_row {
                    continue;
                }
                let (Some(top), Some(bottom)) =
                    (layout.rows.get(top_at), layout.rows.get(top_at + 1))
                else {
                    continue;
                };
                for column in table.start_col..=table.end_col {
                    if column < layout.first_column
                        || recorded.contains(&(row, column))
                        || matches!(merged.get(&(row, column)), Some(Merged::Covered))
                    {
                        continue;
                    }
                    let left_at = (column - layout.first_column) as usize;
                    let (Some(left), Some(right)) =
                        (layout.columns.get(left_at), layout.columns.get(left_at + 1))
                    else {
                        continue;
                    };
                    let Some(fill) = dressed_by_table(sheet, row, column).and_then(|d| d.fill)
                    else {
                        continue;
                    };
                    brush.SetColor(&shade(Some(&fill), 0x00FF_FFFF));
                    target.FillRectangle(&box_of(*left, *top, *right, *bottom), &brush);
                }
            }
        }

        for row in &sheet.rows {
            if row.index < layout.first_row {
                continue;
            }
            let top_at = (row.index - layout.first_row) as usize;
            let (Some(top), Some(bottom)) = (layout.rows.get(top_at), layout.rows.get(top_at + 1))
            else {
                continue;
            };
            if bottom <= top {
                continue;
            }
            for cell in &row.cells {
                if cell.col < layout.first_column {
                    continue;
                }
                let (spans_columns, spans_rows) = match merged.get(&(row.index, cell.col)) {
                    Some(Merged::Covered) => continue,
                    Some(Merged::Anchor { columns, rows }) => (*columns, *rows),
                    None => (0, 0),
                };
                let left_at = (cell.col - layout.first_column) as usize;
                let (Some(left), Some(right)) = (
                    layout.columns.get(left_at),
                    layout.columns.get(left_at + 1 + spans_columns as usize),
                ) else {
                    continue;
                };
                if right <= left {
                    continue;
                }
                let bottom = layout
                    .rows
                    .get(top_at + 1 + spans_rows as usize)
                    .unwrap_or(bottom);
                let cell_box = box_of(*left, *top, *right, *bottom);

                let dress = dressed_by_table(sheet, row.index, cell.col);
                let fill = cell
                    .style
                    .bg_color
                    .as_deref()
                    .or_else(|| dress.as_ref().and_then(|d| d.fill.as_deref()));
                if let Some(fill) = fill {
                    brush.SetColor(&shade(Some(fill), 0x00FF_FFFF));
                    target.FillRectangle(&cell_box, &brush);
                }

                let edges = [
                    (&cell.style.border_top, true, cell_box.top),
                    (&cell.style.border_bottom, true, cell_box.bottom),
                    (&cell.style.border_left, false, cell_box.left),
                    (&cell.style.border_right, false, cell_box.right),
                ];
                for (line, horizontal, at) in edges {
                    let Some(line) = line else { continue };
                    let rule = rule_for(&line.style);
                    brush.SetColor(&shade(line.color.as_deref(), 0));
                    for step in -rule.before..=rule.after {
                        if rule.hollow && step == 0 {
                            continue;
                        }
                        let at = at + step as f32;
                        let (start, stop) = if horizontal {
                            (cell_box.left, cell_box.right)
                        } else {
                            (cell_box.top, cell_box.bottom)
                        };
                        match rule.broken {
                            Broken::Whole => {
                                let edge = if horizontal {
                                    box_of(start, at, stop, at + 1.0)
                                } else {
                                    box_of(at, start, at + 1.0, stop)
                                };
                                target.FillRectangle(&edge, &brush);
                            }
                            broken => {
                                for along in start as i32..stop as i32 {
                                    let (x, y) = if horizontal {
                                        (along as f32, at)
                                    } else {
                                        (at, along as f32)
                                    };
                                    if broken.inked(x as i32, y as i32, along) {
                                        target.FillRectangle(
                                            &box_of(x, y, x + 1.0, y + 1.0), &brush);
                                    }
                                }
                            }
                        }
                    }
                }

                let filtered = has_filter_button(sheet, row.index, cell.col);
                if filtered {
                    let left = cell_box.right - (FILTER_BUTTON + 1) as f32;
                    // The button hangs from the foot of the heading.
                    let top = cell_box.bottom - (FILTER_BUTTON_FOOT + FILTER_BUTTON) as f32;
                    let face = box_of(left, top, cell_box.right - 1.0, top + FILTER_BUTTON as f32);
                    brush.SetColor(&shade(Some("A6ACB3"), 0));
                    target.FillRectangle(&face, &brush);
                    brush.SetColor(&shade(Some("FEFEFE"), 0x00FF_FFFF));
                    target.FillRectangle(
                        &box_of(
                            face.left + 1.0,
                            face.top + 1.0,
                            face.right - 1.0,
                            face.bottom - 1.0,
                        ),
                        &brush,
                    );
                    brush.SetColor(&shade(Some("58595B"), 0));
                    let middle = (face.left + face.right) / 2.0;
                    for step in 0..4 {
                        let half = (3 - step) as f32;
                        target.FillRectangle(
                            &box_of(
                                middle - half,
                                face.top + 7.0 + step as f32,
                                middle + half + 1.0,
                                face.top + 8.0 + step as f32,
                            ),
                            &brush,
                        );
                    }
                }

                let text = cell_text(&cell.value, &cell.style).replace("\r\n", "\n");
                let text = if cell.style.stacked_text {
                    crate::stacked_text(&text)
                } else {
                    text
                };
                if text.is_empty() {
                    continue;
                }
                let points = cell.style.font_size.unwrap_or(11.0);
                let face = HSTRING::from(cell.style.font_name.as_deref().unwrap_or("Calibri"));
                // A table's header row is bold, and no cell inside the range
                // says so in its own style.
                let bold = cell.style.bold || matches!(&dress, Some(dress) if dress.bold);
                let format = writer.CreateTextFormat(
                    PCWSTR(face.as_ptr()),
                    None,
                    if bold {
                        DWRITE_FONT_WEIGHT_BOLD
                    } else {
                        DWRITE_FONT_WEIGHT_NORMAL
                    },
                    if cell.style.italic {
                        DWRITE_FONT_STYLE_ITALIC
                    } else {
                        DWRITE_FONT_STYLE_NORMAL
                    },
                    DWRITE_FONT_STRETCH_NORMAL,
                    em_pixels(points, scale),
                    &HSTRING::from("ja-JP"),
                )?;

                let placed = alignment(&cell.style, &cell.value);
                format.SetTextAlignment(match placed {
                    Align::Left => DWRITE_TEXT_ALIGNMENT_LEADING,
                    Align::Centre => DWRITE_TEXT_ALIGNMENT_CENTER,
                    Align::Right => DWRITE_TEXT_ALIGNMENT_TRAILING,
                    Align::Spread => DWRITE_TEXT_ALIGNMENT_JUSTIFIED,
                })?;
                format.SetWordWrapping(if cell.style.wrap_text {
                    DWRITE_WORD_WRAPPING_WRAP
                } else {
                    DWRITE_WORD_WRAPPING_NO_WRAP
                })?;
                format.SetParagraphAlignment(match cell.style.vertical_align.as_deref() {
                    Some("top") => DWRITE_PARAGRAPH_ALIGNMENT_NEAR,
                    Some("center") | Some("centre") => DWRITE_PARAGRAPH_ALIGNMENT_CENTER,
                    _ => DWRITE_PARAGRAPH_ALIGNMENT_FAR,
                })?;

                // A line of text stands in the box Excel gives it — the row a
                // sheet of this font alone would have — with its baseline where
                // Excel puts it, rather than wherever this engine's own font
                // metrics would land. Measured against Excel's picture across
                // twenty faces; see `line_box`.
                let placed_line = crate::line_box(
                    cell.style.font_name.as_deref().unwrap_or("Calibri"),
                    points,
                    bold,
                    cell.style.italic,
                );
                if let Some((line, baseline)) = placed_line {
                    format.SetLineSpacing(
                        DWRITE_LINE_SPACING_METHOD_UNIFORM,
                        line * scale,
                        baseline * scale,
                    )?;
                }

                // Excel keeps three pixels between the left edge of a cell and
                // its text, and two on the right. The gutter either side is
                // not the same: measured across fourteen faces, right-aligned
                // and centred text both sit a pixel further right than an even
                // gutter puts them, on every face whose glyphs the two engines
                // place alike.
                let gutter = (2.0 * scale).round() + 1.0;
                let mut area = box_of(
                    cell_box.left + gutter,
                    cell_box.top,
                    cell_box.right - (gutter - scale).max(0.0),
                    cell_box.bottom,
                );
                if filtered {
                    area.right -= FILTER_BUTTON as f32;
                }

                // Text too long for its cell runs on over the empty neighbours
                // beside it, which is what Excel shows.
                let body = HSTRING::from(text.as_str());
                // How far the text reaches is Excel's own measurement — its
                // characters' advances added up — and not this engine's, which
                // runs a shade narrower. The difference decides two things:
                // whether the last letter of a spilling cell is drawn or
                // clipped, and where right- and centre-aligned text starts.
                let one_line = !cell.style.wrap_text && !text.contains('\n');
                let reach = one_line
                    .then(|| {
                        crate::run_width(
                            cell.style.font_name.as_deref().unwrap_or("Calibri"),
                            points,
                            bold,
                            cell.style.italic,
                            &text,
                        )
                        .map(|width| width * scale)
                    })
                    .flatten();
                // Text spread across its cell is meant to fill exactly that
                // cell, so it never runs on: widening the area first would
                // spread it over the neighbours.
                if !cell.style.wrap_text
                    && placed != Align::Spread
                    && matches!(cell.value, CellValue::String(_))
                {
                    let width = match reach {
                        Some(width) => width,
                        None => {
                            let held = writer.CreateTextLayout(
                                body.as_wide(),
                                &format,
                                100_000.0,
                                area.bottom - area.top,
                            )?;
                            let mut measured = DWRITE_TEXT_METRICS::default();
                            held.GetMetrics(&mut measured)?;
                            measured.width
                        }
                    };
                    let spare = width - (area.right - area.left);
                    if spare > 0.0 {
                        let (before, after) = match placed {
                            Align::Left | Align::Spread => (0.0, spare),
                            Align::Right => (spare, 0.0),
                            Align::Centre => (spare / 2.0, spare - spare / 2.0),
                        };
                        area.left -=
                            room_before(layout, row, cell.col, before as i32, &merged) as f32;
                        area.right += room_after(
                            layout,
                            row,
                            cell.col + spans_columns,
                            after as i32,
                            &merged,
                        ) as f32;
                    }
                }

                // Right- and centre-aligned text is started where Excel's own
                // measurement puts it, rather than where this engine's would:
                // a pixel of difference in the run's width is a pixel of
                // difference on every number in the column.
                if let Some(width) = reach {
                    let room = area.right - area.left;
                    let start = match placed {
                        Align::Right => Some(area.right - width),
                        // Excel gives the odd pixel to the left of the text.
                        Align::Centre => Some(area.left + ((room - width) / 2.0).ceil()),
                        Align::Left | Align::Spread => None,
                    };
                    // Text with no room to spare is left to the engine, which
                    // anchors it the way Excel does and clips the end that
                    // overflows.
                    if let Some(start) = start.filter(|start| *start >= area.left) {
                        format.SetTextAlignment(DWRITE_TEXT_ALIGNMENT_LEADING)?;
                        area.left = start;
                    }
                }

                let header = matches!(&dress, Some(dress) if dress.white_text);
                brush.SetColor(&if header {
                    shade(None, 0x00FF_FFFF)
                } else {
                    shade(cell.style.font_color.as_deref(), 0)
                });

                if placed == Align::Spread {
                    // Excel spreads a distributed cell across its whole width,
                    // one line or many — which is how a Japanese sheet sets a
                    // heading: 第 ３ 表, not 第３表. DirectWrite's justification
                    // leaves a last line alone, so the gap is put between the
                    // letters instead.
                    let laid = writer.CreateTextLayout(
                        body.as_wide(),
                        &format,
                        area.right - area.left,
                        area.bottom - area.top,
                    )?;
                    let mut measured = DWRITE_TEXT_METRICS::default();
                    laid.GetMetrics(&mut measured)?;
                    let letters = text.chars().count() as f32;
                    let spare = (area.right - area.left) - measured.width;
                    if letters > 1.0 && spare > 0.0 {
                        if let Ok(spaced) = laid.cast::<IDWriteTextLayout1>() {
                            let gap = spare / (letters - 1.0);
                            // Half either side of every letter comes to one gap
                            // between each pair, and half a gap at each end,
                            // which is where Excel puts them.
                            spaced.SetCharacterSpacing(
                                gap / 2.0,
                                gap / 2.0,
                                0.0,
                                DWRITE_TEXT_RANGE {
                                    startPosition: 0,
                                    length: body.as_wide().len() as u32,
                                },
                            )?;
                        }
                    }
                    target.DrawTextLayout(
                        D2D_POINT_2F {
                            x: area.left,
                            y: area.top,
                        },
                        &laid,
                        &brush,
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                    );
                } else if cell.runs.is_empty() {
                    target.DrawText(
                        body.as_wide(),
                        &format,
                        &area,
                        &brush,
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                        measuring_mode(),
                    );
                } else {
                    // Parts of the text are dressed differently — an 8pt aside
                    // inside an 11pt cell, a raised footnote marker. A layout
                    // takes the dressing over character ranges; drawing the
                    // whole cell in one format makes the text wider than
                    // Excel's and wraps where Excel does not.
                    let laid = writer.CreateTextLayout(
                        body.as_wide(),
                        &format,
                        area.right - area.left,
                        area.bottom - area.top,
                    )?;
                    let mut at = 0u32;
                    for run in &cell.runs {
                        let length = run.text.encode_utf16().count() as u32;
                        let span = DWRITE_TEXT_RANGE {
                            startPosition: at,
                            length,
                        };
                        at += length;
                        // A raised or lowered run is drawn at 65% of the size
                        // it would otherwise be — measured against Excel, where
                        // 20pt capitals stand 20px tall and their superscript
                        // stands 13.
                        let raised = run.vert_align.is_some();
                        let points = run.size.unwrap_or(points);
                        if run.size.is_some() || raised {
                            let shrunk = if raised { points * 0.65 } else { points };
                            laid.SetFontSize(em_pixels(shrunk, scale), span)?;
                        }
                        if let Some(font) = run.font.as_deref() {
                            laid.SetFontFamilyName(&HSTRING::from(font), span)?;
                        }
                        if run.bold {
                            laid.SetFontWeight(DWRITE_FONT_WEIGHT_BOLD, span)?;
                        }
                        if run.italic {
                            laid.SetFontStyle(DWRITE_FONT_STYLE_ITALIC, span)?;
                        }
                        if run.underline {
                            laid.SetUnderline(true, span)?;
                        }
                    }
                    target.DrawTextLayout(
                        D2D_POINT_2F {
                            x: area.left,
                            y: area.top,
                        },
                        &laid,
                        &brush,
                        D2D1_DRAW_TEXT_OPTIONS_CLIP,
                    );
                }
            }
        }

        target.EndDraw(None, None)?;

        let mut pixels = vec![0u8; (width * height * 4) as usize];
        bitmap.CopyPixels(std::ptr::null(), width * 4, &mut pixels)?;
        let mut canvas = image::RgbImage::new(width, height);
        for y in 0..height {
            for x in 0..width {
                let at = ((y * width + x) * 4) as usize;
                canvas.put_pixel(x, y, image::Rgb([pixels[at + 2], pixels[at + 1], pixels[at]]));
            }
        }
        Ok(canvas)
    }
}
