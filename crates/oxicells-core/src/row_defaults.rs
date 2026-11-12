// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The height Excel gives an untouched row, per font — measured data shared
//! by the renderer and by the editor's `AutoFit`.
//!
//! The height Excel gives an untouched row, per font. A sheet that does not
//! pin its default row height (customHeight) has the stated number thrown
//! away; the height is derived from the standard font of the columns. These
//! values are measured from Excel itself — no formula over the font's own
//! metric tables has reproduced the full set, so the measurements stand as
//! the data, the way the font width tables do.
//!
//! Keyed by face name and quarter-points (11pt → 44); the value is the row
//! height in 96‑dpi pixels (0.75pt each).

use crate::ir::CellUnderlineKind;

/// (face, size in quarter-points, row height in pixels at 96 dpi, how far
/// down that height the baseline sits)
static FONT_DEFAULT_ROW_PX: &[(&str, u16, u16, u16)] = &[
    ("Aptos Narrow", 104, 46, 36),
    ("Arial", 104, 44, 36),
    ("Arial", 160, 66, 54),
    ("Yu Gothic", 104, 57, 41),
    ("游ゴシック", 104, 57, 41),
    ("AR P丸ゴシック体E", 44, 18, 16),
    ("Aptos Narrow", 24, 11, 9),
    ("Aptos Narrow", 32, 15, 12),
    ("Aptos Narrow", 36, 16, 13),
    ("Aptos Narrow", 40, 18, 14),
    ("Aptos Narrow", 42, 19, 15),
    ("Aptos Narrow", 44, 20, 16),
    ("Aptos Narrow", 48, 21, 17),
    ("Aptos Narrow", 50, 23, 18),
    ("Aptos Narrow", 56, 25, 20),
    ("Aptos Narrow", 64, 28, 22),
    ("Aptos Narrow", 72, 32, 25),
    ("Aptos Narrow", 80, 35, 28),
    ("Aptos Narrow", 96, 42, 33),
    ("Aptos Narrow", 112, 48, 38),
    ("Aptos Narrow", 144, 62, 49),
    ("Aptos Narrow", 192, 85, 66),
    ("Aptos Narrow", 288, 125, 98),
    ("Arial", 24, 11, 9),
    ("Arial", 32, 15, 12),
    ("Arial", 36, 16, 13),
    ("Arial", 40, 17, 14),
    ("Arial", 42, 18, 15),
    ("Arial", 44, 19, 16),
    ("Arial", 48, 20, 17),
    ("Arial", 50, 22, 18),
    ("Arial", 56, 24, 20),
    ("Arial", 64, 27, 22),
    ("Arial", 72, 31, 25),
    ("Arial", 80, 34, 28),
    ("Arial", 96, 40, 33),
    ("Arial", 112, 46, 38),
    ("Arial", 144, 59, 49),
    ("Arial", 192, 79, 65),
    ("Arial", 200, 83, 68),
    ("Arial", 208, 85, 70),
    ("Arial", 288, 120, 98),
    ("Arial", 320, 133, 109),
    ("Arial", 480, 199, 163),
    ("Arial Unicode MS", 44, 19, 16),
    ("BIZ UDPゴシック", 24, 12, 10),
    ("BIZ UDPゴシック", 32, 15, 13),
    ("BIZ UDPゴシック", 36, 16, 14),
    ("BIZ UDPゴシック", 40, 16, 14),
    ("BIZ UDPゴシック", 42, 17, 15),
    ("BIZ UDPゴシック", 44, 18, 16),
    ("BIZ UDPゴシック", 48, 19, 17),
    ("BIZ UDPゴシック", 56, 22, 20),
    ("BIZ UDPゴシック", 64, 25, 22),
    ("BIZ UDPゴシック", 72, 28, 25),
    ("BIZ UDPゴシック", 80, 31, 28),
    ("BIZ UDPゴシック", 96, 37, 33),
    ("BIZ UDPゴシック", 144, 55, 49),
    ("BIZ UDゴシック", 24, 12, 10),
    ("BIZ UDゴシック", 32, 15, 13),
    ("BIZ UDゴシック", 36, 16, 14),
    ("BIZ UDゴシック", 40, 16, 14),
    ("BIZ UDゴシック", 42, 17, 15),
    ("BIZ UDゴシック", 44, 18, 16),
    ("BIZ UDゴシック", 48, 19, 17),
    ("BIZ UDゴシック", 56, 22, 20),
    ("BIZ UDゴシック", 64, 25, 22),
    ("BIZ UDゴシック", 72, 28, 25),
    ("BIZ UDゴシック", 80, 31, 28),
    ("BIZ UDゴシック", 96, 37, 33),
    ("BIZ UDゴシック", 144, 55, 49),
    ("Calibri", 24, 11, 9),
    ("Calibri", 32, 15, 12),
    ("Calibri", 36, 16, 13),
    ("Calibri", 40, 17, 14),
    ("Calibri", 42, 19, 15),
    ("Calibri", 44, 20, 16),
    ("Calibri", 48, 21, 17),
    ("Calibri", 50, 23, 18),
    ("Calibri", 56, 25, 20),
    ("Calibri", 64, 28, 22),
    ("Calibri", 72, 31, 25),
    ("Calibri", 80, 35, 28),
    ("Calibri", 96, 42, 33),
    ("Calibri", 112, 48, 38),
    ("Calibri", 144, 62, 49),
    ("Calibri", 192, 82, 65),
    ("Calibri", 288, 123, 97),
    ("Cambria", 32, 14, 12),
    ("Cambria", 38, 17, 14),
    ("Cambria", 42, 18, 15),
    ("Cambria", 44, 19, 16),
    ("Cambria", 48, 21, 17),
    ("Cambria", 50, 22, 18),
    ("Cambria", 56, 24, 20),
    ("Cambria", 64, 27, 22),
    ("Cambria", 72, 30, 25),
    ("Cambria", 80, 34, 28),
    ("Cambria", 96, 40, 33),
    ("Cambria", 112, 46, 38),
    ("Cambria", 144, 60, 49),
    ("Cambria", 192, 79, 65),
    ("Cambria", 288, 118, 97),
    ("Consolas", 32, 15, 12),
    ("Consolas", 38, 17, 14),
    ("Consolas", 42, 19, 15),
    ("Consolas", 44, 20, 16),
    ("Consolas", 48, 21, 17),
    ("Consolas", 50, 22, 18),
    ("Consolas", 56, 25, 20),
    ("Consolas", 64, 27, 22),
    ("Consolas", 72, 31, 25),
    ("Consolas", 80, 35, 28),
    ("Consolas", 96, 41, 33),
    ("Consolas", 112, 47, 38),
    ("Consolas", 144, 61, 49),
    ("Consolas", 192, 81, 65),
    ("Consolas", 288, 121, 97),
    ("Courier New", 24, 11, 9),
    ("Courier New", 32, 15, 12),
    ("Courier New", 38, 18, 14),
    ("Courier New", 42, 19, 15),
    ("Courier New", 44, 20, 16),
    ("Courier New", 48, 21, 17),
    ("Courier New", 50, 23, 18),
    ("Courier New", 56, 25, 20),
    ("Courier New", 64, 28, 22),
    ("Courier New", 72, 32, 25),
    ("Courier New", 80, 35, 28),
    ("Courier New", 96, 42, 33),
    ("Courier New", 112, 48, 38),
    ("Courier New", 144, 62, 49),
    ("Courier New", 192, 83, 65),
    ("Courier New", 288, 123, 97),
    ("Century", 24, 16, 11),
    ("Century", 28, 17, 12),
    ("Century", 36, 18, 14),
    ("Century", 40, 17, 14),
    ("Century", 42, 18, 15),
    ("Century", 44, 19, 16),
    ("Century", 48, 21, 17),
    ("Century", 56, 24, 20),
    ("Century", 60, 25, 21),
    ("HGP創英角ｺﾞｼｯｸUB", 48, 19, 17),
    ("HGP教科書体", 44, 18, 16),
    ("HGP行書体", 32, 14, 12),
    ("HGP行書体", 36, 15, 13),
    ("HGP行書体", 40, 16, 14),
    ("HGP行書体", 44, 18, 16),
    ("HGP行書体", 48, 19, 17),
    ("HGPｺﾞｼｯｸM", 24, 12, 10),
    ("HGPｺﾞｼｯｸM", 32, 14, 12),
    ("HGPｺﾞｼｯｸM", 36, 15, 13),
    ("HGPｺﾞｼｯｸM", 40, 16, 14),
    ("HGPｺﾞｼｯｸM", 44, 18, 17),
    ("HGPｺﾞｼｯｸM", 48, 19, 17),
    ("HGPｺﾞｼｯｸM", 56, 23, 21),
    ("HGS行書体", 32, 14, 12),
    ("HGS行書体", 36, 15, 13),
    ("HGS行書体", 40, 16, 14),
    ("HGS行書体", 44, 18, 16),
    ("HGS行書体", 48, 19, 17),
    ("HGSｺﾞｼｯｸ", 36, 15, 13),
    ("HGSｺﾞｼｯｸE", 36, 15, 13),
    ("HGSｺﾞｼｯｸM", 32, 14, 12),
    ("HGSｺﾞｼｯｸM", 36, 15, 13),
    ("HGSｺﾞｼｯｸM", 44, 18, 17),
    ("HG創英角ﾎﾟｯﾌﾟ体", 40, 16, 14),
    ("HG明朝B", 80, 32, 28),
    ("HG行書体", 44, 18, 16),
    ("MS Gothic", 40, 16, 14),
    ("MS P ゴシック", 36, 15, 13),
    ("MS PGothic", 44, 18, 16),
    ("MS UI Gothic", 24, 12, 10),
    ("MS UI Gothic", 32, 14, 12),
    ("MS UI Gothic", 36, 15, 13),
    ("MS UI Gothic", 40, 16, 14),
    ("MS UI Gothic", 42, 17, 15),
    ("MS UI Gothic", 44, 18, 16),
    ("MS UI Gothic", 48, 19, 17),
    ("MS UI Gothic", 56, 23, 20),
    ("MS UI Gothic", 64, 25, 22),
    ("MS UI Gothic", 72, 28, 25),
    ("MS UI Gothic", 80, 32, 28),
    ("MS UI Gothic", 96, 38, 33),
    ("MS UI Gothic", 144, 56, 49),
    ("MS ゴシック", 40, 16, 13),
    ("MS　Pゴシック", 44, 18, 16),
    ("Meiryo UI", 24, 11, 10),
    ("Meiryo UI", 32, 16, 13),
    ("Meiryo UI", 36, 16, 13),
    ("Meiryo UI", 40, 19, 15),
    ("Meiryo UI", 42, 20, 16),
    ("Meiryo UI", 44, 21, 17),
    ("Meiryo UI", 48, 22, 18),
    ("Meiryo UI", 52, 24, 19),
    ("Meiryo UI", 56, 26, 21),
    ("Meiryo UI", 60, 27, 22),
    ("Meiryo UI", 64, 28, 23),
    ("Meiryo UI", 72, 32, 26),
    ("Meiryo UI", 80, 38, 30),
    ("Meiryo UI", 96, 44, 35),
    ("Meiryo UI", 144, 65, 52),
    ("Segoe UI", 24, 13, 11),
    ("Segoe UI", 32, 14, 12),
    ("Segoe UI", 36, 16, 13),
    ("Segoe UI", 40, 19, 15),
    ("Segoe UI", 42, 21, 16),
    ("Segoe UI", 44, 22, 17),
    ("Segoe UI", 48, 23, 18),
    ("Segoe UI", 56, 27, 21),
    ("Segoe UI", 64, 34, 25),
    ("Segoe UI", 72, 35, 27),
    ("Segoe UI", 80, 41, 31),
    ("Segoe UI", 96, 50, 37),
    ("Segoe UI", 144, 70, 53),
    ("Segoe UI Symbol", 40, 19, 15),
    ("Tahoma", 24, 11, 9),
    ("Tahoma", 32, 14, 12),
    ("Tahoma", 36, 15, 13),
    ("Tahoma", 40, 17, 14),
    ("Tahoma", 42, 18, 15),
    ("Tahoma", 44, 19, 16),
    ("Tahoma", 48, 20, 17),
    ("Tahoma", 56, 24, 20),
    ("Tahoma", 64, 26, 22),
    ("Tahoma", 72, 30, 25),
    ("Tahoma", 80, 34, 28),
    ("Tahoma", 96, 40, 33),
    ("Tahoma", 144, 59, 49),
    ("Terminal", 28, 12, 10),
    ("Terminal", 40, 17, 14),
    ("Terminal", 48, 19, 14),
    ("Terminal", 56, 21, 18),
    ("Times New Roman", 24, 11, 9),
    ("Times New Roman", 32, 15, 12),
    ("Times New Roman", 34, 15, 12),
    ("Times New Roman", 36, 16, 13),
    ("Times New Roman", 40, 17, 14),
    ("Times New Roman", 42, 18, 15),
    ("Times New Roman", 44, 20, 16),
    ("Times New Roman", 48, 21, 17),
    ("Times New Roman", 52, 22, 18),
    ("Times New Roman", 56, 25, 20),
    ("Times New Roman", 60, 26, 21),
    ("Times New Roman", 64, 27, 22),
    ("Times New Roman", 72, 31, 25),
    ("Times New Roman", 76, 32, 26),
    ("Times New Roman", 80, 35, 28),
    ("Times New Roman", 88, 37, 30),
    ("Times New Roman", 96, 41, 33),
    ("Times New Roman", 112, 47, 38),
    ("Times New Roman", 144, 61, 49),
    ("Times New Roman", 192, 82, 66),
    ("Times New Roman", 288, 122, 98),
    ("Verdana", 24, 11, 9),
    ("Verdana", 32, 14, 12),
    ("Verdana", 36, 15, 13),
    ("Verdana", 40, 17, 14),
    ("Verdana", 42, 18, 15),
    ("Verdana", 44, 19, 16),
    ("Verdana", 48, 20, 17),
    ("Verdana", 56, 24, 20),
    ("Verdana", 64, 26, 22),
    ("Verdana", 72, 30, 25),
    ("Verdana", 80, 33, 28),
    ("Verdana", 96, 39, 33),
    ("Verdana", 144, 61, 50),
    ("Yu Gothic UI", 24, 13, 10),
    ("Yu Gothic UI", 32, 14, 12),
    ("Yu Gothic UI", 36, 16, 13),
    ("Yu Gothic UI", 40, 19, 15),
    ("Yu Gothic UI", 42, 21, 16),
    ("Yu Gothic UI", 44, 22, 17),
    ("Yu Gothic UI", 48, 23, 18),
    ("Yu Gothic UI", 56, 27, 21),
    ("Yu Gothic UI", 64, 34, 25),
    ("Yu Gothic UI", 72, 35, 27),
    ("Yu Gothic UI", 80, 41, 31),
    ("Yu Gothic UI", 88, 44, 33),
    ("Yu Gothic UI", 96, 50, 37),
    ("Yu Gothic UI", 144, 70, 53),
    ("メイリオ", 24, 13, 10),
    ("メイリオ", 32, 19, 13),
    ("メイリオ", 36, 20, 14),
    ("メイリオ", 40, 22, 15),
    ("メイリオ", 42, 23, 16),
    ("メイリオ", 44, 25, 17),
    ("メイリオ", 48, 26, 18),
    ("メイリオ", 52, 27, 19),
    ("メイリオ", 56, 30, 21),
    ("メイリオ", 60, 32, 22),
    ("メイリオ", 64, 33, 23),
    ("メイリオ", 72, 38, 26),
    ("メイリオ", 80, 44, 30),
    ("メイリオ", 96, 51, 35),
    ("メイリオ", 112, 58, 40),
    ("メイリオ", 144, 76, 52),
    ("メイリオ", 192, 101, 69),
    ("メイリオ", 288, 151, 103),
    ("明朝", 40, 16, 13),
    ("明朝", 44, 18, 17),
    ("明朝", 72, 28, 26),
    ("標準ゴシック", 28, 13, 11),
    ("標準ゴシック", 32, 14, 12),
    ("標準ゴシック", 36, 15, 13),
    ("標準ゴシック", 40, 16, 13),
    ("標準ゴシック", 42, 17, 14),
    ("標準ゴシック", 44, 18, 16),
    ("游ゴシック", 24, 13, 10),
    ("游ゴシック", 28, 15, 11),
    ("游ゴシック", 32, 17, 13),
    ("游ゴシック", 36, 21, 15),
    ("游ゴシック", 40, 22, 16),
    ("游ゴシック", 42, 23, 17),
    ("游ゴシック", 44, 25, 18),
    ("游ゴシック", 288, 154, 111),
    ("Yu Gothic", 288, 154, 111),
    ("游ゴシック", 48, 26, 19),
    ("游ゴシック", 52, 27, 20),
    ("游ゴシック", 56, 32, 23),
    ("游ゴシック", 60, 33, 24),
    ("游ゴシック", 64, 34, 25),
    ("游ゴシック", 72, 40, 29),
    ("游ゴシック", 80, 44, 32),
    ("游ゴシック", 88, 47, 34),
    ("游ゴシック", 96, 53, 38),
    ("游ゴシック", 112, 59, 43),
    ("游ゴシック", 144, 78, 56),
    ("游ゴシック Light", 72, 40, 28),
    ("游明朝", 24, 13, 10),
    ("游明朝", 32, 17, 13),
    ("游明朝", 36, 21, 15),
    ("游明朝", 40, 22, 16),
    ("游明朝", 42, 23, 17),
    ("游明朝", 44, 24, 18),
    ("游明朝", 48, 26, 19),
    ("游明朝", 56, 32, 23),
    ("游明朝", 64, 34, 25),
    ("游明朝", 72, 40, 29),
    ("游明朝", 80, 44, 32),
    ("游明朝", 96, 52, 38),
    ("游明朝", 144, 77, 56),
    ("ＭＳ ゴシック", 24, 12, 10),
    ("ＭＳ ゴシック", 28, 13, 11),
    ("ＭＳ ゴシック", 32, 14, 12),
    ("ＭＳ ゴシック", 36, 15, 13),
    ("ＭＳ ゴシック", 40, 16, 14),
    ("ＭＳ ゴシック", 42, 17, 15),
    ("ＭＳ ゴシック", 44, 18, 16),
    ("ＭＳ ゴシック", 46, 18, 16),
    ("ＭＳ ゴシック", 47, 19, 17),
    ("ＭＳ ゴシック", 48, 19, 17),
    ("ＭＳ ゴシック", 50, 20, 18),
    ("ＭＳ ゴシック", 52, 20, 18),
    ("ＭＳ ゴシック", 54, 22, 19),
    ("ＭＳ ゴシック", 55, 22, 19),
    ("ＭＳ ゴシック", 56, 23, 20),
    ("ＭＳ ゴシック", 58, 23, 20),
    ("ＭＳ ゴシック", 59, 24, 21),
    ("ＭＳ ゴシック", 60, 24, 21),
    ("ＭＳ ゴシック", 63, 25, 22),
    ("ＭＳ ゴシック", 64, 25, 22),
    ("ＭＳ ゴシック", 70, 27, 24),
    ("ＭＳ ゴシック", 72, 28, 25),
    ("ＭＳ ゴシック", 74, 30, 26),
    ("ＭＳ ゴシック", 80, 32, 28),
    ("ＭＳ ゴシック", 88, 34, 30),
    ("ＭＳ ゴシック", 96, 38, 33),
    ("ＭＳ ゴシック", 112, 43, 38),
    ("ＭＳ ゴシック", 144, 56, 49),
    ("ＭＳ ゴシック", 192, 74, 65),
    ("ＭＳ 明朝", 12, 8, 7),
    ("ＭＳ 明朝", 24, 12, 10),
    ("ＭＳ 明朝", 28, 13, 11),
    ("ＭＳ 明朝", 30, 14, 12),
    ("ＭＳ 明朝", 32, 14, 12),
    ("ＭＳ 明朝", 36, 15, 13),
    ("ＭＳ 明朝", 40, 16, 14),
    ("ＭＳ 明朝", 42, 17, 15),
    ("ＭＳ 明朝", 44, 18, 16),
    ("ＭＳ 明朝", 46, 18, 16),
    ("ＭＳ 明朝", 47, 19, 17),
    ("ＭＳ 明朝", 48, 19, 17),
    ("ＭＳ 明朝", 50, 20, 18),
    ("ＭＳ 明朝", 52, 20, 18),
    ("ＭＳ 明朝", 54, 22, 19),
    ("ＭＳ 明朝", 55, 22, 19),
    ("ＭＳ 明朝", 56, 23, 20),
    ("ＭＳ 明朝", 60, 24, 21),
    ("ＭＳ 明朝", 63, 25, 22),
    ("ＭＳ 明朝", 64, 25, 22),
    ("ＭＳ 明朝", 68, 27, 24),
    ("ＭＳ 明朝", 70, 27, 24),
    ("ＭＳ 明朝", 72, 28, 25),
    ("ＭＳ 明朝", 74, 30, 26),
    ("ＭＳ 明朝", 76, 30, 26),
    ("ＭＳ 明朝", 80, 32, 28),
    ("ＭＳ 明朝", 88, 34, 30),
    ("ＭＳ 明朝", 96, 38, 33),
    ("ＭＳ 明朝", 112, 43, 38),
    ("ＭＳ 明朝", 144, 56, 49),
    ("ＭＳ Ｐゴシック", 24, 12, 10),
    ("ＭＳ Ｐゴシック", 28, 13, 11),
    ("ＭＳ Ｐゴシック", 30, 14, 12),
    ("ＭＳ Ｐゴシック", 32, 14, 12),
    ("ＭＳ Ｐゴシック", 36, 15, 13),
    ("ＭＳ Ｐゴシック", 37, 15, 13),
    ("ＭＳ Ｐゴシック", 40, 16, 14),
    ("ＭＳ Ｐゴシック", 42, 17, 15),
    ("ＭＳ Ｐゴシック", 43, 17, 15),
    ("ＭＳ Ｐゴシック", 44, 18, 16),
    ("ＭＳ Ｐゴシック", 48, 19, 17),
    ("ＭＳ Ｐゴシック", 50, 20, 18),
    ("ＭＳ Ｐゴシック", 52, 20, 18),
    ("ＭＳ Ｐゴシック", 54, 22, 19),
    ("ＭＳ Ｐゴシック", 56, 23, 20),
    ("ＭＳ Ｐゴシック", 60, 24, 21),
    ("ＭＳ Ｐゴシック", 64, 25, 22),
    ("ＭＳ Ｐゴシック", 68, 27, 24),
    ("ＭＳ Ｐゴシック", 72, 28, 25),
    ("ＭＳ Ｐゴシック", 80, 32, 28),
    ("ＭＳ Ｐゴシック", 88, 34, 30),
    ("ＭＳ Ｐゴシック", 96, 38, 33),
    ("ＭＳ Ｐゴシック", 112, 43, 38),
    ("ＭＳ Ｐゴシック", 144, 56, 49),
    ("ＭＳ Ｐ明朝", 24, 12, 10),
    ("ＭＳ Ｐ明朝", 28, 13, 11),
    ("ＭＳ Ｐ明朝", 30, 14, 12),
    ("ＭＳ Ｐ明朝", 32, 14, 12),
    ("ＭＳ Ｐ明朝", 34, 14, 12),
    ("ＭＳ Ｐ明朝", 36, 15, 13),
    ("ＭＳ Ｐ明朝", 40, 16, 14),
    ("ＭＳ Ｐ明朝", 42, 17, 15),
    ("ＭＳ Ｐ明朝", 44, 18, 16),
    ("ＭＳ Ｐ明朝", 48, 19, 17),
    ("ＭＳ Ｐ明朝", 56, 23, 20),
    ("ＭＳ Ｐ明朝", 64, 25, 22),
    ("ＭＳ Ｐ明朝", 72, 28, 25),
    ("ＭＳ Ｐ明朝", 76, 30, 26),
    ("ＭＳ Ｐ明朝", 80, 32, 28),
    ("ＭＳ Ｐ明朝", 88, 34, 30),
    ("ＭＳ Ｐ明朝", 96, 38, 33),
    ("ＭＳ Ｐ明朝", 144, 56, 49),
];

/// Where a bold face does not stand where its regular does.
///
/// Most do. Swept over nine faces at six sizes and nine row heights — 1674
/// rows through `_xlsx_valign_pixels.py` — these two are the whole of the
/// difference, and both show it only where the line is placed against the
/// middle or the foot of its row: 游ゴシック's bold line box is a pixel
/// shorter than its regular's, and Calibri's bold baseline sits a pixel
/// higher in a box of the same height. A top-aligned bold cell is right
/// either way, which `_xlsx_bold_baseline.py` confirms across sixty more.
static FONT_BOLD_ROW_PX: &[(&str, u16, u16, u16)] = &[
    ("Yu Gothic", 104, 56, 41),
    ("游ゴシック", 104, 56, 41),
    ("Calibri", 40, 17, 13),
    ("游ゴシック", 44, 24, 18),
];

/// Style-specific line boxes that differ from the regular face.
static FONT_STYLE_LINE_BOXES: &[(&str, u16, bool, bool, u16, u16)] = &[
    ("Arial", 32, true, true, 14, 12),
    ("Arial", 44, true, false, 20, 16),
    ("Arial", 48, true, false, 21, 17),
    ("Arial", 56, false, true, 25, 20),
    ("Arial", 56, true, true, 25, 20),
    ("Arial", 80, true, false, 35, 28),
    ("Arial", 96, false, true, 41, 33),
    ("Arial", 112, false, true, 47, 38),
    ("Arial", 112, true, false, 47, 38),
    ("Arial", 144, false, true, 60, 49),
    ("Arial", 144, true, false, 60, 49),
    ("Arial", 192, false, true, 80, 65),
    ("Arial", 192, true, false, 80, 65),
    ("Arial", 192, true, true, 80, 65),
    ("Arial", 288, false, true, 121, 98),
    ("Arial", 288, true, false, 121, 98),
    ("Arial", 288, true, true, 121, 98),
    ("Courier New", 32, false, true, 16, 12),
    ("Courier New", 44, false, true, 21, 16),
    ("Courier New", 44, true, false, 21, 16),
    ("Courier New", 44, true, true, 21, 16),
    ("Courier New", 48, false, true, 22, 17),
    ("Courier New", 48, true, false, 22, 17),
    ("Courier New", 48, true, true, 22, 17),
    ("Courier New", 56, true, false, 26, 20),
    ("Courier New", 56, true, true, 26, 20),
    ("Courier New", 80, true, false, 36, 28),
    ("Courier New", 80, true, true, 36, 28),
    ("Courier New", 96, false, true, 41, 33),
    ("Courier New", 96, true, false, 43, 33),
    ("Courier New", 96, true, true, 43, 33),
    ("Courier New", 112, true, false, 49, 38),
    ("Courier New", 112, true, true, 49, 38),
    ("Courier New", 144, true, false, 63, 49),
    ("Courier New", 144, true, true, 63, 49),
    ("Courier New", 192, true, false, 84, 65),
    ("Courier New", 192, true, true, 84, 65),
    ("Courier New", 288, true, false, 126, 97),
    ("Courier New", 288, true, true, 126, 97),
    ("Times New Roman", 32, true, false, 14, 12),
    ("Times New Roman", 38, true, true, 18, 14),
    ("Times New Roman", 42, true, true, 19, 15),
    ("Times New Roman", 44, true, false, 19, 16),
    ("Times New Roman", 50, true, true, 23, 18),
    ("Times New Roman", 56, true, true, 26, 20),
    ("Times New Roman", 72, true, false, 30, 25),
    ("Times New Roman", 80, true, false, 34, 28),
    ("Times New Roman", 80, true, true, 34, 28),
    ("Times New Roman", 96, true, false, 40, 33),
    ("Times New Roman", 96, true, true, 40, 33),
    ("Times New Roman", 112, true, false, 46, 38),
    ("Times New Roman", 112, true, true, 46, 38),
    ("Times New Roman", 144, true, false, 60, 49),
    ("Times New Roman", 144, true, true, 60, 49),
    ("Times New Roman", 192, true, false, 81, 66),
    ("Times New Roman", 192, true, true, 81, 66),
    ("Times New Roman", 288, true, false, 120, 98),
    ("Times New Roman", 288, true, true, 120, 98),
];

/// Own line boxes for raised and lowered text, in 96-dpi pixels.
static FONT_POSITION_LINE_BOXES: &[(&str, u16, bool, bool, u16, u16, u16, u16)] = &[
    ("Arial", 32, false, false, 15, 12, 15, 12),
    ("Arial", 32, true, false, 15, 12, 15, 12),
    ("Arial", 32, false, true, 15, 12, 15, 12),
    ("Arial", 32, true, true, 14, 12, 16, 13),
    ("Arial", 44, false, false, 22, 19, 25, 19),
    ("Arial", 44, true, false, 23, 19, 22, 17),
    ("Arial", 44, false, true, 22, 19, 25, 19),
    ("Arial", 44, true, true, 21, 18, 23, 18),
    ("Arial", 48, false, false, 24, 21, 26, 20),
    ("Arial", 48, true, false, 25, 21, 25, 19),
    ("Arial", 48, false, true, 23, 20, 26, 20),
    ("Arial", 48, true, true, 24, 21, 24, 19),
    ("Arial", 96, false, false, 46, 39, 48, 37),
    ("Arial", 96, true, false, 45, 38, 48, 37),
    ("Arial", 96, false, true, 47, 39, 47, 36),
    ("Arial", 96, true, true, 46, 39, 48, 37),
    ("Arial", 288, false, false, 140, 118, 144, 110),
    ("Arial", 288, true, false, 140, 117, 145, 110),
    ("Arial", 288, false, true, 141, 118, 145, 110),
    ("Arial", 288, true, true, 140, 117, 145, 110),
    ("Calibri", 32, false, false, 17, 14, 17, 13),
    ("Calibri", 32, true, false, 17, 14, 17, 13),
    ("Calibri", 32, false, true, 17, 14, 17, 13),
    ("Calibri", 32, true, true, 17, 14, 17, 13),
    ("Calibri", 44, false, false, 23, 19, 24, 18),
    ("Calibri", 44, true, false, 23, 19, 24, 18),
    ("Calibri", 44, false, true, 23, 19, 24, 18),
    ("Calibri", 44, true, true, 23, 19, 24, 18),
    ("Calibri", 48, false, false, 24, 20, 25, 19),
    ("Calibri", 48, true, false, 24, 20, 25, 19),
    ("Calibri", 48, false, true, 24, 20, 25, 19),
    ("Calibri", 48, true, true, 24, 20, 25, 19),
    ("Calibri", 96, false, false, 48, 39, 48, 36),
    ("Calibri", 96, true, false, 48, 39, 48, 36),
    ("Calibri", 96, false, true, 48, 39, 48, 36),
    ("Calibri", 96, true, true, 48, 39, 48, 36),
    ("Calibri", 288, false, false, 141, 115, 143, 107),
    ("Calibri", 288, true, false, 141, 115, 143, 107),
    ("Calibri", 288, false, true, 141, 115, 143, 107),
    ("Calibri", 288, true, true, 141, 115, 143, 107),
    ("Times New Roman", 32, false, false, 15, 12, 17, 13),
    ("Times New Roman", 32, true, false, 14, 12, 18, 14),
    ("Times New Roman", 32, false, true, 15, 12, 17, 13),
    ("Times New Roman", 32, true, true, 15, 12, 17, 13),
    ("Times New Roman", 44, false, false, 24, 20, 22, 17),
    ("Times New Roman", 44, true, false, 22, 19, 23, 18),
    ("Times New Roman", 44, false, true, 24, 20, 22, 17),
    ("Times New Roman", 44, true, true, 22, 18, 22, 17),
    ("Times New Roman", 48, false, false, 25, 21, 25, 19),
    ("Times New Roman", 48, true, false, 25, 21, 23, 18),
    ("Times New Roman", 48, false, true, 25, 21, 25, 19),
    ("Times New Roman", 48, true, true, 25, 21, 25, 19),
    ("Times New Roman", 96, false, false, 47, 39, 47, 36),
    ("Times New Roman", 96, true, false, 46, 39, 48, 37),
    ("Times New Roman", 96, false, true, 46, 38, 47, 36),
    ("Times New Roman", 96, true, true, 46, 39, 48, 37),
    ("Times New Roman", 288, false, false, 142, 118, 144, 109),
    ("Times New Roman", 288, true, false, 139, 117, 144, 110),
    ("Times New Roman", 288, false, true, 142, 118, 144, 109),
    ("Times New Roman", 288, true, true, 139, 117, 144, 110),
    ("Courier New", 32, false, false, 15, 12, 17, 13),
    ("Courier New", 32, true, false, 15, 12, 17, 13),
    ("Courier New", 32, false, true, 16, 12, 16, 12),
    ("Courier New", 32, true, true, 15, 12, 17, 13),
    ("Courier New", 44, false, false, 23, 19, 24, 18),
    ("Courier New", 44, true, false, 26, 21, 23, 17),
    ("Courier New", 44, false, true, 24, 19, 23, 17),
    ("Courier New", 44, true, true, 25, 20, 23, 17),
    ("Courier New", 48, false, false, 26, 22, 25, 19),
    ("Courier New", 48, true, false, 28, 23, 24, 18),
    ("Courier New", 48, false, true, 27, 22, 26, 19),
    ("Courier New", 48, true, true, 27, 22, 24, 18),
    ("Courier New", 96, false, false, 48, 39, 48, 36),
    ("Courier New", 96, true, false, 50, 40, 47, 35),
    ("Courier New", 96, false, true, 48, 40, 49, 37),
    ("Courier New", 96, true, true, 49, 39, 47, 35),
    ("Courier New", 288, false, false, 145, 119, 145, 108),
    ("Courier New", 288, true, false, 147, 118, 144, 106),
    ("Courier New", 288, false, true, 145, 119, 145, 108),
    ("Courier New", 288, true, true, 147, 118, 144, 106),
    ("メイリオ", 32, false, false, 19, 13, 19, 13),
    ("メイリオ", 32, true, false, 19, 13, 19, 13),
    ("メイリオ", 32, false, true, 19, 13, 19, 13),
    ("メイリオ", 32, true, true, 19, 13, 19, 13),
    ("メイリオ", 44, false, false, 27, 19, 25, 17),
    ("メイリオ", 44, true, false, 27, 19, 25, 17),
    ("メイリオ", 44, false, true, 27, 19, 25, 17),
    ("メイリオ", 44, true, true, 27, 19, 25, 17),
    ("メイリオ", 48, false, false, 29, 21, 26, 18),
    ("メイリオ", 48, true, false, 29, 21, 26, 18),
    ("メイリオ", 48, false, true, 29, 21, 26, 18),
    ("メイリオ", 48, true, true, 29, 21, 26, 18),
    ("メイリオ", 96, false, false, 55, 39, 51, 35),
    ("メイリオ", 96, true, false, 55, 39, 51, 35),
    ("メイリオ", 96, false, true, 55, 39, 51, 35),
    ("メイリオ", 96, true, true, 55, 39, 51, 35),
    ("メイリオ", 288, false, false, 165, 117, 151, 103),
    ("メイリオ", 288, true, false, 165, 117, 151, 103),
    ("メイリオ", 288, false, true, 165, 117, 151, 103),
    ("メイリオ", 288, true, true, 165, 117, 151, 103),
    ("Aptos Narrow", 32, false, false, 16, 13, 17, 13),
    ("Aptos Narrow", 32, true, false, 16, 13, 17, 13),
    ("Aptos Narrow", 32, false, true, 16, 13, 17, 13),
    ("Aptos Narrow", 32, true, true, 16, 13, 17, 13),
    ("Aptos Narrow", 44, false, false, 22, 18, 24, 18),
    ("Aptos Narrow", 44, true, false, 22, 18, 24, 18),
    ("Aptos Narrow", 44, false, true, 22, 18, 24, 18),
    ("Aptos Narrow", 44, true, true, 22, 18, 24, 18),
    ("Aptos Narrow", 48, false, false, 24, 20, 25, 19),
    ("Aptos Narrow", 48, true, false, 24, 20, 25, 19),
    ("Aptos Narrow", 48, false, true, 24, 20, 25, 19),
    ("Aptos Narrow", 48, true, true, 24, 20, 25, 19),
    ("Aptos Narrow", 96, false, false, 47, 38, 48, 36),
    ("Aptos Narrow", 96, true, false, 47, 38, 48, 36),
    ("Aptos Narrow", 96, false, true, 47, 38, 48, 36),
    ("Aptos Narrow", 96, true, true, 47, 38, 48, 36),
    ("Aptos Narrow", 288, false, false, 141, 114, 145, 108),
    ("Aptos Narrow", 288, true, false, 141, 114, 145, 108),
    ("Aptos Narrow", 288, false, true, 141, 114, 145, 108),
    ("Aptos Narrow", 288, true, true, 141, 114, 145, 108),
];

/// Differences between the measured line descent and the device descent.
static FONT_DEVICE_DESCENT_OFFSETS: &[(&str, u16, i16)] = &[
    ("Aptos Narrow", 192, 1),
    ("Aptos Narrow", 288, 1),
    ("Arial", 288, 1),
    ("Times New Roman", 192, 1),
    ("Times New Roman", 288, 1),
    ("メイリオ", 24, -1),
    ("メイリオ", 32, 1),
    ("メイリオ", 36, 1),
    ("メイリオ", 40, 1),
    ("メイリオ", 42, 1),
    ("メイリオ", 44, 1),
    ("メイリオ", 48, 1),
    ("メイリオ", 52, 1),
    ("メイリオ", 56, 1),
    ("メイリオ", 60, 1),
    ("メイリオ", 64, 1),
    ("メイリオ", 72, 1),
    ("メイリオ", 80, 2),
    ("メイリオ", 96, 2),
    ("メイリオ", 112, 2),
    ("メイリオ", 144, 3),
    ("メイリオ", 192, 4),
    ("メイリオ", 288, 6),
];

/// The measured default row height for a font, in 96-dpi pixels.
pub fn font_default_row_px(face: &str, size: f32) -> Option<u16> {
    found(face, size).map(|(_, _, px, _)| *px)
        .or_else(|| font_default_geometry_px(face,size).map(|(_,row)|row)).or_else(|| {
        // Past the sizes the full table was measured at, the row alone:
        // the height depends on the em's whole pixels and nothing else.
        (face == "Yu Gothic" || face == "游ゴシック").then(|| {
            let em = em_pixels(size);
            YU_GOTHIC_ROW_PX_BY_EM.iter().find(|(held, _)| *held == em).map(|(_, px)| *px)
        })?
    })
}

/// The row 游ゴシック asks for, in pixels, by its em in whole pixels at
/// 96 dpi. Measured in Excel's VBA by autofitting a row holding one cell at
/// every half point from 1 to 72: sizes landing on the same em ask the same
/// row (17 and 17.5 points both 37 pixels), and below an em of 16 the
/// sheet's own 25-pixel row is all that shows.
const YU_GOTHIC_ROW_PX_BY_EM: &[(i32, u16)] = &[
    (16, 26), (17, 27), (18, 30), (19, 32), (20, 33), (21, 34), (22, 36), (23, 37), (24, 40),
    (25, 42), (26, 43), (27, 44), (28, 45), (29, 47), (30, 48), (31, 51), (32, 53), (33, 54),
    (34, 55), (35, 57), (36, 58), (37, 59), (38, 60), (39, 62), (40, 65), (41, 66), (42, 68),
    (43, 69), (44, 70), (45, 72), (46, 73), (47, 76), (48, 78), (49, 79), (50, 80), (51, 81),
    (52, 83), (53, 86), (54, 87), (55, 89), (56, 90), (57, 91), (58, 93), (59, 96), (60, 97),
    (61, 98), (62, 100), (63, 101), (64, 102), (65, 104), (66, 107), (67, 108), (68, 110),
    (69, 111), (70, 112), (71, 113), (72, 117), (73, 118), (74, 119), (75, 121), (76, 122),
    (77, 123), (78, 127), (79, 128), (80, 129), (81, 130), (82, 132), (83, 133), (84, 134),
    (85, 138), (86, 139), (87, 140), (88, 142), (89, 143), (90, 144), (91, 148), (92, 149),
    (93, 150), (94, 151), (95, 153), (96, 154),
];

/// The box a line of this font stands in, and how far down it the baseline
/// sits.
///
/// How far below the top of the box Excel puts the baseline was measured the
/// same way the heights were and for the same reason: the device's own descent
/// accounts for two thirds of the table and nothing accounts for the rest, so
/// the measurements stand as the data. The weight is taken into account here
/// and not in the row height, which is not measured for it.
pub fn font_line_box(face: &str, size: f32, bold: bool) -> Option<(u16, u16)> {
    if bold {
        let size_q = (size * 4.0).round() as u16;
        if let Some((_, _, px, baseline)) = FONT_BOLD_ROW_PX
            .iter()
            .find(|(f, s, _, _)| *f == face && *s == size_q)
        {
            return Some((*px, *baseline));
        }
    }
    found(face, size).map(|(_, _, px, baseline)| (*px, *baseline))
}

fn found(face: &str, size: f32) -> Option<&'static (&'static str, u16, u16, u16)> {
    let face = if face.eq_ignore_ascii_case("Meiryo") { "メイリオ" } else { face };
    let size_q = (size * 4.0).round() as u16;
    let exact = FONT_DEFAULT_ROW_PX
        .iter()
        .find(|(f, s, _, _)| *f == face && *s == size_q);
    if exact.is_some() {
        return exact;
    }
    // A size the table has not been asked about before — the three quarters
    // of a point a cell shrinks by, say — draws exactly like any other size
    // that lands on the same whole pixel of em, because that pixel is all the
    // device is given. So the entry for any such size is this size's entry.
    let em = em_pixels(size);
    FONT_DEFAULT_ROW_PX
        .iter()
        .find(|(f, s, _, _)| *f == face && em_pixels(*s as f32 / 4.0) == em)
}

/// The line a cell shares with the sheet's Normal font.
pub fn font_cell_line_box(
    face: &str,
    size: f32,
    bold: bool,
    italic: bool,
    normal_face: &str,
    normal_size: f32,
    merged_wrap: bool,
) -> Option<(u16, u16)> {
    let canonical = if face.eq_ignore_ascii_case("Meiryo") { "メイリオ" } else { face };
    let style = FONT_STYLE_LINE_BOXES.iter().find(|(held, points, held_bold, held_italic, _, _)| {
        *held == canonical && *held_bold == bold && *held_italic == italic
            && em_pixels(*points as f32 / 4.0) == em_pixels(size)
    });
    let (height, baseline) = style.map(|(_, _, _, _, height, baseline)| (*height, *baseline))
        .or_else(|| font_line_box(face, size, bold))?;
    if normal_face.is_empty() {
        return Some((height, baseline));
    }
    let (normal_height, normal_baseline) = font_line_box(normal_face, normal_size, false)?;
    let descent = height.saturating_sub(baseline);
    let normal_descent = normal_height.saturating_sub(normal_baseline);
    let device_descent = if merged_wrap {
        let offset = FONT_DEVICE_DESCENT_OFFSETS.iter()
            .find(|(held, points, _)| *held == canonical
                && em_pixels(*points as f32 / 4.0) == em_pixels(size))
            .map_or(0, |(_, _, offset)| *offset);
        (i32::from(descent) - i32::from(offset)).clamp(0, i32::from(u16::MAX)) as u16
    } else { descent };
    let height = baseline.saturating_add(device_descent.max(normal_descent))
        .max(normal_baseline.saturating_add(descent.max(normal_descent)));
    Some((height, baseline.max(normal_baseline)))
}


/// The shared line of a row. Wrapped horizontal merge anchors contribute
/// descent; the ordinary and covered cells establish the row's baseline.
pub fn font_row_line_box<'a>(
    normal_face: &str,
    normal_size: f32,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool)>,
) -> Option<(u16, u16)> {
    font_position_row_line_box(
        normal_face,
        normal_size,
        cells.into_iter().map(|(face, size, bold, italic, anchor)| {
            (face, size, bold, italic, anchor, None)
        }),
    )
}

/// A run's own line, including its raised or lowered position.
pub fn font_run_line_box(
    face: &str,
    size: f32,
    bold: bool,
    italic: bool,
    position: Option<&str>,
) -> Option<(u16, u16)> {
    let superscript = match position {
        None | Some("baseline") => {
            return font_cell_line_box(face, size, bold, italic, "", 0.0, false);
        }
        Some("superscript") => true,
        Some("subscript") => false,
        _ => return None,
    };
    let canonical = if face.eq_ignore_ascii_case("Meiryo") { "メイリオ" } else { face };
    FONT_POSITION_LINE_BOXES.iter().find(|(held, points, held_bold, held_italic, ..)| {
        *held == canonical && *held_bold == bold && *held_italic == italic
            && f32::from(*points) / 4.0 == size
    }).map(|(_, _, _, _, up_height, up_baseline, down_height, down_baseline)| {
        if superscript { (*up_height, *up_baseline) } else { (*down_height, *down_baseline) }
    })
}

/// The shared line of ordinary, raised and lowered runs in a row.
pub fn font_position_row_line_box<'a>(
    normal_face: &str,
    normal_size: f32,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool, Option<&'a str>)>,
) -> Option<(u16, u16)> {
    font_text_row_line_box(
        normal_face, normal_size,
        cells.into_iter().map(|(face, size, bold, italic, anchor, position)| {
            (face, size, bold, italic, anchor, position, 1)
        }),
    )
}

/// Wrapped plain text shares the Normal font at its last baseline.
pub fn font_text_row_line_box<'a>(
    normal_face: &str,
    normal_size: f32,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool, Option<&'a str>, u32)>,
) -> Option<(u16, u16)> {
    let (height, mut baseline) = font_line_box(normal_face, normal_size, false)?;
    let mut descent = height.saturating_sub(baseline);
    for (face, size, bold, italic, wrapped_merge_anchor, position, lines) in cells {
        let (height, cell_baseline) = font_run_line_box(face, size, bold, italic, position)?;
        descent = descent.max(height.saturating_sub(cell_baseline));
        if !wrapped_merge_anchor {
            let extra = u16::try_from(lines.saturating_sub(1)).unwrap_or(u16::MAX);
            baseline = baseline.max(cell_baseline.saturating_add(height.saturating_mul(extra)));
        }
    }
    Some((baseline.saturating_add(descent), baseline))
}


// Delta-coded measured row heights and baseline bounds, one entry per em.
const FONT_ROW_PROFILE_0: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRARRARRARRARRASRARRARRARRASRARRARRASRAQRASRARRARRARRASRARRARRARRARRASRARRARRASRARRA",
    "RRARRARRASRARRARRARRARRARRASRARRARRASRARRARRARRASRARRARRARRARRARRARRASRARRARRARRASRARRARRASRARRARRARRARRARRARRAUSARRA",
    "RRARRASRARRARRASRARRARRARRASRARRAQRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRAQRARRARRASRARRARRASRARRARRARRASRARRA",
    "RRASRARRARRARRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRATSAQRASRARRARRARRASRARRARRASRARRARRARRAQRASRARRASRARRARRA",
    "RRARRARRARRASRARRARRARRASRASRARRARRARRARRARRARRARRARRARRARRASRARRARRASRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRA",
    "RRASRARRATSASRARRARRARRARRARRARRARRARRARRASRASRARRARRASRARRARRARRARRARRARRARRASRARRASRARRARRARRARRARRARRASRARRARRASRA",
    "RRARRARRASRARRARRARRARRARRARRASRARRARRASRARRARRARRARRASRARRAQRARRARRASRARRATSARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRARRASRARRATSARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_1: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRASRARRARRARRARRARRARRASRASRARRARRARRARRA",
    "RRARRASRASRARRARRARRARRARRARRASRARRARRASRARRARRARRARRASRARRARRARRASRARRARRASRARRARRARRARRARRARRASRASRARRARRARRATSARRA",
    "RRARRATRAQRARRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRARRASRARRASRARRARRARRARRARRARRASRARRARRASRARRARRARRA",
    "RRASRARRARRARRARRASRARRARRASRARRARRARRARRARRASRARRASRARRARRATSARRARRARRASRARRARRARRASRARRARRASRARRAQRARRARRARRARRARRA",
    "TRARRARRARRARRAQRARRASRARRARRASRASRAQRARRARRASRARRARRARRARRARRASRARRASRARRARRARRARRARRASRARRARRASRARRARRARRARRARRASRA",
    "RRARRARRATSASRARRARRASRARRARRARRARRARRASRARRASRARRARRARRARRARRARRASRARRARRASRARRARRARRARRASRARRARRARRARRARRASRARRASRA",
    "RRARRARRARRARRARRASRARRASRARRARRARRARRARRASRARRARRARRARRASRAQRARRARRASRARRATSARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRARRASRARRATSARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_2: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRARRARRARRARRASRASRARRARRARRASRARRARRARRARRARRARRASRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRASRARRARRARRARRARRARRATRARRARRARRARRARRARRASRARRARRARRARRASRARRASRAQRARRARRASRARRARRASRARRARRARRATSARRA",
    "RRARRASRARRASRARRARRASRARRARRARRARRARRARRARRARRATRARRARRARRASRARRARRARRARRARRARRARRARRARRARRASRARRASRARRARRARRARRARRA",
    "RRARRASRARRARRARRARRASRARRASRARRARRARRARRARRASRARRARRARRARRAUSARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRARRA",
    "RRASRARRARRASRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRATRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRASRA",
    "SRARRARRATSARRARRARRARRASRARRARRARRARRASRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRARRARRASRASRARRARRARRASRA",
    "RRARRARRARRARRARRARRARRASRARRARRASRASRARRARRARRARRARRARRARRARRARRARRASRARRATSARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRARRASRARRATSARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_3: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRARRASRARRARRARRARRASRARRASRARRARRARRARRASRAQRASRARRARRASRARRARRARRARRASRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRARRARRASRARRASRARRARRARRARRASRARRARRARRARRARRASRARRASRARRARRARRARRARRASRARRARRARRASRARRARRATSARRA",
    "SRARRARRARRARRASRARRARRASRARRARRARRARRARRARRASRARRASRARRARRARRARRARRASRARRARRASRAQRASRARRASRARRARRARRARRARRARRASRARRA",
    "RRASRARRARRARRARRASRARRARRARRARRARRASRARRASRARRARRARRARRASRATSARRARRASRARRARRARRARRARRASRARRARRARRAQRASRARRASRARRARRA",
    "RRARRARRARRASRAQRASRARRARRARRARRATRAQRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRARRASRARRARRARRA",
    "SRARRARRAUSARRARRARRARRARRARRASRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRASRARRARRARRARRARRARRASRARRASRARRARRARRA",
    "RRARRASRARRARRARRARRASRARRASRARRARRARRARRARRARRARRASRARRASRAQRARRARRASRARRATSARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRARRASRARRATSARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_4: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRARRASRARRATSASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRATSASRARRA",
    "RRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRA",
    "RRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRATSASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRARRAUSARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_5: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRA",
    "RRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRA",
    "RRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRA",
    "RRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRA",
    "RRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRA",
    "RRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_6: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRASRAYRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_7: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRA",
    "RRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRA",
    "RRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRA",
    "RRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRA",
    "SRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_8: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRATRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRASRARRA",
    "RRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRATRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRASRARRARRARRATRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRASRASRARRARRARRA",
    "RRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRASRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRASRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRASRARRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRATRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRA",
    "RRARRARRASRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRAZRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRA",
    "RRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRA",
    "RRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRASRARRARRARRA",
    "SRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRA",
    "RRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRARRARRASRARRARRASRARRARRASRARRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_9: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARSARQASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRAUSARRARRARRASRARRARRARRASRARRARRARRA",
    "RRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRAPRASRARRATSARRASRA",
    "RRARRARRASRARRARRARRASRARRARRARRASRARRARRAQRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRA",
    "RRASRARRARRARRATSASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRA",
    "SRARRARRARRARRARRARRARRASRARRARRARRASRATSARRARRASRARRARRARRAPRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRATSARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRATSARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRATSARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRA",
    "RRARRARRASRARRARRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_10: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARSARQARRASRARRARRARRASRARRARRASRARRARRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRARRASRATSARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRARRASRARRA",
    "SRAQRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRATSARRARRASRARRARRASRAQRASRARRARRASRA",
    "QRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRASRAQRARRASRARRARRARRAUSAQRASRA",
    "RRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRASRARRARRASRAQRASRARRARRARRARRASRARRARRARRASRARRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRATSASRARRARRARRASRARRARRASRAQRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRA",
    "RRARRARRARRASRARRARRASRAQRASRARRARRARRAUSARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRATSARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRATSARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRATSARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRA",
    "RRARRARRASRARRARRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_11: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRATSARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRAUSARRARRARRASRARRARRARRASRARRARRARRA",
    "RRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRAPRASRARRATSARRASRA",
    "RRARRARRASRARRARRARRASRARRARRARRASRARRARRAQRARRASRARRARRARRASRARRARRARRARRARRARRARRARRARRASRARRASRAQRASRAQRASRARRARRA",
    "RRASRASRARRARRATSASRARRARRARRASRARRARRARRASRARRARRAQRASRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRASRA",
    "SRARRARRARRARRARRARRAQRASRARRARRARRASRATSASRAQRASRARRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRATSARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRATSARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRATSARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRA",
    "RRARRARRASRARRARRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_12: &str = concat!(
    "XVCQQCQQCRRCRRCQQCSSAQQARRARRASRARRASRARRARRARRASRARRASRARRAQRARRASRARRARRARRARRARRARRARRASRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRATSARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRA",
    "RRARRARRASRARRARRASRARRARRASRARRAQRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRATSARRARRASRARRARRASRARRARRARRARRARRA",
    "SRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRATSARRASRA",
    "RRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRASRARRARRARRAUSARRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRASRARRARRARRARRARRASRARRARRARRASRA",
    "RRARRARRARRASRARRARRARRARRASRARRASRARRATSARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRA",
    "RRARRASRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRATSARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRA",
    "SRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRATSARRARRASRARRARRARRARRASRARRA",
    "RRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRATSARRA",
    "SRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRARRARRARRASRARRARRARRARRASRARRARRARRASRARRA",
    "RRARRARRASRARRARRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_13: &str = concat!(
    "XVCQQCRRCQQCRRCRRBRRARQBRSARRAQQARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRARRA",
    "SRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRA",
    "RRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRA",
    "SRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRA",
    "RRARRARRASRARRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRA",
    "SRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRARRASRARRARRA",
    "RRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRA",
    "SRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRA",
    "RRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRA",
    "SRARRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRA",
    "RRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRARRASRARRARRARRARRARRA",
    "RRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRARRARRARRARRASRARRARRA",
    "RRARRARRARRASRARRARRARRARRARRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_14: &str = concat!(
    "XVCQQCQQCRRCSRDRRCRRBRRBSRARRARRAUSARRARRASRARRARRATRASSARRARRASRARRATSASQARRARSARRASRARRATRASSARQARSASRARRARRARRASRA",
    "TSARRASRARRARRASRARRATSASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRASRATSARRASRARRARRARRAUSARRARRASRARRARRAUSA",
    "RRARRARRASRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRASRARRARRATSASRARRARRASRARRATSASRARRARRARRASRATSARRA",
    "SRARRARRASRARRATSARRASRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRASRATSARRASRARRARRARRAUSARRARRASRARRARRAUSARRARRA",
    "RRASRARRARRAUSARRARRASRARRARRATSASRARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRASRARRATSASRARRARRARRASRARRATSASRARRA",
    "RRASRARRATSARRASRARRARRASRATSARRASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRAUSARRARRASRARRARRASRATSARRARRASRA",
    "RRARRAUSARRARRASRARRARRATSASRARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRASRARRATSARRASRARRARRASRARRATSASRARRARRASRA",
    "RRATSARRASRARRARRASRARRATSASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRASRATSARRASRARRARRARRAUSARRARRASRARRARRA",
    "USARRARRASRARRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRASRARRARRARSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSA",
    "QRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRA",
    "QRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRA",
    "QRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRA",
    "QRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_15: &str = concat!(
    "XVCQQCQQCRRCRRCSRCRRBRRBSRARRARRAUSARRARRARRASRARRATRASSARRARRASRARRATSARQASRARSARRASRARRATRASSARQARSARRASRATRARRASRA",
    "RSARRASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRAUSARRARRASRARRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRATSA",
    "SRARRARRASRARRARRATSASRARRARRASRARRATSARRASRARRARRASRATSARRASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRAUSARRA",
    "RRASRARRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRASRARRARRATSASRARRARRASRARRATSASRARRARRARRASRATSARRASRA",
    "RRARRASRARRATSARRASRARRARRASRATSARRASRARRARRARRASRATSARRASRARRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRARRAUSARRARRA",
    "SRARRARRATSASRARRARRASRARRATSASRARRARRARRASRARRATSASRARRARRASRARRATSARRASRARRARRASRATSARRASRARRARRARRASRATSARRASRARRA",
    "RRASRATSARRARRASRARRARRAUSARRARRASRARRARRARRAUSARRARRASRARRARRAUSARRARRARRASRARRATSASRARRARRARRASRARRATSASRARRARRASRA",
    "RRATSARRASRARRARRASRARRATSASRARRARRARRASRATSARRASRARRARRASRATSARRARRASRARRARRASRATSARRASRARRARRARRAUSARRARRASRARRARRA",
    "USARRARRARRASRARRARRAUSARRARRASRARRARRATSASRARRARRASRARRATSASRARRARRARRASRARRASSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSA",
    "QRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRA",
    "QRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRA",
    "QRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRA",
    "QRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRA",
    "QRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRA",
);

const FONT_ROW_PROFILE_16: &str = concat!(
    "XVCQQCQQCQQCRRCSRDRRCSRATSARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRASRARRARRASRARRASRARRA",
    "SRARRATSASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRA",
    "SRARRASRARRASRARRARRASRARRASRARRASRARRATSASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRARRA",
    "SRARRASRARRASRARRASRATSARRASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRATSASRARRASRARRASRARRASRARRARRASRARRASRARRASRA",
    "RRASRATSARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRATSASRARRASRA",
    "RRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRASRARRA",
    "RRASRARRASRARRASRARRATSASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRARRASRARRASRARRASRARRA",
    "SRATSARRASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRATSASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRA",
    "SRARRASRARRARRASRARRASRARRASRARRASRATSARRASRARRASRARRASRARRASRARRARRASRARRASRARRASRARRATSASRARRASRARRASRARRASRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRA",
    "QSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQRAQSAQRAQRAQRAQRA",
);

static FONT_ROW_PROFILES: &[(&str, u8, &str)] = &[
    ("Arial", 0, FONT_ROW_PROFILE_0),
    ("Arial", 1, FONT_ROW_PROFILE_1),
    ("Arial", 2, FONT_ROW_PROFILE_2),
    ("Arial", 3, FONT_ROW_PROFILE_3),
    ("Aptos Narrow", 0, FONT_ROW_PROFILE_4),
    ("Aptos Narrow", 1, FONT_ROW_PROFILE_4),
    ("Aptos Narrow", 2, FONT_ROW_PROFILE_4),
    ("Aptos Narrow", 3, FONT_ROW_PROFILE_4),
    ("Calibri", 0, FONT_ROW_PROFILE_5),
    ("Calibri", 1, FONT_ROW_PROFILE_5),
    ("Calibri", 2, FONT_ROW_PROFILE_5),
    ("Calibri", 3, FONT_ROW_PROFILE_5),
    ("Courier New", 0, FONT_ROW_PROFILE_6),
    ("Courier New", 1, FONT_ROW_PROFILE_7),
    ("Courier New", 2, FONT_ROW_PROFILE_8),
    ("Courier New", 3, FONT_ROW_PROFILE_7),
    ("Times New Roman", 0, FONT_ROW_PROFILE_9),
    ("Times New Roman", 1, FONT_ROW_PROFILE_10),
    ("Times New Roman", 2, FONT_ROW_PROFILE_11),
    ("Times New Roman", 3, FONT_ROW_PROFILE_12),
    ("ＭＳ Ｐゴシック", 0, FONT_ROW_PROFILE_13),
    ("ＭＳ Ｐゴシック", 1, FONT_ROW_PROFILE_13),
    ("ＭＳ Ｐゴシック", 2, FONT_ROW_PROFILE_13),
    ("ＭＳ Ｐゴシック", 3, FONT_ROW_PROFILE_13),
    ("ＭＳ 明朝", 0, FONT_ROW_PROFILE_13),
    ("ＭＳ 明朝", 1, FONT_ROW_PROFILE_13),
    ("ＭＳ 明朝", 2, FONT_ROW_PROFILE_13),
    ("ＭＳ 明朝", 3, FONT_ROW_PROFILE_13),
    ("游ゴシック", 0, FONT_ROW_PROFILE_14),
    ("游ゴシック", 1, FONT_ROW_PROFILE_15),
    ("游ゴシック", 2, FONT_ROW_PROFILE_14),
    ("游ゴシック", 3, FONT_ROW_PROFILE_15),
    ("メイリオ", 0, FONT_ROW_PROFILE_16),
    ("メイリオ", 1, FONT_ROW_PROFILE_16),
    ("メイリオ", 2, FONT_ROW_PROFILE_16),
    ("メイリオ", 3, FONT_ROW_PROFILE_16),
];

// Accounting underline row bounds: tiny baseline limits and height deltas.
const ACCOUNTING_ROW_PROFILE_0: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCCCECCCKCCEAECCCECCCCKCCECCCCECCCCIECCECCCECCCCCIECCCECCECCCCCIICCCECCECCCEIAECCCEC",
    "CECCCEICEACCECCECCCEICECCCCCCECCCKCCECCCEGAECCCKCCECCCAECECCCICCECCCEECCCCICCCCCECCEECCCICCECCCCCCECEIGECCCCCCCC",
    "CEEICECCCCCCCECECICCCCECCECCCEICCCCCECCECCCCKCACCECGCECCCCKCCCCECCCCECCIECCCCECCCCECCIECCCCECCCCECCICECCCECCCCEC",
    "MCCECCCECCCCECICCECCCCECCCECICCECCCCECCCEICCCECCCCECCCCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_1: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCCCECCCECCEAECCCECCCCKCCECCCCECCCCCECCKCCCECCCCCCECCCECCKCCCCCCICCCECCECCCKCFKCCCEC",
    "CECCCECCKFCCKCCECCCECCECCCCCCKCCCECCECCCEGAECCCKCCECCCAECECCCCCCECCCKECCCCCCCCCCECCEKCCCCCCECCCCCCECECGKCCCCCCCC",
    "CEECCECCCCCCCKCECCCCCCECCECCCKCCCCCCECCECCCCECACCECGCKCCCCECCCCECCCCECCCKCCCCECCCCECCCECCCCKCCCCECCCCECCCECCCCKC",
    "GCCECCCECCCCECCCCKCCCCECCCECCCCECCCCKCCCECCCCECCCCECCCCKCCCECCGCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_2: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCKCCCCECCCKCCECCCCCCKECCCCCCEKCCCCCCECCKCCCCECCCKCCECCCCCCKECCCGCCCMACCCCCECCIECCCE",
    "CCCCIECECCCCCIECCECCCCEICCCECCECICCCECECCGICCECCCECIECACCCCCGICCCACECCKEACCECCCCzBgCCECCCCCKCCECCCCCEICCGECCECIC",
    "CCECECCCICCECCECCICECCCCCECKCCCCCCECKCCCCCECCCIEACCECGCKCCCCECCCCKCCCCECCCEICCCECCCCKCCCECCCCEICCCECCCCKCCCECCCC",
    "EIGCCECCCEICCCECCCCEICCCECCCEICCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAA",
);

const ACCOUNTING_ROW_PROFILE_3: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCECCCCECCCECCECCCCCCKECCCCCCEECCCCCCKCCECCCCECCCECCKCCCCCCEECCCGCCCMFCCCCCKCCCECCCE",
    "CCCCCECKCCCCCCECCECCCCECCCCKCCECCCCCECECCGCCCKCCCECCECACCCCCGCCCCACECCKEACCECCCCCECECCCCCKCCECCCCCECCCGECCKCCCCC",
    "ECECCCCCCECCKCCCCECCCCCECECCCCCCKCECCCCCECCCCEACCECGCKCCCCECCCCECCCCECCCKCCCCECCCCECCCECCCCKCCCCECCCCECCCECCCCKC",
    "GCCECCCECCCCECCCCKCCCCECCCECCCCECCCCKCCCECCCCECCCCECCCCKCCCECCGCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_4: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCCCEECCIECCCCCCECCECCICECCCECCCECCICCCGCCCCCCECCICECEACCECCECCIGCCCECECCECCICCCCCGC",
    "CCECCCICCCCCCECECCCCICCECCCCECECCICCECCCCICCCECICECCCCECCCCCCKCCECCECCCCCCKCCCCCCGCCCCCEICCCCCECCCEECIGCCCCECCCC",
    "ECCKCCCCECCCCECCCICCEECCCECCCCICCCECCEECCCCCICCCCECGCECCCCKCCCCECCCCECCIECCCCECCCCECCIECCCCECCCCECCICECCCECCCCEC",
    "MCCECCCECCCCECICCECCCCECCCECICCECCCCECCCEICCCECCCCECCCCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_5: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCCCEECCCECCCCCCECCKCCCCECCCECCCECCCCCCMCCCCCCECCCCECKFCCKCCECCCGCCCECECCKCCCCCCCCGC",
    "CCECCCCCCCCCCKCECCCCCCCECCCCECKCCCCCECCCCICCCECCCKCCCCECCCCCCECCECCKCCCCCCECCCCCCGCCCCCKCCCCCCECCCEECCGCCCCKCCCC",
    "ECCECCCCECCCCKCCCCCCEECCCECCCCCCCCKCCEECCCCCCCCCCECGCKCCCCECCCCECCCCECCCKCCCCECCCCECCCECCCCKCCCCECCCCECCCECCCCKC",
    "GCCECCCECCCCECCCCKCCCCECCCECCCCECCCCKCCCECCCCECCCCECCCCKCCCECCGCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_6: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCCECCCCKCECCCCKLQCCECCCCKCCCECCCEICCCCCECECICCECCCCCKCECCCCCECICECCGCECICCECCECCCICCECE",
    "CCCCIECCEAECEICCCCCECCEICCCECCCCIECECCCCEGICECCCCCEICCAECECCCICCEAECCCzBiCACCCECCCKCCCCECCCKCCCCCECCCKCCICCCCC/B",
    "sCCCECCCCEICCCECCECICCCCECECCICCECCCCEIECCCCCCCEIEACCECGCKCCCCECCCCKCCCCECCCEICCCECCCCKCCCECCCCEICCCECCCCKCCCECC",
    "CCEIGCCECCCEICCCECCCCEICCCECCCEICCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAA",
);

const ACCOUNTING_ROW_PROFILE_7: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCCECCCCECECCCCEAECCECCCCKCCCECCCECCCCCCECKCCCCECCCCCECECCCCCKCCCECCGCECCCCECCKCCCCCCECE",
    "CCCCCECCKFKCECCCCCCECCECCCCKCCCCCECECCCCEGCCKCCCCCECCCAECECCCCCCEAECCCCMACCCECCCECCCCECCCKCCCCCECCCECCICCCCCCKCC",
    "ECCCCECCCCECCKCCCCCCECECCCCCECCCCKCECCCCCCCECEACCECGCKCCCCECCCCECCCCECCCKCCCCECCCCECCCECCCCKCCCCECCCCECCCECCCCKC",
    "GCCECCCECCCCECCCCKCCCCECCCECCCCECCCCKCCCECCCCECCCCECCCCKCCCECCGCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_8: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECECCCECCCECCECCCECCECCCECCCECCECCCECCCECGECCCECCCEICECCCECCECCCECCCECCECCCECCCECCECCC",
    "ECCCECIECCCECCECCCECCCECCECCCECCCECCECCCECCCECCKCCCECCECCGECCCECCECCCECCCECCECCCECCCECCEICCECCECCCECCCECCECCCECC",
    "CECCECCCECCCECCECICECCECCCECCCECCECCCECCCECCECCGECCCECCECCIECCCECCECCCECCECCCECCCECCECCCECCCECCECCCKCCCECCECCCEC",
    "CECCCECCCECCECCCECCCECCECCCEICCECCECCCICCECCCECCCECCECCCECCCECCECCCECICECCECCCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_9: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCECCECCCECICECCECCCECCCECCIECCECCCECCCECCKCCCECCCECCCECCKCCCECCCECCCECCEICCECCCECCEC",
    "CCEICCECCCECCECCCEICCECCCECCECCCEICCECCECCCECCCECICECCECCCECCCECIECCCECCCECCCECIECCCECCCECCCECCKCCCECCCECCECCCKC",
    "CCECCCECCECCCKCCCECCCECCECCCKCCCECCECCCECCCEICCECCECCCECCCEICECCCECCCECCCEICECCCECCCECCCECIECCCECCCECCECCIECCCEC",
    "CCECCECCIECCCECCECCCECCIECCCECCECCCECCCKCCCECCECCCECCAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_10: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCECCECCCECCCECCKCCCECCCECCCECCKCCCECCCECCECCCKCCCECCCECCECCCKCCCECCCECCECCCKCCCECCEC",
    "CCECCCKCCCECCECCCECCCKCCCECCECCCECCCKCCECCCECCCECCCKCCECCCECCCECCKCCCECCCECCCECCKCCCECCCECCCECCKCCCECCCECCECCCKC",
    "CCECCCECCECCCKCCCECCCECCECCCKCCCECCECCCECCCKCCCECCECCCECCCKCCECCCECCCECCCKCCECCCECCCECCCKCCECCCECCCECCKCCCECCCEC",
    "CCECCKCCCECCCECCECCCKCCCECCCECCECCCKCCCECCCECCECCCIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_11: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECECCEECCCECCGCCCECCCEECECGCECCECCCGCCCECCEECCECCCICCECCEECCCECCEECCECCCEECGCCEECCCECC",
    "GCCCECCCEECECGCECCECCCGCCCECCEECCECCCEGCECCEECCCECCEECCECCCEECGCCEECCCECCGCCCECCCEECECGCECCECCCGCCCECCEECCECCCEG",
    "CECCEECCCECCEECCECCCEECGCECECCCECCGCCCECCCEECEQGCCECCECCGCCCECCEECECCCECGECCEECCECCECCGCCCECCEECGCECECCECCEECCEC",
    "CECCGCCCEGCECCECCEECCECCEECECCCECGECCEECCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_12: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECECCCGCCAGCCGCCCECCCEECECCCICBICCCGCCCECCEECCECCCICCECCCGCCAGCCEECCECCCEECGCCGCCCECCC",
    "GCCCECCCEECECGCECCECCCGCCCECCEECCECCCEGGACCGCCCECCCEECCECCCEECGCCKBCCIBCCGCCCECCCEECECKCACGACCCGCCCECCEECCECCCEG",
    "KDCCKBCCIBCCEECCECCCEECGCOCFCCMFCCGCCCECCCEECEQGCCECCECCGCCCECCEECECCCECGECCEECCECCECCGCCCECCEECGCECECCECCEECCEC",
    "CECCGCCCEGCECCECCEECCECCEECECCCECGECCEECCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_13: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCEECCKDCCKDCKCDKCCCECCKDCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCEC",
    "CCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCC",
    "ECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKC",
    "CECCECCCECCKCCECCCECCECCKCCECCCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_14: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCEECCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCEC",
    "CCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCC",
    "ECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKCCECCECCCECCKCCECCCECCECCKCCCECCECCECCCKC",
    "CECCECCCECCKCCECCCECCECCKCCECCCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_15: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCGCCCEECCCECCCGCCCECCCCGCCCIECCECCCCGCCCECCCGCCCECCGCGCCCGCCCECCCCGCCCECCCGCECGECCCECCC",
    "GCCCECCCCGCCCICCEECCCCGCCCECCCGCCCECCCGECECGCCCECCCCGCCCECCCGCECEEECCECCCGCCCECCCCGCCCICCCEECCCGCCCECCCGCCCECCCG",
    "ECCCICCCECCCCGCCCECCCGCECGCCCEECCCGCCCECCCCGCCSGCCECCECCGCCCECCEECECCCECGECCEECCECCECCGCCCECCEECGCECECCECCEECCEC",
    "CECCGCCCEGCECCECCEECCECCEECECCCECGECCEECCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_16: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCGCCCEECCCECCCGCCCECCCCGCCCEICCECCCCGCCCECCCGCCCECCGCGCCCGCCCECCCCGCCCECCCGCCCIECCCECCC",
    "GCCCECCCCGCCCICCEECCCCGCCCECCCGCCCECCCKECACGCCCECCCCGCCCECCCGCGCCKBCCECCCGCCCECCCCGCCCMCCCEACCCGCCCECCCGCCCECCCO",
    "ECCCACCCECCCCGCCCECCCGCKCKCCCFECCCGCCCECCCCGCCSGCCECCECCGCCCECCEECECCCECGECCEECCECCECCGCCCECCEECGCECECCECCEECCEC",
    "CECCGCCCEGCECCECCEECCECCEECECCCECGECCEECCECCECCGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_17: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCECCCECCCECCCECCCIECCCECCCECCCECCCECCCKCCCECCGCECCCECCCECCIECCCECCCECCCECCCCECCIECC",
    "CECCCECCCICCCECCIECCCCECCCECCCECCCECCIECCCECCCECCCCECCBECMCECCCECCCECCCECCACECICECCCECCCECCCECCCECICECCCGECCCECC",
    "CECCCECICECCCECCCECCCCECCCEICCCCCCECCCEGCCECCCBIECCCCECCCECCCCECCCCEICCECCCCECCCECCCCEGCICECCCECCCCECCCCECCCEICC",
    "CECCCECCCCECCCCECICECGCCECCCCECCCECCCCKCCCECCCCECCCCECCCECICCECCCCECEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_18: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCECCCECCCECCCECCCCKCCCECCCECCCECCCECCCKCCCECCGCECCCECCCECCIECCCECCCKCCCBCCICBCCIECC",
    "CKCCCECCCICCCBCCIKCCCCECCCECCCECCCBCCOECCCECCCECCCCKCCHECGCKCCCECCCECCCECCACECCCKCCCECCCECCCECCCKCCCECCCGECCCECC",
    "CKCCCECCCECCCECCCKCCCCECCCECCCCCCCECCCKGCCECCCHCKCCCCECCCECCCCECCCCKCCCECCCCECCCECCCCKGCCCECCCECCCCECCCCKCCCECCC",
    "CECCCECCCCKCCCCECCCECGCCECCCCKCCCECCCCECCCECCCCKCCCCECCCECCCCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_19: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCCECCCKCCECCCCCCEICCECCCCECICECCCCECCCKCCCECCCCCKGCCCECCCECICECCCCCECKAECCCCECCCIECCCEC",
    "CCCKCCCCEGCCECIEAECCEAECICECCCCECCCIECCCECCCEICCECEACECCIIAECCCECCICECCCCECCEICEAECCCCEICCECCCCCECICECCCGECCIECC",
    "EAECCCCKCCCECCCCEICCCECCCECCICECCEAECCIICCCCECCCEICCCECCCECICCECCCCECCIECCCCECCCKCCCCEGCCCEICCECCCCECICCECCCECCC",
    "IECCCECCCCKCCCCECCCECMCCECCCCECICAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_20: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCCECCCECCECCCCCCECCCECCCCKCCCECCCCECCCECCCKCCCCCEGCCCECCCECCCKCCCCCECEAECCCCECCCCKCCCEC",
    "CCCECCCCEGCCKCCEAECCEAECCCECCCCKCCCCECCCECCCECCCKCEACECCCIAECCCECCCCKCCCCECCECCEAECCCCKCCCECCCCCECCCECCCGKCCCECC",
    "EAECCCCECCCKCCCCECCCCECCCECCCCKCCEAECCCICCCCECCCKCCCCECCCECCCCECCCCKCCCECCCCECCCECCCCKGCCCECCCECCCCECCCCKCCCECCC",
    "CECCCECCCCKCCCCECCCECGCCECCCCKCCCECCCCECCCECCCCKCCCCECCCECCCCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_21: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCECCCECCCECCCECCCIECCCECCCECCCECCCECCCKCCCECCGCECCCECCCECCIECCCECCCECCCECCCCECCIECC",
    "CECCCECCCICCCECCIECCCCECCCECCCECCCECCIECCCECCCECCCCECCBECMCECCCECCCECCCECCACECICECCCCCCCCCECEAEAECICEECCGECCCECC",
    "CECCAECICECCCECCCECCCCECCEEICCCCCAECCCEGEAECCCCIECCCCECCCECCCCECCCCEICCECCCCECCCECCCCEGCICECCCECCCCECCCCECCCEICC",
    "CECCCECCCCECCCCECICECGCCECCCCECCCECCCCKCCCECCCCECCCCECCCECICCECCCCECEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_22: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECCCECCCECCCECCCECCCCKCCCECCCECCCECCCECCCKCCCECCGCECCCECCCECCIECCCECCCKCCCBCCICBCCIECC",
    "CKCCCECCCICCCBCCIKCCCCECCCECCCECCCBCCOECCCECCCECCCCKCCHECGCKCCCECCCECCCECCACECCCKCCCCCCCCCECEAEAECCCEKCCGECCCECC",
    "CECCAECCCKCCCECCCECCCCECCKECCCCCCAECCCEGEAECCCCCKCCCCECCCECCCCECCCCKCCCECCCCECCCECCCCKGCCCECCCECCCCECCCCKCCCECCC",
    "CECCCECCCCKCCCCECCCECGCCECCCCKCCCECCCCECCCECCCCKCCCCECCCECCCCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_23: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECECCIECECACECCCICCCECCCECICCECCCECCCKCCCCECCCEIGCCCECCCCEICCECCCCECICCECCECCECGECCCEC",
    "CCCKCCCCEGCCECIECCCCCECCICECCCCECCCIECCCECCCCKCCCECCCCECIGCECCCECCICECCCECCCEICCCECCCCEICCCECCCCECICECCCICCCICEC",
    "CCECCCCKCCCECCCECICCCECCCECCICECCCCECEIGCCCCECCCEICCCECCCECICCECCCCECCIECCCCECCCKCCCCEGCCCEICCECCCCECICCECCCECCC",
    "IECCCECCCCKCCCCECCCECMCCECCCCECICAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_24: &str = concat!(
    "FFFFFFGGHHHHJJJJUAACCAEACCECECCCECECACECCCCCCCECCCKCCCCECCCECCCECCCCKCCCECGCCCECCCCECCCKCCCCECCCCECCECCKCFKCCCEC",
    "CCCECCCCEGCCKCCECCCCCECCCCECCCCKCCCCECCCECCCCECCCKCCCCECCGCECCCECCCCKCCCECCCECCCCECCCCKCCCCECCCCECCCECCCOCCCCCEC",
    "CCECCCCECCCKCCCECCCCCECCCECCCCKCCCCECECGCCCCECCCKCCCCECCCECCCCECCCCKCCCECCCCECCCECCCCKGCCCECCCECCCCECCCCKCCCECCC",
    "CECCCECCCCKCCCCECCCECGCCECCCCKCCCECCCCECCCECCCCKCCCCECCCECCCCECCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_25: &str = concat!(
    "FFFFGGGGHHIIJJKKUACACCCCCCACCCCCCECCICDCKCCCCCCECICCCJCQCCCCCCEICCCCPWCCCCCCKCCCCCCECCCCCCKCCCCCCECCCCCPcCCCCCCE",
    "CCCCCVCiBCCCCCCECCCVCCiBCCCCCCECCbCCCoBCCCCCCECCIChBCoBCCCCCCECICCCnBuBCCCCCCCEICCCCnBuBCCCCCCKCCCCCCECCCCCCKCCC",
    "CCCECCCCCtB6BCCCCCCECCCCCzBgCCCCCCCCECCCzBCCgCCCCCCCECCC5BCCmCCCCCCCECCI/BCCmCCCCCCCECCIClCCsCCCCCCCECICCClCCsCC",
    "CCCCCEICCCCrCyCCCCCCCKCCCCCCECCCCCCKCCCCCCECCCCCrC4CCCCCCCECCCCCxCC+CCCCCCCECCCxCCC+CCCCCCCECC3CCCCkDCCCCCCECCIC",
    "9CC8CAAAAAAAAAAAA7B8BAAAAAAAAAAAAATUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
);

const ACCOUNTING_ROW_PROFILE_26: &str = concat!(
    "FFFFGGGGHHIIJJKKUACACCCCCCACCCCCCECCCCCCECCCCCCECCCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCE",
    "CCCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCC",
    "CECCCCCCECCCCCCKCCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCCECCCCCCECCCCCCKCCCCCCECCC",
    "CCCECCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCKCCCCCCECCCCCCECCCCCCECCCCCCCKCCCCCCEC",
    "CCCCCEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_27: &str = concat!(
    "FFFFGGGGHHIIJJKKUACACCCCCCACCCCCCECCCCCCECCCCCCKCCCCCDCKCCCCCCECCCCCCKCCCCDCKCCCCCCECCCCCCKCCCCJCQCCCCCCECCCCCCK",
    "CCCCCJCQCCCCCCECCCCCCKCCCCJCQCCCCCCECCCCCCKCCCCJCQCCCCCCECCCCCCKCCCCCPCWCCCCCCECCCCCCKCCCCPCWCCCCCCECCCCCCKCCCCP",
    "CWCCCCCCECCCCCCKCCCCCPCWCCCCCCECCCCCCKCCCCVCcCCCCCCECCCCCCKCCCCVCcCCCCCCECCCCCCKCCCCCVCcCCCCCCECCCCCCKCCCCVCcCCC",
    "CCCECCCCCCKCCCCbCiBCCCCCCECCCCCCKCCCCCbCiBCCCCCCECCCCCCKCCCCbCiBCCCCCCECCCCCCKCCCCbCiBCCCCCCECCCCCCKCCCCChBCoBCA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAA",
);

const ACCOUNTING_ROW_PROFILE_28: &str = concat!(
    "FFFFFFGGHHIIJJKKUAACECCCECCICCECCGECCECGECCCEIGECCECCCEGCECCECGECICEGCECCEGCCECCEGCEICCICCECCICCCECCICCEICICCCEC",
    "GECCECCGECCEIGECCCEGCECCECGCECCKGCECCEGCCECCEGCECCCOCCECCICCCECCICCECCGKCCECCICCCECGECCECGEICCECGECCECGCECCEGCKC",
    "CCEGCECCEGCCECCICCKCCEGCCECCICCECCGECCKCCICCCECGECCECGCECIECGECCECGCECCECGECCIEGCECCEGCCECCEGCECCIICAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_29: &str = concat!(
    "FFFFFFGGHHIIJJKKUAACCECCECCICCCECGECCECGCECCEIGECCCEGCECCECCCEGCEICEGCCECCICCECCCICCKCCICCCECGECCECCGECCKCGCECCE",
    "GCECCCEGCECCKGCCECCICCECCCICCECIICCCECGECCECCGECCECMECCCEGCECCECGCECCEGIECCCEGCECCCICCECCICICECCICCECCGECCECGEIC",
    "CECGECCECGCECCEGCEICCEGCECCEGCCECCICCEICCICCECCICCCECGECCIECGECCECGCECCECGECCIEGCECCEGCCECCEGCECCIICCAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

const ACCOUNTING_ROW_PROFILE_30: &str = concat!(
    "FFFFFFFFGGHHIIJJUAAACECEGCECECECCECECECEGCECECKCECCECECECGECECECECIECECECEGCECECECCECEIECEGCECECECECCECECEIGECEC",
    "ECECCECECECEGCKCECECCECECECEGCECECKCECCECECECGECECECECIECECECEGCECECECCECEIECEGCECECECECCECECEIGECECECECCECECECE",
    "GCKCECECCECECECEGCECECKCECCECECECGECECECECIECECECEGCECECECCECECKCEGCECECECECCECECECMECECECECCECECECEGCEIECECCECA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AA",
);

static ACCOUNTING_ROW_PROFILES: &[(&str, u8, CellUnderlineKind, &str)] = &[

    ("Arial", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_0),

    ("Arial", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_1),

    ("Arial", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_2),

    ("Arial", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_3),

    ("Arial", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_4),

    ("Arial", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_5),

    ("Arial", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_6),

    ("Arial", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_7),

    ("Aptos Narrow", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Narrow", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Aptos Display", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_8),

    ("Calibri", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_9),

    ("Calibri", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_10),

    ("Calibri", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_9),

    ("Calibri", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_10),

    ("Calibri", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_9),

    ("Calibri", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_10),

    ("Calibri", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_9),

    ("Calibri", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_10),

    ("Courier New", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_11),

    ("Courier New", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_12),

    ("Courier New", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_13),

    ("Courier New", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_14),

    ("Courier New", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_15),

    ("Courier New", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_16),

    ("Courier New", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_13),

    ("Courier New", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_14),

    ("Times New Roman", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_17),

    ("Times New Roman", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_18),

    ("Times New Roman", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_19),

    ("Times New Roman", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_20),

    ("Times New Roman", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_21),

    ("Times New Roman", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_22),

    ("Times New Roman", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_23),

    ("Times New Roman", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_24),

    ("ＭＳ Ｐゴシック", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_25),

    ("ＭＳ Ｐゴシック", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ Ｐゴシック", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_25),

    ("ＭＳ Ｐゴシック", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ Ｐゴシック", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_25),

    ("ＭＳ Ｐゴシック", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ Ｐゴシック", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_25),

    ("ＭＳ Ｐゴシック", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ 明朝", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_27),

    ("ＭＳ 明朝", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ 明朝", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_27),

    ("ＭＳ 明朝", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ 明朝", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_27),

    ("ＭＳ 明朝", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("ＭＳ 明朝", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_27),

    ("ＭＳ 明朝", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_26),

    ("游ゴシック", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_28),

    ("游ゴシック", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_28),

    ("游ゴシック", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_29),

    ("游ゴシック", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_29),

    ("游ゴシック", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_28),

    ("游ゴシック", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_28),

    ("游ゴシック", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_29),

    ("游ゴシック", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_29),

    ("メイリオ", 0, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 0, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 1, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 1, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 2, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 2, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 3, CellUnderlineKind::SingleAccounting, ACCOUNTING_ROW_PROFILE_30),

    ("メイリオ", 3, CellUnderlineKind::DoubleAccounting, ACCOUNTING_ROW_PROFILE_30),

];

// Measured row minima and effective upper/lower extents, delta-coded per pixel em.
const RUN_ROW_PROFILE_0: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCGGACCAEEACCAEEAECCCCACCACCAECCEEAAAAIGCDBBCACEEAGGAAAAGECCCACCAAAAEEAECCCCA",
    "EEAGECEEAEEABBAEEAECCCCAEEAAAAAAACCAIGCCCACCAECCCCAEEAAAAECCCCAEEACCAEEAEEACCACACEEAAAAEEAECCCCAEEAGECAAAGGA",
    "AAAEEACCAEEAEAEGGAAAACCAECCIIAAAAGECBBAEEACCAECCAAAEGBECCAAACCAGGACACEEACCAGECAAACCAAAAGECEEACCAECCACBCCAEEA",
    "ECCCCACCAGECAAACCAGGAECCAAAEEAECCEEAAAACCACCAEEAAAAGECGGAAAAAAAGECAAAGGAECCCCAAAAEEACACKICACBGECAAAEEABBAIGC",
    "CCACCACACIIAAAACCAACBGECCCAECCAAAEEAEEAAAAEEACCAECCEEACCACCAECCECCAAAEEAEEACCAEEAAAACCAEEACCACCAGECCCAAAAECC",
    "GECEEACCACCAAAAGGAAAAECCCCAEEAAAAGGACCACCAECCEEACACEEAGECECCCCAEEAAAAEEACCACCAEEAEEABBAECCGECCCACCAIGCCCAAAA",
    "CCACCAEEAEEAAAAECCEEACACEEAEEACCACCACCACCACACGGAAAAGECAAAEEAEEAECCAAAGGAAAACCACCAGGAABCEEACCAGECCCAEEAAAAGGA",
    "ABCCCADBBEEAAAAIGCCCAGECCCAGECAAAEEAAAACCAECCEEAAAAIIACCAECCAAAEEAAAAEEAECCCCAAAAIIACACEEACCACCAAAAGECAAAEEA",
    "CCACCAGECEEAAAAEEAECCCCAAAAEEAAAAGECEEACCACCAEEACACEEACCACCACCAECCAAAGGAEEAECCCCACCAAAAEEAECCCCAGECEEAAAAIGC",
    "CCACCACCAECCAAAEEACCACCAGECCCAAAAGGACCAECCCCACCAAAAEEAECCEEACCAGGACACCCACCACCACCAGECAAAEEACCACCAGECEEAAAACCA",
    "ECCCCACCAEEACCAECCEEACCACCAEEACACCCACCACCAEEAGECAAACCAEEAECCC1crFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_1: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECAAAGECGECCCAGECCCAAAACCACCAGECAAACCACCAAAAMIEAAACCACCACCAECCGECCCACCACCAECCGEC",
    "BABAAAGECGECCCACCAAAACCAKGECCACCACCAAAAGECCCAECCCCACCAGECAAAGECCCACCAGECCCACCAAAAGECGECCCAAAACCACCAECCCCACCA",
    "GECCCAGECCCAECCGECCCAGECAAACCACCAECCGECCCACCAAAACCAIGCAAACCACCAGECECCCCACCAAAACCAKGECCAAAACCACCAIEEEEACCACCA",
    "AAAGECCCAAAACCACCAKGEAAACCACCAAAAGECGECCCACCACCAGECECCCCACCACCAECCCCAGECAAACCACCAGECAAAGECEEAECCCCAGECCCAAAA",
    "GECCCAECCCCACCAGECEEAAAAGECAAAGECCCAGECCCABABGECAAACCAGECCCAECCAAACCAGECCCAGECCCACCACCAGECGECCCAAAAGECCCAECC",
    "AAACCACCACCAGECCCACCAAAAGECGECCCACCACCACCAIEECCAAAACCAKGEECCCCACCACCACCAKGECCACCABABCCAIEEAAACCACCAAAAGECGEC",
    "CCACCACCAKGECCAAAACCAAAAGECBABCCAGECCCAGECECCCCACCAAAAGECCCACCAAAACCAGECGECCCACCACCAECCCCAGECAAAGECGECCCACCA",
    "AAACCAIGCCCACCAAAACCAKGECCAAAACCACCAKGECCAAAACCACCAKGEGECAAACCACCAGECCCAAAACCACCAGECAAAGECCCACCAGECAAACCACCA",
    "CCAGECAAAGECCCACCAIEECCACCACCACCAIEECCACCACCACCAECCCCACCACCACCAECCCCACCACCAAAAKGEGECCCACCAAAAGECKGECCACCAAAA",
    "GECCCACCAAAACCAGECCCACCAECCCCAGECGECCCAAAACCAGECCCACCAAAACCAGECGECAAACCACCAGECGECAAACCACCAGECCCAECCCCACCAGEC",
    "AAACCACCAC/ZhIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_2: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCGGACCAEEAECCEEACCACCACCAECCCCACCAAAAGECAAAEEAGECEEAAAAEEAAAACCAEEAGECECCCCA",
    "CCACCACCAEEAAAAGECECCEEAAAAEEAAAAEEAEEACACCCACCAECCGGACCACCACCAGECBBAEEACCAGECCCACCACACCCAGGACCACCAEEACCAECC",
    "CACCCAEEAGGAECCCCAAAACCAIEECEBAAAEEAEEAAAAAAAIGCAAAGGACCACACAAAGGAAAAIGCCCACCAEEAAAAAAAGECEEAECCCCAAAAAAAIIA",
    "CCACCAECCEEAAAAECCCCACCAGGAEEACACCCAAAAEEACCAGECAAAGGAABCCCAEEAEEACCAEEAECCAAAECCEEAAAAKICCCAEEABBAGECAAAEEA",
    "EEAECCCCAAAAECCGGAACBCCACCAEEABBAEEAGCEEEAEEAEEABBAACBEEAECCCCAEEAECCECCACBCCACCAIGCAAAEEABBACCAEEAGECAAAIGC",
    "EEAAAAAAAEEAAAAIGCCCACCAABCIIAAAAEEACCACCAGECAAAAAAEEAIGCECCEEACCAABCGGACCACCACCAEEACACCCAGECCCAEEAEEAAAACCA",
    "AAAGECCCAEEACACIIABBACCACCAGECCCAEEACCAAAACCAIGCAAAGECCCACCAAAAEEAAAAEEAIGCCCAECCAAACCAEEACCACCAECCCCAAAAGGA",
    "CCAGECKMBEEAAAAECCEEAGECCCAIGCAAAEEAAAACCAECCEEAAAAGGAAAAECCEEAEEAAAAEEAECCCCAAAAGGACACEEAEEACCAAAAGECAAAEEA",
    "CCACCAGECCCAAAAGGAECCCCACCACCAAAAGECEEACCACCAGGACACCCACCACCACCAGECAAAGGACCAECCCCAEEAAAAEEACACCCAGECGGAAAAIGC",
    "CCACCAAAAGECAAAEEACCACCAECCGGAAAAEEACCAECCCCACCAAAAEEAECCEEAEEACCACACEEACCACCACCAGECAAAGGACCACCAECCEEAAAAEEA",
    "CACCCACCAEEACCAIGCAAACCACCAEEACACEEACCACCAEEAC5cnFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_3: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECAAAGECCCACCAECCGECCCACCACCAECCCCAGECCCAAAAGECCCAAAACCACCAKGECCACCACCAAAAECCGEC",
    "CCACCACCAKGECCAAAAAAACCAGECCCACCACCACCAECCGECGECAAACCAGECCCACCAAAACCAGECCCAAAAGECGECECCCCACCACCACCAGECCCAAAA",
    "AAACCAKGECCAGECCCAGECGECBBAEEACCACCAGECCCAGECAAACCAGECCCAAAAGECCCAGECAAAGECCCACCAGECCCAAAACCAAAAGECCCACCAGEC",
    "GECGECAAACCACCAAAAGECCCACCACCAAAAGECKGECCABABAAAGECCCAAAACCACCAGECCCAGECAAACCAECCCCACCAKGECCAGECCCAAAACCAGEC",
    "GECAAACCACCAAAAGECEEACCAGECCCAGECGECBBACCACCAGECCCAEEACCAAAAGECGECAAAAAAEEAGECGECAAACCACCAGECCCACCAECCCCAECC",
    "CCACCACCACCAGECECCCCACCAAAAKGECCACCACCACCAECCCCACCACCAKGEECCCCAGECAAACCAGECCCACCACCAAAAGECAAACCAGECCCAGECCCA",
    "CCAAAAGECGECAAACCAGECCCAGECDBBCCACCACCAGECCCAAAACCAAAAOIGCCACCACCACCAGECAAACCADBBCCAGECGECCCACCAAAAKGECCAGEC",
    "BABAAAIGCCCACCAECCCCAKGECCAAAACCACCAGECCCAECCCCACCAKGECCAAAACCACCAGECCCAAAACCAGECGECAAACCAGECCCAGECAAACCACCA",
    "CCAGECAAAGECCCACCAECCCCACCACCACCAIEECCACCACCAGECECCCCACCACCACCAECCCCACCAGECAAAGECGECCCACCAAAAGECGECCCACCAAAA",
    "GECCCAGECAAACCAGECGECCCAAAACCAGECCCACCAAAAGECGECCCACCAAAACCAGECCCAAAAGECCCAGECGECAAACCACCAGECCCAAAACCACCAGEC",
    "ECCCCACCAC/ZhIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_4: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCGGACCAEEACCACCAECCGECCCACCACCAGECBBAIIAAAACCACCAGGAABCEEAAAAGECCCAGGAAAACCA",
    "ECCCCACCAGGACACCCACCAEEAECCEEAAAACCACCACCACCAIEECCAEEACCACCAEEAEEACACCCAEEACCACCAECCCCAECCACBCCAEEAGECAAAEEA",
    "GECCCACCACCAECCGGAAAACCAECCGGACACEEAAAAECCCCAEEAAAAGGAAAAAAACCAGGAEAEGGAAAAEEAECCEEAAAAEEACCACCACCACCACCAEEA",
    "AAAECCEEAGECAAAGGAEEAAAAAAACCACCAGECCCACCACCAEEACACEEAECCCCACCAEEAAAAGGAECCCCAAAAGGABBAMIECCAEEAAAAGECAAAIIA",
    "AAAECCBBAGGAAAAEEAGECEEAAAACCAAAAEEAEEAECCCCACCAECCEEACCAECCEEACCABBAEEAEEACCAECCCCACCAGGAAAACCAGGAGCEAAACCA",
    "CCAEEACCAECCAAAGGACCACCACCAEEACACEEACCAEEAECCGECAAACCAGECCCACCAEEAAAAECCEEACCAEEAEEACACCCAEEAECCAAAEEACCACCA",
    "GECCCAEEAEEAAAAECCEEACCAAAACCACCACCAGECECCCCAGGAAAAGECBBAEEACCAGGAAAAGGAAAAAAAECCGGAAAAGECCACCCACCAGGAAAAGGA",
    "AAACCADDAEEAAAAGECCCAGECCCAIGCAAAEEAAAACCAECCEEAAAAIIACCAECCAAAEEAAAAEEAECCCCAAAAGGACACEEACCACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCCCACCAEEAAAAGECEEACCACCAEEACACEEACCACCACCAECCAAAGGACCAECCCCACCAAAAEEAGECCCAGECEEAAAAGEC",
    "CCACCACCAECCAAAGGACCACCAGECCCAAAAEEAEEAECCCCACCAAAAEEAECCEEACCAEEACACCCACCACCACCAIGCAAAEEACCACCAECCEEAAAACCA",
    "ECCCCAEEAEEACCAECCCCACCACCAEEACACEEACCACCAEEAGECAAACCACCAECCCCAC1crFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_5: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECAAAKGECCACCAGECCCAAAAAAACCAGECCCAAAACCACCAKGECCAGECCCAAAAGECCCAAAAGECCCAGECCCA",
    "AAACCACCAGECAAACCAGECCCAECCCCACCAGECCCAGECCCABBACCACCAKGECCACCACCAAAAGECGECCCACCAAAAGECAAAIGCCCACCAECCCCACCA",
    "AAACCAGECCCAGECKGECCAGECAAACCAAAACCAGECAAACCAGECCCAGECCCACCACCACCACACGECCCACCAECCCCACCACCACCACCAKGECCACCACCA",
    "GECECCCCAAAACCACCAGECCCACCACCACCAECCGECCCACCAGECECCCCAAAACCAGECGECCCACCAAAACCAGECCCACCAECCCCAGECCCAIEECCACCA",
    "GECAAACCAGECCCAGECAAACCACCACCAGECCCACCAAAACCAGECAAACCACCAAAAKGECCACCAGECCCAGECAAACCACCACCAGECCCACCABBACCAKGE",
    "CCACCACCAAAAGECCCAGECCCACCAKGEAAACCACCACCAECCAAACCACCAKGEGECCCACCACCAAAAGECCCAGECCCAAAAGECCCAAAAGECCCAGECCCA",
    "AAACCAGECGECCCAAAACCACCAGECGECCCACCAAAAECCCCACCACCAAAAGECCCAGECCCACCAGECGECCCAAAACCAGECECCAAACCACCAGECCCACCA",
    "CCACCAGECGECCCAAAACCAKGECCAAAACCACCAGECCCAIEECCACCAGECCCAAAACCACCAGECGECAAACCACCAGECAAACCACCACCAGECECCCCACCA",
    "CCAGECAAAGECCCAGECECCCCACCACCACCAECCCCACCACCACCAECCGECCCACCACCAECCCCACCAGECAAAGECCCACCAGECAAAGECGECCCACCAAAA",
    "KGECCACCAAAACCAGECCCACCAECCCCAGECCCACCAAAACCAGECCCACCAECCCCAKGECCAAAACCACCAGECGECAAACCACCAGECCCAAAACCACCAGEC",
    "AAACCACCAG9Z/HAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_6: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGABBAIGCCCAEEAAAAGGAECCCCAECCCCACCAEEAAAAGECBABECCCCAGGACACGGAAAACCAAAAIGCAAAEEA",
    "CCAGECCCAEEAAAAGECCCACCAAAAEEABBAEEAIGCCCACACGGAAAAEEAAAAECCCCAEEAAAAGGAEEACACAAAGECAAAGGACCACCAAAAKICAAAEEA",
    "AAAECCEEACCAECCCCAGECCCACCAGGAAAAGECAAACCAECCGGAAAAEEABBACCAGGAGECAAAGECEEAAAAAAAEEAAAAIGCEEACCAABCEGBCACEEA",
    "ECCCCAEEACCABBAGGAEEAECCCCACCACACEEACCACCACCAIGCAAAEEAAAACCAEEAGECAAAECCEEAAAACCAGGAABCKICAAACCAECCEEAAAAEEA",
    "AAACCAIGCEEABBAEEACEBECCAAAGECAAAEEAEEACCAAAAGGACACCEBCACCCAEEACCAAAAIEEEGBAAACCACCACACGGACCACCACACGGAAAAEEA",
    "CCAECCEEACCAAAAECCEEACCACCAGGABBAIGCAAACCACCAIGCAAACCAGCECCAGGAEEAAAAEEAEEAECCAAAEEACACGGAAAACCACCAIGCAAAEEA",
    "CCAAAAGECCCAAAAGECGGACCAAAACCAAAAGGAECCCCAECCEEAAAAEEACCACCAGECEEABBACCAEEAECCCCAIGCAAAEEAAAACCACCAGGAAAAGEC",
    "BBAECCMOBEEAAAAGECEEAGECAAAGECAAAEEAEEACCACACEEAAAAGGACCAECCCCACCAAAAGGAECCCCACCAEEACACEEACCACCACCAIGCAAACCA",
    "CCACCAGECEEAAAACCAECCCCAEEAEEAAAAGECCCACCACCAEEACACEEACCACCACCAGECAAAGGACCAECCAAAGGAAAAEEAECCCCAECCGGAAAAGEC",
    "CCACCAEEAECCAAAEEACCACCAGECCCAAAAEEAEEAECCCCAEEAAAACCAECCEEACCAEEACACEEACCACCACCAGECAAAGGAAAACCAECCGGAAAAEEA",
    "CACCCACCAEEACCAGECCCACCACCAEEACACEEACCACCACCAG1cnFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_7: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECCCAECCCCACCAGECCCAAAAGECAAAGECCCACCACCAAAAMIEAAACCAGECAAAGECCCACCACCAAAAGECGEC",
    "BABAAAGECGECCCAAAAGECCCAGECCCACCACCAAAAGECECCCCACCACCAGECAAAGECCCACCAGECCCAAAAGECAAAGECCCACCACCACCAECCCCACCA",
    "GECAAAGECGECGECCCAAAAGECCCACCACCAECCGECCCAAAACCACCAGECCCACCAGECAAAGECECCCCACCACCAGECCCAAAACCACCAIEEEEAAAACCA",
    "AAAGECGECCCACCACCAGECAAACCAGECAAAGECCCACCACCAAAAGECGECCCACCACCAIEECCAAAACCACCAGECCCAAAAGECCCAGECECCCCACCAGEC",
    "GECCCAAAACCACCAGECEEAAAAGECAAAGECCCAGECCCABABKGEAAAEEAECCCCAGECCCACCABBAEEAGECCCACCAAAAGECGECCCAAAAGECCCAGEC",
    "CCAAAACCACCAGECECCCCACCACCAGECCCAAAACCACCAKGEAAACCAGECECCGECCCACCACCACCAKGEAAACCACCAAAAGECGECCCACCAAAAGECCCA",
    "CCACCAAAAKGECCAAAACCACCAGECBABCCAGECAAAGECECCCCACCACCAGECCCAAAACCACCAGECGECAAABABECCGECCCACCACCACCAGECCCAIEE",
    "BABAAAIGCCCACCAAAAGECKGECCAAAACCACCAKGECCAAAACCACCAGECCCAAAAGECCCAGECCCAAAACCAGECGECAAACCACCACCAGECAAACCAGEC",
    "CCAGECAAACCACCAGECECCCCACCAGECCCAECCCCACCACCACCAECCCCACCACCAGECECCCCACCAGECAAAGECCCACCACCAAAAGECKGECCACCAAAA",
    "GECCCACCAECCCCAGECCCACCAECCCCAGECCCACCAAAACCAGECCCACCAECCCCAGECGECAAACCACCAGECCCAAAACCACCAKGECCAAAACCACCAGEC",
    "AAACCAGECC/ZhIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_8: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGACACEEAECCCCACCAEEACACEEACCACCAECCEEAAAAGECCCACCACCAGECAAAEEAECCCCACCAEEACACEEA",
    "CCACCAECCEEAAAAGECCCACCACCAGECAAAEEACCAECCCCAGECCACEEACCACCAECCEEAAAAEEAECCCCACCAGECAAAEEACCAECCCCAEEACACEEA",
    "CCACCAECCEEAAAAGGAECCCCACCAGECAAAEEACCAECCCCAEEAAAAGECCCACCAECCEEAAAAEEAECCCCACCAEEACACEEACCAECCCCAEEAAAAGEC",
    "CCACCAECCEEAAAAEEAECCCCACCAEEACACEEACCAECCCCAEEAAAAGECCCACCACCAGECAAAEEAECCCCACCAEEACACEEACCACCAECCEEAAAAGEC",
    "CCACCACCAGECAAAEEAECCCCACCAGECCACEEACCACCAECCEEAAAAGECCCACCACCAGECAAAEEACCAECCCCAEEACACEEACCACCAECCEEAAAAEEA",
    "ECCCCACCAGECAAAEEACCAECCCCAEEACACEEACCACCAECCEEAAAAEEAECCCCACCAGECAAAEEACCAECCCCAEEAAAAGECCCACCAECCEEAAAAEEA",
    "ECCCCACCAEEACACEEACCAECCCCAEEAAAAGECCCACCAGECEEAAAAEEAECCCCACCAEEACACEEACCAECCCCAEEAAAAGECCCACCACCAGECAAAEEA",
    "ECCCCACCAGECCACEEACCACCAECCEEAAAAGECCCACCACCAGECAAAEEACCAECCCCAEEACACEEACCACCAECCEEAAAAGECCCACCACCAGECAAAEEA",
    "CCAECCCCAEEACACEEACCACCAECCEEAAAAEEAECCCCACCAGECAAAEEACCAECCCCAEEAAAAGECCCACCAECCEEAAAAEEAECCCCACCAGECAAAEEA",
    "CCAECCCCAEEAAAAGECCCACCAECCEEAAAAEEAECCCCACCAEEACACEEACCAECCCCAEEAAAAGECCCACCACCAGECAAAEEAECCCCACCAEEAEAEEEA",
    "CCAECCCCAEEAAAAGECCCACCACCAGECAAAEEAECCCrb1GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_9: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCCCAAAACCAKGECCAAAACCAGECGECAAACCACCAAAAKGECCACCAAAACCAGECECCCCACCACCAECCGEC",
    "CCACCAAAAGECCCAECCCCACCAGECECCCCACCACCAECCCCAGECAAACCAGECCCAAAAGECCCAGECAAACCAGECAAAGECCCACCAAAAGECGECAAACCA",
    "CCACCAIEECCACCACCAECCGECCCAAAACCACCAKGEAAACCACCACCAECCGECCCAAAACCAGECGECAAACCACCAGECAAAGECCCAAAAGECGECCCAAAA",
    "CCAGECAAAGECCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCAAAACCAKGECCAAAACCAGECGECAAACCACCAAAA",
    "KGECCACCAAAACCAGECECCCCACCACCAECCGECCCACCAAAAGECCCAECCCCACCAGECECCCCACCACCAECCCCAGECAAACCAGECCCAAAAGECCCAGEC",
    "AAACCAGECAAAGECCCACCAAAAGECGECAAACCACCACCAIEECCACCACCAECCGECCCAAAACCACCAKGEAAACCACCACCAECCGECCCAAAACCAGECGEC",
    "AAACCACCAGECAAAGECCCAAAAGECGECCCAAAACCAGECAAAGECCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCA",
    "AAACCAKGECCAAAACCAGECGECAAACCACCAAAAKGECCACCAAAACCAGECGECAAACCACCAECCGECCCACCAAAAGECCCAECCCCACCAGECECCCCACCA",
    "CCAECCCCAGECAAACCAGECCCAAAAGECCCAGECAAACCAGECAAAGECCCACCAAAAGECGECCCAAAACCACCAIEECCACCACCAECCGECCCAAAACCACCA",
    "KGEAAACCACCACCAECCGECCCAAAACCAGECGECAAACCACCAGECAAAGECCCAAAAGECGECCCAAAACCAGECCCAECCCCACCAECCCCACCAGECAAAGEC",
    "CCAAAAGECCCAGnZ1IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_10: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAECCEEACCAECCCCACCACACEEACCACCAECCEEAAAAEEAECCCCACCAGECAAAEEACCAGECCCAEEAAAAGEC",
    "CCACCACCAGECAAAEEAECCCCACCAEEACACCCACCACCAECCEEACCAGECCCACCACCAGECAAAEEACCAECCCCAEEAAAAGECCCACCAECCEEAAAAEEA",
    "ECCEEACCAEEACACEEAAAACCAECCEEAAAAGECCCACCACCAGECAAAEEACCAECCCCAEEAECCEEACCACCAECCEEAAAAEEAECCCCACCAEEACACEEA",
    "CCAECCAAAEEAAAAGECEEACCACCAGECAAAEEACCAECCCCAEEACACEEACCACCAECCEEAAAAEEAECCCCACCAIGCAAAEEACCAECCCCACCAAAAGEC",
    "CCACCACCAGECAAAEEAECCCCACCAEEACACEEAEEACCAECCEEAAAAGECCCACCACCAGECAAAEEACCAECCCCACCAAAAGECCCACCAECCGGAAAAEEA",
    "ECCCCACCAEEACACEEACCACCAECCEEAAAAGECCCACCACCAGECAAAEEAEEAECCCCAEEACACCCACCACCAECCEEAAAAEEAECCCCACCAEEACACEEA",
    "CCAECCCCAGGAAAAGECCCACCACCAGECAAAEEACCAECCCCAEEACACEEAAAACCAECCEEAAAAEEAGECCCACCAGECAAAEEACCAECCCCAEEAAAAGEC",
    "CCACCACCAGECAAAEEAECCCCACCAGGACACEEACCACCACACEEAAAAGECCCACCACCAGECAAAEEACCAECCCCAEEAAAAGECEEACCAECCEEAAAAEEA",
    "ECCCCACCAEEACACEEACCACCAECCCCAAAAGECCCACCACCAIGCAAAEEACCAECCCCAEEACACEEACCACCAECCEEAAAAEEAECCCCACCAEEACACGGA",
    "CCAECCCCAEEAAAAECCCCACCACCAGECAAAEEAECCCCACCAEEACACEEACCACCAGECEEAAAAEEAECCCCACCAGECAAAEEACCAECCCCAEEAAAAGEC",
    "AAACCACCAGECAAAGGAECCCCACCAEEACACEEACCACCAECCC3bpGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_11: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCCCACCAAAAKGECCAAAACCACCAGECECCCCACCACCAIEECCACCAAAACCAGECGECAAACCACCAGECAAA",
    "GECCCACCAECCCCAGECAAACCAGECCCAAAAGECCCAGECAAACCACCAECCGECCCACCAAAACCAKGECCAAAACCAGECGECAAACCACCAAAAKGECCACCA",
    "AAACCAGECGECAAACCACCAGECECCCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAECCCCACCACCAECCGECCCACCAECCCCAGECCCAAAACCA",
    "GECECCCCACCACCAAAAKGECCACCAAAACCAGECGECAAACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECAAACCAGECCCAECCCCAGECCCAAAA",
    "GECCCACCAECCCCAGECAAACCACCAGECECCCCACCACCAAAAKGECCAAAACCACCAKGEAAACCACCACCAECCGECCCACCAAAAGECCCAECCCCACCAGEC",
    "AAACCAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCAAAAGECGECCCAAAACCACCAKGEAAACCACCACCAECCGEC",
    "CCAAAACCAGECCCAECCCCACCAGECECCCCACCACCAECCCCAGECAAACCAGECCCAAAAGECCCAGECAAACCAGECAAAGECCCACCAAAAGECGECCCAAAA",
    "CCACCAKGEAAACCACCAAAAGECGECCCAAAACCAKGECCAAAACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAECCCCACCAGEC",
    "AAAGECCCACCAAAAGECGECCCAAAACCACCAIEECCACCACCAECCGECCCACCAAAACCAKGEAAACCACCACCAECCGECCCACCAAAAGECCCAGECAAACCA",
    "GECAAAGECCCACCAECCCCACCAGECAAAGECCCAAAACCAGECGECAAACCACCACCAIEECCACCACCAECCGECCCAAAACCACCAKGEAAACCACCACCAECC",
    "GECCCACCAAAAGECErZzIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_12: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGACCAGECCCACCACCAECCAAAEEACCACCAECCEEAAAAEEAECCCCACCAEEAAAAGECEEACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCCCACCACCAAAAGECCCACCAEEAGECAAAEEACCACCAECCEEAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECCCAEEA",
    "CCACCAECCCCAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECAAAEEAEEACCAECCEEAAAAEEAECCCCACCAEEAAAAGECCCACCACCAECCAAAEEA",
    "CCACCAGECEEAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECAAAEEACCACCAECCEEACCAEEAECCCCACCACCAAAAGECCCACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAGECCCACCAEEAAAAGECCCACCACCAGECAAAEEACCACCAECCCCAAAAEEAECCCCAEEAEEAAAAGECCCACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCCCACCAEEACCAGECCCACCACCAECCAAAEEACCACCAECCEEAAAAEEAECCCCACCAEEAAAAGECEEACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCCCACCACCAAAAGECCCACCAEEAGECAAAEEACCACCAECCEEAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECCCAEEA",
    "CCACCAECCCCAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECAAAEEACCAEEAECCEEAAAAEEAECCCCACCAEEAAAAGECCCACCACCAECCAAAEEA",
    "CCACCAECCGGAAAAEEAECCCCACCAEEAAAAGECCCACCACCAGECAAAEEACCACCAECCEEAAAAGGAECCCCACCACCAAAAGECCCACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCEEACCAEEAAAAGECCCACCACCAGECAAAEEACCACCAECCCCAAAAEEAECCCCACCAGGAAAAGECCCACCACCAGECAAAEEA",
    "CCACCAECCEEAAAAEEAECCCCACCAEEAAAAIGCCCACCACCAECCAAAEEACCACCAECCEEAAAAEEAECCCCACCAEzcrFAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_13: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECGECAAACCACCAGECAAACCAGECCCAGECAAACCACCAGECECCCCACCACCACCAECCGECCCACCAAAAGECCCA",
    "GECCCAAAAGECCCACCAAAACCAKGECCACCAAAACCAGECGECAAACCACCAGECCCAAAAGECCCAGECAAACCACCAGECGECAAACCACCACCAECCCCAGEC",
    "CCACCAECCCCACCAGECAAAGECCCACCACCAAAAKGECCACCAAAACCAGECGECCCAAAACCAGECCCAAAACCAGECGECCCAAAACCACCAKGEAAACCACCA",
    "CCAGECAAAGECCCACCAECCCCACCAGECCCAECCCCACCACCAAAAGECGECCCACCAAAAGECCCAGECAAACCAGECCCACCAAAAGECGECCCAAAACCACCA",
    "KGECCAAAACCACCAGECAAACCAGECCCAGECAAACCACCAGECECCCCACCACCACCAECCGECCCACCAAAAGECCCAGECCCAAAAGECCCACCAAAACCAKGE",
    "CCACCAAAACCAGECGECAAACCACCAGECCCAAAAGECCCAGECAAACCACCAGECGECAAACCACCACCAECCCCAGECCCACCAECCCCACCAGECAAAGECCCA",
    "CCACCAAAAKGECCACCAAAACCAGECGECCCAAAACCAGECCCAAAACCAGECGECCCAAAACCACCAKGEAAACCACCACCAGECAAAGECCCACCAECCCCACCA",
    "GECCCAECCCCACCACCAAAAGECGECCCACCAAAAGECCCAGECAAACCAGECCCACCAAAAGECGECCCAAAACCACCAKGECCAAAACCACCAGECAAACCAGEC",
    "CCAGECAAACCACCAGECECCCCACCACCACCAECCGECCCACCAAAAGECCCAGECCCAAAAGECCCACCAAAACCAKGECCACCAAAACCAGECGECAAACCACCA",
    "GECCCAAAAGECCCAGECAAACCACCAGECGECAAACCACCACCAECCCCAGECCCACCAECCCCACCAGECAAAGECCCACCACCAAAAKGECCACCAAAACCAGEC",
    "GECCCAAAACCAGECCCAAAACCAGECE/Z/HAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_14: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCCCAKICCCACCAGGACACGGAAAACCAGECCCAAAAGECAAAEEAAAAIGCAAAEEACCACACEEAEEACACGGA",
    "GGACCACACGGABBAGECCCACCAEEAECCAAAIIAAAAECCBBAEEAECCCCAAAAEEAECCGGAAAAAAAIGCCCAAAAIGCAAAEEACCAECCCCAEEAAAAGEC",
    "CCACCAGECIIACCACCACACEEACCAEEAECCCCAEEAABCBBAGGAAAAIGCCCACCAEEAGECAAACCAECCCCAEEACCACACGGACCAEEACACEEACCAECC",
    "CCAEEACCAGECAAACCAECCCCAEEAGGACACAAACCACCAECCCCAAAAIGCAAACCAIIAIGCCCAAAACCAGECCCACCAECCEEACCACCAECCEEAAAAEEA",
    "GECCCAEEAECCAAAGGAAAAECCCCAAAACCAECCEEAEEAECCAAACCAEEAECCEEAAAAGGACACCCACCAECCEEAEEAAAAIGCCCACCACCAECCAAAIIA",
    "ECCEEACCAEEAECCEEAAAACACAAAEEACCAGECCCACCACCAECCAAAGGAGECCCACCAEEACACGGABBACCAGECCCACCAGECCCAEEACCAECCCCAEEA",
    "CCAECCCCAEEACACAAACCACCACACIIAAAAGGACACCCAIIAECCBBAGGACCAGECAAAEEACCAIGCAAACCAECCGGAAAACCAGECCCADDAGGACACGGA",
    "CCAECCANOEEAAAACCAECCEEAAAAGECAAAEEACACEEAEEAEEACACCCACCAGECAAAEEACACEEAAAAEEAECCGGAAAAIGCCCAEEAECCCCAAAAEEA",
    "CACEEACCAGECAAAEEAECCEEACCACCACACEEAAAAGECCCAEEACACCCAEEAEEAECCCCAAAAGECAAAEEAECCEEAAAACCAECCEEAEEAECCAAAEEA",
    "CACEEACCAEEACACCCACCAGECGGAEEACACEEACCAEEACACEEAAAAECCCCAEEAECCCCAAAAGECEEAEEAAAAGECAAACCAECCEEACCAECCAAAEEA",
    "CCAGECCCACnb5GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_15: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCCCAECCCCAGECCCAAAACCAGECGECAAACCACCAAAAKGECCACCAAAACCAGECGECAAACCACCAECCGEC",
    "CCACCAAAAGECCCAECCCCACCAGECAAACCAGECCCAECCCCAGECAAACCAGECCCAAAAGECCCAGECAAACCAGECAAAGECCCACCAAAAGECGECCCAAAA",
    "CCACCAIEECCACCACCAECCGECCCACCAAAACCAKGEAAACCACCACCAECCGECCCACCAAAAGECGECAAACCACCAGECAAAGECCCACCAECCCCACCAECC",
    "CCAGECCCAECCCCACCAECCCCACCAGECAAAGECCCACCAAAAGECGECAAACCACCAGECECCCCACCACCAAAAKGECCAAAACCACCAGECECCCCACCACCA",
    "IEECCACCAAAACCAGECGECAAACCACCAGECAAAGECCCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCACCAECCCCAGEC",
    "AAACCACCAGECECCCCACCAAAAGECGECCCAAAACCACCAKGEAAACCACCAAAAGECGECCCAAAACCAKGECCAAAACCACCAECCGECCCACCAAAAGECCCA",
    "GECAAACCAGECAAAGECCCACCAECCCCACCAGECAAAGECCCAAAACCAGECGECAAACCAGECCCAECCCCACCAAAAGECGECCCAAAACCACCAKGEAAACCA",
    "CCAAAADHECCACCACCAAAAGECGECAAACCACCAECCCCAGECCCAAAAGECGECAAACCACCAECCCCAGECCCAAAAGECCCAECCCCACCAECCCCACCAGEC",
    "AAAGECCCAAAACCAGECECCCCACCAGECAAAGECCCAAAACCAGECECCCCACCACCAAAAKGECCAAAACCACCAIEECCACCACCAAAAGECGECAAACCACCA",
    "ECCCCAGECCCAAAAGECGECAAACCACCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAECCCCAGECCCAAAACCAGECECCCCACCAAAACCAKGE",
    "CCAAAACCAGECECCCCACCAAAACCAKGECCAAAAClZ7IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_16: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCIGCEEAEEAGECEEABBAEEAECCCCACCAECCBBAGGAGECCCACCAEEACACEEACCAECCCCAIIAABCCCA",
    "EEAAAACACGGAAAAECCEEACCAECCEEAAAAEEAGECCCAAAAIGCAAAGGACACCCAGGAGGAABCEEACCAECCAAACCACACGGACCACCAECCEEAAAAGEC",
    "AAACCAGECEEAAAAGGAECCCCACCAGECCCACCAABCEEACCACCAECCCCACCAGECCCAGGACACEEACCAEEACACEEACCAGECAAAEEACACEEACCAEEA",
    "ECCEEACCAECCBBAEEAECCCCAAAAGGACACEEACCAECCCCAEEAABCGGAGGACCACACGGAAAAKICCCACCACACGGAAAABBAGECCCACCAIGCAAAEEA",
    "ECCCCACCAEEACACGGAAAAECCEEAEEAABCEEAEEACCAGECCCAAAAIGCBBACCACACEEAAAAGGAGECCCAEEAGECAAAGGACACCCACCAEEAECCEEA",
    "AAAGECCCAEEACACCCAEEAAAAECCCCACCAGECCCAEEAABCEEACCAEEACACEEACCAIGCCCACCACACEEAGGACCAECCEEAEEAECCDDAGGAABCEEA",
    "EEACCAECCEEAAAAGECEEACCAABCGGAAAACCACACCCACCAKICAAACCAECCCCACCACCACACEEAEEAECCCCAIIAABCGGACCACCAGECEEABBAGEC",
    "CCACCAKICCCAAAAEEAECCEEAAAAIGCAAAEEACACEEACCACCACACEEACCAGECAAAEEACACEEACCAEEAECCEEAAAAIGCCCAEEAECCCCAAAAEEA",
    "ECCEEAAAAIGCAAAEEACACEEACCAEEACACCCACCAGECCCACCACACEEAEEAEEACACEEAAAAGECAAAEEAECCCCAAAAEEAIGCEEAAAAIGCAAAEEA",
    "CACEEACCAEEACACCCACCAGECCCACCACACEEAEEAEEACACEEAAAAIGCCCAEEAECCEEAAAAECCCCAEEACCAGECAAAEEAECCEpb1GAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_17: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCECCCCACCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCA",
    "GECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAA",
    "GECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGEC",
    "AAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCA",
    "ECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGEC",
    "CCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCA",
    "GECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAA",
    "GECCCAECCCCACCAGECAAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGEC",
    "AAAGECCCAECCCCACCAECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAGECAAAGECCCAECCCCACCA",
    "ECCCCACCAGECAAAGECCCAAAAGECCCAECCCCAGECCCAAAAGECCCAAAAGECCCAECCCCACCAECCCCAGECCCAECCCCACCAECCCCACCAECCCCAGEC",
    "CCAAAAGECC/YhJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_18: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGACBECCAGGAEEAECCGGAAAAGGACCAECCAAACCACCAGECAAAEEAAAAIGCBBAEEAEEACCAGECCCAAAAGGA",
    "GECABCCCAEEACACCCAEEAEEAAAAGECAAAEEACCAECCEEAEEABBAGECCCACCAIIAECCCCAEEAAAAEEACBEEEACCACCAGECEEABBAGGACACEEA",
    "CCACCACCAGECBBAEEAGGAECCCCAEEAAAAIGCBBACCAEEACBEAAAGGAEEAEEACACGGACCAEEACACEEACCAGGACACAAAGGACCAAAAIGCAAAEEA",
    "CCACACEEAIIACACGECEEACCAEEADDACACGGAAAAEEAECCEEACCAAAAGECEEACCACCAECCIIAEEAAAAAAAIGCAAAECCEEAECCEEACCAAAAIGC",
    "CCACCACCAEEACACCCABBAEEAGECCCACCAEEAECCCCAAAAEEAECCGGAECCCCAEEAGECAAAAAAEEACACAAAGGAAAAIGCAAACCAKKACCACACEEA",
    "CCACCAECCEEABBAEEAABCGECEEAEEACCAECCCCAEEACCAECCAAAEEACCAECCEEAGGAAAACACEEACCAEEACCACACGGACCACCAGCEEEACCACCA",
    "ECCAAACCACCACCAECCCCAEEACCAMKCAAACCAEEACACCCAEEAAAAGECEEACCAEEAGECABCGGAAAACCAECCCCAAAAGGAECCCCAAAAEEACCACAC",
    "AAAEEAKFQEEAAAACCAECCEEAAAAGECAAAEEACACEEAEEAEEACACCCACCAGECCCACCACACEEAEEAEEAECCGGAAAAECCCCAEEAECCCCAAAAEEA",
    "ECCEEAAAAGECAAAEEAECCEEACCACCACACEEACCAGECAAAEEACACGGAAAAEEAECCCCAAAAGECCCAEEACACEEAAAAEEAECCEEACCAECCAAAEEA",
    "ECCEEAEEAEEACACEEAAAAGECEEAEEACACCCACCAEEACACEEAAAAGECAAAEEAGECEEAAAAECCCCAEEACCAECCAAAEEACACEEACCAIGCAAACCA",
    "CCAEnb3GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_19: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECCACCCAGECCCAECCGECCCACCACCAECCCCAGECCCAAAAGECCCACCAAAACCAKGECCACCAAAACCAGECGEC",
    "AAAAAACCAGECAAAGECCCACCAGECAAACCACCAGECECCCCACCACCAAAAKGECCAGECAAACCAGECGECCCABBACCAGECCCAAAACCAGECGECAAACCA",
    "CCACCAKGEAAACCACCACCAECCGECCCACCAAAAGECCCAGECBBACCAGECCCACCAAAAKGEGECCCAAAACCACCAKGEAAACCACCACCAGECECCCCACCA",
    "CCAECCCCACCAAAAECCGECCCACCACCAAAAKGECCACCAAAACCAGECGECAAACCACCAGECAAAKGECCACCAGECAAACCAAAAGECECCCCACCACCAAAA",
    "KGECCACCACCAAAAGECGECCCAAAACCAGECCCAAAACCAGECGECAAACCAAAACCAKGEAAACCACCAGECECCGECCCACCAAAAGECCCACCAGECAAAGEC",
    "CCACCAAAAGECGECCCAAAAAAACCAKGECCAAAACCACCAGECECCCCACCACCAECCCCACCACCAECCGECCCAGECCCAAAAKGECCACCABBACCAGECGEC",
    "AAACCACCAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCACCAAAAKGECCACCAAAAAAAGECGECCCAAAAGECGECCCAAAACCAGECGECCCAAAA",
    "CCACCAFJECCACCACCAAAAGECGECAAACCACCAECCCCACCAGECAAAGECCCAAAACCAGECECCCCACCACCAAAAKGECCAECCCCACCAECCCCACCAGEC",
    "AAAGECCCAAAACCAGECECCCCACCACCAAAAKGECCAAAACCACCAECCGECCCACCAAAAGECCCAAAAGECCCAECCCCACCACCAIEEGECCCAAAACCACCA",
    "IEECCACCACCAAAAGECCCAAAAGECCCAECCCCACCACCAECCGECCCAAAACCACCAIEECCACCAAAACCAGECGECAAACCACCAECCCCAGECECCCCAGEC",
    "CCAAAACCAGECECCCCACCAAAACCAKGECCAAAAClZ7IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_20: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCIGCEEACCAGECEEAAAAGGACACCCAEEAECCBBAEEAECCEEAAAAGGACACEEACCACACCCAIIACACEEA",
    "EEABBACACGGAAAACACEEACCAECCEEAAAAGGAGECCCAAAAGECAAAGGACACCCAIIAEEACACCCACCAECCGGABBACACGGACCACCACACEEAAAAIGC",
    "AAACCAGECEEAAAAEEAECCCCACCAGECCCAEEACACEEAAAAEEAABCEEACCAGECCCACCAECCGGACCAEEACACEEACCAGECCCAEEAECCCCACCACCA",
    "ECCEEACCAECCCCAAAAECCCCACCAGGACACEEACCAECCCCACCACACEEAEEACCACACIIAAAAGECEEACCAECCGGAAAACCACACCCACCAECCAAAGGA",
    "GECCCACCAEEACACGGABBAECCEEAEEACACEEACCACCAIGCAAAAAAIGCCCACCAABCEEAAAAGGACACAAAIIAGECAAAGGAECCCCACCAEEACACEEA",
    "AAAGECCCAEEAECCEEACCACCAECCBBACCAGECCCAEEACACEEAAAAEEAECCEEACCAGECCCACCAECCCCACCAGGAECCEEAEEAECCCCAGGADFCEEA",
    "CCACCAECCEEAAAAGECEEAAAACACIIAAAAEEAECCCCACCAIGCAAAAAAGECCCACCAEEACACCCAEEAECCAAAEEACACGGAEEACCAGECEEACCAGEC",
    "CCACCAKICCCAAAAEEAECCEEAAAAGECAAAEEACACEEACCAEEACACEEACCAGECAAAEEACACEEAAAAEEAECCEEAAAAGECCCAEEAIGCCCAAAAEEA",
    "ECCEEAAAAGECAAAEEAECCEEACCAEEACACCCACCAGECCCACCACACEEACCAEEAECCEEAAAAGECAAAEEAECCCCAAAAEEAECCEEAAAAMKCAAAEEA",
    "CACEEACCAEEACACCCACCAGECCCACCACACGGACCAEEACACEEAAAAGECAAAEEAECCEEAAAAECCCCAEEAEEAECCAAAEEAECCEpb1GAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_21: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCGGAGGACCAECCCCACCAEEAECCCCACCAAAAECCGGABBAEEAECCAAACCAEEACCAGECEEAAAAAAAIGC",
    "EEACCACCAECCAAAGGAAAAECCCCAEEACCAECCCCACCAGGAECCAAAEEACCAECCCCAAAAGECGGAGECCCACCACCAECCEEACCACCAGECAAAAAAGGA",
    "ECCCCACCACCACACGGACCAEEACACCCACCAEEAECCCCAGGAAAAAAAGECCCACCACCAGECAAAGGAAAAGECCCACCAAAAIGCAAACCAGGAIEEAAACCA",
    "EEAECCCCAEEAAAAGECCCAEEACCAEEAABCEEACCACCAGECEEABBAEEAGECCCACCACCACACGGACCACCAGECCCABBAGGAECCCCACCAGGAABCEEA",
    "CCAEEAAAAGECAAAGGABCDECCCCAIGCBBAGECCCACCAEEAIGCBBAEEAAAAGECCCAEEABBAIGCCCACCAACBGGAABCCCAEEACCAECCGGAAAACCA",
    "ECCCCAEEAEEAABCEEAEEACCAECCEEABBAEEAECCEEACCACCAECCEEAAAACCAIGCECCAAAEEAEEAECCAAAGGAAAAGECAAAEEAAAAIGCAAAGGA",
    "AAAECCCCAEEAAAAGECCCACCACCAGECCCAEEABBAEEAECCGGAAAAGGAABCCCACCAGGAAAACCAEEACCAECCEEACCACCACACGECEEAGGACACEEA",
    "AAACCACGDEEACACCCAEEACCACCAGECAAAEEACCAECCCCAEEAAAAEEAGECCCACCACCAAAAGECEEACCAAAAGECCCAEEACCACCACACGGAAAAEEA",
    "ECCEEAAAAEEAAAAIGCECCCCACCAGGAECCEEAAAACCAECCEEAAAAEEACCAGECCCAEEAAAACCAECCCCACCAGGAECCCCACCACCACCAGECAAACCA",
    "EEAECCEEACCAAAAEEAECCCCACCAEEAAAAIGCCCACCACCAECCAAAIGCEEACCAECCEEAAAAEEACCAECCAAAGGAAAAIGCCCACCAAAAEEACACEEA",
    "CCACCAGECEEAAAAEEAAAAECCCCAGGAAAAGGACzctFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_22: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCCCACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAGECCCAAAACCAGECGECAAA",
    "CCACCACCAIEECCACCACCAAAAGECCCACCAECCCCAGECCCAAAACCAGECGECAAACCACCAGECKGEAAACCACCACCAECCGECCCACCAAAAGECCCAGEC",
    "AAACCAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAGECECCCCACCA",
    "GECECCCCACCACCAAAAKGECCACCACCAAAAGECGECCCAAAACCAGECCCAECCCCACCAGECAAACCACCACCAIEECCACCACCAAAAGECGECCCAAAACCA",
    "GECCCAGECAAACCAGECGGAAAAGECGECGECAAACCACCAGECECCCCACCACCAAAAKGECCACCAAAACCAGECIGCCCAAAACCAGECCCAAAAGECCCAGEC",
    "AAACCACCAGECECCCCACCACCAAAAKGECCACCAAAACCAGECGECAAACCACCAGECGECECCCCACCAGECAAACCAGECCCAECCCCACCACCAECCGECCCA",
    "CCAAAACCAGECCCAECCCCACCAGECHDDCCAGECCCAGECAAACCACCAGECECCCCACCACCACCAKGECCACCAAAACCAGECGECAAAGECCCAGECAAAGEC",
    "CCACCAKICBABAAACCAGECGECCCAAAACCACCAKGEAAACCACCACCAGECECCCCACCACCAGECAAAGECCCACCAECCCCACCAGECCCAECCCCACCACCA",
    "ECCGECCCACCACCAAAAOIGCCACCACCAAAAGECCCACCAECCCCAGECCCACCAAAAGECGECCCACCAAAACCAKGECCAAAACCACCAGECGECAAACCACCA",
    "GECAAAGECCCACCAGECAAACCAGECCCAGECAAACCACCAGECECCCCAGECCCACCAECCGECCCACCACCAECCCCAGECCCAAAAGECCCACCAC5ZnIAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_23: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGABBAKICEEAAAACCAGECCCAEEACACCCAEEACCACCAEEADDAGECCCAEEACCAECCCCAEEACCACCACACEEA",
    "GGACCAECCEEABBAEEAAAAECCEEAEEACCAECCCCACCACCAGECAAAEEAEEACCAAAAECCGECEEAEEAEEACACEEACCAEEAECCCCACCACCACACGGA",
    "AAACCACCAGGACACEEACACCEBECCCCACCACCAEEAECCEEACCAAAACCAGECCCAAAAIIAABCEEAAAAEEACCAECCAAAGGACCACCAGECGECAAACCA",
    "IGCCCAAAAGECACBECCCCAEEACACEGBCACEEACCAAAAECCEEAAAAGGACCAECCCCACCACCAEEACACCCAGGACCACACEEACCACCACACIIAAAACCA",
    "ECCEEACACCEBAAAIGCCCACCACCAIEEBABGECCCACCAEEAGECAAAEEAAAAEEACACGGAAAAGGAAAAECCAAAGGACACCCAGGACACACBGECCCACCA",
    "AAACCAIGCEEAAAAEEACACCCACCAGGAAAACCAECCEEACCACCAECCCCAEEACCAIGCECCBBAEEAEEAECCCCAEEACACCEBECCEEAAAACCAAAAKIC",
    "CCACCAAAAGECAAAEEAEEACCACACEEACCAEEAAAAECCEEAEEAAAAIGCAAAAAACCAGGACACEEACCAECCACBGECCCAAAAEEAIEEEEAEEAAAACCA",
    "ECCCCAOOAGGACACCCACCACCACCAGECAAAGGACCAECCCCACCAAAAEEAIGCCCACCACCAAAAGECCCACCAAAAIGCCCAEEACCACCACACEEAAAAGGA",
    "ECCEEAAAAEEAAAAGECECCCCAEEAEEAECCEEAAAACCAECCGGAAAAEEAAAAGECCCAEEAAAAEEAECCCCACCAEEAECCCCACCACCAEEAGECAAACCA",
    "CCAECCEEAEEAAAAEEAECCCCACCACCAAAAIGCEEACCACCAECCAAAIGCCCACCAGECEEAAAAEEACCAECCAAAEEAAAAKICCCACCAAAAEEACACEEA",
    "CCACCAGECEEAAAAEEAAAAECCEEAExctFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_24: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECGECAAACCACCAGECAAACCAGECAAAGECCCACCACCAGECGECAAACCAGECCCAECCCCACCACCACCAECCCCA",
    "CCACCAECCGECCCACCACCAAAAKGECCACCAAAACCAGECCCAECCCCACCAGECCCACCAECCGECGECCCACCAAAACCAGECGECAAACCACCAGECAAAGEC",
    "CCACCAGECCCAAAAGECAAAIGCAAACCACCACCAGECAAAGECCCACCAGECAAACCACCAGECECCCCACCACCACCAIEECCACCACCACCAECCKGECCACCA",
    "AAAGECCCAAAAEEAAAAKGECCAAAAIGCAAACCAGECCCAAAACCAGECCCACCAAAAGECGECCCACCAAAACCAGECGECAAACCACCAGECAAAGECCCACCA",
    "ECCCCAAAAEEACCAIEECCACCACCAECCIGCAAAGECCCACCAECCCCACCAGECCCAECCCCACCAGECBABECCGECCCAAAACCAKGEAAAACBECCCCAGEC",
    "CCACCAAAAGECGECCCAAAACCACCAGECCCAGECAAACCAGECCCAAAACCAGECGECGECAAACCACCAKGEAAACCACCAAAAIGCECCCCACCACCAGECAAA",
    "GECCCACCAIEECCABABCCACCAIEECCACCACCACCAECCGECBABCCAECCGECCCACCACCAAAAKGECCAAAAEEAAAAGECCCAGECECCCCAGECCCACCA",
    "AAACCAKGECCAAAACCAGECGECCCAAAACCAGECCCAAAAGECCCACCAGECAAACCACCAGECGECAAACCACCACCAIEECCACCACCACCAECCCCACCACCA",
    "ECCGECCCACCACCAAAAOIGCCACCACCAAAAGECCCACCAECCCCAGECCCACCAAAAGECGECCCACCAECCCCACCAGECAAACCACCAGECCCAAAACCAGEC",
    "GECAAACCACCACCAKGEAAACCACCACCAGECAAAGECCCACCAECCCCAGECCCACCAIEECCACCACCACCAECCGECCCACCAAAAGECCCACCAGECAAAGEC",
    "C7ZlIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_25: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCEEAEEACCAIGCCCACCAEEAECCCCACCAAAAECCGGABBAEEAECCAAAAAAEEACCAGECEEAAAAAAAIGC",
    "EEACCACCAECCAAAIIACCAECCCCAEEACCAECCEEAAAAEEAGECAAAEEACCAECCCCACCAECCGGAGECCCAAAACCAECCEEACCACCAGECAAAAAAGGA",
    "CACCCAEEAEEACACEEACCAEEACACCCACCAEEAECCCCAEEACCAAAAGECCCACCACCAGECAAAGGAAAAGECCCACCAAAAIGCAAACCAGGAIEEAAACCA",
    "EEAECCCCAEEAAAAGECCCAEEACCAEEAABCEEACCACCAGECEEABBAEEAGECCCACCACCACACGGACCACCAGECCCABBAGGAECCCCACCAGGAABCEEA",
    "CCAEEAAAAGECAAAGGABCDECCCCAIGCBBAGECCCACCAEEAIGCBBAEEAAAAGECCCAEEABBAIGCCCACCAACBGGAABCCCAEEACCAECCGGAAAACCA",
    "CCACCAEEAEEABBAEEAGECCCAECCCEBABCCEBECCEEACCAEEACACGECCCACCAIGCECCBBAEEAGGAECCAAAGGABBAGECCCAEEABABIGCAAAEEA",
    "CCAECCCCAEEAAAAECCEEACCACCAGECCCACCAAAAEEAECCGGAAAAGECCACCCACCAGGAAAAAAAGGAACBECCEEACCAAAAECCGECGECEGBCACCCA",
    "CCACCAEEAEEACACCCAEEACCACCAGECAAAEEACCAECCCCAEEAAAAEEAGECCCACCACCAAAAGECEEACCAAAAGECCCAEEACCACCACACGGAAAAEEA",
    "ECCEEAAAAEEAAAAGECGECCCACCAEEAECCEEAAAACCAECCGGAAAAEEAAAAGECCCAEEAAAACCAGECCCACCAEEAECCCCACCACCACCAGECAAAEEA",
    "CCAECCEEACCAAAAEEAECCCCAEEACCAAAAIGCCCACCACCAECCAAAKICCCACCAECCEEAAAAEEACCAECCCCAEEAAAAIGCCCACCAAAAEEACACGGA",
    "AAACCAGECEEAAAAEEAAAAECCEEAEEAAAAGGACACCxcvFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_26: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCCCACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAGECCCAAAACCAGECGECAAA",
    "CCACCACCAIEECCACCACCAAAAGECCCACCAECCCCAGECCCAAAACCAGECGECAAACCACCAGECKGEAAACCACCACCAECCGECCCACCAAAAGECCCAGEC",
    "AAACCAGECCCAAAAGECCCAGECAAACCACCAGECECCCCACCACCACCAECCGECCCACCAAAAGECCCAGECAAACCAGECCCAAAAGECCCAGECECCCCACCA",
    "GECECCCCACCACCAAAAKGECCACCACCAAAAGECGECCCAAAACCAGECCCAECCCCACCAGECAAACCACCACCAIEECCACCACCAAAAGECGECCCAAAACCA",
    "GECCCAGECAAACCAGECGGAAAAGECGECGECAAACCACCAGECECCCCACCACCAAAAKGECCACCAAAACCAGECIGCCCAAAACCAGECCCAAAAGECCCAGEC",
    "CCACCACCAGECGECCCAAAACCAAAAMIEAAAEEAAAACCAGECGECAAAAAACCAGECGECECCCCACCAGECAAACCAGECCCAECCCCACCAEEAECCGECCCA",
    "CCAAAACCAGECCCAECCCCACCAGECHDDCCAGECCCAGECAAACCACCAECCECCCCACCACCACCAKGECCAEEAAAACCAGECGECAAAGECAAAIGCAAAGEC",
    "CCACCAGECBABAAACCAGECGECCCAAAACCACCAKGEAAACCACCACCAGECAAACCACCACCAGECAAACCACCACCAIEECCACCAGECCCAAAAGECCCABAB",
    "ECCGECCCACCACCAAAAOIGCCAGECCCAAAAGECCCACCAECCCCAGECCCACCAAAAGECGECCCACCAAAACCAKGECCAAAACCABABGECGECAAACCACCA",
    "GECAAAGECCCACCAGECAAACCAGECCCAGECAAACCACCAGECECCCCAGECCCACCAECCGECCCACCACCAECCGECGECCCAAAAGECCCACCACCAAAAE5Z",
    "lIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_27: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAABCGGAGECCCAAAAGGAECCEEAECCCCABABEEAECCCCABBAEEACCACCAAAAEEAGECEEAAAACCACACIIA",
    "CCACCACCAGECAAAGGAAAAECCCCAGGACCAECCGGABBACCAGGAABCEEAEEACCAECCEEAECCEEAEEAAAACACGGACCAEEAAAAECCCCACCAAAAGEC",
    "AAACCAGGAGGACACCCACCAEEAAAAECCCCAEEAECCCCACCAGECAAACEBCACCCAEEAGGACACCCAAAAEEACCAIGCBBAGGAAAACCAGECIGCAAAEEA",
    "ECCCCAAAAGECCCACCACCAEEACCAIGCAAACCACCACCAGECEEAAAACCAEEAECCCCAAAACCAEEAECCCCAEEAEEACACEEACCAAAAAAAIGCAAAGGA",
    "CCAGECAAAEEAAAAEEAGECCCACCAIGCBBAGECCCAAAACCAIGCAAAEEACCAEEAECCGGABBAGGACACAAACCAGGACACCCAEEACCAAAAGECCCAEEA",
    "CCACCAGECEEABBAEEAAAAECCEEAEEAAAAEEAECCEEAEEAAAAECCEEAAAACCAKGECCAAAAEEAEEAAAAECCGGAAAAGGAECCEEAAAACCAAAAGEC",
    "CCACCACCAGECAAAEEAEEAECCAAAEEACCAEEABBAGECCCAGGAAAAIGCBBACCAEEAGGACACCCAEEACCAAAAECCCCAECCCCAGECEEAGGAAAAEEA",
    "ECCAAAKKAEEACACCCACCACCACCAGECAAAEEAEEAECCCCACCAAAAEEAGECCCACCACCAAAAGECEEACCAAAAIGCCCAEEACCACCACACEEAAAAEEA",
    "ECCEEAAAAEEAAAAIGCECCCCACCAGGAECCEEAAAACCAECCEEAAAAEEAAAAGECEEAEEAAAACCAECCCCACCAEEAECCCCACCACCAEEAGECAAAEEA",
    "CCAECCEEACCAAAAEEAECCCCACCACCAAAAIGCEEACCACCAECCAAAIGCEEACCAECCEEAAAAEEACCAECCAAAGGAAAAIGCCCACCAAAAEEACACEEA",
    "AAACCAGECGGAAAAGGAAAAECCCCAEEAAAAGvctFAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_28: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECECCGECDBBCCAGECGECAAACCAECCGECEEACCAAAACCAKGECCACCAGECCCAGECDBBCCACCAGECECCCCA",
    "CCACCACCAECCCCACCACCAAAAGECGECCCAAAACCAGECCCAGECAAACCAGECCCAAAACCAGECKGECCACCAAAACCAGECGECCCAAAACCAGECCCAECC",
    "CCACCAGECCCAAAACCAGECGECCCAAAACCACCAIEECCACCAAAACCAIGCAAACCACCAGECECCCCACCACCACCAECCCCAGECCCACCAECCGECCCAGEC",
    "AAAGECCCAAAACCACCAKGECCACCAAAACCAGECCCACCAECCCCAGECCCAGECAAACCAGECCCAGECAAACCACCAGECAAACCACCAGECCCAECCCCACCA",
    "GECAAACCACCACCAKGEAAACCACCAGECGECAAAGECCCACCAECCCCACCAGECCCAECCCCACCAGECAAAGECCCACCAAAACCAGECCCAGECAAACCAGEC",
    "CCACCAAAAGECGECCCACCAAAACCAGECCCAGECAAACCAGECCCAAAACCAGECGECECCCCACCACCAKGECCAAAACCACCAGECAAACCAGECCCAGECAAA",
    "CCACCAGECECCCCACCACCAAAAGECGECCCACCACCAECCCCACCACCAECCGECCCACCACCAAAAKGECCACCACCAAAAGECAAAGECGECCCAGECCCACCA",
    "ECCCCAGECCCAAAACCAGECGECCCAAAACCACCAGECAAAGECCCACCAGECAAACCACCAGECGECAAACCACCACCAECCCCAGECCCACCAECCCCACCACCA",
    "ECCGECCCACCACCAAAAOIGCCACCACCAAAAGECGECCCAAAACCAGECCCACCAAAACCAKGECCACCAAAACCAGECGECAAACCACCAGECCCAAAACCAGEC",
    "GECAAACCACCACCAKGEAAACCACCACCAGECECCCCACCACCAECCCCAGECCCACCAIEECCACCACCACCAECCGECCCACCAAAAGECCCACCAGECAAAGEC",
    "C7ZlIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_29: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAIAIEEAGECCCAECCEEAAAAGECCCAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCAECCEEAAAAGEC",
    "EEAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAGECEEACACEEACCAECCCCAGECAAAGEC",
    "CCAECCECCEEACACEEAECCCCAECCEEACACEEACCAECCCCAIGCAAAGECCCAECCECCEEACACEEAECCCCAECCEEAAAAGECCCAECCCCAGECAAAGEC",
    "ECCCCAECCGGACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAECCEEACACEEAECCCCACCAGECAAAIGCCCAECCCCAGECCACEEA",
    "ECCCCAECCEEACACEEACCAECCCCAGECAAAGECCCAECCECCEEACACEEAGECCCAECCEEAAAAGECCCAECCCCAGECAAAGECECCCCAECCEEACACEEA",
    "ECCCCAECCEEAAAAGECEEAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAGECEEACACEEA",
    "CCAECCCCAGECAAAGECCCAECCECCEEACACEEAECCCCAECCEEACACEEACCAECCCCAIGCAAAGECCCAECCECCEEACACEEAECCCCAECCEEAAAAGEC",
    "CCAECCCCAGECAAAGECECCCCAECCGGACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAECCEEACACEEAECCCCACCAGECAAAIGC",
    "CCAECCCCAGECCACEEAECCCCAECCEEACACEEACCAECCCCAGECAAAGECCCAECCECCEEACACEEAGECCCAECCEEAAAAGECCCACjY9JAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_30: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECICGCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECC",
    "CCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCA",
    "ECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCA",
    "ECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCA",
    "CCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECC",
    "CCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAEjX7KAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_31: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAGCEKGECCACCAECCEEAAAAIGCECCCCACCAGECAAAIGCECCCCACCAEEACACEEAGECECCCCAGGACACEEA",
    "CCACCAECCIGCAAAGECCCACCAECCEEAECCGECCCACCAAAAGECECCEEAECCCCACCAGECECCEEACCAECCCCAEEACACIGCCCAECCCCAEEAAAAKGE",
    "CCACCAECCEEAAAAKGECCACCACCAGECAAAEEAIEECCACCAGECAAAEEAIEECCACCAEEACACEEAGECECCCCAGGACACEEACCAGECECCEEAAAAGEC",
    "CCAGECECCEEAAAAEEAECCGECCCAGECAAAEEAECCCCAGECEEACACEEACCAECCGECEEACACEEACCAECCGECEEAAAAGECCCACCAECCIGCAAAGEC",
    "CCACCACCAKGEAAACCAECCCCACCAKGEAAAEEACCAECCCCAEEAGCEEEACCAECCCCAGGAECCGECCCACCAECCEEAAAAKGECCACCACCAGECAAAIGC",
    "ECCCCACCAGECAAAIGCECCCCACCAEEACACEEAGECECCCCAEEACACEEAGECCCAECCEEAAAAGECGECCCAECCEEAAAAEEAECCGECCCAGECAAAEEA",
    "ECCGECCCAEEACACEEACCAIEECCAEEACACEEACCAECCGECGGAAAAGECCCACCAIEEEEAAAAGECCCACCAGECGECAAAEEAECCCCAAAAKGEAAAEEA",
    "CCAECCCCAIGCCACEEACCAECCCCAIGCAAAGECCCACCAECCEEAECCGECCCACCAECCEEAECCEEAECCCCACCAGECAAAIGCECCCCACCAEEACACIGC",
    "CCAECCCCAEEACACIGCCCACCAECCGGAAAAGECGECCCAECCEEAAAAEpZ1IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_32: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECIEEECCCCACCAIEECCACCAGECAAAGECGECECCCCACCAECCCCAGECGECAAAGECCCAAAAGECGECECCCCA",
    "GECCCAAAAKGECCAAAAGECCCAECCCCAGECECCCCAGECCCAECCGECCCAECCCCACCAECCGECGECCCAAAAGECCCAECCGECGECAAACCAGECCCAIEE",
    "CCACCAECCCCACCAIEECCAGECCCAECCCCACCAIEECCAGECAAACCAGECECCCCAGECCCAECCCCAKGEAAACCAGECAAAGECCCAGECECCCCACCAECC",
    "CCAKGEAAAGECCCACCAECCGECGECAAACCAGECAAACCAKGECCAECCCCAGECAAAGECGECAAAGECCCAAAAKGEGECCCAAAACCAGECAAAKGECCAAAA",
    "GECCCAGECECCCCAGECECCCCACCAECCGECCCAGECAAACCAGECECCGECCCAAAAGECGECGECAAACCAGECAAAGECCCAECCGECCCAGECAAACCAKGE",
    "ECCCCACCAAAAGECGECECCCCACCAGECAAAGECGECAAAGECGECAAACCAGECGECECCCCACCAAAAKGECCAECCCCACCAGECECCGECCCAAAAGECGEC",
    "AAAGECCCAGECAAAGECCCAECCGECGECAAACCACCAECCKGECCACCAAAAGECCCAIEECCACCAECCGECCCAGECAAAGECGECAAACCACCAIEECCAGEC",
    "CCAAAAGECKGEAAACCACCAECCGECGECCCAAAAGECCCAECCCCAGECECCGECCCAAAACCAKGEGECAAACCACCAIEECCAGECAAACCAGECGECAAAGEC",
    "CCAECCGECCCAAAAGECKGECCAAAACCACCAECCIvXrKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_33: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGAGCEKGECCACCACCAGECAAAIGCECCCCACCAGECAAAIGCCCAECCCCAEEACACEEAGECECCCCAEEAAAAGEC",
    "GECCCAECCEEAAAAGECEEACCACCAGECECCEEAECCCCACCAGECECCEEACCAECCCCAEEAGCEEEACCAECCCCACCAAAAKGECCACCAECCEEAAAAKGE",
    "CCACCACCAGECAAAIGCECCCCACCAGECAAAEEAGECECCCCAEEACACEEAGECCCAECCEEAAAAGECGECCCAECCEEAAAAEEAECCGECCCAGECAAAEEA",
    "ECCGECCCAEEACACEEACCAIEECCAEEACACEEACCACCAIEEEEAAAAGECCCACCAKGEEEAAAAEEAECCCCAGECGECAAAEEAECCCCACCAIGCCACEEA",
    "CCAECCCCAIGCCACEEACCACCAECCIGCAAAGECCCACCAECCEEAECCEEAECCCCACCAGECECCEEAECCCCACCAEEACACIGCCCAECCCCAEEAAAAKGE",
    "CCACCAECCEEAAAAKGECCACCACCAGECAAAEEAGCECCACCAGECAAAEEAGECECCCCAEEACACEEAGECECCCCAEEAAAAGECCCAGECECCGGAAAAGEC",
    "CCAGECCCAGECAAAEEAECCGECCCAGECAAAEEACCAECCGECEEACACEEACCAECCGECEEAAAAGECCCACCAIEEEEAAAAGECCCACCACCAKGEAAAEEA",
    "ECCCCACCAKGEAAAEEACCAECCCCAIGCCACEEACCACCAECCEEAECCGECCCACCAECCEEAECCEEAECCCCACCAGECAAAIGCECCCCACCAEEACACIGC",
    "CCAECCCCAEEACACIGCCCACCAECCEEACCAGECGECCCAECCEEAAAAEEACtZzIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_34: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECIEEECCCCACCAKGEAAACCAGECAAAGECGECAAACCAGECGECAAAGECGECAAAGECCCAAAAGECGECGECAAA",
    "CCACCAECCKGECCAAAAGECCCAGECAAAGECGECAAAGECCCAAAAGECGECGECAAACCAGECECCGECCCAAAAGECCCAGECECCCCACCAECCGECCCAECC",
    "GECCCAGECAAACCAKGEAAAGECCCAAAACCAGECKGEAAACCAGECAAAGECGECCCAECCCCAGECAAAGECCCAECCGECCCACCAECCGECGECAAACCAGEC",
    "AAAKGECCACCAAAAGECGECECCCCAGECAAAGECCCACCAIEECCAGECAAACCACCAIEEGECCCACCAECCCCAKGEAAACCACCAECCGECCCAGECAAAGEC",
    "GECAAACCAKGEAAAGECCCACCAAAAKGEGECAAACCACCAECCGECGECCCAECCCCAGECAAAGECCCAECCGECCCACCAAAAKGEGECAAACCAGECCCAIEE",
    "CCACCAAAAGECGECECCCCACCAGECECCCCACCAIEECCAGECAAACCACCAKGEECCCCACCAAAAGECKGEAAACCAGECCCAECCCCAGECAAAGECGECAAA",
    "CCAGECGECECCCCACCAECCGECGECAAACCACCAGECECCGECCCAAAAGECGECAAAGECGECCCAECCCCACCAECCGECGECAAACCACCAGECIEECCACCA",
    "ECCCCAGECECCCCACCAGECECCCCAGECAAAGECGECCCAAAACCAGECIEECCACCAAAAGECKGECCAAAACCAGECECCCCAGECAAACCAKGECCAAAAGEC",
    "GECECCCCACCAAAAKGEGECCCAAAACCACCAIEEGxXrKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_35: &str = concat!(
    "QOCAAAAAAAAACAAAAAGGCAAAAAAGGACCAEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEACCAECCEEAEEAAAAEEA",
    "AAACCACCAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAIGCCCACCACCAEEAAAACCAGECCCACCAEEAAAAEEACCAGECCCACCAAAAEEA",
    "CCACCAGECEEAAAAEEACCACCAAAAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAGECEEACCACCAEEAAAACCACCAGECCCAEEAAAAEEA",
    "CCACCAGECCCAAAAEEACCACCACCAIGCAAAEEACCACCAAAAEEAECCEEACCACCACCAEEAAAAIGCAAACCACCAEEAAAAEEAGECCCACCAEEAAAAEEA",
    "AAAGECCCAEEAAAAEEACCACCAGECEEAAAACCACCACCACCAGECCCAEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEA",
    "CCAGECCCAEEAAAAEEAAAACCAGECEEAAAAEEACCACCACCAIGCAAACCACCACCACCAEEAECCEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEA",
    "GECCCAAAAEEAAAAEEACCAECCEEAEEAAAAEEAAAACCACCAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAIGCCCACCACCAEEAAAACCA",
    "GECCCACCAEEAAAAEEACCAGECCCACCAAAAEEACCACCAGECEEAAAAEEACCACCAAAAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAGEC",
    "EEACCACCAEEAAAACCACCAGECCCAEEAAAAEEACCACCAGECCCAAAAEEACCACCACCAIGCAAAEEACCACCAAAAEEAECCEEACCACCACCAEEAAAAIGC",
    "AAACCACCAEEAAAAEEAGECCCACCAEEAAAAEEAAAAGECCCAEEAAAAEEACCACCAGECEEAAAACCACCACCACCAGECCCAEEACCACCACCACCAAAAIGC",
    "CCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEACCAGECCCAEEAAAAEEAAAACCAGECEEAAAAEEACCACCACCAIGCAAACCACCACCACCAEEAECCEEA",
    "CCACCACCACCAAAAIrevDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_36: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCAGECCCACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCAGECCCACCAGECCCAAAACCACCAGECCCA",
    "GECCCACCAECCCCACCACCACCAGECCCAAAAGECCCAGECCCACCACCAAAAGECCCACCACCACCAKGEAAACCACCACCAGECCCACCAAAACCAKGECCACCA",
    "CCACCAECCCCACCACCACCAGECGECAAACCACCAGECCCACCACCAAAAGECGECCCACCACCAGECAAACCACCACCAGECCCAGECCCAAAAGECCCACCACCA",
    "CCAGECAAAGECCCACCAGECCCACCAAAACCAGECCCACCAGECCCAECCCCACCACCACCAGECCCAAAAGECCCAGECCCACCACCAAAAGECCCACCACCAGEC",
    "GECAAACCACCACCAGECCCACCAAAACCAGECGECCCACCACCAECCCCACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCAGECCCACCAGEC",
    "CCAAAACCACCAGECCCAGECCCAAAAGECCCACCACCACCAGECAAACCAGECCCAGECCCACCAAAACCAGECCCACCAGECCCAECCCCACCACCACCAGECCCA",
    "AAACCAGECGECCCACCACCAAAAGECCCACCACCAGECGECCCAAAACCACCAGECCCACCACCAAAAKGECCACCACCACCAGECAAACCACCACCAGECCCAGEC",
    "AAACCAGECCCACCACCACCAECCCCAGECCCACCAGECCCAAAACCACCAGECCCACCAGECAAAGECCCACCACCACCAGECAAACCAGECCCAGECCCACCAAAA",
    "CCAGECCCACCACCAGECGECAAACCACCACCAGECCCACCAAAAGECGECCCACCACCACCAECCCCACCACCACCAKGECCAAAACCACCAGECCCACCACCAAAA",
    "KGECCACCACCACCAGECAAACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCACCAGECCCAGECCCAAAACCACCAGECCCACCAGECCCAECC",
    "CCACCACCACCAGECCCAAAACCAGECGECCCACCACCAAAAGECCCACCAClb7GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_37: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCAECCCCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCA",
    "CCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCAGECECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCAECCCCA",
    "CCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECC",
    "CCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECC",
    "CCACCACCAECCCCACCAECCCCACCAGECECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCA",
    "ECCCCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCA",
    "ECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCA",
    "ECCCCACCAGECECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCA",
    "CCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCAECCCCACCA",
    "CCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAIEECCA",
    "CCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCA",
    "CCACCAECCCCACCAECCCCACCACCAECCCCACCACCAECCCCACCAECCCCACCACCAECCCCACCACCAExatHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_38: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCACCACCAECCCCACCACCAECCCCACCA",
    "CCACCAECCCCACCACCAECCCCACCACCACCAECCCCAC7blGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_39: &str = concat!(
    "OMCAAAAAACCAAAAGCECCAECCECCEEACACEEAGECCCAECCEEAAAAGECCCAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCAECCEEAAAAGEC",
    "EEAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAGECEEACACEEACCAECCCCAGECAAAGEC",
    "CCAECCECCEEACACEEAECCCCAECCEEACACEEACCAECCCCAIGCAAAGECCCAECCECCEEACACEEAECCCCAECCEEAAAAGECCCAECCCCAGECAAAGEC",
    "ECCCCAECCGGACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAECCEEACACEEAECCCCACCAGECAAAIGCCCAECCCCAGECCACEEA",
    "ECCCCAECCEEACACEEACCAECCCCAGECAAAGECCCAECCECCEEACACEEAGECCCAECCEEAAAAGECCCAECCCCAGECAAAGECECCCCAECCEEACACEEA",
    "ECCCCAECCEEAAAAGECEEAECCCCAGECAAAGECECCCCAECCEEACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAGECEEACACEEA",
    "CCAECCCCAGECAAAGECCCAECCECCEEACACEEAECCCCAECCEEACACEEACCAECCCCAIGCAAAGECCCAECCECCEEACACEEAECCCCAECCEEAAAAGEC",
    "CCAECCCCAGECAAAGECECCCCAECCGGACACEEAECCCCACCAGECAAAGECCCAECCCCAGECCACEEAECCCCAECCEEACACEEAECCCCACCAGECAAAIGC",
    "CCAECCCCAGECCACEEAECCCCAECCEEACACEEACCAECCCCAGECAAAGECCCAECCECCEEACACEEAGECCCAECCEEAAAAGECCCACjY9JAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_40: &str = concat!(
    "OMCAAAAAAAAACCAEAECCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECC",
    "CCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCA",
    "ECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCA",
    "ECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCA",
    "CCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECC",
    "CCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAEjX7KAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_41: &str = concat!(
    "OMCAAAAAAECACAEEEACCACCAECCEEAAAAKGECCACCAECCEEAAAAIGCECCCCACCAGECAAAIGCECCCCACCAEEACACEEAGECECCCCAGGACACEEA",
    "CCACCAECCIGCAAAGECCCACCAECCEEAECCGECCCACCAAAAGECECCEEAECCCCACCAGECECCEEACCAECCCCAEEACACIGCCCAECCCCAEEAAAAKGE",
    "CCACCAECCEEAAAAKGECCACCACCAGECAAAEEAIEECCACCAGECAAAEEAIEECCACCAEEACACEEAGECECCCCAGGACACEEACCAGECECCEEAAAAGEC",
    "CCAGECECCEEAAAAEEAECCGECCCAGECAAAEEAECCCCAGECEEACACEEACCAECCGECEEACACEEACCAECCGECEEAAAAGECCCACCAECCIGCAAAGEC",
    "CCACCACCAKGEAAACCAECCCCACCAKGEAAAEEACCAECCCCAEEAGCEEEACCAECCCCAGGAECCGECCCACCAECCEEAAAAKGECCACCACCAGECAAAIGC",
    "ECCCCACCAGECAAAIGCECCCCACCAEEACACEEAGECECCCCAEEACACEEAGECCCAECCEEAAAAGECGECCCAECCEEAAAAEEAECCGECCCAGECAAAEEA",
    "ECCGECCCAEEACACEEACCAIEECCAEEACACEEACCAECCGECGGAAAAGECCCACCAIEEEEAAAAGECCCACCAGECGECAAAEEAECCCCAAAAKGEAAAEEA",
    "CCAECCCCAIGCCACEEACCAECCCCAIGCAAAGECCCACCAECCEEAECCGECCCACCAECCEEAECCEEAECCCCACCAGECAAAIGCECCCCACCAEEACACIGC",
    "CCAECCCCAEEACACIGCCCACCAECCGGAAAAGECGECCCAECCEEAAAAEpZ1IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_42: &str = concat!(
    "OMCAAAAAACCAEAECCACCAGECAAAGECCCAECCCCACCAIEECCACCAGECAAAGECGECECCCCACCAECCCCAGECGECAAAGECCCAAAAGECGECECCCCA",
    "GECCCAAAAKGECCAAAAGECCCAECCCCAGECECCCCAGECCCAECCGECCCAECCCCACCAECCGECGECCCAAAAGECCCAECCGECGECAAACCAGECCCAIEE",
    "CCACCAECCCCACCAIEECCAGECCCAECCCCACCAIEECCAGECAAACCAGECECCCCAGECCCAECCCCAKGEAAACCAGECAAAGECCCAGECECCCCACCAECC",
    "CCAKGEAAAGECCCACCAECCGECGECAAACCAGECAAACCAKGECCAECCCCAGECAAAGECGECAAAGECCCAAAAKGEGECCCAAAACCAGECAAAKGECCAAAA",
    "GECCCAGECECCCCAGECECCCCACCAECCGECCCAGECAAACCAGECECCGECCCAAAAGECGECGECAAACCAGECAAAGECCCAECCGECCCAGECAAACCAKGE",
    "ECCCCACCAAAAGECGECECCCCACCAGECAAAGECGECAAAGECGECAAACCAGECGECECCCCACCAAAAKGECCAECCCCACCAGECECCGECCCAAAAGECGEC",
    "AAAGECCCAGECAAAGECCCAECCGECGECAAACCACCAECCKGECCACCAAAAGECCCAIEECCACCAECCGECCCAGECAAAGECGECAAACCACCAIEECCAGEC",
    "CCAAAAGECKGEAAACCACCAECCGECGECCCAAAAGECCCAECCCCAGECECCGECCCAAAACCAKGEGECAAACCACCAIEECCAGECAAACCAGECGECAAAGEC",
    "CCAECCGECCCAAAAGECKGECCAAAACCACCAECCIvXrKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_43: &str = concat!(
    "OMCAAAAAAECAAAAGEECCACCAECCEEAAAAKGECCACCACCAGECAAAIGCECCCCACCAGECAAAIGCCCAECCCCAEEACACEEAGECECCCCAEEAAAAGEC",
    "GECCCAECCEEAAAAGECEEACCACCAGECECCEEAECCCCACCAGECECCEEACCAECCCCAEEAGCEEEACCAECCCCACCAAAAKGECCACCAECCEEAAAAKGE",
    "CCACCACCAGECAAAIGCECCCCACCAGECAAAEEAGECECCCCAEEACACEEAGECCCAECCEEAAAAGECGECCCAECCEEAAAAEEAECCGECCCAGECAAAEEA",
    "ECCGECCCAEEACACEEACCAIEECCAEEACACEEACCACCAIEEEEAAAAGECCCACCAKGEEEAAAAEEAECCCCAGECGECAAAEEAECCCCACCAIGCCACEEA",
    "CCAECCCCAIGCCACEEACCACCAECCIGCAAAGECCCACCAECCEEAECCEEAECCCCACCAGECECCEEAECCCCACCAEEACACIGCCCAECCCCAEEAAAAKGE",
    "CCACCAECCEEAAAAKGECCACCACCAGECAAAEEAGCECCACCAGECAAAEEAGECECCCCAEEACACEEAGECECCCCAEEAAAAGECCCAGECECCGGAAAAGEC",
    "CCAGECCCAGECAAAEEAECCGECCCAGECAAAEEACCAECCGECEEACACEEACCAECCGECEEAAAAGECCCACCAIEEEEAAAAGECCCACCACCAKGEAAAEEA",
    "ECCCCACCAKGEAAAEEACCAECCCCAIGCCACEEACCACCAECCEEAECCGECCCACCAECCEEAECCEEAECCCCACCAGECAAAIGCECCCCACCAEEACACIGC",
    "CCAECCCCAEEACACIGCCCACCAECCEEACCAGECGECCCAECCEEAAAAEEACtZzIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_44: &str = concat!(
    "OMCAAAAAACCACAAECECCACCAECCGECCCAECCCCACCAKGEAAACCAGECAAAGECGECAAACCAGECGECAAAGECGECAAAGECCCAAAAGECGECGECAAA",
    "CCACCAECCKGECCAAAAGECCCAGECAAAGECGECAAAGECCCAAAAGECGECGECAAACCAGECECCGECCCAAAAGECCCAGECECCCCACCAECCGECCCAECC",
    "GECCCAGECAAACCAKGEAAAGECCCAAAACCAGECKGEAAACCAGECAAAGECGECCCAECCCCAGECAAAGECCCAECCGECCCACCAECCGECGECAAACCAGEC",
    "AAAKGECCACCAAAAGECGECECCCCAGECAAAGECCCACCAIEECCAGECAAACCACCAIEEGECCCACCAECCCCAKGEAAACCACCAECCGECCCAGECAAAGEC",
    "GECAAACCAKGEAAAGECCCACCAAAAKGEGECAAACCACCAECCGECGECCCAECCCCAGECAAAGECCCAECCGECCCACCAAAAKGEGECAAACCAGECCCAIEE",
    "CCACCAAAAGECGECECCCCACCAGECECCCCACCAIEECCAGECAAACCACCAKGEECCCCACCAAAAGECKGEAAACCAGECCCAECCCCAGECAAAGECGECAAA",
    "CCAGECGECECCCCACCAECCGECGECAAACCACCAGECECCGECCCAAAAGECGECAAAGECGECCCAECCCCACCAECCGECGECAAACCACCAGECIEECCACCA",
    "ECCCCAGECECCCCACCAGECECCCCAGECAAAGECGECCCAAAACCAGECIEECCACCAAAAGECKGECCAAAACCAGECECCCCAGECAAACCAKGECCAAAAGEC",
    "GECECCCCACCAAAAKGEGECCCAAAACCACCAIEEGxXrKAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_45: &str = concat!(
    "QOCAAAAAAECAAAACCCCCACCACCAEEAAAAEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEACCAECCEEAEEAAAAEEA",
    "AAACCACCAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAIGCCCACCACCAEEAAAACCAGECCCACCAEEAAAAEEACCAGECCCACCAAAAEEA",
    "CCACCAGECEEAAAAEEACCACCAAAAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAGECEEACCACCAEEAAAACCACCAGECCCAEEAAAAEEA",
    "CCACCAGECCCAAAAEEACCACCACCAIGCAAAEEACCACCAAAAEEAECCEEACCACCACCAEEAAAAIGCAAACCACCAEEAAAAEEAGECCCACCAEEAAAAEEA",
    "AAAGECCCAEEAAAAEEACCACCAGECEEAAAACCACCACCACCAGECCCAEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEA",
    "CCAGECCCAEEAAAAEEAAAACCAGECEEAAAAEEACCACCACCAIGCAAACCACCACCACCAEEAECCEEACCACCACCACCAAAAIGCCCACCACCAEEAAAAEEA",
    "GECCCAAAAEEAAAAEEACCAECCEEAEEAAAAEEAAAACCACCAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAIGCCCACCACCAEEAAAACCA",
    "GECCCACCAEEAAAAEEACCAGECCCACCAAAAEEACCACCAGECEEAAAAEEACCACCAAAAIGCAAAEEACCACCACCAEEAECCEEAAAACCACCAEEAAAAGEC",
    "EEACCACCAEEAAAACCACCAGECCCAEEAAAAEEACCACCAGECCCAAAAEEACCACCACCAIGCAAAEEACCACCAAAAEEAECCEEACCACCACCAEEAAAAIGC",
    "AAACCACCAEEAAAAEEAGECCCACCAEEAAAAEEAAAAGECCCAEEAAAAEEACCACCAGECEEAAAACCACCACCACCAGECCCAEEACCACCACCACCAAAAIGC",
    "CCACCACCAEEAAAAEEAGECCCAAAAEEAAAAEEACCAGECCCAEEAAAAEEAAAACCAGECEEAAAAEEACCACCACCAIGCAAACCACCACCACCAEEAECCEEA",
    "CCACCACCACCAAAAIrevDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_46: &str = concat!(
    "OMCAAACCAAAACAACCACCCCCACCAGECAAACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCAGECCCACCAGECCCAAAACCACCAGECCCA",
    "GECCCACCAECCCCACCACCACCAGECCCAAAAGECCCAGECCCACCACCAAAAGECCCACCACCACCAKGEAAACCACCACCAGECCCACCAAAACCAKGECCACCA",
    "CCACCAECCCCACCACCACCAGECGECAAACCACCAGECCCACCACCAAAAGECGECCCACCACCAGECAAACCACCACCAGECCCAGECCCAAAAGECCCACCACCA",
    "CCAGECAAAGECCCACCAGECCCACCAAAACCAGECCCACCAGECCCAECCCCACCACCACCAGECCCAAAAGECCCAGECCCACCACCAAAAGECCCACCACCAGEC",
    "GECAAACCACCACCAGECCCACCAAAACCAGECGECCCACCACCAECCCCACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCAGECCCACCAGEC",
    "CCAAAACCACCAGECCCAGECCCAAAAGECCCACCACCACCAGECAAACCAGECCCAGECCCACCAAAACCAGECCCACCAGECCCAECCCCACCACCACCAGECCCA",
    "AAACCAGECGECCCACCACCAAAAGECCCACCACCAGECGECCCAAAACCACCAGECCCACCACCAAAAKGECCACCACCACCAGECAAACCACCACCAGECCCAGEC",
    "AAACCAGECCCACCACCACCAECCCCAGECCCACCAGECCCAAAACCACCAGECCCACCAGECAAAGECCCACCACCACCAGECAAACCAGECCCAGECCCACCAAAA",
    "CCAGECCCACCACCAGECGECAAACCACCACCAGECCCACCAAAAGECGECCCACCACCACCAECCCCACCACCACCAKGECCAAAACCACCAGECCCACCACCAAAA",
    "KGECCACCACCACCAGECAAACCACCACCAGECGECCCAAAACCAGECCCACCACCACCAECCCCACCAGECCCAGECCCAAAACCACCAGECCCACCAGECCCAECC",
    "CCACCACCACCAGECCCAAAACCAGECGECCCACCACCAAAAGECCCACCAClb7GAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_47: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCAMEICCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECC",
    "CCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCA",
    "ECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCA",
    "ECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCA",
    "CCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECC",
    "CCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECC",
    "CCAECCCCAGECECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECCGECCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAECC",
    "GECCCAECCCCAECCCCAECCCCAECCCCACCAECCCCAECCCCAECCCCAGECECCCCAECCCCAECCCCAEjX7KAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_48: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCAIEEIEECCACCAECCCCACCAGECECCCCACCAECCCCAGECECCCCACCACCAECCCCAGECECCCCACCAECCCCA",
    "CCACCAECCGECCCAECCCCACCAECCCCAGECECCCCACCACCAECCGECCCAECCCCACCAECCGECCCACCAECCCCACCAECCGECCCAECCCCACCACCAIEE",
    "CCACCAECCCCACCAIEECCACCACCAECCCCACCAIEECCACCAECCCCACCAIEECCACCACCAECCCCAGECECCCCACCAECCCCACCAGECECCCCACCAECC",
    "CCAGECECCCCACCACCAECCGECCCAECCCCACCAECCCCAGECCCAECCCCACCAECCGECCCAECCCCACCAECCGECCCACCAECCCCACCAECCGECCCAECC",
    "CCACCACCAIEECCACCAECCCCACCAIEECCACCACCAECCCCACCAIEECCACCAECCCCACCAGECECCCCACCAECCCCACCAIEECCACCACCAECCCCAGEC",
    "ECCCCACCAECCCCAGECECCCCACCACCAECCCCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCGECCCAECCCCACCACCAECCGECCCAECCCCACCA",
    "ECCGECCCACCAECCCCACCAIEECCACCAECCCCACCAECCGECCCACCAECCCCACCAIEECCACCAECCCCACCAGECECCCCACCAECCCCACCAIEECCACCA",
    "CCAECCCCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCCCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCCCAGECECCCCACCACCAECCGEC",
    "CCAECCCCACCAECCGECCCACCAECCCCACCAECCGECCCAECCCCACCACCAIEECCACCAECCCCACCAIEECCACCAECCCCACCACCAIEECCACCAECCCCA",
    "CCAIEECCACCACCAECCCCAGECECCCCACCAECCCCACCACtYzJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_49: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCAIEEIEECCACCACCAECCCCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCCCAGECECCCCACCACCAECC",
    "GECCCAECCCCACCAECCCCACCACCAECCGECCCAECCCCACCAECCGECCCACCAECCCCACCAIEECCACCAECCCCACCACCAIEECCACCAECCCCACCAIEE",
    "CCACCACCAECCCCAGECECCCCACCAECCCCACCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCGECCCAECCCCACCACCAECCGECCCAECCCCACCA",
    "ECCGECCCACCAECCCCACCAIEECCACCAECCCCACCACCAIEECCACCAECCCCACCAIEECCACCACCAECCCCAGECECCCCACCAECCCCACCAGECECCCCA",
    "CCAECCCCAGECECCCCACCACCAECCGECCCAECCCCACCAECCCCAGECCCAECCCCACCAECCGECCCAECCCCACCACCAECCGECCCAECCCCACCACCAIEE",
    "CCACCAECCCCACCAIEECCACCACCAECCCCACCAIEECCACCAECCCCACCAGECECCCCACCAECCCCAGECECCCCACCACCAECCCCAGECECCCCACCAECC",
    "CCAGECCCAECCCCACCAECCGECCCAECCCCACCACCAECCGECCCAECCCCACCAECCGECCCACCAECCCCACCAIEECCACCAECCCCACCACCAIEECCACCA",
    "ECCCCACCAIEECCACCACCAECCCCAGECECCCCACCACCAECCCCAGECECCCCACCAECCCCAGECCCAECCCCACCAECCCCAGECECCCCACCACCAECCGEC",
    "CCAECCCCACCAECCGECCCACCAECCCCACCAECCGECCCAECCCCACCACCAIEECCACCAECCCCACCAIEECCACCACCAECCCCACCAIEECCACCAECCCCA",
    "CCAGECECCCCACCAECCCCAGECECCCCACCACCAECCCCAEtYxJAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_50: &str = concat!(
    "OMCAAAAAACCACAAAAAEECAAACCACCACCACCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCA",
    "CCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCA",
    "CCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCACCAECCCCACCACCACCA",
    "CCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCA",
    "CCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCA",
    "CCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCA",
    "ECCCCACCACCACCACCACCAECCCCACCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCA",
    "ECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECC",
    "CCACCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECC",
    "CCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCACCAECC",
    "CCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCA",
    "CCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCA",
    "CCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCACCACCACCACCAECCCCACCA",
    "CCACCACCACCAECCCCACCACCACCAC7dlEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA",
    "AAAAAAAAAAAAAAAAAAAA",
);

const RUN_ROW_PROFILE_51: &str = "OMCAAAAAACCACAAAAAEECAAA";

const RUN_ROW_PROFILE_52: &str = "OMCAAACACAAACCACCACCACCA";

const RUN_ROW_PROFILE_53: &str = "OMCAAAAAACACECCCCACCACCA";

const RUN_ROW_PROFILE_54: &str = "OMCAAAAAACACCCAECCCCACCA";

const RUN_ROW_PROFILE_55: &str = "OMCAAAAAAAAACACECCCCAECC";

static RUN_ROW_PROFILES: &[(&str, u8, u8, &str)] = &[
    ("Arial", 0, 1, RUN_ROW_PROFILE_0),
    ("Arial", 0, 2, RUN_ROW_PROFILE_1),
    ("Arial", 1, 1, RUN_ROW_PROFILE_2),
    ("Arial", 1, 2, RUN_ROW_PROFILE_3),
    ("Arial", 2, 1, RUN_ROW_PROFILE_4),
    ("Arial", 2, 2, RUN_ROW_PROFILE_5),
    ("Arial", 3, 1, RUN_ROW_PROFILE_6),
    ("Arial", 3, 2, RUN_ROW_PROFILE_7),
    ("Aptos", 0, 1, RUN_ROW_PROFILE_8),
    ("Aptos", 0, 2, RUN_ROW_PROFILE_9),
    ("Aptos", 1, 1, RUN_ROW_PROFILE_8),
    ("Aptos", 1, 2, RUN_ROW_PROFILE_9),
    ("Aptos", 2, 1, RUN_ROW_PROFILE_8),
    ("Aptos", 2, 2, RUN_ROW_PROFILE_9),
    ("Aptos", 3, 1, RUN_ROW_PROFILE_8),
    ("Aptos", 3, 2, RUN_ROW_PROFILE_9),
    ("Aptos Display", 0, 1, RUN_ROW_PROFILE_8),
    ("Aptos Display", 0, 2, RUN_ROW_PROFILE_9),
    ("Aptos Display", 1, 1, RUN_ROW_PROFILE_8),
    ("Aptos Display", 1, 2, RUN_ROW_PROFILE_9),
    ("Aptos Display", 2, 1, RUN_ROW_PROFILE_8),
    ("Aptos Display", 2, 2, RUN_ROW_PROFILE_9),
    ("Aptos Display", 3, 1, RUN_ROW_PROFILE_8),
    ("Aptos Display", 3, 2, RUN_ROW_PROFILE_9),
    ("Aptos Narrow", 0, 1, RUN_ROW_PROFILE_8),
    ("Aptos Narrow", 0, 2, RUN_ROW_PROFILE_9),
    ("Aptos Narrow", 1, 1, RUN_ROW_PROFILE_8),
    ("Aptos Narrow", 1, 2, RUN_ROW_PROFILE_9),
    ("Aptos Narrow", 2, 1, RUN_ROW_PROFILE_8),
    ("Aptos Narrow", 2, 2, RUN_ROW_PROFILE_9),
    ("Aptos Narrow", 3, 1, RUN_ROW_PROFILE_8),
    ("Aptos Narrow", 3, 2, RUN_ROW_PROFILE_9),
    ("Calibri", 0, 1, RUN_ROW_PROFILE_10),
    ("Calibri", 0, 2, RUN_ROW_PROFILE_11),
    ("Calibri", 1, 1, RUN_ROW_PROFILE_10),
    ("Calibri", 1, 2, RUN_ROW_PROFILE_11),
    ("Calibri", 2, 1, RUN_ROW_PROFILE_10),
    ("Calibri", 2, 2, RUN_ROW_PROFILE_11),
    ("Calibri", 3, 1, RUN_ROW_PROFILE_10),
    ("Calibri", 3, 2, RUN_ROW_PROFILE_11),
    ("Cambria", 0, 1, RUN_ROW_PROFILE_12),
    ("Cambria", 0, 2, RUN_ROW_PROFILE_13),
    ("Cambria", 1, 1, RUN_ROW_PROFILE_12),
    ("Cambria", 1, 2, RUN_ROW_PROFILE_13),
    ("Cambria", 2, 1, RUN_ROW_PROFILE_12),
    ("Cambria", 2, 2, RUN_ROW_PROFILE_13),
    ("Cambria", 3, 1, RUN_ROW_PROFILE_12),
    ("Cambria", 3, 2, RUN_ROW_PROFILE_13),
    ("Courier New", 0, 1, RUN_ROW_PROFILE_14),
    ("Courier New", 0, 2, RUN_ROW_PROFILE_15),
    ("Courier New", 1, 1, RUN_ROW_PROFILE_16),
    ("Courier New", 1, 2, RUN_ROW_PROFILE_17),
    ("Courier New", 2, 1, RUN_ROW_PROFILE_18),
    ("Courier New", 2, 2, RUN_ROW_PROFILE_19),
    ("Courier New", 3, 1, RUN_ROW_PROFILE_20),
    ("Courier New", 3, 2, RUN_ROW_PROFILE_17),
    ("Times New Roman", 0, 1, RUN_ROW_PROFILE_21),
    ("Times New Roman", 0, 2, RUN_ROW_PROFILE_22),
    ("Times New Roman", 1, 1, RUN_ROW_PROFILE_23),
    ("Times New Roman", 1, 2, RUN_ROW_PROFILE_24),
    ("Times New Roman", 2, 1, RUN_ROW_PROFILE_25),
    ("Times New Roman", 2, 2, RUN_ROW_PROFILE_26),
    ("Times New Roman", 3, 1, RUN_ROW_PROFILE_27),
    ("Times New Roman", 3, 2, RUN_ROW_PROFILE_28),
    ("Meiryo", 0, 1, RUN_ROW_PROFILE_29),
    ("Meiryo", 0, 2, RUN_ROW_PROFILE_30),
    ("Meiryo", 1, 1, RUN_ROW_PROFILE_29),
    ("Meiryo", 1, 2, RUN_ROW_PROFILE_30),
    ("Meiryo", 2, 1, RUN_ROW_PROFILE_29),
    ("Meiryo", 2, 2, RUN_ROW_PROFILE_30),
    ("Meiryo", 3, 1, RUN_ROW_PROFILE_29),
    ("Meiryo", 3, 2, RUN_ROW_PROFILE_30),
    ("Yu Gothic", 0, 1, RUN_ROW_PROFILE_31),
    ("Yu Gothic", 0, 2, RUN_ROW_PROFILE_32),
    ("Yu Gothic", 1, 1, RUN_ROW_PROFILE_33),
    ("Yu Gothic", 1, 2, RUN_ROW_PROFILE_34),
    ("Yu Gothic", 2, 1, RUN_ROW_PROFILE_31),
    ("Yu Gothic", 2, 2, RUN_ROW_PROFILE_32),
    ("Yu Gothic", 3, 1, RUN_ROW_PROFILE_33),
    ("Yu Gothic", 3, 2, RUN_ROW_PROFILE_34),
    ("MS Mincho", 0, 1, RUN_ROW_PROFILE_35),
    ("MS Mincho", 0, 2, RUN_ROW_PROFILE_36),
    ("MS Mincho", 1, 1, RUN_ROW_PROFILE_35),
    ("MS Mincho", 1, 2, RUN_ROW_PROFILE_36),
    ("MS Mincho", 2, 1, RUN_ROW_PROFILE_35),
    ("MS Mincho", 2, 2, RUN_ROW_PROFILE_36),
    ("MS Mincho", 3, 1, RUN_ROW_PROFILE_35),
    ("MS Mincho", 3, 2, RUN_ROW_PROFILE_36),
    ("MS PGothic", 0, 1, RUN_ROW_PROFILE_35),
    ("MS PGothic", 0, 2, RUN_ROW_PROFILE_36),
    ("MS PGothic", 1, 1, RUN_ROW_PROFILE_35),
    ("MS PGothic", 1, 2, RUN_ROW_PROFILE_36),
    ("MS PGothic", 2, 1, RUN_ROW_PROFILE_35),
    ("MS PGothic", 2, 2, RUN_ROW_PROFILE_36),
    ("MS PGothic", 3, 1, RUN_ROW_PROFILE_35),
    ("MS PGothic", 3, 2, RUN_ROW_PROFILE_36),
    ("Aptos", 0, 0, RUN_ROW_PROFILE_37),
    ("Aptos", 1, 0, RUN_ROW_PROFILE_37),
    ("Aptos", 2, 0, RUN_ROW_PROFILE_37),
    ("Aptos", 3, 0, RUN_ROW_PROFILE_37),
    ("Cambria", 0, 0, RUN_ROW_PROFILE_38),
    ("Cambria", 1, 0, RUN_ROW_PROFILE_38),
    ("Cambria", 2, 0, RUN_ROW_PROFILE_38),
    ("Cambria", 3, 0, RUN_ROW_PROFILE_38),
    ("メイリオ", 0, 1, RUN_ROW_PROFILE_39),
    ("メイリオ", 0, 2, RUN_ROW_PROFILE_40),
    ("メイリオ", 1, 1, RUN_ROW_PROFILE_39),
    ("メイリオ", 1, 2, RUN_ROW_PROFILE_40),
    ("メイリオ", 2, 1, RUN_ROW_PROFILE_39),
    ("メイリオ", 2, 2, RUN_ROW_PROFILE_40),
    ("メイリオ", 3, 1, RUN_ROW_PROFILE_39),
    ("メイリオ", 3, 2, RUN_ROW_PROFILE_40),
    ("游ゴシック", 0, 1, RUN_ROW_PROFILE_41),
    ("游ゴシック", 0, 2, RUN_ROW_PROFILE_42),
    ("游ゴシック", 1, 1, RUN_ROW_PROFILE_43),
    ("游ゴシック", 1, 2, RUN_ROW_PROFILE_44),
    ("游ゴシック", 2, 1, RUN_ROW_PROFILE_41),
    ("游ゴシック", 2, 2, RUN_ROW_PROFILE_42),
    ("游ゴシック", 3, 1, RUN_ROW_PROFILE_43),
    ("游ゴシック", 3, 2, RUN_ROW_PROFILE_44),
    ("ＭＳ 明朝", 0, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ 明朝", 0, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ 明朝", 1, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ 明朝", 1, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ 明朝", 2, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ 明朝", 2, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ 明朝", 3, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ 明朝", 3, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ Ｐゴシック", 0, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ Ｐゴシック", 0, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ Ｐゴシック", 1, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ Ｐゴシック", 1, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ Ｐゴシック", 2, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ Ｐゴシック", 2, 2, RUN_ROW_PROFILE_46),
    ("ＭＳ Ｐゴシック", 3, 1, RUN_ROW_PROFILE_45),
    ("ＭＳ Ｐゴシック", 3, 2, RUN_ROW_PROFILE_46),
    ("Meiryo", 0, 0, RUN_ROW_PROFILE_47),
    ("Meiryo", 1, 0, RUN_ROW_PROFILE_47),
    ("Meiryo", 2, 0, RUN_ROW_PROFILE_47),
    ("Meiryo", 3, 0, RUN_ROW_PROFILE_47),
    ("Yu Gothic", 0, 0, RUN_ROW_PROFILE_48),
    ("Yu Gothic", 1, 0, RUN_ROW_PROFILE_49),
    ("Yu Gothic", 2, 0, RUN_ROW_PROFILE_48),
    ("Yu Gothic", 3, 0, RUN_ROW_PROFILE_49),
    ("MS Mincho", 0, 0, RUN_ROW_PROFILE_50),
    ("MS Mincho", 1, 0, RUN_ROW_PROFILE_50),
    ("MS Mincho", 2, 0, RUN_ROW_PROFILE_50),
    ("MS Mincho", 3, 0, RUN_ROW_PROFILE_50),
    ("MS PGothic", 0, 0, RUN_ROW_PROFILE_50),
    ("MS PGothic", 1, 0, RUN_ROW_PROFILE_50),
    ("MS PGothic", 2, 0, RUN_ROW_PROFILE_50),
    ("MS PGothic", 3, 0, RUN_ROW_PROFILE_50),
];

static TINY_ROW_PROFILES: &[(&str, u8, &str)] = &[
    ("Arial", 0, RUN_ROW_PROFILE_51),
    ("Arial", 1, RUN_ROW_PROFILE_51),
    ("Arial", 2, RUN_ROW_PROFILE_51),
    ("Arial", 3, RUN_ROW_PROFILE_51),
    ("Aptos Narrow", 0, RUN_ROW_PROFILE_51),
    ("Aptos Narrow", 1, RUN_ROW_PROFILE_51),
    ("Aptos Narrow", 2, RUN_ROW_PROFILE_51),
    ("Aptos Narrow", 3, RUN_ROW_PROFILE_51),
    ("Calibri", 0, RUN_ROW_PROFILE_51),
    ("Calibri", 1, RUN_ROW_PROFILE_51),
    ("Calibri", 2, RUN_ROW_PROFILE_51),
    ("Calibri", 3, RUN_ROW_PROFILE_51),
    ("Courier New", 0, RUN_ROW_PROFILE_51),
    ("Courier New", 1, RUN_ROW_PROFILE_51),
    ("Courier New", 2, RUN_ROW_PROFILE_51),
    ("Courier New", 3, RUN_ROW_PROFILE_51),
    ("Times New Roman", 0, RUN_ROW_PROFILE_51),
    ("Times New Roman", 1, RUN_ROW_PROFILE_51),
    ("Times New Roman", 2, RUN_ROW_PROFILE_51),
    ("Times New Roman", 3, RUN_ROW_PROFILE_51),
    ("ＭＳ Ｐゴシック", 0, RUN_ROW_PROFILE_52),
    ("ＭＳ Ｐゴシック", 1, RUN_ROW_PROFILE_52),
    ("ＭＳ Ｐゴシック", 2, RUN_ROW_PROFILE_52),
    ("ＭＳ Ｐゴシック", 3, RUN_ROW_PROFILE_52),
    ("ＭＳ 明朝", 0, RUN_ROW_PROFILE_52),
    ("ＭＳ 明朝", 1, RUN_ROW_PROFILE_52),
    ("ＭＳ 明朝", 2, RUN_ROW_PROFILE_52),
    ("ＭＳ 明朝", 3, RUN_ROW_PROFILE_52),
    ("游ゴシック", 0, RUN_ROW_PROFILE_53),
    ("游ゴシック", 1, RUN_ROW_PROFILE_54),
    ("游ゴシック", 2, RUN_ROW_PROFILE_53),
    ("游ゴシック", 3, RUN_ROW_PROFILE_54),
    ("メイリオ", 0, RUN_ROW_PROFILE_55),
    ("メイリオ", 1, RUN_ROW_PROFILE_55),
    ("メイリオ", 2, RUN_ROW_PROFILE_55),
    ("メイリオ", 3, RUN_ROW_PROFILE_55),
    ("Aptos Display", 0, RUN_ROW_PROFILE_51),
    ("Aptos Display", 1, RUN_ROW_PROFILE_51),
    ("Aptos Display", 2, RUN_ROW_PROFILE_51),
    ("Aptos Display", 3, RUN_ROW_PROFILE_51),
];

#[derive(Clone, Copy)]
struct LineHeightBounds {
    height: u16,
    baseline_min: u16,
    baseline_max: u16,
    descent_min: u16,
    descent_max: u16,
    row_only: bool,
}

fn row_profile(face: &str, bold: bool, italic: bool) -> Option<&'static str> {
    let face = if face.eq_ignore_ascii_case("Meiryo") { "メイリオ" }
        else if face.eq_ignore_ascii_case("Yu Gothic") { "游ゴシック" }
        else if face.eq_ignore_ascii_case("MS PGothic") { "ＭＳ Ｐゴシック" }
        else if face.eq_ignore_ascii_case("MS Mincho") { "ＭＳ 明朝" }
        else if face.eq_ignore_ascii_case("Aptos Display") { "Aptos Narrow" }
        else { face };
    let style = u8::from(bold) | (u8::from(italic) << 1);
    FONT_ROW_PROFILES.iter().find(|(held, flags, _)| {
        held.eq_ignore_ascii_case(face) && *flags == style
    }).map(|(_, _, packed)| *packed)
}

fn row_profile_value(byte: u8) -> i32 {
    match byte {
        b'A'..=b'Z' => i32::from(byte - b'A'),
        b'a'..=b'z' => i32::from(byte - b'a') + 26,
        b'0'..=b'9' => i32::from(byte - b'0') + 52,
        b'+' => 62,
        b'/' => 63,
        _ => unreachable!("invalid generated row profile"),
    }
}

fn measured_line_height(face: &str, size: f32, bold: bool, italic: bool) -> Option<LineHeightBounds> {
    let em = em_pixels(size);
    if !(1..=546).contains(&em) { return None; }
    let packed = row_profile(face, bold, italic)?;
    let (mut height, mut baseline, mut width) = (0, 0, 0);
    for point in packed.as_bytes().chunks_exact(3).take(em as usize) {
        height += row_profile_value(point[0]) - 16;
        baseline += row_profile_value(point[1]) - 16;
        width = row_profile_value(point[2]);
    }
    Some(LineHeightBounds {
        height: height as u16,
        baseline_min: baseline as u16,
        baseline_max: (baseline + width) as u16,
        descent_min: (height - baseline - width).max(0) as u16,
        descent_max: (height - baseline).max(0) as u16,
        row_only: width != 0,
    })
}

fn read_run_row_delta(text: &[u8], cursor: &mut usize) -> i32 {
    let (mut value, mut shift) = (0u32, 0);
    loop {
        let byte = row_profile_value(text[*cursor]) as u32;
        *cursor += 1;
        value |= (byte & 31) << shift;
        if byte & 32 == 0 { break; }
        shift += 5;
    }
    if value & 1 == 0 { (value >> 1) as i32 } else { -((value >> 1) as i32) - 1 }
}

fn decoded_run_row_bounds(packed: &str, em: i32, row_only: bool) -> LineHeightBounds {
    let (mut cursor, mut height, mut baseline, mut descent) = (0, 0i32, 0i32, 0i32);
    for _ in 0..em {
        height += read_run_row_delta(packed.as_bytes(), &mut cursor);
        baseline += read_run_row_delta(packed.as_bytes(), &mut cursor);
        descent += read_run_row_delta(packed.as_bytes(), &mut cursor);
    }
    LineHeightBounds {
        height: height as u16,
        baseline_min: baseline as u16,
        baseline_max: baseline as u16,
        descent_min: descent as u16,
        descent_max: descent as u16,
        row_only,
    }
}

fn measured_run_row_bounds(
    face: &str, size: f32, bold: bool, italic: bool, position: u8,
) -> Option<LineHeightBounds> {
    let em = em_pixels(size);
    if !(1..=546).contains(&em) { return None; }
    let flags = u8::from(bold) | (u8::from(italic) << 1);
    if position == 0 && em <= 8 {
        if let Some((_, _, packed)) = TINY_ROW_PROFILES.iter().find(|(held, style, _)| {
            held.eq_ignore_ascii_case(face) && *style == flags
        }) {
            let mut bounds = decoded_run_row_bounds(packed, em, true);
            // A row minimum is not a measured multi-line advance.
            // Retain an existing exact ordinary advance only when it is unchanged.
            if let Some(old) = measured_line_height(face, size, bold, italic) {
                bounds.row_only = old.row_only || old.height != bounds.height
                    || old.baseline_min != bounds.baseline_min
                    || old.descent_min != bounds.descent_min;
            }
            return Some(bounds);
        }
    }
    let (_, _, _, packed) = RUN_ROW_PROFILES.iter().find(|(held, style, kind, _)| {
        held.eq_ignore_ascii_case(face) && *style == flags && *kind == position
    })?;
    Some(decoded_run_row_bounds(packed, em, position != 0 || em <= 8))
}

fn legacy_run_height_bounds(face: &str, size: f32, bold: bool, italic: bool,
    position: Option<&str>) -> Option<LineHeightBounds> {
    let kind = match position {
        None | Some("baseline") => 0,
        Some("superscript") => 1,
        Some("subscript") => 2,
        _ => return None,
    };
    if let Some(bounds) = measured_run_row_bounds(face, size, bold, italic, kind) {
        return Some(bounds);
    }
    if kind == 0 {
        if let Some(bounds) = measured_line_height(face, size, bold, italic) { return Some(bounds); }
    }
    font_run_line_box(face, size, bold, italic, position).map(|(height, baseline)| {
        LineHeightBounds {
            height, baseline_min: baseline, baseline_max: baseline,
            descent_min: height.saturating_sub(baseline),
            descent_max: height.saturating_sub(baseline), row_only: false,
        }
    })
}


fn run_height_bounds(face: &str, size: f32, bold: bool, italic: bool,
    position: Option<&str>) -> Option<LineHeightBounds> {
    let mut bounds = legacy_run_height_bounds(face, size, bold, italic, position)?;
    let kind = match position {
        None | Some("baseline") => 0,
        Some("superscript") => 1,
        Some("subscript") => 2,
        _ => return None,
    };
    let em = em_pixels(size);
    if (1..=8).contains(&em) {
        let flags = u8::from(bold) | (u8::from(italic) << 1);
        if let Some((_, _, _, packed)) = TINY_FLOW_PROFILES.iter().find(|(held, style, held_kind, _)| {
            held.eq_ignore_ascii_case(face) && *style == flags && *held_kind == kind
        }) {
            let flow = decoded_run_row_bounds(packed, em, false);
            bounds.baseline_min = flow.baseline_min;
            bounds.baseline_max = flow.baseline_max;
            bounds.descent_min = flow.descent_min;
            bounds.descent_max = flow.descent_max;
        }
    } else if bounds.baseline_min == bounds.baseline_max
        && bounds.descent_min == bounds.descent_max {
        let extra = 2u16.saturating_sub(bounds.descent_min);
        bounds.baseline_min = bounds.baseline_min.saturating_sub(extra);
        bounds.baseline_max = bounds.baseline_max.saturating_sub(extra);
        bounds.descent_min = bounds.descent_min.saturating_add(extra);
        bounds.descent_max = bounds.descent_max.saturating_add(extra);
    }
    Some(bounds)
}

fn decorated_run_height_bounds(
    face: &str, size: f32, bold: bool, italic: bool,
    position: Option<&str>, underline: Option<CellUnderlineKind>,
) -> Option<LineHeightBounds> {
    let plain = run_height_bounds(face, size, bold, italic, position)?;
    let kind = match underline {
        Some(kind @ (CellUnderlineKind::SingleAccounting | CellUnderlineKind::DoubleAccounting)) => kind,
        _ => return Some(plain),
    };
    if !matches!(position, None | Some("baseline")) { return None; }
    let em = em_pixels(size);
    if !(1..=546).contains(&em) { return None; }
    let face = if face.eq_ignore_ascii_case("Meiryo") { "メイリオ" }
        else if face.eq_ignore_ascii_case("Yu Gothic") { "游ゴシック" }
        else if face.eq_ignore_ascii_case("MS PGothic") { "ＭＳ Ｐゴシック" }
        else if face.eq_ignore_ascii_case("MS Mincho") { "ＭＳ 明朝" }
        else { face };
    let flags = u8::from(bold) | (u8::from(italic) << 1);
    let packed = ACCOUNTING_ROW_PROFILES.iter().find(|(held, style, rule, _)| {
        held.eq_ignore_ascii_case(face) && *style == flags && *rule == kind
    })?.3.as_bytes();
    let mut cursor = 16;
    let mut height = 0i32;
    for _ in 0..em {
        let mut value = 0u32;
        let mut shift = 0;
        loop {
            let byte = row_profile_value(packed[cursor]) as u32;
            cursor += 1;
            value |= (byte & 31) << shift;
            if byte & 32 == 0 { break; }
            shift += 5;
        }
        height += if value & 1 == 0 { (value >> 1) as i32 }
            else { -((value >> 1) as i32) - 1 };
    }
    let (baseline_min, baseline_max) = if em <= 8 {
        let at = (em as usize - 1) * 2;
        (row_profile_value(packed[at]) as u16, row_profile_value(packed[at + 1]) as u16)
    } else { (plain.baseline_min, plain.baseline_max) };
    Some(LineHeightBounds {
        height: height as u16, baseline_min, baseline_max,
        descent_min: (height as u16).saturating_sub(baseline_max),
        descent_max: (height as u16).saturating_sub(baseline_min),
        row_only: plain.row_only,
    })
}

/// The height of a shared text line, without requiring an exact glyph baseline.
/// Small-font bounds can be sufficient when a larger neighboring font contains
/// them. An uncertain baseline must not discard that neighbor's known height.
pub fn font_text_row_height<'a>(
    normal_face: &str,
    normal_size: f32,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool, Option<&'a str>, u32)>,
) -> Option<u16> {
    font_decorated_text_row_height(normal_face, normal_size,
        cells.into_iter().map(|(face, size, bold, italic, anchor, position, lines)| {
            (face, size, bold, italic, anchor, position, lines, None)
        }))
}


// Measured tiny multiline advances and their shared-line extents.
const TINY_FLOW_PROFILE_0: &str = "OKEAAAAAACCACCAAAAEEAAAA";
const TINY_FLOW_PROFILE_1: &str = "QMEAAAAAAAAACCAAAAGGAAAA";
const TINY_FLOW_PROFILE_2: &str = "MIEAAACCAAAACCAECCCCAECC";
const TINY_FLOW_PROFILE_3: &str = "OKEAAAAAACCAAAAGECCCAECC";
const TINY_FLOW_PROFILE_4: &str = "MIEAAACCACCAECCCCACCACCA";
const TINY_FLOW_PROFILE_5: &str = "OKEAAAAAAEEACACEEACCACCA";
const TINY_FLOW_PROFILE_6: &str = "MIEAAACCACCAECCCCACCAGEC";
const TINY_FLOW_PROFILE_7: &str = "MIEAAACCACCACCAECCCCACCA";
const TINY_FLOW_PROFILE_8: &str = "OKEAAAAAAEEAAAAGECCCACCA";
const TINY_FLOW_PROFILE_9: &str = "OKEAAACCAAAACCACCACCACCA";
const TINY_FLOW_PROFILE_10: &str = "QMEAAAAAAEEAAAACCACCACCA";

static TINY_FLOW_PROFILES: &[(&str, u8, u8, &str)] = &[
    ("Aptos", 0, 0, TINY_FLOW_PROFILE_0),
    ("Aptos", 0, 1, TINY_FLOW_PROFILE_1),
    ("Aptos", 0, 2, TINY_FLOW_PROFILE_0),
    ("Aptos", 1, 0, TINY_FLOW_PROFILE_0),
    ("Aptos", 1, 1, TINY_FLOW_PROFILE_1),
    ("Aptos", 1, 2, TINY_FLOW_PROFILE_0),
    ("Aptos", 2, 0, TINY_FLOW_PROFILE_0),
    ("Aptos", 2, 1, TINY_FLOW_PROFILE_1),
    ("Aptos", 2, 2, TINY_FLOW_PROFILE_0),
    ("Aptos", 3, 0, TINY_FLOW_PROFILE_0),
    ("Aptos", 3, 1, TINY_FLOW_PROFILE_1),
    ("Aptos", 3, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 0, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 0, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Display", 0, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 1, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 1, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Display", 1, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 2, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 2, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Display", 2, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 3, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Display", 3, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Display", 3, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 0, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 0, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Narrow", 0, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 1, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 1, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Narrow", 1, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 2, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 2, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Narrow", 2, 2, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 3, 0, TINY_FLOW_PROFILE_0),
    ("Aptos Narrow", 3, 1, TINY_FLOW_PROFILE_1),
    ("Aptos Narrow", 3, 2, TINY_FLOW_PROFILE_0),
    ("Arial", 0, 0, TINY_FLOW_PROFILE_0),
    ("Arial", 0, 1, TINY_FLOW_PROFILE_1),
    ("Arial", 0, 2, TINY_FLOW_PROFILE_0),
    ("Arial", 1, 0, TINY_FLOW_PROFILE_0),
    ("Arial", 1, 1, TINY_FLOW_PROFILE_1),
    ("Arial", 1, 2, TINY_FLOW_PROFILE_0),
    ("Arial", 2, 0, TINY_FLOW_PROFILE_0),
    ("Arial", 2, 1, TINY_FLOW_PROFILE_1),
    ("Arial", 2, 2, TINY_FLOW_PROFILE_0),
    ("Arial", 3, 0, TINY_FLOW_PROFILE_0),
    ("Arial", 3, 1, TINY_FLOW_PROFILE_1),
    ("Arial", 3, 2, TINY_FLOW_PROFILE_0),
    ("Calibri", 0, 0, TINY_FLOW_PROFILE_0),
    ("Calibri", 0, 1, TINY_FLOW_PROFILE_1),
    ("Calibri", 0, 2, TINY_FLOW_PROFILE_0),
    ("Calibri", 1, 0, TINY_FLOW_PROFILE_0),
    ("Calibri", 1, 1, TINY_FLOW_PROFILE_1),
    ("Calibri", 1, 2, TINY_FLOW_PROFILE_0),
    ("Calibri", 2, 0, TINY_FLOW_PROFILE_0),
    ("Calibri", 2, 1, TINY_FLOW_PROFILE_1),
    ("Calibri", 2, 2, TINY_FLOW_PROFILE_0),
    ("Calibri", 3, 0, TINY_FLOW_PROFILE_0),
    ("Calibri", 3, 1, TINY_FLOW_PROFILE_1),
    ("Calibri", 3, 2, TINY_FLOW_PROFILE_0),
    ("Cambria", 0, 0, TINY_FLOW_PROFILE_0),
    ("Cambria", 0, 1, TINY_FLOW_PROFILE_1),
    ("Cambria", 0, 2, TINY_FLOW_PROFILE_0),
    ("Cambria", 1, 0, TINY_FLOW_PROFILE_0),
    ("Cambria", 1, 1, TINY_FLOW_PROFILE_1),
    ("Cambria", 1, 2, TINY_FLOW_PROFILE_0),
    ("Cambria", 2, 0, TINY_FLOW_PROFILE_0),
    ("Cambria", 2, 1, TINY_FLOW_PROFILE_1),
    ("Cambria", 2, 2, TINY_FLOW_PROFILE_0),
    ("Cambria", 3, 0, TINY_FLOW_PROFILE_0),
    ("Cambria", 3, 1, TINY_FLOW_PROFILE_1),
    ("Cambria", 3, 2, TINY_FLOW_PROFILE_0),
    ("Courier New", 0, 0, TINY_FLOW_PROFILE_0),
    ("Courier New", 0, 1, TINY_FLOW_PROFILE_1),
    ("Courier New", 0, 2, TINY_FLOW_PROFILE_0),
    ("Courier New", 1, 0, TINY_FLOW_PROFILE_0),
    ("Courier New", 1, 1, TINY_FLOW_PROFILE_1),
    ("Courier New", 1, 2, TINY_FLOW_PROFILE_0),
    ("Courier New", 2, 0, TINY_FLOW_PROFILE_0),
    ("Courier New", 2, 1, TINY_FLOW_PROFILE_1),
    ("Courier New", 2, 2, TINY_FLOW_PROFILE_0),
    ("Courier New", 3, 0, TINY_FLOW_PROFILE_0),
    ("Courier New", 3, 1, TINY_FLOW_PROFILE_1),
    ("Courier New", 3, 2, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 0, 0, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 0, 1, TINY_FLOW_PROFILE_1),
    ("MS Mincho", 0, 2, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 1, 0, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 1, 1, TINY_FLOW_PROFILE_1),
    ("MS Mincho", 1, 2, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 2, 0, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 2, 1, TINY_FLOW_PROFILE_1),
    ("MS Mincho", 2, 2, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 3, 0, TINY_FLOW_PROFILE_0),
    ("MS Mincho", 3, 1, TINY_FLOW_PROFILE_1),
    ("MS Mincho", 3, 2, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 0, 0, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 0, 1, TINY_FLOW_PROFILE_1),
    ("MS PGothic", 0, 2, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 1, 0, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 1, 1, TINY_FLOW_PROFILE_1),
    ("MS PGothic", 1, 2, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 2, 0, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 2, 1, TINY_FLOW_PROFILE_1),
    ("MS PGothic", 2, 2, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 3, 0, TINY_FLOW_PROFILE_0),
    ("MS PGothic", 3, 1, TINY_FLOW_PROFILE_1),
    ("MS PGothic", 3, 2, TINY_FLOW_PROFILE_0),
    ("Meiryo", 0, 0, TINY_FLOW_PROFILE_0),
    ("Meiryo", 0, 1, TINY_FLOW_PROFILE_1),
    ("Meiryo", 0, 2, TINY_FLOW_PROFILE_0),
    ("Meiryo", 1, 0, TINY_FLOW_PROFILE_0),
    ("Meiryo", 1, 1, TINY_FLOW_PROFILE_1),
    ("Meiryo", 1, 2, TINY_FLOW_PROFILE_0),
    ("Meiryo", 2, 0, TINY_FLOW_PROFILE_0),
    ("Meiryo", 2, 1, TINY_FLOW_PROFILE_1),
    ("Meiryo", 2, 2, TINY_FLOW_PROFILE_0),
    ("Meiryo", 3, 0, TINY_FLOW_PROFILE_0),
    ("Meiryo", 3, 1, TINY_FLOW_PROFILE_1),
    ("Meiryo", 3, 2, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 0, 0, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 0, 1, TINY_FLOW_PROFILE_1),
    ("Times New Roman", 0, 2, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 1, 0, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 1, 1, TINY_FLOW_PROFILE_1),
    ("Times New Roman", 1, 2, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 2, 0, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 2, 1, TINY_FLOW_PROFILE_1),
    ("Times New Roman", 2, 2, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 3, 0, TINY_FLOW_PROFILE_0),
    ("Times New Roman", 3, 1, TINY_FLOW_PROFILE_1),
    ("Times New Roman", 3, 2, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 0, 0, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 0, 1, TINY_FLOW_PROFILE_1),
    ("Yu Gothic", 0, 2, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 1, 0, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 1, 1, TINY_FLOW_PROFILE_1),
    ("Yu Gothic", 1, 2, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 2, 0, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 2, 1, TINY_FLOW_PROFILE_1),
    ("Yu Gothic", 2, 2, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 3, 0, TINY_FLOW_PROFILE_0),
    ("Yu Gothic", 3, 1, TINY_FLOW_PROFILE_1),
    ("Yu Gothic", 3, 2, TINY_FLOW_PROFILE_0),
    ("メイリオ", 0, 0, TINY_FLOW_PROFILE_2),
    ("メイリオ", 0, 1, TINY_FLOW_PROFILE_3),
    ("メイリオ", 0, 2, TINY_FLOW_PROFILE_2),
    ("メイリオ", 1, 0, TINY_FLOW_PROFILE_2),
    ("メイリオ", 1, 1, TINY_FLOW_PROFILE_3),
    ("メイリオ", 1, 2, TINY_FLOW_PROFILE_2),
    ("メイリオ", 2, 0, TINY_FLOW_PROFILE_2),
    ("メイリオ", 2, 1, TINY_FLOW_PROFILE_3),
    ("メイリオ", 2, 2, TINY_FLOW_PROFILE_2),
    ("メイリオ", 3, 0, TINY_FLOW_PROFILE_2),
    ("メイリオ", 3, 1, TINY_FLOW_PROFILE_3),
    ("メイリオ", 3, 2, TINY_FLOW_PROFILE_2),
    ("游ゴシック", 0, 0, TINY_FLOW_PROFILE_4),
    ("游ゴシック", 0, 1, TINY_FLOW_PROFILE_5),
    ("游ゴシック", 0, 2, TINY_FLOW_PROFILE_6),
    ("游ゴシック", 1, 0, TINY_FLOW_PROFILE_7),
    ("游ゴシック", 1, 1, TINY_FLOW_PROFILE_8),
    ("游ゴシック", 1, 2, TINY_FLOW_PROFILE_7),
    ("游ゴシック", 2, 0, TINY_FLOW_PROFILE_4),
    ("游ゴシック", 2, 1, TINY_FLOW_PROFILE_5),
    ("游ゴシック", 2, 2, TINY_FLOW_PROFILE_6),
    ("游ゴシック", 3, 0, TINY_FLOW_PROFILE_7),
    ("游ゴシック", 3, 1, TINY_FLOW_PROFILE_8),
    ("游ゴシック", 3, 2, TINY_FLOW_PROFILE_7),
    ("ＭＳ 明朝", 0, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 0, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ 明朝", 0, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 1, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 1, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ 明朝", 1, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 2, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 2, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ 明朝", 2, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 3, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ 明朝", 3, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ 明朝", 3, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 0, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 0, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ Ｐゴシック", 0, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 1, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 1, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ Ｐゴシック", 1, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 2, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 2, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ Ｐゴシック", 2, 2, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 3, 0, TINY_FLOW_PROFILE_9),
    ("ＭＳ Ｐゴシック", 3, 1, TINY_FLOW_PROFILE_10),
    ("ＭＳ Ｐゴシック", 3, 2, TINY_FLOW_PROFILE_9),
];


fn run_flow_height_bounds(
    face: &str, size: f32, bold: bool, italic: bool, position: Option<&str>,
) -> Option<LineHeightBounds> {
    let kind = match position {
        None | Some("baseline") => 0,
        Some("superscript") => 1,
        Some("subscript") => 2,
        _ => return None,
    };
    let em = em_pixels(size);
    if (1..=8).contains(&em) {
        let flags = u8::from(bold) | (u8::from(italic) << 1);
        let (_, _, _, packed) = TINY_FLOW_PROFILES.iter().find(|(held, style, run_kind, _)| {
            held.eq_ignore_ascii_case(face) && *style == flags && *run_kind == kind
        })?;
        return Some(decoded_run_row_bounds(packed, em, false));
    }
    let mut bounds = run_height_bounds(face, size, bold, italic, position)?;
    if bounds.baseline_min != bounds.baseline_max || bounds.descent_min != bounds.descent_max {
        return None;
    }
    bounds.row_only = false;
    Some(bounds)
}

/// Shared row height including the room below accounting underlines.
pub fn font_decorated_text_row_height<'a>(
    normal_face: &str,
    normal_size: f32,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool, Option<&'a str>, u32, Option<CellUnderlineKind>)>,
) -> Option<u16> {
    font_styled_decorated_text_row_height(normal_face, normal_size, false, false, cells)
}

/// Shared row height using the actual weight and slant of the base font.
pub fn font_styled_decorated_text_row_height<'a>(
    normal_face: &str,
    normal_size: f32,
    normal_bold: bool,
    normal_italic: bool,
    cells: impl IntoIterator<Item = (&'a str, f32, bool, bool, bool, Option<&'a str>, u32, Option<CellUnderlineKind>)>,
) -> Option<u16> {
    let cells = cells.into_iter().collect::<Vec<_>>();
    let flow = cells.iter().any(|cell| cell.6 > 1) && cells.iter().all(|cell| !cell.4);
    let normal = run_height_bounds(normal_face, normal_size, normal_bold, normal_italic, None)?;
    let normal_minimum = normal.height;
    let normal = if flow {
        run_flow_height_bounds(normal_face, normal_size, normal_bold, normal_italic, None)?
    } else {
        normal
    };
    if normal.height >= 546 { return Some(546); }
    let mut low_baseline = normal.baseline_min;
    let mut high_baseline = normal.baseline_max;
    let mut low_descent = normal.descent_min;
    let mut high_descent = normal.descent_max;
    let mut minimum = normal_minimum;
    for (face, size, bold, italic, wrapped_anchor, position, lines, underline) in cells {
        let bounds = decorated_run_height_bounds(face, size, bold, italic, position, underline)?;
        let row_minimum = bounds.height;
        let bounds = if flow {
            if matches!(underline,
                Some(CellUnderlineKind::SingleAccounting | CellUnderlineKind::DoubleAccounting)) {
                return None;
            }
            run_flow_height_bounds(face, size, bold, italic, position)?
        } else if bounds.row_only && (wrapped_anchor || lines > 1)
            && !matches!(position, None | Some("baseline"))
        {
            // Existing exact line boxes still define the supported flow sizes.
            // Row-only samples do not establish new wrapped or merged advances.
            let (height, baseline) = font_run_line_box(face, size, bold, italic, position)?;
            LineHeightBounds {
                height, baseline_min: baseline, baseline_max: baseline,
                descent_min: height.saturating_sub(baseline),
                descent_max: height.saturating_sub(baseline), row_only: false,
            }
        } else { bounds };
        if bounds.height >= 546 {
            return if wrapped_anchor { None } else { Some(546) };
        }
        if lines > 1 && matches!(underline,
            Some(CellUnderlineKind::SingleAccounting | CellUnderlineKind::DoubleAccounting)) {
            return None;
        }
        // The measured tiny-font height contains the minimum single-line row.
        // It does not establish the spacing of several wrapped lines.
        if lines > 1 && (bounds.row_only || bounds.baseline_min != bounds.baseline_max) { return None; }
        low_descent = low_descent.max(bounds.descent_min);
        high_descent = high_descent.max(bounds.descent_max);
        if !wrapped_anchor {
            let extra = u16::try_from(lines.saturating_sub(1)).unwrap_or(u16::MAX);
            let advance = bounds.height.saturating_mul(extra);
            low_baseline = low_baseline.max(bounds.baseline_min.saturating_add(advance));
            high_baseline = high_baseline.max(bounds.baseline_max.saturating_add(advance));
            minimum = minimum.max(bounds.height.saturating_add(advance));
            if flow { minimum = minimum.max(row_minimum); }
        }
    }
    let low = low_baseline.saturating_add(low_descent).max(minimum).min(546);
    let high = high_baseline.saturating_add(high_descent).max(minimum).min(546);
    (low == high).then_some(low)
}

fn em_pixels(points: f32) -> i32 {
    (points * 96.0 / 72.0).round() as i32
}

/// The row a cell asks for when its text holds a raised or lowered run, in
/// 96-dpi pixels.
///
/// A superscript reaches above the line its font would otherwise need, and
/// Excel grows the row for it: 游ゴシック 11 asks for 25 pixels plainly and 27
/// with a raised run, ＭＳ Ｐゴシック 11 for 15 and 21. Measured by
/// `tools\metrics\_xlsx_raised_extra.py` on Excel's own AutoFit, in a
/// workbook whose blank row is short enough not to hide the answer — ﻿    // 83 measurements.
/// A face and size the sweep did not reach is left to its plain line.
static RAISED_ROW_PX: &[(&str, u16, u16)] = &[
    ("ＭＳ 明朝", 42, 20),
    ("ＭＳ 明朝", 44, 21),
    ("ＭＳ 明朝", 48, 22),
    ("ＭＳ 明朝", 56, 27),
    ("ＭＳ 明朝", 64, 29),
    ("ＭＳ 明朝", 72, 33),
    ("ＭＳ 明朝", 80, 37),
    ("ＭＳ 明朝", 96, 44),
    ("ＭＳ ゴシック", 42, 20),
    ("ＭＳ ゴシック", 44, 21),
    ("ＭＳ ゴシック", 48, 22),
    ("ＭＳ ゴシック", 56, 27),
    ("ＭＳ ゴシック", 64, 29),
    ("ＭＳ ゴシック", 72, 33),
    ("ＭＳ ゴシック", 80, 37),
    ("ＭＳ ゴシック", 96, 44),
    ("ＭＳ Ｐゴシック", 42, 20),
    ("ＭＳ Ｐゴシック", 44, 21),
    ("ＭＳ Ｐゴシック", 48, 22),
    ("ＭＳ Ｐゴシック", 56, 27),
    ("ＭＳ Ｐゴシック", 64, 29),
    ("ＭＳ Ｐゴシック", 72, 33),
    ("ＭＳ Ｐゴシック", 80, 37),
    ("ＭＳ Ｐゴシック", 96, 44),
    ("ＭＳ Ｐ明朝", 42, 20),
    ("ＭＳ Ｐ明朝", 44, 21),
    ("ＭＳ Ｐ明朝", 48, 22),
    ("ＭＳ Ｐ明朝", 56, 27),
    ("ＭＳ Ｐ明朝", 64, 29),
    ("ＭＳ Ｐ明朝", 72, 33),
    ("ＭＳ Ｐ明朝", 80, 37),
    ("ＭＳ Ｐ明朝", 96, 44),
    ("游ゴシック", 32, 20),
    ("游ゴシック", 36, 23),
    ("游ゴシック", 40, 24),
    ("游ゴシック", 42, 25),
    ("游ゴシック", 44, 27),
    ("游ゴシック", 48, 29),
    ("游ゴシック", 56, 35),
    ("游ゴシック", 64, 37),
    ("游ゴシック", 72, 44),
    ("游ゴシック", 80, 48),
    ("游ゴシック", 96, 58),
    ("メイリオ", 32, 22),
    ("メイリオ", 36, 22),
    ("メイリオ", 40, 24),
    ("メイリオ", 42, 25),
    ("メイリオ", 44, 27),
    ("メイリオ", 48, 29),
    ("メイリオ", 56, 33),
    ("メイリオ", 64, 36),
    ("メイリオ", 72, 42),
    ("メイリオ", 80, 47),
    ("メイリオ", 96, 55),
    ("Meiryo UI", 40, 21),
    ("Meiryo UI", 42, 22),
    ("Meiryo UI", 44, 23),
    ("Meiryo UI", 48, 25),
    ("Meiryo UI", 56, 29),
    ("Meiryo UI", 64, 31),
    ("Meiryo UI", 72, 36),
    ("Meiryo UI", 80, 41),
    ("Meiryo UI", 96, 48),
    ("Calibri", 40, 20),
    ("Calibri", 42, 22),
    ("Calibri", 44, 23),
    ("Calibri", 48, 24),
    ("Calibri", 56, 28),
    ("Calibri", 64, 31),
    ("Calibri", 72, 35),
    ("Calibri", 80, 39),
    ("Calibri", 96, 48),
    ("游ゴシック Light", 32, 20),
    ("游ゴシック Light", 36, 23),
    ("游ゴシック Light", 40, 24),
    ("游ゴシック Light", 42, 25),
    ("游ゴシック Light", 44, 26),
    ("游ゴシック Light", 48, 29),
    ("游ゴシック Light", 56, 35),
    ("游ゴシック Light", 64, 37),
    ("游ゴシック Light", 72, 44),
    ("游ゴシック Light", 80, 48),
    ("游ゴシック Light", 96, 58),
];

/// The row a raised or lowered run asks for, when the sweep has measured it.
pub fn raised_row_px(face: &str, size: f32) -> Option<u16> {
    let size_q = (size * 4.0).round() as u16;
    if let Some((_, _, px)) = RAISED_ROW_PX
        .iter()
        .find(|(f, s, _)| *f == face && *s == size_q)
    {
        return Some(*px);
    }
    let em = em_pixels(size);
    RAISED_ROW_PX
        .iter()
        .find(|(f, s, _)| *f == face && em_pixels(*s as f32 / 4.0) == em)
        .map(|(_, _, px)| *px)
}

#[cfg(test)]
mod tests {
    #[test]
    fn mixed_line_heights_retain_small_and_capped_runs() {
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("Arial", 1.05, false, false, false, None, 1),
            ("Arial", 12.0, false, false, false, None, 1),
        ]), Some(20));
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("Arial", 409.55, false, false, false, None, 1),
            ("Arial", 8.0, false, false, false, None, 1),
        ]), Some(546));
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("メイリオ", 5.95, false, false, false, None, 1),
            ("Courier New", 9.0, false, false, false, None, 1),
        ]), Some(17));
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("Oxi unmeasured row font", 12.0, false, false, false, None, 1),
        ]), None);
    }

    use super::font_cell_line_box;

    #[test]
    fn mixed_fonts_keep_room_above_and_below_the_baseline() {
        assert_eq!(font_cell_line_box("Arial", 12.0, false, false, "Aptos Narrow", 11.0, false), Some((21, 17)));
        assert_eq!(font_cell_line_box("Meiryo", 8.0, false, false, "Aptos Narrow", 11.0, false), Some((22, 16)));
        assert_eq!(font_cell_line_box("メイリオ", 11.0, false, false, "Arial", 20.0, false), Some((36, 28)));
    }

    #[test]
    fn a_styled_face_keeps_its_own_line_box() {
        assert_eq!(font_cell_line_box("Courier New", 12.0, true, false, "Arial", 8.0, false), Some((22, 17)));
        assert_eq!(font_cell_line_box("Arial", 24.0, false, true, "Arial", 8.0, false), Some((41, 33)));
        assert_eq!(font_cell_line_box("Unmeasured", 12.0, false, false, "Arial", 8.0, false), None);
    }

    #[test]
    fn half_point_and_large_sizes_share_the_normal_font_line() {
        assert_eq!(font_cell_line_box("Arial", 12.5, false, false, "Aptos Narrow", 11.0, false).map(|line| line.0), Some(22));
        assert_eq!(font_cell_line_box("Courier New", 9.5, true, true, "Arial", 8.0, false).map(|line| line.0), Some(18));
        assert_eq!(font_cell_line_box("Times New Roman", 14.0, true, false, "Arial", 20.0, false).map(|line| line.0), Some(34));
        assert_eq!(font_cell_line_box("Meiryo", 48.0, false, false, "Arial", 52.0, false).map(|line| line.0), Some(102));
        assert_eq!(font_cell_line_box("Aptos Narrow", 72.0, false, false, "Arial", 80.0, false).map(|line| line.0), Some(136));
        assert_eq!(font_cell_line_box("Arial", 72.0, false, false, "Arial", 120.0, false).map(|line| line.0), Some(199));
    }

    #[test]
    fn a_row_shares_baselines_and_descents_across_cells() {
        use super::font_row_line_box;
        let height = |fonts: &[(&str, f32, bool, bool, bool)]| {
            font_row_line_box("Arial", 8.0, fonts.iter().copied()).map(|line| line.0)
        };
        assert_eq!(height(&[
            ("Arial", 12.0, false, false, false), ("Meiryo", 8.0, false, false, false),
        ]), Some(23));
        assert_eq!(height(&[
            ("Arial", 12.0, false, false, true), ("Meiryo", 8.0, false, false, false),
        ]), Some(19));
        assert_eq!(height(&[
            ("Meiryo", 72.0, false, false, true), ("Arial", 72.0, false, false, false),
        ]), Some(146));
        assert_eq!(height(&[
            ("Times New Roman", 48.0, true, false, true), ("Arial", 48.0, true, false, false),
        ]), Some(80));
        assert_eq!(height(&[
            ("Arial", 12.0, false, false, false), ("Arial", 72.0, false, false, true),
            ("Arial", 12.0, false, false, false),
        ]), Some(39));
    }

    #[test]
    fn small_fonts_share_the_normal_baseline_in_both_directions() {
        for (bold, italic) in [(false, false), (true, false), (false, true), (true, true)] {
            for face in ["Arial", "Times New Roman", "Courier New"] {
                assert_eq!(
                    font_cell_line_box(face, 6.0, bold, italic, "Arial", 8.0, false)
                        .map(|line| line.0),
                    Some(15),
                    "{face}: bold={bold}, italic={italic}",
                );
            }
        }
        assert_eq!(
            font_cell_line_box("Arial", 12.0, false, false, "Arial", 6.0, false)
                .map(|line| line.0),
            Some(20),
        );
        assert_eq!(
            font_cell_line_box("Times New Roman", 6.0, false, false, "Arial", 6.0, false)
                .map(|line| line.0),
            Some(11),
        );
        assert_eq!(
            super::font_row_line_box("Arial", 6.0, [
                ("Arial", 8.0, false, false, false),
                ("Times New Roman", 6.0, false, false, false),
            ]).map(|line| line.0),
            Some(15),
        );
    }

    #[test]
    fn positioned_runs_share_the_normal_baseline_and_descent() {
        assert_eq!(
            super::font_position_row_line_box("Arial", 20.0, [
                ("Courier New", 12.0, false, true, false, Some("subscript")),
            ]),
            Some((35, 28)),
        );
        assert_eq!(
            super::font_position_row_line_box("Arial", 8.0, [
                ("Arial", 12.0, false, false, false, Some("superscript")),
            ]),
            Some((24, 21)),
        );
        assert_eq!(
            super::font_position_row_line_box("游ゴシック", 72.0, [
                ("Arial", 72.0, false, false, false, Some("superscript")),
            ]),
            Some((161, 118)),
        );
    }

    #[test]
    fn wrapped_positions_share_the_last_baseline_once() {
        assert_eq!(
            super::font_text_row_line_box("Arial", 20.0, [
                ("Arial", 12.0, false, false, false, Some("superscript"), 2),
            ]),
            Some((51, 45)),
        );
        assert_eq!(
            super::font_text_row_line_box("Arial", 20.0, [
                ("Arial", 12.0, false, false, false, Some("subscript"), 3),
            ]),
            Some((78, 72)),
        );
        assert_eq!(
            super::font_text_row_line_box("Arial", 20.0, [
                ("Arial", 12.0, false, false, false, None, 2),
            ]),
            Some((43, 37)),
        );
    }


    #[test]
    fn tiny_rows_and_multiline_text_share_canonical_extents() {
        assert_eq!(super::font_text_row_height("Times New Roman", 6.75, [
            ("Arial", 6.75, false, false, false, None, 1),
        ]), Some(12));
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("メイリオ", 4.5, false, false, false, Some("superscript"), 1),
            ("Times New Roman", 6.75, false, false, false, None, 1),
        ]), Some(13));
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("Times New Roman", 6.75, false, false, false, None, 2),
        ]), Some(25));
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("Times New Roman", 6.75, false, false, false, None, 3),
        ]), Some(36));
    }

    #[test]
    fn tiny_multiline_advances_share_the_last_line_extents() {
        for (normal_size, face, position, lines, expected) in [
            (8.0, "Arial", None, 2, 15),
            (8.0, "Arial", None, 3, 22),
            (8.0, "Arial", Some("superscript"), 3, 25),
            (8.0, "メイリオ", Some("subscript"), 3, 19),
            (20.0, "メイリオ", Some("subscript"), 7, 46),
        ] {
            assert_eq!(super::font_text_row_height("Arial", normal_size, [
                (face, 1.0, false, false, false, position, lines),
            ]), Some(expected));
        }
    }

    #[test]
    fn tiny_script_row_minimum_is_separate_from_its_vertical_extents() {
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("Arial", 3.75, false, false, false, Some("superscript"), 1),
        ]), Some(9));
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("Arial", 3.75, false, false, false, Some("superscript"), 1),
            ("メイリオ", 4.5, false, false, false, Some("subscript"), 1),
        ]), Some(10));
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("Arial", 5.25, false, false, false, Some("superscript"), 1),
            ("メイリオ", 4.5, false, false, false, Some("subscript"), 1),
        ]), Some(13));
    }

    #[test]
    fn tiny_requested_font_names_keep_distinct_row_heights() {
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("Meiryo", 4.5, false, false, false, None, 1),
        ]), Some(9));
        assert_eq!(super::font_text_row_height("Arial", 1.0, [
            ("メイリオ", 4.5, false, false, false, None, 1),
        ]), Some(10));
    }

    #[test]
    fn existing_positioned_line_boxes_remain_available_in_flow() {
        assert_eq!(super::font_text_row_height("Arial", 8.0, [
            ("Arial", 12.0, false, false, true, Some("superscript"), 1),
        ]), Some(15));
        assert_eq!(super::font_text_row_height("Arial", 20.0, [
            ("Arial", 12.0, false, false, false, Some("superscript"), 2),
        ]), Some(51));
    }

}

// Normal-style default geometry measured from untouched worksheets. The
// digit unit comes from the difference between two integer column widths.
static FONT_DEFAULT_GEOMETRY_PX: &[(&str, u16, u16, u16)] = &[
    ("Arial", 160, 29, 66),
    ("AR P丸ゴシック体E", 44, 8, 19),
    ("Aptos Narrow", 24, 5, 11),
    ("Aptos Narrow", 32, 6, 15),
    ("Aptos Narrow", 36, 6, 16),
    ("Aptos Narrow", 40, 7, 18),
    ("Aptos Narrow", 42, 7, 19),
    ("Aptos Narrow", 44, 7, 20),
    ("Aptos Narrow", 48, 8, 21),
    ("Aptos Narrow", 50, 9, 23),
    ("Aptos Narrow", 56, 10, 25),
    ("Aptos Narrow", 64, 11, 28),
    ("Aptos Narrow", 72, 12, 32),
    ("Aptos Narrow", 80, 14, 35),
    ("Aptos Narrow", 96, 16, 42),
    ("Aptos Narrow", 104, 18, 46),
    ("Aptos Narrow", 112, 19, 48),
    ("Aptos Narrow", 144, 24, 62),
    ("Aptos Narrow", 192, 32, 85),
    ("Aptos Narrow", 288, 49, 125),
    ("Arial", 4, 2, 7),
    ("Arial", 16, 3, 9),
    ("Arial", 24, 5, 11),
    ("Arial", 32, 6, 15),
    ("Arial", 36, 7, 16),
    ("Arial", 40, 7, 17),
    ("Arial", 42, 8, 18),
    ("Arial", 44, 8, 19),
    ("Arial", 48, 9, 20),
    ("Arial", 50, 9, 22),
    ("Arial", 52, 9, 22),
    ("Arial", 56, 11, 24),
    ("Arial", 64, 12, 27),
    ("Arial", 72, 13, 31),
    ("Arial", 80, 15, 34),
    ("Arial", 96, 18, 40),
    ("Arial", 104, 19, 44),
    ("Arial", 112, 21, 46),
    ("Arial", 128, 24, 54),
    ("Arial", 144, 27, 59),
    ("Arial", 192, 36, 79),
    ("Arial", 200, 37, 83),
    ("Arial", 208, 38, 85),
    ("Arial", 288, 53, 120),
    ("Arial", 320, 60, 133),
    ("Arial", 480, 89, 199),
    ("Arial", 1636, 303, 546),
    ("Arial Unicode MS", 44, 8, 19),
    ("BIZ UDPゴシック", 24, 5, 11),
    ("BIZ UDPゴシック", 32, 8, 15),
    ("BIZ UDPゴシック", 36, 9, 16),
    ("BIZ UDPゴシック", 40, 10, 16),
    ("BIZ UDPゴシック", 42, 11, 17),
    ("BIZ UDPゴシック", 44, 11, 18),
    ("BIZ UDPゴシック", 48, 12, 19),
    ("BIZ UDPゴシック", 56, 14, 22),
    ("BIZ UDPゴシック", 64, 16, 25),
    ("BIZ UDPゴシック", 72, 18, 28),
    ("BIZ UDPゴシック", 80, 21, 31),
    ("BIZ UDPゴシック", 96, 24, 37),
    ("BIZ UDPゴシック", 144, 36, 55),
    ("BIZ UDゴシック", 24, 5, 11),
    ("BIZ UDゴシック", 32, 6, 15),
    ("BIZ UDゴシック", 36, 6, 16),
    ("BIZ UDゴシック", 40, 7, 16),
    ("BIZ UDゴシック", 42, 7, 17),
    ("BIZ UDゴシック", 44, 8, 18),
    ("BIZ UDゴシック", 48, 8, 19),
    ("BIZ UDゴシック", 56, 10, 22),
    ("BIZ UDゴシック", 64, 11, 25),
    ("BIZ UDゴシック", 72, 12, 28),
    ("BIZ UDゴシック", 80, 14, 31),
    ("BIZ UDゴシック", 96, 16, 37),
    ("BIZ UDゴシック", 144, 24, 55),
    ("Calibri", 24, 5, 11),
    ("Calibri", 32, 6, 15),
    ("Calibri", 36, 6, 16),
    ("Calibri", 40, 7, 17),
    ("Calibri", 42, 7, 19),
    ("Calibri", 44, 7, 20),
    ("Calibri", 48, 8, 21),
    ("Calibri", 50, 9, 23),
    ("Calibri", 52, 9, 23),
    ("Calibri", 56, 10, 25),
    ("Calibri", 64, 11, 28),
    ("Calibri", 72, 12, 31),
    ("Calibri", 80, 14, 35),
    ("Calibri", 96, 16, 42),
    ("Calibri", 112, 19, 48),
    ("Calibri", 144, 24, 62),
    ("Calibri", 192, 32, 82),
    ("Calibri", 288, 49, 123),
    ("Cambria", 32, 6, 14),
    ("Cambria", 38, 7, 17),
    ("Cambria", 42, 8, 18),
    ("Cambria", 44, 8, 19),
    ("Cambria", 48, 9, 21),
    ("Cambria", 50, 9, 22),
    ("Cambria", 56, 11, 24),
    ("Cambria", 64, 12, 27),
    ("Cambria", 72, 13, 30),
    ("Cambria", 80, 15, 34),
    ("Cambria", 96, 18, 40),
    ("Cambria", 112, 20, 46),
    ("Cambria", 144, 27, 60),
    ("Cambria", 192, 35, 79),
    ("Cambria", 288, 53, 118),
    ("Century", 24, 5, 11),
    ("Century", 28, 5, 12),
    ("Century", 36, 7, 18),
    ("Century", 40, 7, 17),
    ("Century", 42, 8, 18),
    ("Century", 44, 8, 19),
    ("Century", 48, 9, 21),
    ("Century", 56, 11, 24),
    ("Century", 60, 11, 25),
    ("Consolas", 32, 6, 15),
    ("Consolas", 38, 7, 17),
    ("Consolas", 42, 8, 19),
    ("Consolas", 44, 8, 20),
    ("Consolas", 48, 9, 21),
    ("Consolas", 50, 9, 22),
    ("Consolas", 56, 10, 25),
    ("Consolas", 64, 12, 27),
    ("Consolas", 72, 13, 31),
    ("Consolas", 80, 15, 35),
    ("Consolas", 96, 18, 41),
    ("Consolas", 112, 20, 47),
    ("Consolas", 144, 26, 61),
    ("Consolas", 192, 35, 81),
    ("Consolas", 288, 53, 121),
    ("Courier New", 24, 5, 11),
    ("Courier New", 32, 7, 15),
    ("Courier New", 38, 8, 18),
    ("Courier New", 42, 8, 19),
    ("Courier New", 44, 9, 20),
    ("Courier New", 48, 10, 21),
    ("Courier New", 50, 10, 23),
    ("Courier New", 52, 10, 23),
    ("Courier New", 56, 11, 25),
    ("Courier New", 64, 13, 28),
    ("Courier New", 72, 14, 32),
    ("Courier New", 80, 16, 35),
    ("Courier New", 96, 19, 42),
    ("Courier New", 112, 22, 48),
    ("Courier New", 144, 29, 62),
    ("Courier New", 192, 38, 83),
    ("Courier New", 288, 58, 123),
    ("HGP創英角ｺﾞｼｯｸUB", 48, 10, 19),
    ("HGP教科書体", 44, 8, 19),
    ("HGP行書体", 32, 6, 15),
    ("HGP行書体", 36, 7, 16),
    ("HGP行書体", 40, 7, 17),
    ("HGP行書体", 44, 8, 19),
    ("HGP行書体", 48, 9, 20),
    ("HGPｺﾞｼｯｸM", 24, 5, 11),
    ("HGPｺﾞｼｯｸM", 32, 6, 15),
    ("HGPｺﾞｼｯｸM", 36, 7, 16),
    ("HGPｺﾞｼｯｸM", 40, 7, 17),
    ("HGPｺﾞｼｯｸM", 44, 8, 19),
    ("HGPｺﾞｼｯｸM", 48, 9, 20),
    ("HGPｺﾞｼｯｸM", 56, 11, 24),
    ("HGS行書体", 32, 6, 15),
    ("HGS行書体", 36, 7, 16),
    ("HGS行書体", 40, 7, 17),
    ("HGS行書体", 44, 8, 19),
    ("HGS行書体", 48, 9, 20),
    ("HGSｺﾞｼｯｸ", 36, 7, 16),
    ("HGSｺﾞｼｯｸE", 36, 7, 15),
    ("HGSｺﾞｼｯｸM", 32, 6, 15),
    ("HGSｺﾞｼｯｸM", 36, 7, 16),
    ("HGSｺﾞｼｯｸM", 44, 8, 19),
    ("HG創英角ﾎﾟｯﾌﾟ体", 40, 7, 16),
    ("HG明朝B", 80, 15, 34),
    ("HG行書体", 44, 8, 19),
    ("MS Gothic", 24, 5, 11),
    ("MS Gothic", 28, 5, 12),
    ("MS Gothic", 32, 6, 14),
    ("MS Gothic", 36, 6, 15),
    ("MS Gothic", 40, 7, 16),
    ("MS Gothic", 42, 7, 17),
    ("MS Gothic", 44, 8, 18),
    ("MS Gothic", 46, 8, 18),
    ("MS Gothic", 47, 8, 19),
    ("MS Gothic", 48, 8, 19),
    ("MS Gothic", 50, 9, 20),
    ("MS Gothic", 52, 9, 20),
    ("MS Gothic", 54, 9, 22),
    ("MS Gothic", 55, 9, 22),
    ("MS Gothic", 56, 10, 23),
    ("MS Gothic", 58, 10, 23),
    ("MS Gothic", 59, 10, 24),
    ("MS Gothic", 60, 10, 24),
    ("MS Gothic", 63, 11, 25),
    ("MS Gothic", 64, 11, 25),
    ("MS Gothic", 70, 12, 27),
    ("MS Gothic", 72, 12, 28),
    ("MS Gothic", 74, 13, 30),
    ("MS Gothic", 80, 14, 32),
    ("MS Gothic", 88, 15, 34),
    ("MS Gothic", 96, 16, 38),
    ("MS Gothic", 112, 19, 43),
    ("MS Gothic", 144, 24, 56),
    ("MS Gothic", 192, 32, 74),
    ("MS Mincho", 12, 3, 8),
    ("MS Mincho", 24, 5, 11),
    ("MS Mincho", 28, 5, 12),
    ("MS Mincho", 30, 6, 13),
    ("MS Mincho", 32, 6, 14),
    ("MS Mincho", 36, 6, 15),
    ("MS Mincho", 40, 7, 16),
    ("MS Mincho", 42, 7, 17),
    ("MS Mincho", 44, 8, 18),
    ("MS Mincho", 46, 8, 18),
    ("MS Mincho", 47, 8, 19),
    ("MS Mincho", 48, 8, 19),
    ("MS Mincho", 50, 9, 20),
    ("MS Mincho", 52, 9, 20),
    ("MS Mincho", 54, 9, 22),
    ("MS Mincho", 55, 9, 22),
    ("MS Mincho", 56, 10, 23),
    ("MS Mincho", 60, 10, 24),
    ("MS Mincho", 63, 11, 25),
    ("MS Mincho", 64, 11, 25),
    ("MS Mincho", 68, 12, 27),
    ("MS Mincho", 70, 12, 27),
    ("MS Mincho", 72, 12, 28),
    ("MS Mincho", 74, 13, 30),
    ("MS Mincho", 76, 13, 30),
    ("MS Mincho", 80, 14, 32),
    ("MS Mincho", 88, 15, 34),
    ("MS Mincho", 96, 16, 38),
    ("MS Mincho", 112, 19, 43),
    ("MS Mincho", 144, 24, 56),
    ("MS P ゴシック", 24, 5, 11),
    ("MS P ゴシック", 28, 5, 12),
    ("MS P ゴシック", 30, 6, 13),
    ("MS P ゴシック", 32, 6, 15),
    ("MS P ゴシック", 36, 7, 16),
    ("MS P ゴシック", 37, 7, 16),
    ("MS P ゴシック", 40, 7, 17),
    ("MS P ゴシック", 42, 8, 18),
    ("MS P ゴシック", 43, 8, 18),
    ("MS P ゴシック", 44, 8, 19),
    ("MS P ゴシック", 48, 9, 20),
    ("MS P ゴシック", 50, 9, 22),
    ("MS P ゴシック", 52, 9, 22),
    ("MS P ゴシック", 54, 10, 23),
    ("MS P ゴシック", 56, 11, 24),
    ("MS P ゴシック", 60, 11, 25),
    ("MS P ゴシック", 64, 12, 27),
    ("MS P ゴシック", 68, 13, 29),
    ("MS P ゴシック", 72, 13, 31),
    ("MS P ゴシック", 80, 15, 34),
    ("MS P ゴシック", 88, 16, 36),
    ("MS P ゴシック", 96, 18, 40),
    ("MS P ゴシック", 112, 21, 46),
    ("MS P ゴシック", 144, 27, 59),
    ("MS PGothic", 24, 5, 11),
    ("MS PGothic", 28, 5, 12),
    ("MS PGothic", 30, 6, 13),
    ("MS PGothic", 32, 6, 14),
    ("MS PGothic", 36, 6, 15),
    ("MS PGothic", 37, 6, 15),
    ("MS PGothic", 40, 7, 16),
    ("MS PGothic", 42, 7, 17),
    ("MS PGothic", 43, 7, 17),
    ("MS PGothic", 44, 8, 18),
    ("MS PGothic", 48, 8, 19),
    ("MS PGothic", 50, 9, 20),
    ("MS PGothic", 52, 9, 20),
    ("MS PGothic", 54, 9, 22),
    ("MS PGothic", 56, 10, 23),
    ("MS PGothic", 60, 10, 24),
    ("MS PGothic", 64, 11, 25),
    ("MS PGothic", 68, 12, 27),
    ("MS PGothic", 72, 12, 28),
    ("MS PGothic", 80, 14, 32),
    ("MS PGothic", 88, 15, 34),
    ("MS PGothic", 96, 16, 38),
    ("MS PGothic", 112, 19, 43),
    ("MS PGothic", 144, 24, 56),
    ("MS PMincho", 24, 5, 11),
    ("MS PMincho", 28, 5, 12),
    ("MS PMincho", 30, 6, 13),
    ("MS PMincho", 32, 6, 14),
    ("MS PMincho", 34, 6, 14),
    ("MS PMincho", 36, 6, 15),
    ("MS PMincho", 40, 7, 16),
    ("MS PMincho", 42, 7, 17),
    ("MS PMincho", 44, 8, 18),
    ("MS PMincho", 48, 8, 19),
    ("MS PMincho", 52, 9, 20),
    ("MS PMincho", 56, 10, 23),
    ("MS PMincho", 64, 11, 25),
    ("MS PMincho", 72, 12, 28),
    ("MS PMincho", 76, 13, 30),
    ("MS PMincho", 80, 14, 32),
    ("MS PMincho", 88, 15, 34),
    ("MS PMincho", 96, 16, 38),
    ("MS PMincho", 144, 24, 56),
    ("MS Pゴシック", 24, 5, 11),
    ("MS Pゴシック", 28, 5, 12),
    ("MS Pゴシック", 30, 6, 13),
    ("MS Pゴシック", 32, 6, 15),
    ("MS Pゴシック", 36, 7, 16),
    ("MS Pゴシック", 37, 7, 16),
    ("MS Pゴシック", 40, 7, 17),
    ("MS Pゴシック", 42, 8, 18),
    ("MS Pゴシック", 43, 8, 18),
    ("MS Pゴシック", 44, 8, 19),
    ("MS Pゴシック", 48, 9, 20),
    ("MS Pゴシック", 50, 9, 22),
    ("MS Pゴシック", 52, 9, 22),
    ("MS Pゴシック", 54, 10, 23),
    ("MS Pゴシック", 56, 11, 24),
    ("MS Pゴシック", 60, 11, 25),
    ("MS Pゴシック", 64, 12, 27),
    ("MS Pゴシック", 68, 13, 29),
    ("MS Pゴシック", 72, 13, 31),
    ("MS Pゴシック", 80, 15, 34),
    ("MS Pゴシック", 88, 16, 36),
    ("MS Pゴシック", 96, 18, 40),
    ("MS Pゴシック", 112, 21, 46),
    ("MS Pゴシック", 144, 27, 59),
    ("MS UI Gothic", 24, 5, 11),
    ("MS UI Gothic", 32, 6, 14),
    ("MS UI Gothic", 36, 6, 15),
    ("MS UI Gothic", 40, 7, 16),
    ("MS UI Gothic", 42, 7, 17),
    ("MS UI Gothic", 44, 8, 18),
    ("MS UI Gothic", 48, 8, 19),
    ("MS UI Gothic", 56, 10, 23),
    ("MS UI Gothic", 64, 11, 25),
    ("MS UI Gothic", 72, 12, 28),
    ("MS UI Gothic", 80, 14, 32),
    ("MS UI Gothic", 96, 16, 38),
    ("MS UI Gothic", 144, 24, 56),
    ("MS ゴシック", 24, 5, 11),
    ("MS ゴシック", 28, 5, 12),
    ("MS ゴシック", 32, 6, 15),
    ("MS ゴシック", 36, 7, 16),
    ("MS ゴシック", 40, 7, 17),
    ("MS ゴシック", 42, 8, 18),
    ("MS ゴシック", 44, 8, 19),
    ("MS ゴシック", 46, 8, 19),
    ("MS ゴシック", 47, 9, 20),
    ("MS ゴシック", 48, 9, 20),
    ("MS ゴシック", 50, 9, 22),
    ("MS ゴシック", 52, 9, 22),
    ("MS ゴシック", 54, 10, 23),
    ("MS ゴシック", 55, 10, 23),
    ("MS ゴシック", 56, 11, 24),
    ("MS ゴシック", 58, 11, 24),
    ("MS ゴシック", 59, 11, 25),
    ("MS ゴシック", 60, 11, 25),
    ("MS ゴシック", 63, 12, 27),
    ("MS ゴシック", 64, 12, 27),
    ("MS ゴシック", 70, 13, 29),
    ("MS ゴシック", 72, 13, 31),
    ("MS ゴシック", 74, 14, 31),
    ("MS ゴシック", 80, 15, 34),
    ("MS ゴシック", 88, 16, 36),
    ("MS ゴシック", 96, 18, 40),
    ("MS ゴシック", 112, 21, 46),
    ("MS ゴシック", 144, 27, 59),
    ("MS ゴシック", 192, 36, 79),
    ("MS　Pゴシック", 44, 8, 19),
    ("Meiryo", 24, 5, 11),
    ("Meiryo", 32, 7, 19),
    ("Meiryo", 36, 7, 20),
    ("Meiryo", 40, 8, 22),
    ("Meiryo", 42, 9, 23),
    ("Meiryo", 44, 9, 25),
    ("Meiryo", 52, 11, 27),
    ("Meiryo", 56, 12, 30),
    ("Meiryo", 60, 12, 32),
    ("Meiryo", 64, 13, 33),
    ("Meiryo", 72, 15, 38),
    ("Meiryo", 80, 17, 44),
    ("Meiryo", 96, 20, 51),
    ("Meiryo", 112, 23, 58),
    ("Meiryo", 144, 30, 76),
    ("Meiryo", 192, 40, 101),
    ("Meiryo", 288, 60, 151),
    ("Meiryo UI", 24, 5, 11),
    ("Meiryo UI", 32, 7, 16),
    ("Meiryo UI", 36, 7, 16),
    ("Meiryo UI", 40, 8, 19),
    ("Meiryo UI", 42, 9, 20),
    ("Meiryo UI", 44, 9, 21),
    ("Meiryo UI", 48, 10, 22),
    ("Meiryo UI", 52, 11, 24),
    ("Meiryo UI", 56, 12, 26),
    ("Meiryo UI", 60, 12, 27),
    ("Meiryo UI", 64, 13, 28),
    ("Meiryo UI", 72, 15, 32),
    ("Meiryo UI", 80, 17, 38),
    ("Meiryo UI", 96, 20, 44),
    ("Meiryo UI", 144, 30, 65),
    ("OxiMissingFont", 52, 9, 22),
    ("Segoe UI", 24, 5, 11),
    ("Segoe UI", 32, 6, 14),
    ("Segoe UI", 36, 6, 16),
    ("Segoe UI", 40, 7, 19),
    ("Segoe UI", 42, 8, 21),
    ("Segoe UI", 44, 8, 22),
    ("Segoe UI", 48, 9, 23),
    ("Segoe UI", 56, 10, 27),
    ("Segoe UI", 64, 11, 34),
    ("Segoe UI", 72, 13, 35),
    ("Segoe UI", 80, 15, 41),
    ("Segoe UI", 96, 17, 50),
    ("Segoe UI", 144, 26, 70),
    ("Segoe UI Symbol", 40, 7, 19),
    ("Tahoma", 24, 5, 11),
    ("Tahoma", 32, 6, 14),
    ("Tahoma", 36, 7, 15),
    ("Tahoma", 40, 7, 17),
    ("Tahoma", 42, 8, 18),
    ("Tahoma", 44, 8, 19),
    ("Tahoma", 48, 9, 20),
    ("Tahoma", 56, 10, 24),
    ("Tahoma", 64, 11, 26),
    ("Tahoma", 72, 13, 30),
    ("Tahoma", 80, 15, 34),
    ("Tahoma", 96, 17, 40),
    ("Tahoma", 144, 26, 59),
    ("Terminal", 28, 5, 12),
    ("Terminal", 40, 5, 15),
    ("Terminal", 48, 12, 19),
    ("Terminal", 56, 10, 21),
    ("Times New Roman", 24, 5, 11),
    ("Times New Roman", 32, 6, 15),
    ("Times New Roman", 34, 6, 15),
    ("Times New Roman", 36, 6, 16),
    ("Times New Roman", 40, 6, 17),
    ("Times New Roman", 42, 7, 18),
    ("Times New Roman", 44, 7, 20),
    ("Times New Roman", 48, 8, 21),
    ("Times New Roman", 52, 9, 22),
    ("Times New Roman", 56, 9, 25),
    ("Times New Roman", 60, 10, 26),
    ("Times New Roman", 64, 11, 27),
    ("Times New Roman", 72, 12, 31),
    ("Times New Roman", 76, 12, 32),
    ("Times New Roman", 80, 14, 35),
    ("Times New Roman", 88, 14, 37),
    ("Times New Roman", 96, 16, 41),
    ("Times New Roman", 112, 19, 47),
    ("Times New Roman", 144, 24, 61),
    ("Times New Roman", 192, 32, 82),
    ("Times New Roman", 288, 48, 122),
    ("Verdana", 24, 5, 11),
    ("Verdana", 32, 7, 14),
    ("Verdana", 36, 8, 15),
    ("Verdana", 40, 8, 17),
    ("Verdana", 42, 9, 18),
    ("Verdana", 44, 10, 19),
    ("Verdana", 48, 10, 20),
    ("Verdana", 56, 12, 24),
    ("Verdana", 64, 13, 26),
    ("Verdana", 72, 15, 30),
    ("Verdana", 80, 17, 33),
    ("Verdana", 96, 20, 39),
    ("Verdana", 144, 31, 61),
    ("Yu Gothic", 24, 5, 11),
    ("Yu Gothic", 28, 5, 12),
    ("Yu Gothic", 32, 6, 17),
    ("Yu Gothic", 36, 7, 21),
    ("Yu Gothic", 40, 7, 22),
    ("Yu Gothic", 42, 8, 23),
    ("Yu Gothic", 44, 8, 25),
    ("Yu Gothic", 48, 9, 26),
    ("Yu Gothic", 56, 11, 32),
    ("Yu Gothic", 60, 11, 33),
    ("Yu Gothic", 64, 12, 34),
    ("Yu Gothic", 72, 13, 40),
    ("Yu Gothic", 80, 15, 44),
    ("Yu Gothic", 88, 16, 47),
    ("Yu Gothic", 96, 18, 53),
    ("Yu Gothic", 104, 19, 57),
    ("Yu Gothic", 112, 21, 59),
    ("Yu Gothic", 144, 27, 78),
    ("Yu Gothic", 288, 53, 154),
    ("Yu Gothic UI", 24, 5, 11),
    ("Yu Gothic UI", 32, 6, 14),
    ("Yu Gothic UI", 36, 6, 16),
    ("Yu Gothic UI", 40, 7, 19),
    ("Yu Gothic UI", 42, 8, 21),
    ("Yu Gothic UI", 44, 8, 22),
    ("Yu Gothic UI", 48, 9, 23),
    ("Yu Gothic UI", 56, 10, 27),
    ("Yu Gothic UI", 64, 11, 34),
    ("Yu Gothic UI", 72, 13, 35),
    ("Yu Gothic UI", 80, 15, 41),
    ("Yu Gothic UI", 88, 16, 44),
    ("Yu Gothic UI", 96, 17, 50),
    ("Yu Gothic UI", 144, 26, 70),
    ("meiryo", 48, 10, 26),
    ("yu gothic", 52, 9, 27),
    ("メイリオ", 24, 5, 11),
    ("メイリオ", 32, 7, 19),
    ("メイリオ", 36, 7, 20),
    ("メイリオ", 40, 8, 22),
    ("メイリオ", 42, 9, 23),
    ("メイリオ", 44, 9, 25),
    ("メイリオ", 48, 10, 26),
    ("メイリオ", 52, 11, 27),
    ("メイリオ", 56, 12, 30),
    ("メイリオ", 60, 12, 32),
    ("メイリオ", 64, 13, 33),
    ("メイリオ", 72, 15, 38),
    ("メイリオ", 80, 17, 44),
    ("メイリオ", 96, 20, 51),
    ("メイリオ", 112, 23, 58),
    ("メイリオ", 144, 30, 76),
    ("メイリオ", 192, 40, 101),
    ("メイリオ", 288, 60, 151),
    ("明朝", 40, 7, 17),
    ("明朝", 44, 8, 19),
    ("明朝", 72, 13, 31),
    ("標準ゴシック", 28, 5, 12),
    ("標準ゴシック", 32, 6, 14),
    ("標準ゴシック", 36, 6, 15),
    ("標準ゴシック", 40, 7, 16),
    ("標準ゴシック", 42, 7, 17),
    ("標準ゴシック", 44, 8, 18),
    ("游ゴシック", 24, 5, 11),
    ("游ゴシック", 28, 5, 12),
    ("游ゴシック", 32, 6, 17),
    ("游ゴシック", 36, 7, 21),
    ("游ゴシック", 40, 7, 22),
    ("游ゴシック", 42, 8, 23),
    ("游ゴシック", 44, 8, 25),
    ("游ゴシック", 48, 9, 26),
    ("游ゴシック", 52, 9, 27),
    ("游ゴシック", 56, 11, 32),
    ("游ゴシック", 60, 11, 33),
    ("游ゴシック", 64, 12, 34),
    ("游ゴシック", 72, 13, 40),
    ("游ゴシック", 80, 15, 44),
    ("游ゴシック", 88, 16, 47),
    ("游ゴシック", 96, 18, 53),
    ("游ゴシック", 104, 19, 57),
    ("游ゴシック", 112, 21, 59),
    ("游ゴシック", 144, 27, 78),
    ("游ゴシック", 288, 53, 154),
    ("游ゴシック Light", 72, 13, 40),
    ("游明朝", 24, 5, 11),
    ("游明朝", 32, 6, 17),
    ("游明朝", 36, 7, 21),
    ("游明朝", 40, 7, 22),
    ("游明朝", 42, 8, 23),
    ("游明朝", 44, 8, 24),
    ("游明朝", 48, 9, 26),
    ("游明朝", 56, 10, 32),
    ("游明朝", 64, 11, 34),
    ("游明朝", 72, 13, 40),
    ("游明朝", 80, 15, 44),
    ("游明朝", 96, 17, 52),
    ("游明朝", 144, 26, 77),
    ("ＭＳ ゴシック", 24, 5, 11),
    ("ＭＳ ゴシック", 28, 5, 12),
    ("ＭＳ ゴシック", 32, 6, 14),
    ("ＭＳ ゴシック", 36, 6, 15),
    ("ＭＳ ゴシック", 40, 7, 16),
    ("ＭＳ ゴシック", 42, 7, 17),
    ("ＭＳ ゴシック", 44, 8, 18),
    ("ＭＳ ゴシック", 46, 8, 18),
    ("ＭＳ ゴシック", 47, 8, 19),
    ("ＭＳ ゴシック", 48, 8, 19),
    ("ＭＳ ゴシック", 50, 9, 20),
    ("ＭＳ ゴシック", 52, 9, 20),
    ("ＭＳ ゴシック", 54, 9, 22),
    ("ＭＳ ゴシック", 55, 9, 22),
    ("ＭＳ ゴシック", 56, 10, 23),
    ("ＭＳ ゴシック", 58, 10, 23),
    ("ＭＳ ゴシック", 59, 10, 24),
    ("ＭＳ ゴシック", 60, 10, 24),
    ("ＭＳ ゴシック", 63, 11, 25),
    ("ＭＳ ゴシック", 64, 11, 25),
    ("ＭＳ ゴシック", 70, 12, 27),
    ("ＭＳ ゴシック", 72, 12, 28),
    ("ＭＳ ゴシック", 74, 13, 30),
    ("ＭＳ ゴシック", 80, 14, 32),
    ("ＭＳ ゴシック", 88, 15, 34),
    ("ＭＳ ゴシック", 96, 16, 38),
    ("ＭＳ ゴシック", 112, 19, 43),
    ("ＭＳ ゴシック", 144, 24, 56),
    ("ＭＳ ゴシック", 192, 32, 74),
    ("ＭＳ 明朝", 12, 3, 8),
    ("ＭＳ 明朝", 24, 5, 11),
    ("ＭＳ 明朝", 28, 5, 12),
    ("ＭＳ 明朝", 30, 6, 13),
    ("ＭＳ 明朝", 32, 6, 14),
    ("ＭＳ 明朝", 36, 6, 15),
    ("ＭＳ 明朝", 40, 7, 16),
    ("ＭＳ 明朝", 42, 7, 17),
    ("ＭＳ 明朝", 44, 8, 18),
    ("ＭＳ 明朝", 46, 8, 18),
    ("ＭＳ 明朝", 47, 8, 19),
    ("ＭＳ 明朝", 48, 8, 19),
    ("ＭＳ 明朝", 50, 9, 20),
    ("ＭＳ 明朝", 52, 9, 20),
    ("ＭＳ 明朝", 54, 9, 22),
    ("ＭＳ 明朝", 55, 9, 22),
    ("ＭＳ 明朝", 56, 10, 23),
    ("ＭＳ 明朝", 60, 10, 24),
    ("ＭＳ 明朝", 63, 11, 25),
    ("ＭＳ 明朝", 64, 11, 25),
    ("ＭＳ 明朝", 68, 12, 27),
    ("ＭＳ 明朝", 70, 12, 27),
    ("ＭＳ 明朝", 72, 12, 28),
    ("ＭＳ 明朝", 74, 13, 30),
    ("ＭＳ 明朝", 76, 13, 30),
    ("ＭＳ 明朝", 80, 14, 32),
    ("ＭＳ 明朝", 88, 15, 34),
    ("ＭＳ 明朝", 96, 16, 38),
    ("ＭＳ 明朝", 112, 19, 43),
    ("ＭＳ 明朝", 144, 24, 56),
    ("ＭＳ Ｐゴシック", 24, 5, 11),
    ("ＭＳ Ｐゴシック", 28, 5, 12),
    ("ＭＳ Ｐゴシック", 30, 6, 13),
    ("ＭＳ Ｐゴシック", 32, 6, 14),
    ("ＭＳ Ｐゴシック", 36, 6, 15),
    ("ＭＳ Ｐゴシック", 37, 6, 15),
    ("ＭＳ Ｐゴシック", 40, 7, 16),
    ("ＭＳ Ｐゴシック", 42, 7, 17),
    ("ＭＳ Ｐゴシック", 43, 7, 17),
    ("ＭＳ Ｐゴシック", 44, 8, 18),
    ("ＭＳ Ｐゴシック", 48, 8, 19),
    ("ＭＳ Ｐゴシック", 50, 9, 20),
    ("ＭＳ Ｐゴシック", 52, 9, 20),
    ("ＭＳ Ｐゴシック", 54, 9, 22),
    ("ＭＳ Ｐゴシック", 56, 10, 23),
    ("ＭＳ Ｐゴシック", 60, 10, 24),
    ("ＭＳ Ｐゴシック", 64, 11, 25),
    ("ＭＳ Ｐゴシック", 68, 12, 27),
    ("ＭＳ Ｐゴシック", 72, 12, 28),
    ("ＭＳ Ｐゴシック", 80, 14, 32),
    ("ＭＳ Ｐゴシック", 88, 15, 34),
    ("ＭＳ Ｐゴシック", 96, 16, 38),
    ("ＭＳ Ｐゴシック", 112, 19, 43),
    ("ＭＳ Ｐゴシック", 144, 24, 56),
    ("ＭＳ Ｐ明朝", 24, 5, 11),
    ("ＭＳ Ｐ明朝", 28, 5, 12),
    ("ＭＳ Ｐ明朝", 30, 6, 13),
    ("ＭＳ Ｐ明朝", 32, 6, 14),
    ("ＭＳ Ｐ明朝", 34, 6, 14),
    ("ＭＳ Ｐ明朝", 36, 6, 15),
    ("ＭＳ Ｐ明朝", 40, 7, 16),
    ("ＭＳ Ｐ明朝", 42, 7, 17),
    ("ＭＳ Ｐ明朝", 44, 8, 18),
    ("ＭＳ Ｐ明朝", 48, 8, 19),
    ("ＭＳ Ｐ明朝", 52, 9, 20),
    ("ＭＳ Ｐ明朝", 56, 10, 23),
    ("ＭＳ Ｐ明朝", 64, 11, 25),
    ("ＭＳ Ｐ明朝", 72, 12, 28),
    ("ＭＳ Ｐ明朝", 76, 13, 30),
    ("ＭＳ Ｐ明朝", 80, 14, 32),
    ("ＭＳ Ｐ明朝", 88, 15, 34),
    ("ＭＳ Ｐ明朝", 96, 16, 38),
    ("ＭＳ Ｐ明朝", 144, 24, 56),
    ("Aptos Narrow", 128, 22, 56),
    ("Calibri", 128, 22, 56),
];

/// Excel's column digit unit and untouched row height, in 96-dpi pixels.
pub fn font_default_geometry_px(face: &str, size: f32) -> Option<(u16, u16)> {
    let size_q = (size * 4.0).round() as u16;
    FONT_DEFAULT_GEOMETRY_PX.iter()
        .find(|(f,s,_,_)| f.eq_ignore_ascii_case(face) && *s == size_q)
        .or_else(|| FONT_DEFAULT_GEOMETRY_PX.iter().find(|(f,s,_,_)|
            f.eq_ignore_ascii_case(face) && em_pixels(*s as f32 / 4.0) == em_pixels(size)))
        .map(|(_,_,digit,row)| (*digit,*row))
}

/// Measured default geometry for a styled Normal font. Unmeasured styles
/// retain the regular-face geometry rather than inventing weight metrics.
pub fn font_styled_default_geometry_px(face: &str, size: f32, bold: bool, italic: bool) -> Option<(u16, u16)> {
    const STYLED: &[(&str, u16, bool, bool, u16, u16)] = &[
        ("Yu Gothic", 104, true, false, 20, 56),
        ("游ゴシック", 104, true, false, 20, 56),
    ];
    let size_q = (size * 4.0).round() as u16;
    STYLED.iter().find(|(held, points, b, i, _, _)|
        held.eq_ignore_ascii_case(face) && *points == size_q && *b == bold && *i == italic)
        .map(|(_, _, _, _, digit, height)| (*digit, *height))
        .or_else(|| font_default_geometry_px(face, size))
}

// Requested names remain in the source model. An unavailable font has its
// own default geometry; the physical fallback face is not a replacement for it.
static UNAVAILABLE_FONT_GEOMETRY_PX: &[(u16,u16,u16)] = &[
    (4, 2, 7),
    (12, 3, 8),
    (16, 3, 9),
    (24, 5, 11),
    (28, 5, 12),
    (30, 6, 13),
    (32, 6, 15),
    (34, 6, 15),
    (36, 7, 16),
    (37, 7, 16),
    (38, 7, 17),
    (40, 7, 17),
    (42, 8, 18),
    (43, 8, 18),
    (44, 8, 19),
    (46, 8, 19),
    (47, 9, 20),
    (48, 9, 20),
    (50, 9, 22),
    (52, 9, 22),
    (54, 10, 23),
    (55, 10, 23),
    (56, 11, 24),
    (58, 11, 24),
    (59, 11, 25),
    (60, 11, 25),
    (63, 12, 27),
    (64, 12, 27),
    (68, 13, 29),
    (70, 13, 29),
    (72, 13, 31),
    (74, 14, 31),
    (76, 14, 31),
    (80, 15, 34),
    (88, 16, 36),
    (96, 18, 40),
    (112, 21, 46),
    (128, 24, 54),
    (144, 27, 59),
    (192, 36, 79),
    (200, 37, 83),
    (208, 38, 85),
    (288, 53, 120),
    (320, 60, 133),
    (480, 89, 199),
    (1636, 303, 546)
];

/// Default geometry when the requested font is unavailable.
pub fn unavailable_font_geometry_px(size: f32) -> Option<(u16,u16)> {
    let q = (size*4.0).round() as u16;
    UNAVAILABLE_FONT_GEOMETRY_PX.iter().find(|(s,_,_)| *s==q)
        .or_else(|| UNAVAILABLE_FONT_GEOMETRY_PX.iter().find(|(s,_,_)|
            em_pixels(*s as f32/4.0)==em_pixels(size)))
        .map(|(_,digit,row)| (*digit,*row))
}
