// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

use crate::ir::{Image, WrapType};

/// A wrapping outline in coordinates relative to the cell's text origin.
#[derive(Clone, Debug)]
pub(super) struct Obstacle {
    points: Vec<(f32, f32)>,
    left_distance: f32,
    right_distance: f32,
    full_width: bool,
    pub bottom: f32,
}

impl Obstacle {
    pub fn from_image(image: &Image, x: f32, y: f32) -> Option<Self> {
        let position = image.position.as_ref()?;
        let kind = image.wrap_type.as_ref()?;
        if matches!(kind, WrapType::None) {
            return None;
        }
        let points = if matches!(kind, WrapType::Tight) && image.wrap_polygon.len() >= 3 {
            image.wrap_polygon.iter().map(|&(px, py)| {
                (x + px * image.width, y + py * image.height)
            }).collect::<Vec<_>>()
        } else {
            let top = y - image.effect_extent_t.max(0.0);
            let bottom = y + image.height + image.effect_extent_b.max(0.0);
            vec![(x, top), (x + image.width, top), (x + image.width, bottom), (x, bottom)]
        };
        let bottom = points.iter().map(|p| p.1).fold(f32::NEG_INFINITY, f32::max);
        Some(Self {
            points,
            left_distance: position.dist_l.unwrap_or(0.0).max(0.0),
            right_distance: position.dist_r.unwrap_or(0.0).max(0.0),
            full_width: matches!(kind, WrapType::TopAndBottom),
            bottom,
        })
    }

    /// Intersect the whole line box, including edges that cross it between vertices.
    fn interval(&self, top: f32, bottom: f32) -> Option<(f32, f32)> {
        let outline_top = self.points.iter().map(|p| p.1).fold(f32::INFINITY, f32::min);
        if top >= self.bottom || bottom <= outline_top {
            return None;
        }
        if self.full_width {
            return Some((f32::NEG_INFINITY, f32::INFINITY));
        }
        let mut left = f32::INFINITY;
        let mut right = f32::NEG_INFINITY;
        for i in 0..self.points.len() {
            let (ax, ay) = self.points[i];
            let (bx, by) = self.points[(i + 1) % self.points.len()];
            if ay >= top && ay <= bottom {
                left = left.min(ax);
                right = right.max(ax);
            }
            if ay != by {
                for edge in [top, bottom] {
                    let t = (edge - ay) / (by - ay);
                    if (0.0..=1.0).contains(&t) {
                        let x = ax + t * (bx - ax);
                        left = left.min(x);
                        right = right.max(x);
                    }
                }
            }
        }
        (left <= right).then_some((left - self.left_distance, right + self.right_distance))
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) struct LineFrame {
    pub left: f32,
    pub width: f32,
    pub top: f32,
    pub gap: f32,
    pub advance: f32,
}

/// The same line boxes are queried by the height estimator and the text emitter.
#[derive(Clone, Debug, Default)]
pub(super) struct ParagraphWrap {
    pub obstacles: Vec<Obstacle>,
    pub top: f32,
    pub heights: Vec<f32>,
}

impl ParagraphWrap {
    pub fn frame(&self, index: usize, width: f32, first_width: f32,
        indent: f32, first_indent: f32) -> LineFrame {
        let mut top = self.top;
        let mut segment = 0;
        loop {
            // Positive first-line indentation follows the available text region
            // when a float moves its left edge. A hanging indent still extends
            // the paragraph's original left boundary.
            let inset = if segment == 0 { first_indent.max(0.0) } else { 0.0 };
            let left = if segment == 0 { (indent + first_indent.min(0.0)).max(0.0) } else { indent };
            let line_width = if segment == 0 { first_width + inset } else { width };
            let original_top = top;
            let height_at = |i| self.heights.get(i).or(self.heights.last()).copied().unwrap_or(0.0);
            let mut height = height_at(segment);
            let intervals = loop {
                let mut intervals = vec![(left, left + line_width)];
                let mut next_bottom = f32::INFINITY;
                for obstacle in &self.obstacles {
                    if let Some((a, b)) = obstacle.interval(top, top + height) {
                        if b <= left || a >= left + line_width { continue; }
                        next_bottom = next_bottom.min(obstacle.bottom);
                        intervals = intervals.into_iter().flat_map(|(l, r)| {
                            let mut pieces = Vec::with_capacity(2);
                            if a > l { pieces.push((l, a.min(r))); }
                            if b < r { pieces.push((b.max(l), r)); }
                            pieces
                        }).filter(|(l, r)| r > l).collect();
                    }
                }
                intervals.sort_by(|a, b| a.0.total_cmp(&b.0));
                if let Some(first) = intervals.first_mut() {
                    first.0 += inset;
                }
                intervals.retain(|(l, r)| r > l);
                if intervals.is_empty() {
                    if next_bottom.is_finite() && next_bottom > top {
                        top = next_bottom;
                        continue;
                    }
                    intervals.push((left + inset, (left + line_width).max(left + inset)));
                }
                intervals.sort_by(|a, b| a.0.total_cmp(&b.0));
                let row_height = (0..intervals.len()).map(|i| height_at(segment + i))
                    .fold(height, f32::max);
                if row_height > height {
                    height = row_height;
                    continue;
                }
                break intervals;
            };
            for (part, &(l, r)) in intervals.iter().enumerate() {
                if segment + part == index {
                    let last = part + 1 == intervals.len() || index + 1 >= self.heights.len();
                    return LineFrame { left: l, width: r - l, top,
                        gap: if part == 0 { top - original_top } else { 0.0 },
                        advance: if last { height } else { 0.0 } };
                }
            }
            segment += intervals.len();
            top += height;
        }
    }
}

#[derive(Default)]
pub(super) struct Measurement {
    pub wrap: Option<ParagraphWrap>,
    pub heights: Vec<f32>,
    pub geometry: (f32, f32, f32, f32),
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rectangle(left: f32, top: f32, right: f32, bottom: f32) -> Obstacle {
        Obstacle { points: vec![(left, top), (right, top), (right, bottom), (left, bottom)],
            left_distance: 0.0, right_distance: 0.0, full_width: false, bottom }
    }

    #[test]
    fn first_line_indent_follows_a_float_but_does_not_repeat_on_later_segments() {
        let wrap = ParagraphWrap { obstacles: vec![rectangle(0.0, 0.0, 30.0, 60.0)],
            top: 0.0, heights: vec![10.0, 10.0] };
        let first = wrap.frame(0, 100.0, 90.0, 0.0, 10.0);
        let next = wrap.frame(1, 100.0, 90.0, 0.0, 10.0);
        assert_eq!((first.left, first.width), (40.0, 60.0));
        assert_eq!((next.left, next.width), (30.0, 70.0));
        let middle = ParagraphWrap { obstacles: vec![rectangle(40.0, 0.0, 60.0, 60.0)], ..wrap };
        assert_eq!(middle.frame(0, 100.0, 90.0, 0.0, 10.0).left, 10.0);
        assert_eq!(middle.frame(1, 100.0, 90.0, 0.0, 10.0).left, 60.0);
    }

    #[test]
    fn text_on_both_sides_of_an_image_shares_one_line_height() {
        let wrap = ParagraphWrap { obstacles: vec![rectangle(40.0, 0.0, 60.0, 60.0)],
            top: 0.0, heights: vec![10.0, 15.0, 10.0] };
        let first = wrap.frame(0, 100.0, 100.0, 0.0, 0.0);
        let second = wrap.frame(1, 100.0, 100.0, 0.0, 0.0);
        let last = wrap.frame(2, 100.0, 100.0, 0.0, 0.0);
        assert_eq!((first.left, first.width, first.top, first.advance), (0.0, 40.0, 0.0, 0.0));
        assert_eq!((second.left, second.width, second.top, second.advance), (60.0, 40.0, 0.0, 15.0));
        assert_eq!((last.top, last.advance), (15.0, 10.0));
    }

    #[test]
    fn lines_intersect_by_height_and_stop_wrapping_at_outline_bottom() {
        let wrap = ParagraphWrap { obstacles: vec![rectangle(80.0, 15.0, 100.0, 40.0)],
            top: 0.0, heights: vec![10.0, 20.0, 10.0, 10.0] };
        assert_eq!(wrap.frame(0, 100.0, 100.0, 0.0, 0.0).width, 100.0);
        assert_eq!(wrap.frame(1, 100.0, 100.0, 0.0, 0.0).width, 80.0);
        assert_eq!(wrap.frame(2, 100.0, 100.0, 0.0, 0.0).width, 80.0);
        assert_eq!(wrap.frame(3, 100.0, 100.0, 0.0, 0.0).width, 100.0);
    }

    #[test]
    fn spanning_obstacles_delay_flow_without_adding_image_height_to_every_line() {
        let wrap = ParagraphWrap { obstacles: vec![rectangle(0.0, 5.0, 100.0, 35.0),
            rectangle(0.0, 30.0, 100.0, 50.0)], top: 0.0, heights: vec![10.0] };
        assert_eq!(wrap.frame(0, 100.0, 100.0, 0.0, 0.0).gap, 50.0);
        assert_eq!(wrap.frame(1, 100.0, 100.0, 0.0, 0.0).top, 60.0);
        assert_eq!(wrap.frame(1, 100.0, 100.0, 0.0, 0.0).gap, 0.0);
    }

    #[test]
    fn sloped_contour_and_two_obstacles_leave_the_correct_text_interval() {
        let triangle = Obstacle { points: vec![(0.0, 0.0), (60.0, 60.0), (0.0, 60.0)],
            left_distance: 0.0, right_distance: 3.0, full_width: false, bottom: 60.0 };
        let wrap = ParagraphWrap { obstacles: vec![triangle, rectangle(90.0, 0.0, 110.0, 60.0)],
            top: 10.0, heights: vec![10.0] };
        let frame = wrap.frame(0, 100.0, 100.0, 0.0, 0.0);
        assert_eq!((frame.left, frame.width), (23.0, 67.0));
    }
}
