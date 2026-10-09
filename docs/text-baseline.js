// Layout supplies an explicit baseline offset for participating horizontal runs.
// Null/absent offsets retain the existing placement of other text.
export function textBaseline(element, fontSize) {
  return element.y + (element.baseline_offset ?? fontSize * 0.85);
}
