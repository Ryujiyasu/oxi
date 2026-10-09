import copy
import unittest

from measure_pagination_oxi import aggregate_dump


class VerticalReadingOrderTests(unittest.TestCase):
    def aggregate(self, fragments):
        return aggregate_dump({"pages": [{"page": 1, "elements": [
            {"type": "text", "para_idx": 0, "vert": True,
             "x": x, "y": y, "text": text}
            for x, y, text in fragments
        ]}]})["1"]

    def test_indented_first_column_still_opens_paragraph(self):
        records = self.aggregate([(200, 60, "first"), (180, 50, "second")])
        self.assertEqual([r["text"] for r in records], ["firstsecond"])

    def test_runs_far_down_one_column_remain_one_paragraph(self):
        records = self.aggregate([(200, 50, "first"), (200, 150, "second"),
                                  (180, 50, "third")])
        self.assertEqual([r["text"] for r in records], ["firstsecondthird"])

    def test_next_band_follows_all_columns_of_previous_band(self):
        records = self.aggregate([(200, 50, "a"), (180, 50, "b"),
                                  (200, 300, "c"), (180, 300, "d")])
        self.assertEqual([r["text"] for r in records], ["abcd"])




class BodyReadingOrderTests(unittest.TestCase):
    def text(self, value, y, x=72, paragraph=1, cell=False):
        return dict(type="text", text=value, y=y, x=x, para_idx=paragraph,
                    cell_para_idx=0 if cell else None,
                    cell_row_idx=0 if cell else None,
                    cell_col_idx=0 if cell else None)

    def aggregate(self, elements):
        return aggregate_dump(dict(pages=[dict(page=1, elements=elements)]))["1"]

    def test_body_column_continuation_preserves_emitted_reading_order(self):
        records = self.aggregate([
            self.text("First ", 72), self.text("then ", 140),
            self.text("second column.", 72, x=320),
        ])
        self.assertEqual([r["text"] for r in records], ["First then second column."])

    def test_repeated_text_with_distinct_body_identity_stays_distinct(self):
        records = self.aggregate([
            self.text("Repeated paragraph.", 72, paragraph=1),
            self.text("Repeated paragraph.", 150, paragraph=2),
        ])
        self.assertEqual(len(records), 2)
        self.assertEqual({r["para_idx"] for r in records}, {1, 2})

    def test_reused_nested_cell_identity_retains_separate_instances(self):
        records = self.aggregate([
            self.text("First nested cell.", 72, cell=True),
            self.text("Second nested cell.", 200, cell=True),
        ])
        self.assertEqual(len(records), 2)

    def test_genuine_page_shift_still_fails_pagination_gate(self):
        from pagination_diff import diff_doc
        value = "A uniquely identifiable paragraph."
        pages = aggregate_dump(dict(pages=[dict(page=1, elements=[]),
            dict(page=2, elements=[self.text(value, 72)])]))
        word = dict(n_pages=2, paragraphs=[dict(i=1, page=1,
            text=value[:30], in_table=False)])
        result = diff_doc("synthetic_shift", word, dict(n_pages=2, pages=pages))
        self.assertEqual(result["n_matched"], 1)
        self.assertEqual(result["delta_histogram"], {1: 1})
        self.assertFalse(result["pass"])


class ParagraphStartProjectionTests(unittest.TestCase):
    def layout(self, text="A paragraph whose source begins at an object"):
        prefix = "\f" + text
        shared = dict(type="text", x=72.0, w=0.0, h=14.0, font_size=12.0,
                      para_idx=3, cell_para_idx=None, cell_row_idx=None, cell_col_idx=None,
                      source_container_idx=0, source_paragraph_prefix=prefix,
                      source_paragraph_chars=len(prefix), source_paragraph_controls=1,
                      source_paragraph_column_controls=0)
        caret = dict(shared, y=500.0, text="", source_boundary_attachment=True)
        painted = dict(shared, y=72.0, text=text, w=240.0)
        return dict(pages=[dict(page=1, elements=[caret]), dict(page=2, elements=[painted])])

    def test_source_start_is_projected_without_moving_painted_elements(self):
        layout = self.layout()
        original = copy.deepcopy(layout)
        records = aggregate_dump(layout)
        self.assertEqual(records["1"][0]["text"], "A paragraph whose source begin")
        self.assertEqual(records["1"][0]["y"], 500.0)
        self.assertEqual(records["2"], [])
        self.assertEqual(layout, original)

    def test_short_paragraph_keeps_existing_word_collector_convention(self):
        layout = self.layout("Short title")
        records = aggregate_dump(layout)
        self.assertEqual(records["1"][0]["text"], "")
        self.assertEqual(records["2"][0]["text"], "Short title")

    def test_projection_uses_explicit_page_identity(self):
        layout = self.layout()
        layout["pages"][0]["page"] = 27
        layout["pages"][1]["page"] = 28
        records = aggregate_dump(layout)
        self.assertTrue(records["27"][0]["text"])
        self.assertEqual(records["28"], [])

    def test_repeated_local_index_in_another_container_does_not_fold(self):
        layout = self.layout()
        layout["pages"][1]["elements"][0]["source_container_idx"] = 1
        records = aggregate_dump(layout)
        self.assertEqual(records["1"][0]["text"], "")
        self.assertTrue(records["2"][0]["text"])

    def test_old_dumps_and_unattached_controls_keep_their_existing_projection(self):
        for omitted in ("source_container_idx", "source_boundary_attachment"):
            with self.subTest(omitted=omitted):
                layout = self.layout()
                for page in layout["pages"]:
                    for element in page["elements"]:
                        element.pop(omitted, None)
                records = aggregate_dump(layout)
                self.assertEqual(records["1"][0]["text"], "")
                self.assertTrue(records["2"][0]["text"])

if __name__ == "__main__":
    unittest.main()
