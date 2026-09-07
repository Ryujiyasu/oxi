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


if __name__ == "__main__":
    unittest.main()
