from __future__ import annotations

import sys
import unittest
from pathlib import Path

from docx import Document

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))

from converter import OcrLine, OcrWord, PdfToWordConverter  # noqa: E402


class FakeTable:
    cells = [(0.0, 0.0, 60.0, 20.0), (60.0, 0.0, 180.0, 20.0)]


class ConverterTests(unittest.TestCase):
    def test_normalize_chinese_ocr_spacing(self) -> None:
        self.assertEqual(PdfToWordConverter._normalize_text("\u667a \u80fd \u5236 \u9020 \uff0c \u5de5 \u7a0b"), "\u667a\u80fd\u5236\u9020\uff0c\u5de5\u7a0b")
        self.assertEqual(PdfToWordConverter._normalize_text("PDF \u8f6c Word"), "PDF \u8f6c Word")

    def test_cell_indices_selects_source_grid_range(self) -> None:
        edges = [0.0, 60.0, 180.0]
        self.assertEqual(PdfToWordConverter._cell_indices(edges, 60.0, 180.0), (1, 1))
        self.assertEqual(PdfToWordConverter._cell_indices(edges, 0.0, 180.0), (0, 1))

    def test_append_table_writes_editable_cell_text(self) -> None:
        document = Document()
        lines = (
            OcrLine("\u59d3\u540d", 10.0, 5.0, 20.0, 8.0, ()),
            OcrLine("\u66fe\u6587\u745e", 75.0, 5.0, 30.0, 8.0, ()),
        )
        PdfToWordConverter()._append_table(document, FakeTable(), lines, 1.0, 1.0)
        self.assertEqual(len(document.tables), 1)
        self.assertEqual(document.tables[0].cell(0, 0).text, "\u59d3\u540d")
        self.assertEqual(document.tables[0].cell(0, 1).text, "\u66fe\u6587\u745e")


    def test_words_split_across_adjacent_cells_by_coordinates(self) -> None:
        line = OcrLine("\u6027\u522b\u7537", 10.0, 5.0, 100.0, 8.0, (
            OcrWord("\u6027", 10.0, 5.0, 12.0, 8.0),
            OcrWord("\u522b", 25.0, 5.0, 12.0, 8.0),
            OcrWord("\u7537", 75.0, 5.0, 12.0, 8.0),
        ))
        converter = PdfToWordConverter()
        self.assertEqual(converter._text_in_box((line,), (0.0, 0.0, 60.0, 20.0), 1.0, 1.0), "\u6027\u522b")
        self.assertEqual(converter._text_in_box((line,), (60.0, 0.0, 120.0, 20.0), 1.0, 1.0), "\u7537")

    def test_normalize_numbered_full_width_parentheses(self) -> None:
        self.assertEqual(PdfToWordConverter._normalize_text("\uff08 2 \uff09"), "\uff082\uff09")

    def test_normalize_recovers_misread_list_marker(self) -> None:
        source = (
            "\uff081\uff09\u5bf9\u7956\u56fd\u7684\u6210\u5c31\u548c\u6587\u5316\u611f\u5230\u81ea\u8c6a\uff0c\u62e5\u6709\u6c11\u65cf\u81ea\u5c0a\u5fc3\u548c\u6c11\u65cf\u81ea\u4fe1\u5fc3"
            "\uff082\uff09\u7ef4\u62a4\u56fd\u5bb6\u5229\u76ca\uff0c\u5fd7\u613f\u4e3a\u56fd\u5bb6\u548c\u793e\u4f1a\u670d\u52a1\u3002 0 \uff09\u5173\u5fc3\u56fd\u5bb6\u5927\u4e8b"
        )
        expected = (
            "\uff081\uff09\u5bf9\u7956\u56fd\u7684\u6210\u5c31\u548c\u6587\u5316\u611f\u5230\u81ea\u8c6a\uff0c\u62e5\u6709\u6c11\u65cf\u81ea\u5c0a\u5fc3\u548c\u6c11\u65cf\u81ea\u4fe1\u5fc3"
            "\uff082\uff09\u7ef4\u62a4\u56fd\u5bb6\u5229\u76ca\uff0c\u5fd7\u613f\u4e3a\u56fd\u5bb6\u548c\u793e\u4f1a\u670d\u52a1\u3002\uff083\uff09\u5173\u5fc3\u56fd\u5bb6\u5927\u4e8b"
        )
        self.assertEqual(PdfToWordConverter._normalize_text(source), expected)

    def test_normalize_keeps_unrelated_zero_parenthesis(self) -> None:
        source = "\u9879\u76ee\u5b8c\u6210\u7387\u4e3a 0 \uff09\u4e0d\u7eb3\u5165\u5217\u8868"
        self.assertEqual(PdfToWordConverter._normalize_text(source), source)

    def test_punctuation_on_a_lower_baseline_stays_with_its_text_row(self) -> None:
        line = OcrLine("", 10.0, 10.0, 80.0, 20.0, (
            OcrWord("\u5b8c", 10.0, 10.0, 12.0, 20.0),
            OcrWord("\u6210", 27.0, 10.0, 12.0, 20.0),
            OcrWord("\uff0c", 43.0, 25.0, 4.0, 5.0),
            OcrWord("\u7ee7", 55.0, 10.0, 12.0, 20.0),
            OcrWord("\u7eed", 72.0, 10.0, 12.0, 20.0),
        ))
        actual = PdfToWordConverter()._text_in_box((line,), (0.0, 0.0, 100.0, 40.0), 1.0, 1.0)
        self.assertEqual(actual, "\u5b8c\u6210\uff0c\u7ee7\u7eed")

    def test_append_table_preserves_uneven_grid_and_row_heights(self) -> None:
        class UnevenTable:
            cells = [
                (84.6, 147.44, 167.9, 182.39),
                (167.9, 147.44, 239.9, 182.39),
                (239.9, 147.44, 539.85, 182.39),
                (84.6, 182.39, 167.9, 420.79),
                (167.9, 182.39, 239.9, 420.79),
                (239.9, 182.39, 539.85, 420.79),
            ]

        document = Document()
        PdfToWordConverter()._append_table(document, UnevenTable(), (), 1.0, 1.0)
        result = document.tables[0]
        widths = [column.w.twips for column in result._tbl.tblGrid.gridCol_lst]

        self.assertGreater(widths[2], widths[0] * 3)
        self.assertGreater(widths[0], widths[1])
        self.assertIsNotNone(result.rows[0].height)
        self.assertGreater(result.rows[1].height.twips, result.rows[0].height.twips * 4)

    def test_new_document_preserves_source_page_size(self) -> None:
        document = PdfToWordConverter._new_document(595.3, 841.9)
        section = document.sections[0]
        self.assertAlmostEqual(section.page_width.pt, 595.3, places=0)
        self.assertAlmostEqual(section.page_height.pt, 841.9, places=0)

if __name__ == "__main__":
    unittest.main()
