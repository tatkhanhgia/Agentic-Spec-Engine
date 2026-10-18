import os
import sys
import tempfile
import unittest


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

import fix_docx_format as formatter
import check_frame_format as checker
from add_error_response import make_response_attributes_table


def grid_widths(table):
    return [
        int(col.get(formatter.qn("w")))
        for col in table.findall(
            formatter.qn("tblGrid") + "/" + formatter.qn("gridCol"))
    ]


def regular_row_widths(row):
    return [
        (
            int(cell.find(
                formatter.qn("tcPr") + "/" + formatter.qn("tcW")
            ).get(formatter.qn("w"))),
            cell.find(
                formatter.qn("tcPr") + "/" + formatter.qn("tcW")
            ).get(formatter.qn("type")),
        )
        for cell in row.findall(formatter.qn("tc"))
    ]


class GeneratedTableWidthTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        template = os.path.join(
            ROOT, "scripts", "GoPaperless_DOCX_Format_Template.md")
        cls.rules = formatter.parse_table_property_rules(template)
        cls.attribute_widths = formatter.parse_column_widths(
            cls.rules, "Generated attribute table column widths", 5)
        cls.connection_widths = formatter.parse_column_widths(
            cls.rules, "Generated connection table column widths", 4)

    def assert_regular_rows_match_grid(self, table, first_regular_row):
        expected = grid_widths(table)
        for row in table.findall(formatter.qn("tr"))[first_regular_row:]:
            self.assertEqual(
                [(width, "dxa") for width in expected],
                regular_row_widths(row),
            )

    def test_path_attributes_cells_match_five_column_grid(self):
        table = formatter.make_path_attributes_table(
            ["pageNo", "recordsPerPage"], self.attribute_widths)

        self.assertEqual(self.attribute_widths, grid_widths(table))
        self.assert_regular_rows_match_grid(table, 1)

        title_width = table.find(
            formatter.qn("tr") + "/" + formatter.qn("tc") + "/"
            + formatter.qn("tcPr") + "/" + formatter.qn("tcW"))
        self.assertEqual(("5000", "pct"), (
            title_width.get(formatter.qn("w")),
            title_width.get(formatter.qn("type")),
        ))

    def test_connection_table_has_four_matching_grid_columns(self):
        table = formatter.make_connection_info_table(
            {}, self.connection_widths)

        self.assertEqual(self.connection_widths, grid_widths(table))
        self.assert_regular_rows_match_grid(table, 0)

    def test_normalizer_repairs_legacy_connection_grid_and_cells(self):
        table = formatter.make_connection_info_table(
            {}, self.connection_widths)
        table.find(formatter.qn("tblGrid")).remove(
            table.findall(
                formatter.qn("tblGrid") + "/" + formatter.qn("gridCol"))[0])
        for row in table.findall(formatter.qn("tr")):
            for cell in row.findall(formatter.qn("tc")):
                tcW = cell.find(
                    formatter.qn("tcPr") + "/" + formatter.qn("tcW"))
                tcW.set(formatter.qn("w"), "1666")
                tcW.set(formatter.qn("type"), "pct")

        root = formatter.etree.Element("root")
        root.append(table)
        formatter.normalize_connection_table_widths(root, self.rules)

        self.assertEqual(self.connection_widths, grid_widths(table))
        self.assert_regular_rows_match_grid(table, 0)

    def test_response_attributes_caller_uses_explicit_widths(self):
        table = make_response_attributes_table(
            [{
                "name": "message",
                "type": "String",
                "presence": "M",
                "description": "Human-readable message.",
            }],
            self.attribute_widths,
        )

        self.assertEqual(self.attribute_widths, grid_widths(table))
        self.assert_regular_rows_match_grid(table, 1)

    def test_normalizer_repairs_legacy_percentage_path_cells(self):
        table = formatter.make_path_attributes_table(
            ["pageNo"], self.attribute_widths)
        for row in table.findall(formatter.qn("tr"))[1:]:
            for cell in row.findall(formatter.qn("tc")):
                tcW = cell.find(
                    formatter.qn("tcPr") + "/" + formatter.qn("tcW"))
                tcW.set(formatter.qn("w"), "1666")
                tcW.set(formatter.qn("type"), "pct")

        root = formatter.etree.Element("root")
        root.append(table)
        formatter.normalize_attribute_table_widths(root)

        self.assert_regular_rows_match_grid(table, 1)

    def test_packaging_gate_rejects_percentage_path_cells(self):
        table = formatter.make_path_attributes_table(
            ["pageNo"], self.attribute_widths)
        tcW = table.findall(formatter.qn("tr"))[1].findall(
            formatter.qn("tc"))[0].find(
                formatter.qn("tcPr") + "/" + formatter.qn("tcW"))
        tcW.set(formatter.qn("w"), "1666")
        tcW.set(formatter.qn("type"), "pct")
        root = formatter.etree.Element("root")
        root.append(table)

        with tempfile.TemporaryDirectory() as temp_dir:
            word_dir = os.path.join(temp_dir, "word")
            os.makedirs(word_dir)
            formatter.etree.ElementTree(root).write(
                os.path.join(word_dir, "document.xml"),
                xml_declaration=True,
                encoding="utf-8",
            )
            output = os.path.join(temp_dir, "invalid.docx")

            with self.assertRaisesRegex(
                    ValueError, "table-width validation failed"):
                formatter.pack_docx(temp_dir, output)
            self.assertFalse(os.path.exists(output))

    def test_check_h_catches_connection_header_width(self):
        table = formatter.make_connection_info_table(
            {}, self.connection_widths)
        header = table.findall(formatter.qn("tr"))[0]
        tcW = header.findall(formatter.qn("tc"))[0].find(
            formatter.qn("tcPr") + "/" + formatter.qn("tcW"))
        tcW.set(formatter.qn("w"), "1666")
        tcW.set(formatter.qn("type"), "pct")

        issue = checker.row_width_issue(
            header, checker.table_grid_widths(table), exact=True)
        self.assertIn("percentage widths are forbidden", issue)

    def test_template_width_count_is_validated(self):
        with self.assertRaises(ValueError):
            formatter.parse_column_widths(
                {"bad": "900, 1800 dxa"}, "bad", 5)


if __name__ == "__main__":
    unittest.main()
