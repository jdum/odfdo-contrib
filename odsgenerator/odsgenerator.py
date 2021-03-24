#!/usr/bin/env python
# Copyright 2021 Jérôme Dumonteil
# Licence: MIT
# Authors: jerome.dumonteil@gmail.com
"""
Quick .ODS generator from json file, using odfdo library
"""

import sys
import json
from odfdo import Document, Table, Row, Cell, Element

DEFAULT_STYLES = [
    {
        "name": "default_table_row",
        "style": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="4.52mm"
            fo:break-before="auto" style:use-optimal-row-height="true"/>
            </style:style>
        """,
    },
    {
        "name": "table_row_1cm",
        "style": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="1cm"
            fo:break-before="auto"/>
            </style:style>
        """,
    },
    {
        "name": "bold",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:text-properties fo:font-weight="bold"
            style:font-weight-asian="bold" style:font-weight-complex="bold"/>
            <style:table-cell-properties style:text-align-source="value-type"/>
            <style:paragraph-properties
            fo:margin-right="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "left",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="start"
            fo:margin-left="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "right",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="end"
            fo:margin-right="1mm"/>
            </style:style>
        """,
    },
    {
        "name": "center",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="center"/>
            </style:style>
        """,
    },
    {
        "name": "bold_left_bg_gray_grid06",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties fo:font-weight="bold"
            style:font-weight-asian="bold" style:font-weight-complex="bold"
            fo:background-color="#dddddd" fo:border="0.06pt solid #000000"
            style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="start"
            fo:margin-left="1.2mm"/>
            </style:style>
        """,
    },
    {
        "name": "grid06",
        "style": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties
            fo:border="0.06pt solid #000000"/>
            <style:paragraph-properties
            fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
            </style:style>
        """,
    },
]

DEFAULTS_DICT = {
    "style_table_row": "default_table_row",
    "style_table_cell": "",
    "style_str": "left",
    "style_int": "right",
    "style_float": "lpod-default-number-style",
    "style_other": "left",
}

BODY = "body"
TABLE = "table"
ROW = "row"
VALUE = "value"
NAME = "name"
WIDTH = "width"
STYLE = "style"
STYLES = "styles"
DEFAULTS = "defaults"
DEFAULT_TAB_PREFIX = "Tab"


class ODSGenerator:
    def __init__(self, content):
        self.doc = Document("spreadsheet")
        self.doc.body.clear()
        self.tab_counter = 0
        self.defaults = DEFAULTS_DICT
        self.styles_elements = {}
        self.used_styles = set()
        self.parse(content)

    def save(self, path):
        self.doc.save(path)

    def parse_styles(self, opt):
        for s in DEFAULT_STYLES:
            style = Element.from_tag(s["style"])
            style.name = s["name"]
            self.styles_elements[s["name"]] = style
        styles = opt.get(STYLES, [])
        for s in styles:
            name = s.get("name")
            definition = s.get("style")
            style = Element.from_tag(definition)
            style.name = name
            self.styles_elements[name] = style

    def insert_style(self, name):
        if name and name not in self.used_styles and name in self.styles_elements:
            style = self.styles_elements[name]
            self.doc.insert_style(style, automatic=True)
            self.used_styles.add(name)

    def guess_style(self, opt, family, default):
        style_list = opt.get(STYLE, [])
        if not isinstance(style_list, list):
            style_list = [style_list]
        for style_name in style_list:
            if style_name:
                style = self.styles_elements.get(style_name)
                if style and style.family == family:
                    return style_name
        if default:
            style = self.styles_elements.get(default)
            if style and style.family == family:
                return default
        return None

    @staticmethod
    def split(item, key):
        if isinstance(item, dict):
            inner = item.pop(key, [])
            return (inner, item)
        # item can be list or value
        return (item, {})

    def parse(self, content):
        body, opt = self.split(content, BODY)
        self.defaults.update(opt.get(DEFAULTS, {}))
        self.parse_styles(opt)
        for table_content in body:
            self.parse_table(table_content)

    def parse_table(self, table_content):
        rows, opt = self.split(table_content, TABLE)
        self.tab_counter += 1
        table = Table(opt.get(NAME, f"{DEFAULT_TAB_PREFIX} {self.tab_counter}"))
        style_table_row = self.guess_style(
            opt, "table-row", self.defaults["style_table_row"]
        )
        style_table_cell = self.guess_style(
            opt, "table-cell", self.defaults["style_table_cell"]
        )
        for row_content in rows:
            self.parse_row(table, row_content, style_table_row, style_table_cell)
        self.parse_width(table, opt)
        self.doc.body.append(table)

    def parse_row(self, table, row_content, style_table_row, style_table_cell):
        cells, opt = self.split(row_content, ROW)
        style_table_row = self.guess_style(opt, "table-row", style_table_row)
        self.insert_style(style_table_row)
        row = Row(style=style_table_row)
        style_table_cell = self.guess_style(opt, "table-cell", style_table_cell)
        for cell_content in cells:
            self.parse_cell(row, cell_content, style_table_cell)
        table.append(row)

    def parse_cell(self, row, cell_content, style_table_cell):
        value, opt = self.split(cell_content, VALUE)
        if style_table_cell:
            default = style_table_cell
        elif isinstance(value, str):
            default = self.defaults["style_str"]
        elif isinstance(value, int):
            default = self.defaults["style_int"]
        elif isinstance(value, float):
            default = self.defaults["style_float"]
        else:
            default = self.defaults["style_other"]
        style = self.guess_style(opt, "table-cell", default)
        self.insert_style(style)
        row.append(Cell(value=value, style=style))

    def _column_width_style(self, width):
        """width format: "10.5mm"""
        return self.doc.insert_style(
            Element.from_tag(
                f"""
                    <style:style style:family="table-column">
                    <style:table-column-properties fo:break-before="auto"
                    style:column-width="{width}"/>
                    </style:style>
                """
            ),
            automatic=True,
        )

    def parse_width(self, table, opt):
        width_opt = opt.get(WIDTH)
        if not width_opt:
            return
        if isinstance(width_opt, list):
            for position, width in enumerate(width_opt):
                if width:
                    column = table.get_column(position)
                    column.style = self._column_width_style(width)
                    table.set_column(position, column)
            return
        for position, column in enumerate(table.get_columns()):
            column.style = self._column_width_style(width_opt)
            table.set_column(position, column)


def content_to_ods(data, dest_path):
    doc = ODSGenerator(data)
    doc.save(dest_path)


def main(param_file, dest_path):
    with open(param_file, mode="r", encoding="utf8") as f:
        content = json.load(f)
    content_to_ods(content, dest_path)


def test():
    data_with_minimal_structure = [
        [
            ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"],
            [0, 10, 20, 30, 40, 50, 60, 70, 80, 90],
            [1, 11, 21, 31, 41, 51, 61, 71, 81, 91],
            [2, 12, 22, 32, 42, 52, 62, 72, 82, 92],
        ],
        [
            ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"],
            [100, 110, 120, 130, 140, 150, 160, 170, 180, 190],
            [101, 111, 121, 131, 141, 151, 161, 171, 181, 191],
            [102, 112, 122, 132, 142, 152, 162, 172, 182, 192],
        ],
    ]

    data_with_many_parameters = {
        "styles": [
            {
                "name": "bg_yellow",
                "style": """
                    <style:style style:family="table-cell"
                    style:parent-style-name="Default">
                    <style:table-cell-properties fo:background-color="#fff9ae"/>
                    <style:paragraph-properties
                    fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
                    </style:style>
                """,
            },
            {
                "name": "bg_yellow_grid06",
                "style": """
                    <style:style style:family="table-cell"
                    style:parent-style-name="Default">
                    <style:table-cell-properties fo:background-color="#fff9ae"
                    fo:border="0.06pt solid #000000"/>
                    <style:paragraph-properties
                    fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
                    </style:style>
                """,
            },
        ],
        "defaults": {
            "styles_str": "bold",
        },
        "body": [
            {
                "name": "first tab",
                "table": [
                    ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"],
                    [0, 10, 20, 30, 40, 50, 60, 70, 80, 90],
                    [1, 11, 21, 31, 41, 51, 61, 71, 81, 91],
                    [2, 12, 22, 32, 42, 52, 62, 72, 82, 92],
                ],
                "style": ["number", "table_row_1cm"],
                "width": ["2cm", "4cm", "2cm", "4cm", "5cm"],
            },
            {
                "name": "second tab",
                "table": [
                    {
                        "row": ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j"],
                        "style": "bold_left_bg_gray_grid06",
                    },
                    {
                        "row": [
                            100.01,
                            110.02,
                            "hop",
                            130,
                            140,
                            150,
                            160,
                            170,
                            180,
                            190,
                        ],
                        "style": "grid06",
                    },
                    {
                        "row": [
                            101,
                            111,
                            {"value": 121, "style": "bg_yellow_grid06"},
                            131,
                            141,
                            151,
                            161,
                            171,
                            181,
                            191,
                        ],
                        "style": "grid06",
                    },
                    [102.314, 112, 122, 132, 142, 152, 162, 172, 182, 192],
                ],
                "width": "2.5cm",
            },
        ],
    }
    content_to_ods(data_with_minimal_structure, "data_with_minimal_structure.ods")
    content_to_ods(data_with_many_parameters, "data_with_many_parameters.ods")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Missing parameters.")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
