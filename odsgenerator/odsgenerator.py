#!/usr/bin/env python
# Copyright 2021 Jérôme Dumonteil
# Licence: MIT
# Authors: jerome.dumonteil@gmail.com
"""
Simple .ODS generator from json file, using odfdo library

Usage:
    odsgenerator.py <input.json> <output.ods>"

Principle:
    - a document is a list of tabs,
    - a tab is a list of of rows,
    - a row is a list of cells.

A cell can be:
    - int, float or str
    - a dict, with the following keys (only the 'value' key is mandatory)
        - value: int, float or str
        - style: str or list of str, a style name or a list of style names
        - text: str, a string representation of the value (for ODF readers
          who use it).

A row can be:
    - a list of cells
    - a dict, with the following keys (only the 'row' key is mandatory)
        - row: a list of cells, see above
        - style: str or list of str, a style name or a list of style names

A tab can be:
    - a list of rows
    - a dict, with the following keys (only the 'table' key is mandatory)
        - table: a list of rows,
        - width: a list containing the width of each column of the table
        - name: str, the name of the tab
        - style: str or list of str, a style name or a list of style names

A document can be:
    - a list of tabs
    - a dict, with the following keys (only the 'body' key is mandatory)
        - body: a list of tabs
        - styles: a list of dict of styles defintitions
        - defaults: a dict, for the defaults styles

A style definition is a dict with 2 items:
    - name: str, the name of the style.
    - an XML definition of the ODF style, see DEFAULT_STYLES below.

The styles provided for a row or a table can be of family table-row or
table-cell, they apply to row or and below cells. A style defined at a
lower level (cell for instance) has priority over the style defined above
(row for instance).

In short, if you don't need custom styles, this is a valid document
description:
    "[ [ ["a", "b", "c" ] ] ]"
This json string will create a document with only one tab (name will
be "Tab 1" by default), containing one row of 3 values "a", "b", "c".

Styles:
    - the DEFAULT_STYLES defined below are always available, they can be
      called by their name for cells or rows.
    - To add a custom style, use the "styles" category of the document dict.
      A style is a dict with 2 keys, "definition" and "name".
"""

import sys
import json
import odfdo
from odfdo import Document, Table, Row, Cell, Element

__version__ = 1.2

DEFAULT_STYLES = [
    {
        "name": "default_table_row",
        "definition": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="4.52mm"
            fo:break-before="auto" style:use-optimal-row-height="true"/>
            </style:style>
        """,
    },
    {
        "name": "table_row_1cm",
        "definition": """
            <style:style style:family="table-row">
            <style:table-row-properties style:row-height="1cm"
            fo:break-before="auto"/>
            </style:style>
        """,
    },
    {
        "name": "bold",
        "definition": """
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
        "name": "bold_center",
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:text-properties fo:font-weight="bold"
            style:font-weight-asian="bold" style:font-weight-complex="bold"/>
            <style:table-cell-properties style:text-align-source="vfix"/>
            <style:paragraph-properties fo:text-align="center"/>
            </style:style>
        """,
    },
    {
        "name": "left",
        "definition": """
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
        "definition": """
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
        "definition": """
            <style:style style:family="table-cell"
            style:parent-style-name="Default">
            <style:table-cell-properties style:text-align-source="fix"/>
            <style:paragraph-properties fo:text-align="center"/>
            </style:style>
        """,
    },
    {
        "name": "dec1",
        "definition": """
            <number:number-style><number:number number:decimal-places="1"
            loext:min-decimal-places="1" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "dec2",
        "definition": """
            <number:number-style><number:number number:decimal-places="2"
            loext:min-decimal-places="2" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "dec4",
        "definition": """
            <number:number-style><number:number number:decimal-places="4"
            loext:min-decimal-places="4" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "dec6",
        "definition": """
            <number:number-style><number:number number:decimal-places="6"
            loext:min-decimal-places="6" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "integer",
        "definition": """
            <number:number-style><number:number number:decimal-places="0"
            loext:min-decimal-places="0" number:min-integer-digits="1"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "integer_no0",
        "definition": """
            <number:number-style><number:number number:decimal-places="0"
            loext:min-decimal-places="0" number:min-integer-digits="0"
            number:grouping="false"/>
            </number:number-style>
        """,
        "always_insert": True,
    },
    {
        "name": "grid06",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "bold_left_bg_gray_grid06",
        "definition": """
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
        "name": "bold_center_bg_gray_grid06",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties fo:font-weight="bold"
             style:font-weight-asian="bold" style:font-weight-complex="bold"
             fo:background-color="#dddddd" fo:border="0.06pt solid #000000"
             style:text-align-source="fix"/>
             <style:paragraph-properties fo:text-align="center"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_right",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default">
             <style:table-cell-properties style:text-align-source="fix"
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-right="1.2mm" fo:text-align="end"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_int",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="integer">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_center_int",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="integer_no0">
             <style:table-cell-properties fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties fo:text-align="center"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_dec1",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="dec1">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_dec2",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="dec2">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_dec4",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="dec4">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "grid06_dec6",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="dec6">
             <style:table-cell-properties
             fo:border="0.06pt solid #000000"/>
             <style:paragraph-properties
             fo:margin-left="1.2mm" fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
    {
        "name": "cell_dec2",
        "definition": """
             <style:style style:family="table-cell"
             style:parent-style-name="Default" style:data-style-name="dec2">
             <style:paragraph-properties
             fo:margin-right="1.2mm"/>
             </style:style>
         """,
    },
]

DEFAULTS_DICT = {
    "style_table_row": "default_table_row",
    "style_table_cell": "",
    "style_str": "left",
    "style_int": "right",
    "style_float": "right",
    "style_other": "left",
}

BODY = "body"
TABLE = "table"
ROW = "row"
VALUE = "value"
TEXT = "text"
NAME = "name"
DEFINITION = "definition"
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
            try:
                style = Element.from_tag(s[DEFINITION])
            except Exception:
                print("-" * 80)
                print(s)
                print("-" * 80)
                raise
            style.name = s[NAME]
            self.styles_elements[s[NAME]] = style
            if s.get("always_insert"):
                self.insert_style(s[NAME])
        styles = opt.get(STYLES)
        if isinstance(styles, list):
            for s in styles:
                name = s.get(NAME)
                definition = s.get(DEFINITION)
                style = Element.from_tag(definition)
                style.name = name
                self.styles_elements[name] = style
                if s.get("always_insert"):
                    self.insert_style(name)

    def insert_style(self, name, automatic=True):
        if name and name not in self.used_styles and name in self.styles_elements:
            style = self.styles_elements[name]
            self.doc.insert_style(style, automatic=automatic)
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
        row.append(Cell(value=value, style=style, text=opt.get(TEXT)))

    def column_width_style(self, width):
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
                    column.style = self.column_width_style(width)
                    table.set_column(position, column)
            return
        for position, column in enumerate(table.get_columns()):
            column.style = self.column_width_style(width_opt)
            table.set_column(position, column)


def content_to_ods(data, dest_path):
    doc = ODSGenerator(data)
    doc.save(dest_path)


def main(param_file, dest_path):
    with open(param_file, mode="r", encoding="utf8") as f:
        content = json.load(f)
    content_to_ods(content, dest_path)


def check_odfdo_version():
    if tuple(int(x) for x in odfdo.__version__.split(".")) > (3, 3, 0):
        return True
    print("Error: I need odfdo version >= 3.3.0")
    return False


if __name__ == "__main__":
    if not check_odfdo_version():
        sys.exit(1)
    if len(sys.argv) != 3:
        print("Usage: odsgenerator.py <input.json> <output.ods>")
        sys.exit(1)
    main(sys.argv[1], sys.argv[2])
