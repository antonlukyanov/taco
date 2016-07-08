#!/usr/bin/env python3

import argparse
import collections
import re
import os.path as path
import yaml

from core import types
from core import converters


ALLOWED_TYPES = ['txt', 'pdf', 'tex', 'docx']


def get_ext(filepath):
    m = re.match('.+\.([^\.]+)$', filepath)
    if m is None or m.group(1) not in ALLOWED_TYPES:
        raise RuntimeError('Export to specified file type is not supported. See help.')
    return m.group(1)


def load_converter(ext, theme):
    if ext == 'pdf':
        cvt = converters.Latex2PdfConverter('tex', theme)
    elif ext == 'docx':
        cvt = converters.DocxConverter(ext, theme)
    else:
        cvt = converters.ListConverter(ext, theme)

    return cvt


def load_settings(filepath):
    settings = None
    if path.isfile(filepath):
        with open(filepath) as f:
            settings = yaml.load(f)
    return settings


def load_data(filepaths, ext, theme=None):
    yaml_tables_lst = []
    for fp in filepaths:
        with open(fp) as f:
            d = yaml.load(f)
            # Support only one table per file.
            for k, v in d.items():
                yaml_tables_lst.append({k: v})
                break

    layout = None
    layout_settings = None

    first = {'type': 'layout'}
    tables = [first]

    for i, yaml_table in enumerate(yaml_tables_lst):
        for tp, data in yaml_table.items():
            fp, fp_ext = path.splitext(filepaths[i])
            settings_path = fp + '_settings' + fp_ext
            settings = load_settings(settings_path)
            settings = settings[theme][ext] if theme else settings[ext]

            if tp == 'layout':
                layout = data
                layout_settings = settings
                continue

            cls = tp[:-1].capitalize()
            if not hasattr(types, cls):
                cls = types.Row
            else:
                cls = getattr(types, cls)

            rows = []
            for row in data['rows']:
                if not isinstance(row, dict):
                    row = dict(zip(data['fields'], row))
                for key in data['fields']:
                    row.setdefault(key)
                rows.append(cls(**row))

            table = types.Table(data['title'], data['description'], rows)
            if data['sort']:
                table.sort(data['sort'])

            tables.append({
                'type': tp,
                'table': table,
                'settings': settings
            })

    first['table'] = layout
    first['settings'] = types.PageLayout(layout_settings)

    return tables


parser = argparse.ArgumentParser(
    description='Simple converter of table-like data to some common data types like pdf or docx.'
)
parser.add_argument(
    'input',
    nargs='+',
    help='Input files in YAML format.'
)
parser.add_argument(
    'output', help='Supported file types are: %s.' % (', '.join(ALLOWED_TYPES))
)
parser.add_argument(
    '-t', '--theme',
    help='Use specified theme is a collection of templates located in the folder '
         'templates/<ext>/<theme>.',
    action='store'
)
args = parser.parse_args()

ext = get_ext(args.output)
cvt = load_converter(ext, args.theme)
tables = load_data(args.input, ext, args.theme)
cvt.convert(tables, args.output)
