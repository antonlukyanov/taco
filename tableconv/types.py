import re
from builtins import property
from operator import attrgetter

from docx.shared import Cm

class Row(dict):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.__dict__ = self


class Publication:
    def __init__(self, title, authors, coauthorship, date, where, format, size, status, url):
        self.title = title
        self._authors = authors
        self.coauthorship = coauthorship
        self.date = date
        self.where = where
        self.format = format
        self.size = size
        self.status = status
        self.url = url

    def __getitem__(self, key):
        return getattr(self, key)

    @property
    def authors_tex(self):
        return ', '.join([('\\t{%s}' % a.strip()) for a in self._authors.split(',')])

    @property
    def authors(self):
        return self._authors.strip()

    @authors.setter
    def authors(self, value):
        self._authors = value


class Table:
    def __init__(self, title, description=None, rows=None):
        self.title = title
        self.description = description
        self.rows = rows
        self._sort_order = None

    def sort(self, order=None):
        if order:
            self._sort_order = order
            for pair in reversed(order.split(',')):
                field, field_order = pair.strip().split(' ')
                is_desc = field_order == 'desc'
                self.rows.sort(key=attrgetter(field), reverse=is_desc)
        else:
            return self._sort_order


class PageLayout:
    # Width x height.
    paper_size = {
        'a0': (841, 1189),
        'a1': (594, 841),
        'a2': (420, 594),
        'a3': (297, 420),
        'a4': (210, 297),
        'a5': (148, 210),
        'a6': (105, 148),
    }

    def __init__(self, settings):
        self._settings = {
            **{
                'paper_setup': {
                    'size': 'A4',
                    'orientation': 'landscape',
                    'margins': {
                        'left': '20mm',
                        'right': '10mm',
                        'top': '10mm',
                        'bottom': '10mm',
                    }
                }
            },
            **settings
        }

        ps = self._settings['paper_setup']
        ps['size'] = ps['size'].lower()
        if ps['size'] not in self.paper_size:
            raise RuntimeError(
                'Incorrect paper size: %s. Supported paper sizes are: %s.' %
                (ps['size'], ', '.join(self.paper_size.keys()))
            )

        if ps['orientation'] not in ['landscape', 'portrait']:
            raise RuntimeError('Orientation must be either "landscape" or "portrait".')

        for side in ['left', 'right', 'top', 'bottom']:
            length, unit = self.parse_length(ps['margins'][side])
            ps['margins'][side] = Cm(length).mm if unit == 'cm' else length

    def parse_length(self, length):
        length = length.lower()
        m = re.match(
            '^([0-9]{1,4}(\.[0-9]{1,3})?)(mm|cm)$',
            length
        )
        if m is not None:
            value = float(m.group(1))
            unit = m.group(3)
            if unit not in ['mm', 'cm']:
                raise RuntimeError('Unit %s is not suppported.' % unit)
            if value <= 0:
                raise RuntimeError('Length must be positive number.')
        else:
            raise RuntimeError('Could not parse length.')
        return float(value), unit

    @property
    def size(self):
        return self._settings['paper_setup']['size']

    @property
    def width(self):
        return self.paper_size[self.size][0]

    @property
    def height(self):
        return self.paper_size[self.size][1]

    @property
    def margins(self):
        return self._settings['paper_setup']['margins']

    @property
    def margin_left(self):
        return self.margins['left']

    @property
    def margin_right(self):
        return self.margins['right']

    @property
    def margin_top(self):
        return self.margins['top']

    @property
    def margin_bottom(self):
        return self.margins['bottom']

    @property
    def orientation(self):
        return self._settings['paper_setup']['orientation']
