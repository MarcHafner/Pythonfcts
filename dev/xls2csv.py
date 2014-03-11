#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
import os.path as op
import csv
import re
import errno as er
import imp
import platform as pl

#
# for setting tsv (from the command line):
# FIELDDELIMITER=$'\t' FILEEXT='.tsv' python xls2csv.py <file>
#

# ---------------------------------------------------------------------------

import setparams as _sg
_params = dict(
    FIELDDELIMITER = u',',
    RECORDDELIMITER = u'\r\n' if pl.system == 'Windows' else u'\n',
    FILEEXT = '.csv',
    VERBOSE = False,
    ENCODING = 'utf8',
    SUBDIR = '',
)

_sg.setparams(_params)
del _sg, _params

# ---------------------------------------------------------------------------

assert len(FIELDDELIMITER) > 0
assert len(RECORDDELIMITER) > 0

FIELDDELIMITER = FIELDDELIMITER.encode(ENCODING)
RECORDDELIMITER = RECORDDELIMITER.encode(ENCODING)

assert RECORDDELIMITER != FIELDDELIMITER


def _makedirs(path):
    try:
        os.makedirs(path)

    except OSError, e:
        if e.errno != er.EEXIST: raise


def _mk_escape(fd=FIELDDELIMITER, rd=RECORDDELIMITER, enc=ENCODING):
    # this function can be called only once...
    # ...since it self-destructs:
    del globals()['_mk_escape']

    special  = set(('\r', fd, rd))
    dspecial = [c.decode(enc) for c in special]
    _esc = lambda m: u''.join(ur'\%03o' % ord(c) for c in m.group(0))
    _re  = re.compile(u'(%s)' % u'|'.join(dspecial))

    def escape(v):
        try:
            return _re.sub(_esc, unicode(v))
        except:
            return _re.sub(_esc, v.decode(enc))

    return escape



def _encodeval(val, _enc=ENCODING, _escape=_mk_escape()):
    return val if val is None else _escape(val).encode(_enc)


_sentinel = object()
def _encode(row, extra=_sentinel):
    rowlist = [_encodeval(value) for value in row]
    if extra != _sentinel and extra > 0:
        rowlist += [''] * extra
    return tuple(rowlist)

def _tostring(row):
    return _encodeval(''.join([v for v in row if v is not None]))

class _meta(type):
    def __new__(mcls, name, bases, dct, _functype=type(lambda: 0)):
        for k, v in dct.items():
            if isinstance(v, _functype): dct[k] = classmethod(v)

        return super(_meta, mcls).__new__(mcls, name, bases, dct)

    def __call__(this, *args, **kwargs):
        this.write(*args, **kwargs)

def _warn(msg):
    if VERBOSE:
        print >> sys.stderr, msg

class _x2csv(object):
    __metaclass__ = _meta

    WSRE = re.compile('\S')

    nsheets = lambda cl, wb: len(cl.sheets(wb))
    def write(this, path, worksheet_to_outpath=None):
        if worksheet_to_outpath is None:
            outdir = SUBDIR if SUBDIR else op.splitext(path)[0]
            def _ws2pth(sh, _fileext=FILEEXT):
                return op.join(outdir, '%s%s' % (this.name(sh), _fileext))
            worksheet_to_outpath = _ws2pth

        wsre = this.WSRE
        with this.open_wb(path) as wb:
            for sh in this.sheets(wb):
                for row in this.rows(sh):
                    if wsre.search(_tostring(row)): break
                else:
                    # we're in a whitespace-only sheet, so we skip it
                    continue

                outpath = worksheet_to_outpath(sh)
                _makedirs(op.dirname(outpath))
                with open(outpath, 'wb') as f:
                    (csv.writer(f,
                                delimiter=FIELDDELIMITER,
                                lineterminator=RECORDDELIMITER)
                     .writerows(this.rows(sh)))

    def rows(this, sheet):
        _rows = this._rows(sheet)

        rown = 1
        for row in _rows:
            assert isinstance(row, list)
            rown += 1
            ncols = len(row)
            if ncols:
                yield _encode(row)
                break
            yield []
        else:
            return

        the_sheet = "sheet '%s'" % this.name(sheet)
        for row in _rows:
            assert isinstance(row, list)
            nc = len(row)
            extra = ncols - nc
            if extra < 0:
                raise ValueError('Too many columns in row %d of %s' %
                                 (rown, the_sheet))

            if extra:
                _warn('Only %d columns in row %d of %s (expected %d)' %
                      (nc, rown, the_sheet, ncols))

            yield _encode(row, extra)
            rown += 1


class _xls2csv(_x2csv):
    name = lambda cl, sh: sh.name
    sheets = lambda cl, wb: wb.sheets()

    def _rows(this, sheet):
        for i in range(sheet.nrows):
            yield sheet.row_values(i)

    import xlrd as xl
    # NOTE: xlrd reads what may look like integers in the xlsx file as
    # floats (e.g. a numeric value of 1 in the xls file will appear as
    # 1.0 in the corresponding cell of the csv file).

    def open_wb(this, path):
        return this.xl.open_workbook(path)

try:
    imp.find_module('xlsxrd')
except ImportError, e:
    class _wbwrapper(object):
        def __init__(self, wb): self.__dict__['_wb'] = wb
        __enter__ = lambda s: s._wb
        __exit__ = lambda *a: None
        __getattr__ = lambda s, a: getattr(s._wb, a)
        def __setattr__(self, attr, value): raise Exception, 'internal error'


    class _xlsx2csv(_x2csv):
        name = lambda cl, sh: sh.title
        sheets = lambda cl, wb: wb.worksheets
        def _rows(this, sheet):
            for r in sheet.rows:
                yield [cell.value for cell in r]

        def open_wb(this, path):
            import openpyxl as xl
            # NOTE: openpyxl renders (what may look like)
            # "whole-number floats" (i.e. floats whose fractional part
            # is zero) as integers (e.g. a numeric value that is
            # displayed as 1.0 in the xlsx file will appear as 1 in
            # the corresponding cell of the csv file); to put it
            # differently, openpyxl preserves only non-zero fractional
            # parts.
            return _wbwrapper(xl.load_workbook(path))
else:
    class _xlsx2csv(_xls2csv): import xlsxrd as xl
    # NOTE: xlsxrd reads what may look like integers in the xlsx file
    # as floats (see note for xlrd above).

# ------------------------------------------------------------

def xls2csv(path, worksheet_to_outpath=None):
    """
    Convert xls file to csv.

    After
    
    xls2csv('src/test/data/Protein_Profiling_Data.xls')

    one gets:

    src/test/data
    ├── Protein_Profiling_Data
    │   ├── All_pg-cell.csv
    │   ├── All_pg-cell_noEpCAM.csv
    │   ├── All_pg-cell_noSrc.csv
    │   ├── pERK comparison_pg-ml.csv
    │   ├── phospho_pg-cell.csv
    │   ├── phospho_pg-ml.csv
    │   ├── phospho updated 111213.csv
    │   ├── total_pg-cell.csv
    │   ├── total_pg-ml.csv
    │   └── total updated 111213.csv
    └── Protein_Profiling_Data.xls
    """

    orig_ext = op.splitext(path)[1][1:]
    ext = orig_ext.lower()
    if ext == 'xls':
        write = _xls2csv
    elif ext == 'xlsx':
        write = _xlsx2csv
    else:
        raise ValueError, 'unsupported file type: %s' % orig_ext

    write(path, worksheet_to_outpath)

if __name__ == '__main__':

    for p in sys.argv[1:]:
        xls2csv(p)
