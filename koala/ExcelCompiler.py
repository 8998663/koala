# cython: profile=True

import os.path
import textwrap
from math import *

import networkx
from networkx.algorithms import number_connected_components

from koala.reader import read_archive, read_named_ranges, read_cells
from koala.excellib import *
from koala.utils import *
from koala.ast import graph_from_seeds, shunting_yard, build_ast, prepare_volatile
from koala.ExcelError import *
from koala.Cell import Cell
from koala.Range import RangeFactory
from koala.Spreadsheet import Spreadsheet


class ExcelCompiler(object):
    """Class responsible for taking cells and named_range and create a graph
       that can be serialized to disk, and executed independently of excel.
    """

    def __init__(self, file, ignore_sheets = [], ignore_hidden = False, debug = False):
        print "___### Initializing Excel Compiler ###___"

        file_name = os.path.abspath(file)
        # Decompose subfiles structure in zip file
        archive = read_archive(file_name)
        # Parse cells
        self.cells = read_cells(archive, ignore_sheets, ignore_hidden)
        # Parse named_range { name (ExampleName) -> address (Sheet!A1:A10)}
        self.named_ranges = read_named_ranges(archive)
        self.Range = RangeFactory(self.cells)
        self.volatiles = set()
        self.debug = debug

    def clean_volatile(self, subset, orig_sp):
        G = orig_sp.G
        cells = orig_sp.cellmap
        named_ranges = orig_sp.named_ranges

        sp = Spreadsheet(G,cells, named_ranges, debug = self.debug)

        cleaned_cells, cleaned_ranged_names = sp.clean_volatile(subset)
        self.cells = cleaned_cells
        self.named_ranges = cleaned_ranged_names
        self.volatiles = set()
            
