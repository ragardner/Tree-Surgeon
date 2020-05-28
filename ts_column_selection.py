#Copyright © 2020 R. A. Gardner

from ts_extra_vars_and_funcs import *
from ts_toplevels import *
from ts_widgets import *
from ts_classes_d import *

import tkinter as tk
from tkinter import ttk, filedialog
from tksheet import Sheet
import os
from sys import argv
from platform import system as get_os
import datetime
import io
from ast import literal_eval
import re
from operator import itemgetter
from itertools import chain, islice, repeat, cycle
from collections import defaultdict, deque, Counter
from math import floor
import json
import pickle
import zlib
import lzma
from base64 import urlsafe_b64encode as b64e, urlsafe_b64decode as b64d
from base64 import b32encode as b32e, b32decode as b32d
import csv as csv_module
from openpyxl import Workbook, load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
import builtins
from fastnumbers import isint, isintlike, isfloat, isreal


class columnselection(tk.Frame):
    def __init__(self,parent,C):
        tk.Frame.__init__(self,parent)
        self.C = C
        self.parent_cols = []
        self.rowlen = 0
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.grid_columnconfigure(1,weight=1)
        
        self.flattened_choices = flattened_base_ids_choices(self,command=self.flattened_mode_toggle)
        self.flattened_choices.grid(row=1,column=0,pady=(0,5),sticky="wnse")
        self.flattened_selector = flattened_column_selector(self)
        self.selector = id_and_parent_column_selector(self)
        self.selector.grid(row=1,column=0,sticky="wnse")
        self.sheetdisplay = Sheet(self,
                                  theme = self.C.theme,
                                  header_font = ("Calibri", 13, "normal"),
                                  row_drag_and_drop_perform = False,
                                  column_drag_and_drop_perform = False,
                                  outline_thickness=1)
        self.sheetdisplay.enable_bindings("enable_all")
        self.sheetdisplay.extra_bindings([("row_index_drag_drop", self.drag_row),
                                          ("column_header_drag_drop", self.drag_col),
                                          ("ctrl_x", self.ctrl_x_in_sheet),
                                          ("delete_key", self.del_in_sheet),
                                          ("rc_delete_column", self.del_in_sheet),
                                          ("rc_delete_row", self.del_in_sheet),
                                          ("rc_insert_column", self.reset_selectors),
                                          ("rc_insert_row", self.reset_selectors),
                                          ("ctrl_v", self.ctrl_v_in_sheet),
                                          ("ctrl_z", self.ctrl_z_in_sheet),
                                          ("edit_cell", self.edit_cell_in_sheet)
                                          ])
        self.sheetdisplay.grid(row=0,column=1,rowspan=3,sticky="nswe")
        
        self.cont_ = button(self,
                            text="Build tree with selections     ",
                            style="TF.Std.TButton",command=self.try_to_build_tree)
        self.cont_.grid(row=2,column=0,sticky="wns",padx=10,pady=(10, 50))
        self.cont_.config(width=40)

        self.flattened_selector.grid(row=0,column=0,pady=(0,9),sticky="nswe")
        self.selector.grid_forget()
        self.selector.grid(row=0,column=0,sticky="nswe")
        self.flattened_selector.grid_forget()

    def flattened_mode_toggle(self):
        x = self.flattened_choices.get_choices()[0]
        if x:
            self.flattened_selector.grid(row=0,column=0,pady=(0,9),sticky="nswe")
            self.selector.grid_forget()
        else:
            self.selector.grid(row=0,column=0,sticky="nswe")
            self.flattened_selector.grid_forget()

    def drag_col(self, selected_cols, c):
        c = int(c)
        colsiter = list(selected_cols)
        colsiter.sort()
        stins = colsiter[0]
        endins = colsiter[-1] + 1
        totalcols = len(colsiter)
        if stins > c:
            for rn in range(len(self.C.treeframe.sheet)):
                self.C.treeframe.sheet[rn] = (self.C.treeframe.sheet[rn][:c] +
                                              self.C.treeframe.sheet[rn][stins:stins + totalcols] +
                                              self.C.treeframe.sheet[rn][c:stins] +
                                              self.C.treeframe.sheet[rn][stins + totalcols:])
        else:
            for rn in range(len(self.C.treeframe.sheet)):
                self.C.treeframe.sheet[rn] = (self.C.treeframe.sheet[rn][:stins] +
                                              self.C.treeframe.sheet[rn][stins + totalcols:c + 1] +
                                              self.C.treeframe.sheet[rn][stins:stins + totalcols] +
                                              self.C.treeframe.sheet[rn][c + 1:])
        self.sheetdisplay.MT.data_ref = self.C.treeframe.sheet
        self.selector.set_columns([h for h in self.C.treeframe.sheet[0]])
        self.flattened_selector.set_columns([h for h in self.C.treeframe.sheet[0]])
        self.selector.detect_id_col()
        self.selector.detect_par_cols()

    def drag_row(self, selected_rows, r):
        r = int(r)
        rowsiter = list(selected_rows)
        rowsiter.sort()
        stins = rowsiter[0]
        endins = rowsiter[-1] + 1
        totalrows = len(rowsiter)
        if stins > r:
            self.C.treeframe.sheet = (self.C.treeframe.sheet[:r] +
                                      self.C.treeframe.sheet[stins:stins + totalrows] +
                                      self.C.treeframe.sheet[r:stins] +
                                      self.C.treeframe.sheet[stins + totalrows:])
        else:
            self.C.treeframe.sheet = (self.C.treeframe.sheet[:stins] +
                                      self.C.treeframe.sheet[stins + totalrows:r + 1] +
                                      self.C.treeframe.sheet[stins:stins + totalrows] +
                                      self.C.treeframe.sheet[r + 1:])
        self.sheetdisplay.MT.data_ref = self.C.treeframe.sheet
        if endins == 0 or r == 0 or stins == 0:
            self.selector.set_columns([h for h in self.C.treeframe.sheet[0]])
            self.flattened_selector.set_columns([h for h in self.C.treeframe.sheet[0]])
            self.selector.detect_id_col()
            self.selector.detect_par_cols()

    def reset_selectors(self, event = None):
        idcol = self.selector.get_id_col()
        parcols = self.selector.get_par_cols()
        ancparcols = self.flattened_selector.get_par_cols()
        self.selector.set_columns([h for h in self.C.treeframe.sheet[0]] if self.C.treeframe.sheet else [])
        self.flattened_selector.set_columns([h for h in self.C.treeframe.sheet[0]] if self.C.treeframe.sheet else [])
        if idcol is not None and self.C.treeframe.sheet:
            self.selector.set_id_col(idcol)
        if parcols and self.C.treeframe.sheet:
            self.selector.set_par_cols(parcols)
        if ancparcols and self.C.treeframe.sheet:
            self.flattened_selector.set_par_cols(ancparcols)

    def del_in_sheet(self, event = None):
        self.reset_selectors()

    def ctrl_x_in_sheet(self, event = None):
        self.reset_selectors()

    def ctrl_v_in_sheet(self, event = None):
        self.reset_selectors()

    def ctrl_z_in_sheet(self, event = None):
        self.reset_selectors()

    def edit_cell_in_sheet(self, event = None):
        idcol = self.selector.get_id_col()
        parcols = self.selector.get_par_cols()
        ancparcols = self.flattened_selector.get_par_cols()
        if event[1] == idcol or event[1] in parcols or event[1] in ancparcols or event[0] == 0:
            self.reset_selectors()

    def enable_widgets(self):
        self.selector.enable_me()
        self.flattened_selector.enable_me()
        self.flattened_choices.enable_me()
        self.cont_.config(state="normal")
        self.sheetdisplay.basic_bindings(True)
        self.sheetdisplay.enable_bindings("enable_all")
        self.sheetdisplay.extra_bindings([("row_index_drag_drop", self.drag_row),
                                          ("column_header_drag_drop", self.drag_col),
                                          ("ctrl_x", self.ctrl_x_in_sheet),
                                          ("delete_key", self.del_in_sheet),
                                          ("rc_delete_column", self.del_in_sheet),
                                          ("rc_delete_row", self.del_in_sheet),
                                          ("rc_insert_column", self.reset_selectors),
                                          ("rc_insert_row", self.reset_selectors),
                                          ("ctrl_v", self.ctrl_v_in_sheet),
                                          ("ctrl_z", self.ctrl_z_in_sheet),
                                          ("edit_cell", self.edit_cell_in_sheet)
                                          ])

    def disable_widgets(self):
        self.selector.disable_me()
        self.flattened_selector.disable_me()
        self.flattened_choices.disable_me()
        self.cont_.config(state="disabled")
        self.sheetdisplay.basic_bindings(False)
        self.sheetdisplay.disable_bindings("disable_all")
        self.sheetdisplay.extra_bindings("unbind_all")
        self.sheetdisplay.unbind("<Delete>")
        
    def populate(self, columns, non_tsrgn_xl_file = False):
        self.sheetdisplay.deselect("all")
        self.non_tsrgn_xl_file = non_tsrgn_xl_file
        self.rowlen = len(columns)
        self.selector.set_columns([h for h in self.C.treeframe.sheet[0]])
        self.flattened_selector.set_columns([h for h in self.C.treeframe.sheet[0]])
        self.C.treeframe.sheet = self.sheetdisplay.set_sheet_data(data = self.C.treeframe.sheet,
                                                                  redraw=True)
        self.sheetdisplay.headers(newheaders=0)
        if len(self.C.treeframe.sheet)  < 3000:
            self.sheetdisplay.set_all_cell_sizes_to_text()
        self.selector.detect_id_col()
        self.selector.detect_par_cols()
        self.flattened_selector.detect_par_cols()
        self.C.show_frame("columnselection")
        
    def try_to_build_tree(self):
        baseids, order, delcols = self.flattened_choices.get_choices()
        if baseids:
            hiers = list(self.flattened_selector.get_par_cols())
            if len(hiers) < 2:
                return
        else:
            hiers = list(self.selector.get_par_cols())
            if not hiers:
                return
            idcol = self.selector.get_id_col()
            if idcol in hiers or idcol is None:
                return
        self.C.status_bar.change_text("Loading...   ")
        self.C.disable_at_start()
        self.C.treeframe.sheet = self.sheetdisplay.set_sheet_data(data = self.C.treeframe.sheet)
        if baseids:
            if order == "Order: Base → Top":
                idcol = hiers.pop(0)
            elif order == "Order: Top → Base":
                idcol = hiers.pop(len(hiers) - 1)
            self.C.treeframe.sheet[:] = [row + list(repeat("",self.rowlen - len(row))) if len(row) < self.rowlen else row for row in self.C.treeframe.sheet]
            self.C.treeframe.sheet, self.rowlen, idcol, newpc = self.C.treeframe.treebuilder.convert_flattened_to_normal(data = self.C.treeframe.sheet,
                                                                                                                          idcol = idcol,
                                                                                                                          parcols = hiers,
                                                                                                                          rowlen = self.rowlen,
                                                                                                                          order = order,
                                                                                                                          delcols = delcols,
                                                                                                                         warnings = self.C.treeframe.warnings)
            hiers = [newpc]
        self.C.treeframe.headers = [Header(name) for name in self.C.treeframe.fix_heads(self.C.treeframe.sheet.pop(0),self.rowlen)]
        self.C.treeframe.ic = idcol
        self.C.treeframe.hiers = hiers
        self.C.treeframe.pc = hiers[0]
        self.C.treeframe.row_len = int(self.rowlen)
        self.C.treeframe.set_metadata(headers = True)
        self.C.treeframe.sheet, self.C.treeframe.nodes, self.C.treeframe.warnings = self.C.treeframe.treebuilder.build(self.C.treeframe.sheet,
                                                                                                                            self.C.treeframe.new_sheet,
                                                                                                                            self.C.treeframe.row_len,
                                                                                                                            self.C.treeframe.ic,
                                                                                                                            self.C.treeframe.hiers,
                                                                                                                            self.C.treeframe.nodes,
                                                                                                                            self.C.treeframe.warnings)
        self.C.treeframe.populate(non_tsrgn_xl_file = self.non_tsrgn_xl_file)
        self.C.treeframe.show_warnings(str(self.C.open_dict['filepath']),str(self.C.open_dict['sheet']))
