#Copyright © 2020 R. A. Gardner

from ts_extra_vars_and_funcs import *
from ts_toplevels import *
from ts_widgets import *

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


class wb_sh_sel(tk.Frame): 
    def __init__(self,parent,C):
        tk.Frame.__init__(self,parent)
        self.C = C
        self.sheets_label = label(self,text="Workbook sheets:",font=EF, theme = self.C.theme)
        self.sheets_label.grid(row=0,column=0,padx=10,pady=(10,20),sticky="nswe")
        self.sheet_select = ez_dropdown(self,TF,60)
        self.sheet_select.bind("<<ComboboxSelected>>",lambda focus: self.focus_set())
        self.sheet_select.grid(row=0,column=1,pady=(10,20),sticky="nswe")
        self.run_with_sheet = button(self,text="Read data from selected sheet",
                                     style="TF.Std.TButton",command=self.cont)
        self.run_with_sheet.grid(row=1,column=1,sticky="nswe")

    def enable_widgets(self):
        self.sheet_select.config(state="readonly")
        self.run_with_sheet.config(state="normal")

    def disable_widgets(self):
        self.sheet_select.config(state="disabled")
        self.run_with_sheet.config(state="disabled")

    def updatesheets(self,sheets):
        self.run_with_sheet.config(state="normal")
        self.run_with_sheet.update_idletasks()
        self.sheet_select.set_my_value(sheets[0])
        self.sheet_select['values'] = sheets

    def cont(self):
        self.C.disable_at_start()
        self.C.open_dict['sheet'] = self.sheet_select.get_my_value()
        self.C.wb_sheet_has_been_selected(self.sheet_select.get_my_value())
