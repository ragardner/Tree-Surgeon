#Copyright © 2020 R. A. Gardner

#Required libraries
from tksheet import Sheet
from openpyxl import Workbook, load_workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment

from ts_column_selection import *
from ts_compare import *
from ts_extra_vars_and_funcs import *
from ts_sheet_selection import *
from ts_toplevels import *
from ts_widgets import *
from ts_classes_d import *

import tkinter as tk
from tkinter import ttk, filedialog
from math import floor
from collections import defaultdict, deque, Counter
import os
from sys import argv
from platform import system as get_os
import datetime
import io
import base64
from ast import literal_eval
import re
from operator import itemgetter
from itertools import chain, islice, repeat, cycle
import json
import pickle
import zlib
import lzma
import csv as csv_module
from warnings import simplefilter as ignorew_

try:
    from ts_d import *
    try:
        run_app(argv)
    except Exception as e:
        with open("treesurgeon_start_error.txt", "w") as f:
            f.write(f"{e}")
except Exception as e:
    with open("treesurgeon_import_error.txt", "w") as f:
        f.write(f"{e}")


