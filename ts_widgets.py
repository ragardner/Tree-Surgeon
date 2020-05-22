#Copyright © 2020 R. A. Gardner

from ts_extra_vars_and_funcs import *
import tkinter as tk
from tkinter import ttk, filedialog
import csv as csv_module
from collections import defaultdict, deque, Counter
import os
import datetime
import re
from itertools import islice, repeat, chain, cycle
import json
import pickle
import zlib
import lzma
from base64 import b32encode as b32e, b32decode as b32d
import csv as csv_module
from openpyxl import Workbook, load_workbook
import io
import datetime
from tksheet import Sheet
from math import floor
from ast import literal_eval
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import PatternFill
from openpyxl.styles import Alignment
from platform import system as get_os
from fastnumbers import isint, isintlike, isfloat, isreal


class id_and_parent_column_selector(tk.Frame):
    def __init__(self,
                 parent,
                 headers=[[]],
                 show_disp_1=True,
                 show_disp_2=True,
                 theme = "dark blue",
                 expand = False):
        tk.Frame.__init__(self,parent,background=theme_bg(theme))
        self.grid_propagate(False)
        self.grid_rowconfigure(1,weight=1)
        if show_disp_1:
            self.grid_columnconfigure(0,weight=1,uniform="x")
        if show_disp_2:
            self.grid_columnconfigure(1,weight=1,uniform="x")
        self.C = parent
        self.headers = headers
        self.id_col = None
        self.par_cols = set()
        self.id_col_display = readonly_entry_with_scrollbar(self, font = EFB, theme = theme)
        self.id_col_display.set_my_value("   ID column:   ")
        if show_disp_1:
            self.id_col_display.grid(row=0,column=0,sticky="nswe")
        self.id_col_selection = Sheet(self,
                                      height = 280 if not expand else None,
                                      width = 250 if not expand else None,
                                      theme = theme,
                                      show_selected_cells_border = False,
                                              align="center",
                                              header_align="center",
                                              row_index_align="center",
                                            selected_cells_background="#0078d7",
                                      selected_cells_foreground = "white",
                                      header_select_foreground="white",
                                         row_index_select_foreground="white",
                                              header_select_background="#0078d7",
                                              row_index_select_background="#0078d7",
                                              header_font = ("Calibri", 13, "normal"),
                                              column_width=170,
                                              row_index_width=60,
                                              headers=["SELECT ID"])
        self.id_col_selection.data_reference(newdataref=self.headers)
        self.id_col_selection.enable_bindings(("single",
                                                "column_width_resize",
                                                "double_click_column_resize"))
        self.id_col_selection.extra_bindings([("cell_select",self.id_col_selection_B1),
                                              ("deselect",self.id_col_selection_B1)])
        if show_disp_1:
            self.id_col_selection.grid(row=1,column=0,sticky="nswe")
        self.par_col_selection = Sheet(self,
                                       height = 280 if not expand else None,
                                       width = 250 if not expand else None,
                                        theme = theme,
                                              align="center",
                                       show_selected_cells_border = False,
                                              header_align="center",
                                              row_index_align="center",
                                              selected_cells_background="#8cba66",
                                       selected_cells_foreground = "white",
                                       header_select_foreground="white",
                                         row_index_select_foreground="white",
                                              header_select_background="#8cba66",
                                              row_index_select_background="#8cba66",
                                              header_font = ("Calibri", 13, "normal"),
                                              column_width=170,
                                              row_index_width=60,
                                              headers=["SELECT PARENTS"])
        self.par_col_selection.data_reference(newdataref=self.headers)
        self.par_col_selection.extra_bindings([("cell_select",self.par_col_selection_B1),
                                               ("deselect",self.par_col_deselection_B1)])
        self.par_col_selection.enable_bindings(("toggle",
                                                "column_width_resize",
                                                "double_click_column_resize"))
        if show_disp_2:
            self.par_col_selection.grid(row=1,column=1,sticky="nswe")
        self.par_col_display = readonly_entry_with_scrollbar(self, font = EFB, theme = theme)
        self.par_col_display.set_my_value("   Parent columns:   ")
        if show_disp_2:
            self.par_col_display.grid(row=0,column=1,sticky="nswe")
        self.detect_id_col_button = button(self,text="Detect ID column",style="EFB.Std.TButton",
                                           command=self.detect_id_col)
        if show_disp_1:
            self.detect_id_col_button.grid(row=2,column=0,padx=2,pady=2,sticky="ns")
        self.detect_par_cols_button = button(self,text="Detect parent columns",style="EFB.Std.TButton",
                                             command=self.detect_par_cols)
        if show_disp_2:
            self.detect_par_cols_button.grid(row=2,column=1,padx=2,pady=2,sticky="ns")

    def reset_size(self,width=500,height=350):
        self.config(width=width,height=height)
        self.update_idletasks()
        self.par_col_selection.refresh()
        self.id_col_selection.refresh()

    def set_columns(self,columns):
        self.clear_displays()
        self.set_par_cols([])
        self.headers = [[h] for h in columns]
        self.id_col_selection.data_reference(newdataref=self.headers,redraw=True)
        self.par_col_selection.data_reference(newdataref=self.headers,redraw=True)

    def disable_me(self):
        self.id_col_selection.basic_bindings(enable=False)
        self.par_col_selection.basic_bindings(enable=False)
        self.detect_id_col_button.config(state="disabled")
        self.detect_par_cols_button.config(state="disabled")

    def enable_me(self):
        self.id_col_selection.basic_bindings(enable=True)
        self.par_col_selection.basic_bindings(enable=True)
        self.detect_id_col_button.config(state="normal")
        self.detect_par_cols_button.config(state="normal")

    def detect_id_col(self):
        for i,e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x == "id" or x.startswith("id"):
                self.set_id_col(i)
                break

    def detect_par_cols(self):
        parent_cols = []
        for i,e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x.startswith("parent"):
                parent_cols.append(i)
        if parent_cols:
            self.set_par_cols(parent_cols)

    def id_col_selection_B1(self,event=None):
        if event is not None:
            self.id_col = tuple(tup[0] for tup in self.id_col_selection.get_selected_cells())
            if self.id_col:
                self.id_col = self.id_col[0]
                self.id_col_display.set_my_value(f"   ID column:   {self.id_col + 1}")
            else:
                self.id_col = None
                self.id_col_display.set_my_value("   ID column:   ")

    def par_col_selection_B1(self,event=None):
        if event is not None:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))

    def par_col_deselection_B1(self,event=None):
        if event is not None:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))

    def clear_displays(self):
        self.headers = [[]]
        self.id_col = None
        self.id_col_selection.deselect("all")
        self.id_col_display.set_my_value("   ID column:   ")
        self.par_cols = set()
        self.par_col_selection.deselect("all")
        self.par_col_display.set_my_value("   Parent columns:   ")
        self.id_col_selection.data_reference(newdataref=[[]],redraw=True)
        self.par_col_selection.data_reference(newdataref=[[]],redraw=True)

    def set_id_col(self,col):
        self.id_col = col
        self.id_col_selection.deselect("all")
        self.id_col_selection.refresh()
        self.id_col_selection.select_cell(row=col,column=0,redraw=True)
        self.id_col_selection.see(row=col,column=0)
        self.id_col_display.set_my_value("   ID column:   " + str(col + 1))

    def set_par_cols(self, cols):
        self.par_col_selection.deselect("all")
        self.par_col_selection.refresh()
        if cols:
            self.par_cols = set(cols)
            for r in cols:
                self.par_col_selection.toggle_select_cell(r, 0, redraw = False)
            self.par_col_selection.see(row = cols[0], column = 0)
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([f"{n}" for n in sorted(p + 1 for p in self.par_cols)]))

    def get_id_col(self):
        return self.id_col

    def get_par_cols(self):
        return tuple(sorted(self.par_cols))

    def change_theme(self, theme = "dark blue"):
        self.id_col_selection.change_theme(theme)
        self.par_col_selection.change_theme(theme)
        self.id_col_selection.set_options(selected_cells_background="#0078d7",
                                            selected_cells_foreground = "white",
                                            header_select_foreground="white",
                                            row_index_select_foreground="white",
                                            header_select_background="#0078d7",
                                            row_index_select_background="#0078d7")
        self.par_col_selection.set_options(selected_cells_background="#8cba66",
                                            selected_cells_foreground = "white",
                                            header_select_foreground="white",
                                            row_index_select_foreground="white",
                                            header_select_background="#8cba66",
                                            row_index_select_background="#8cba66")
        self.config(background=theme_bg(theme))
        self.id_col_display.my_entry.config(background = theme_entry_bg(theme),
                                  foreground = theme_entry_fg(theme),
                                  disabledbackground = theme_entry_dbg(theme),
                                  disabledforeground = theme_entry_dfg(theme),
                                  insertbackground = theme_entry_cursor(theme),
                                  readonlybackground = theme_entry_dbg(theme)
                                  )
        self.par_col_display.my_entry.config(background = theme_entry_bg(theme),
                                  foreground = theme_entry_fg(theme),
                                  disabledbackground = theme_entry_dbg(theme),
                                  disabledforeground = theme_entry_dfg(theme),
                                  insertbackground = theme_entry_cursor(theme),
                                  readonlybackground = theme_entry_dbg(theme)
                                  )

class flattened_base_ids_choices(tk.Frame):
    def __init__(self,
                 parent,
                 command,
                 theme = "dark blue"):
        tk.Frame.__init__(self,parent,background = theme_bg(theme))
        self.C = parent
        self.extra_func = command
        self.only_base_ids_button = x_checkbutton(self,
                                                  text="Sheet is flattened with base IDs   ",
                                                  style="x_button.Std.TButton",
                                                  compound="right",
                                                  command=self.flattened_mode_toggle,
                                                  checked=False)
        self.only_base_ids_button.grid(row=0,column=0,sticky="nswe",padx=(0,5),pady=(0,5))
        self.order_dropdown = ez_dropdown(self,font=EF)
        self.order_dropdown.bind("<<ComboboxSelected>>",lambda focus: self.focus_set())
        self.order_dropdown['values'] = ["Order: Top → Base","Order: Base → Top"]
        self.order_dropdown.set_my_value("Order: Top → Base")
        self.order_dropdown.grid(row=0,column=1,sticky="nswe")
        self.flattened_mode_toggle(func=False)

    def disable_me(self):
        self.only_base_ids_buttons.config(state="disabled")
        self.order_dropdown.config(state="disabled")

    def enable_me(self):
        self.only_base_ids_buttons.config(state="normal")
        self.flattened_mode_toggle(func=False)

    def flattened_mode_toggle(self,event=None,func=True):
        if self.only_base_ids_button.get_checked():
            self.order_dropdown.config(state="readonly")
        else:
            self.order_dropdown.config(state="disabled")
        if func:
            self.extra_func()

    def set_checked(self,s=True):
        self.only_base_ids_button.set_checked(s)

    def get_choices(self):
        return self.only_base_ids_button.get_checked(), self.order_dropdown.get_my_value(), True

    def change_theme(self, theme = "dark blue"):
        self.config(bg = theme_bg(theme))


class flattened_column_selector(tk.Frame):
    def __init__(self,
                 parent,
                 headers=[[]],
                 theme = "dark blue"):
        tk.Frame.__init__(self, parent, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.grid_rowconfigure(1,weight=1)
        self.grid_columnconfigure(0,weight=1)
        self.C = parent
        self.headers = headers
        self.par_cols = set()
        self.par_col_display = readonly_entry_with_scrollbar(self,font=EFB, theme = theme)
        self.par_col_display.set_my_value("   Hierarchy columns:   ")
        self.par_col_display.grid(row=0,column=0,sticky="nswe")
        self.par_col_selection = Sheet(self,
                                       width = 500,
                                       height = 300,
                                       theme = theme,
                                              align="center",
                                       show_selected_cells_border = False,
                                              header_align="center",
                                              row_index_align="center",
                                              selected_cells_background="#8cba66",
                                              header_select_background="#8cba66",
                                              row_index_select_background="#8cba66",
                                             header_select_foreground="white",
                                             row_index_select_foreground="white",
                                             selected_cells_foreground="white",
                                              header_font = ("Calibri", 13, "normal"),
                                              column_width=350,
                                              row_index_width=60,
                                              headers=["SELECT ALL HIERARCHY COLUMNS"])
        self.par_col_selection.data_reference(newdataref=self.headers)
        self.par_col_selection.extra_bindings([("cell_select",self.par_col_selection_B1),
                                               ("deselect",self.par_col_deselection_B1)])
        self.par_col_selection.enable_bindings(("toggle",
                                                "column_width_resize",
                                                "double_click_column_resize"))
        self.par_col_selection.grid(row=1,column=0,sticky="nswe")
        
    def set_columns(self,columns):
        self.clear_displays()
        self.set_par_cols([])
        self.headers = [[h] for h in columns]
        self.par_col_selection.data_reference(newdataref=self.headers,redraw=True)

    def disable_me(self):
        self.par_col_selection.basic_bindings(enable=False)
        self.par_col_selection.extra_bindings([("cell_select",None),
                                               ("deselect",None)])

    def enable_me(self):
        self.par_col_selection.basic_bindings(enable=True)
        self.par_col_selection.extra_bindings([("cell_select",self.par_col_selection_B1),
                                               ("deselect",self.par_col_deselection_B1)])

    def detect_par_cols(self):
        parent_cols = []
        for i,e in enumerate(self.headers):
            if not e:
                continue
            x = e[0].lower().strip()
            if x.startswith("parent"):
                parent_cols.append(i)
        if parent_cols:
            self.set_par_cols(parent_cols)

    def par_col_selection_B1(self,event=None):
        if event is not None:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))

    def par_col_deselection_B1(self,event=None):
        if event is not None:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))
            
    def clear_displays(self):
        self.headers = [[]]
        self.par_col_selection.data_reference(newdataref=[[]],redraw=True)
        self.par_cols = set()
        self.par_col_selection.deselect("all")
        self.par_col_display.set_my_value("   Hierarchy columns:   ")

    def set_par_cols(self,cols):
        self.par_col_selection.deselect("all")
        self.par_col_selection.refresh()
        if cols:
            self.par_cols = set(cols)
            for r in cols:
                self.par_col_selection.toggle_select_cell(r, 0, redraw = False)
            self.par_col_selection.see(row=cols[0],column=0)
            self.par_col_display.set_my_value("   Hierarchy columns:   " + ",".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))
        self.par_col_selection.refresh()

    def get_par_cols(self):
        return tuple(sorted(self.par_cols))

    def change_theme(self, theme = "dark blue"):
        self.config(bg = theme_bg(theme))
        self.par_col_selection.change_theme(theme)
        self.par_col_selection.set_options(selected_cells_background="#8cba66",
                                              header_select_background="#8cba66",
                                              row_index_select_background="#8cba66",
                                             header_select_foreground="white",
                                             row_index_select_foreground="white",
                                             selected_cells_foreground="white")
        self.par_col_display.my_entry.config(background = theme_entry_bg(theme),
                                  foreground = theme_entry_fg(theme),
                                  disabledbackground = theme_entry_dbg(theme),
                                  disabledforeground = theme_entry_dfg(theme),
                                  insertbackground = theme_entry_cursor(theme),
                                  readonlybackground = theme_entry_dbg(theme)
                                  )


class single_column_selector(tk.Frame):
    def __init__(self,
                 parent,
                 headers=[[]],
                 width=250,
                 height=350,
                 theme = "dark blue"):
        tk.Frame.__init__(self,parent,width=width,height=height, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.grid_rowconfigure(1,weight=1)
        self.grid_columnconfigure(0,weight=1)
        self.C = parent
        self.headers = headers
        self.col = None
        self.col_display = readonly_entry_with_scrollbar(self, font = EFB, theme = theme)
        self.col_display.set_my_value("  Column:   ")
        self.col_display.grid(row=0,column=0,sticky="nswe")
        self.col_selection = Sheet(self,
                                   theme = theme,
                                   show_selected_cells_border = False,
                                         align="center",
                                          header_align="center",
                                          row_index_align="center",
                                          selected_cells_background="#8cba66",
                                          header_select_background="#8cba66",
                                          row_index_select_background="#8cba66",
                                         header_select_foreground="white",
                                         row_index_select_foreground="white",
                                         selected_cells_foreground="white",
                                          header_font = ("Calibri", 13, "normal"),
                                          column_width=180,
                                          row_index_width=50,
                                          headers=["SELECT A COLUMN"])
        self.col_selection.data_reference(newdataref=self.headers)
        self.col_selection.extra_bindings([("cell_select",self.col_selection_B1),
                                           ("deselect", self.col_deselect)])
        self.col_selection.enable_bindings(("single",
                                            "column_width_resize",
                                            "double_click_column_resize"))
        self.col_selection.grid(row=1,column=0,sticky="nswe")
        
    def set_columns(self,columns):
        self.clear_displays()
        self.headers = [[h] for h in columns]
        self.col_selection.data_reference(newdataref=self.headers,redraw=True)

    def disable_me(self):
        self.col_selection.basic_bindings(enable=False)

    def enable_me(self):
        self.col_selection.basic_bindings(enable=True)

    def col_selection_B1(self,event=None):
        if event is not None:
            self.col = event[1]
            self.col_display.set_my_value("   Column:   " + str(event[1] + 1))

    def col_deselect(self, event = None):
        if event is not None:
            self.col = None
            self.col_display.set_my_value("  Column:   ")

    def par_col_deselection_B1(self,event=None):
        if event is not None:
            self.par_cols = set(tup[0] for tup in self.par_col_selection.get_selected_cells())
            self.par_col_display.set_my_value("   Parent columns:   " + ", ".join([str(n) for n in sorted(p + 1 for p in self.par_cols)]))

    def clear_displays(self):
        self.headers = [[]]
        self.col_selection.data_reference(newdataref=[[]],redraw=True)
        self.col = 0
        self.col_selection.deselect("all")
        self.col_display.set_my_value("   Hierarchy columns:   ")

    def set_col(self,col = None):
        if col is not None:
            self.col = int(col)
            self.col_selection.deselect("all")
            self.col_selection.select_cell(col, 0, redraw = False)
            self.col_selection.see(row=col,column=0)
            self.col_selection.refresh()
            self.col_display.set_my_value("   Column:   " + str(col + 1))

    def get_col(self):
        return int(self.col)


class x_checkbutton(ttk.Button):
    def __init__(self,parent,text="",style="Std.TButton",command=None,state="normal",checked=False,compound="right"):
        button.__init__(self,
                        parent,
                        text=text,
                        style=style,
                        command=command,
                        state=state)
        self.image_compound = compound
        self.on_icon = tk.PhotoImage(format="gif",data=checked_icon)
        self.off_icon = tk.PhotoImage(format="gif",data=unchecked_icon)
        self.checked = checked
        if checked:
            self.config(image=self.on_icon,compound=compound)
        else:
            self.config(image=self.off_icon,compound=compound)
        self.bind("<1>",self.B1)
        
    def set_checked(self,state="toggle"):
        if state == "toggle":
            self.checked = not self.checked
            if self.checked:
                self.config(image=self.on_icon,compound=self.image_compound)
            else:
                self.config(image=self.off_icon,compound=self.image_compound)
        elif state:
            self.checked = True
            self.config(image=self.on_icon,compound=self.image_compound)
        elif not state:
            self.checked = False
            self.config(image=self.off_icon,compound=self.image_compound)
            
    def get_checked(self):
        return bool(self.checked)
    
    def B1(self,event):
        x = str(self['state'])
        if "normal" in x:
            self.checked = not self.checked
            if self.checked:
                self.config(image=self.on_icon,compound=self.image_compound)
            else:
                self.config(image=self.off_icon,compound=self.image_compound)
        self.update_idletasks()
        
    def change_text(self,text):
        self.config(text=text)
        self.update_idletasks()


class auto_add_condition_num_frame(tk.Frame):
    def __init__(self,
                 parent,
                 col_sel,
                 sheet,
                 theme = "dark blue"):
        tk.Frame.__init__(self, parent, height = 200, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.col_sel = col_sel
        self.sheet_ref = sheet
        self.grid_columnconfigure(1,weight=1)
        self.grid_columnconfigure(3,weight=1)
        self.grid_rowconfigure(3,weight=1)
        self.min_label = label(self,text="Min:",font=EFB, theme = theme)
        self.min_label.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(5,0))
        self.min_entry = numerical_entry_with_scrollbar(self, theme = theme)
        self.min_entry.grid(row=0,column=1,sticky="nswe",padx=(20,0),pady=(20,0))
        self.max_label = label(self,text="Max:",font=EFB, theme = theme)
        self.max_label.grid(row=0,column=2,sticky="nswe",padx=(10,10),pady=(5,0))
        self.max_entry = numerical_entry_with_scrollbar(self, theme = theme)
        self.max_entry.grid(row=0,column=3,sticky="nswe",padx=(10,20),pady=(20,0))
        self.get_col_min = button(self,text="Get column minimum",style="EF.Std.TButton",command=self.get_col_min_val)
        self.get_col_min.grid(row=1,column=1,sticky="nswe",padx=(20,0))
        self.get_col_max = button(self,text="Get column maximum",style="EF.Std.TButton",command=self.get_col_max_val)
        self.get_col_max.grid(row=1,column=3,sticky="nswe",padx=(10,20))
        self.asc_desc_dropdown = ez_dropdown(self,font=EF)
        self.asc_desc_dropdown['values'] = ("ASCENDING","DESCENDING")
        self.asc_desc_dropdown.set_my_value("ASCENDING")
        self.asc_desc_dropdown.grid(row=2,column=1,sticky="nswe",padx=(20,0))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=3,column=0,columnspan=4,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Save conditions",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(20,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(20,20))
        self.result = False
        self.min_val = ""
        self.max_val = ""
        self.order = ""
        self.min_entry.place_cursor()

    def get_col_min_val(self,event=None):
        c = self.col_sel
        try:
            self.min_entry.set_my_value(str(min(float(r[c]) for r in self.sheet_ref if isreal(r[c]))))
        except:
            pass

    def get_col_max_val(self,event=None):
        c = self.col_sel
        try:
            self.max_entry.set_my_value(str(max(float(r[c]) for r in self.sheet_ref if isreal(r[c]))))
        except:
            pass

    def confirm(self,event=None):
        self.result = True
        self.min_val = self.min_entry.get_my_value()
        self.max_val = self.max_entry.get_my_value()
        self.order = self.asc_desc_dropdown.get_my_value()
        self.destroy()

    def cancel(self,event=None):
        self.destroy()


class auto_add_condition_date_frame(tk.Frame):
    def __init__(self,
                 parent,
                 col_sel,
                 sheet,
                 DATE_FORM,
                 theme = "dark blue"):
        tk.Frame.__init__(self, parent, height = 225, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(0,weight=1)
        self.grid_columnconfigure(2,weight=1)
        self.grid_rowconfigure(3,weight=1)
        self.col_sel = col_sel
        self.sheet_ref = sheet
        self.DATE_FORM = DATE_FORM
        if DATE_FORM == "%d/%m/%Y":
            label_form = "DD/MM/YYYY"
        elif DATE_FORM == "%m/%d/%Y":
            label_form = "MM/DD/YYYY"
        elif DATE_FORM == "%Y/%m/%d":
            label_form = "YYYY/MM/DD"
        else:
            self.DATE_FORM = "DD/MM/YYYY"
            label_form = "DD/MM/YYYY"
        self.min_label = label(self,text="Min  " + label_form,font=EFB, theme = theme)
        self.min_label.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(5,0))
        self.min_entry = date_entry(self,DATE_FORM, theme = theme)
        self.min_entry.grid(row=0,column=1,sticky="nswe",padx=(20,0),pady=(20,0))
        self.max_label = label(self,text="Max  " + label_form,font=EFB, theme = theme)
        self.max_label.grid(row=0,column=2,sticky="nswe",padx=(10,10),pady=(5,0))
        self.max_entry = date_entry(self,DATE_FORM, theme = theme)
        self.max_entry.grid(row=0,column=3,sticky="nswe",padx=(10,20),pady=(20,0))
        self.get_col_min = button(self,text="Get column minimum",style="EF.Std.TButton",command=self.get_col_min_val)
        self.get_col_min.grid(row=1,column=1,sticky="nswe",padx=(20,0))
        self.get_col_max = button(self,text="Get column maximum",style="EF.Std.TButton",command=self.get_col_max_val)
        self.get_col_max.grid(row=1,column=3,sticky="nswe",padx=(10,20))
        self.asc_desc_dropdown = ez_dropdown(self,font=EF)
        self.asc_desc_dropdown['values'] = ("ASCENDING","DESCENDING")
        self.asc_desc_dropdown.set_my_value("ASCENDING")
        self.asc_desc_dropdown.grid(row=2,column=1,sticky="nswe",padx=(20,0))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=3,column=0,columnspan=4,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Save conditions",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(20,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(20,20))
        self.result = False
        self.min_val = ""
        self.max_val = ""
        self.order = ""
        self.min_entry.place_cursor()

    def detect_date_form(self,date):
        forms = []
        for form in ("%d/%m/%Y","%m/%d/%Y","%Y/%m/%d"):
            try:
                datetime.datetime.strptime(date,form).date()
                forms.append(form)
            except:
                pass
        if len(forms) == 1:
            return forms[0]
        return False

    def get_col_min_val(self,event=None):
        c = self.col_sel
        try:
            self.min_entry.set_my_value(datetime.datetime.strftime(min(datetime.datetime.strptime(r[c],self.DATE_FORM)
                                                                       for r in self.sheet_ref if self.detect_date_form(r[c]) == self.DATE_FORM),self.DATE_FORM))                                                   
        except:
            pass

    def get_col_max_val(self,event=None):
        c = self.col_sel
        try:
            self.max_entry.set_my_value(datetime.datetime.strftime(max(datetime.datetime.strptime(r[c],self.DATE_FORM)
                                                                       for r in self.sheet_ref if self.detect_date_form(r[c]) == self.DATE_FORM),self.DATE_FORM))
        except:
            pass

    def confirm(self,event=None):
        self.result = True
        self.min_val = self.min_entry.get_my_value()
        self.max_val = self.max_entry.get_my_value()
        self.order = self.asc_desc_dropdown.get_my_value()
        self.destroy()

    def cancel(self,event=None):
        self.destroy()
        

class edit_condition_frame(tk.Frame):
    def __init__(self,
                 parent,
                 condition,
                 colors,
                 color=None,
                 coltype="Text Detail",
                 confirm_text="Save condition",
                 theme = "dark blue"):
        tk.Frame.__init__(self,parent, height = 160, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(1,weight=1)
        if coltype in ("ID", "Parent", "Text Detail"):
            self.if_cell_label = label(self,text="If cell text is exactly:",font=EFB, theme = theme)
        else:
            self.if_cell_label = label(self,text="If cell value is:",font=EFB, theme = theme)
        self.if_cell_label.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(20,40))
        self.condition_display = condition_entry_with_scrollbar(self,coltype=coltype, theme = theme)
        self.condition_display.set_my_value(condition)
        self.condition_display.grid(row=0,column=1,sticky="nswe",pady=(20,20),padx=(0,0))
        self.color_dropdown = ez_dropdown(self,EF)
        self.color_dropdown['values'] = colors
        if color is None:
            self.color_dropdown.set_my_value(colors[0])
        else:
            self.color_dropdown.set_my_value(color)
        self.color_dropdown.grid(row=0,column=2,sticky="nswe",pady=(20,20),padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=1,column=0,columnspan=3,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text=confirm_text,style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(0,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(0,20))
        self.result = False
        self.new_condition = ""
        self.color_dropdown.bind("<<ComboboxSelected>>",self.disable_cancel)
        self.condition_display.place_cursor()
        
    def disable_cancel(self,event=None):
        try:
            self.cancel_button.config(state="disabled")
            self.after(300,self.enable_cancel)
        except:
            pass
    def enable_cancel(self,event=None):
        self.cancel_button.config(state="normal")
        
    def confirm(self,event=None):
        self.result = True
        self.new_condition = self.condition_display.get_my_value()
        self.color = self.color_dropdown.get_my_value()
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()


class edit_formula_frame(tk.Frame):
    def __init__(self,parent,colname,formula,type_,formula_apply_only,theme = "dark blue"):
        tk.Frame.__init__(self,parent,height=215, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.type_ = type_
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(2,weight=1)
        self.col_name_display = readonly_entry_with_scrollbar(self, theme = theme)
        self.col_name_display.set_my_value(colname)
        self.col_name_display.grid(row=0,column=0,columnspan=4,sticky="nswe",pady=(20,20),padx=(20,20))
        self.equals_label = label(self,text=" = ",font=EFB, theme = theme)
        self.equals_label.grid(row=1,column=0,sticky="nswe",padx=(20,10),pady=(0,20))
        self.formula_display = formula_entry_with_scrollbar(self, theme = theme)
        self.formula_display.set_my_value(formula)
        self.formula_display.grid(row=1,column=1,sticky="nswe",pady=(0,20),padx=(0,5))
        self.formula_only_apply = label(self,text="Only apply formula\nif no cells are blank",font=EFB, theme = theme)
        self.formula_only_apply.grid(row=1,column=2,sticky="nswe",padx=(0,5),pady=(0,20))
        self.formula_only_apply = ez_dropdown(self,EF)
        self.formula_only_apply['values'] = ["True","False"]
        self.formula_only_apply.set_my_value(f"{formula_apply_only}")
        self.formula_only_apply.grid(row=1,column=3,sticky="nswe",pady=(0,20),padx=(0,20))
        self.formula_only_apply.bind("<<ComboboxSelected>>",lambda event: self.formula_display.focus_set())
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=2,column=0,columnspan=4,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Save formula",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(0,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(0,20))
        self.result = False
        self.new_formula = ""
        self.formula_display.place_cursor()
        
    def confirm(self, event = None):
        self.result = True
        self.formula_only_apply_result = self.formula_only_apply.get_my_value()
        self.new_formula = self.formula_display.get_my_value()
        self.destroy()

    def cancel(self, event = None):
        self.destroy()


class edit_validation_frame(tk.Frame):
    def __init__(self,parent,coltype,colname,validation, theme = "dark blue"):
        tk.Frame.__init__(self, parent, height = 210)
        self.grid_propagate(False)
        self.C = parent
        self.config(bg = theme_bg(theme))
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(2,weight=1)
        self.col_name_display = readonly_entry_with_scrollbar(self, theme = theme)
        self.col_name_display.set_my_value(colname)
        self.col_name_display.grid(row=0,column=0,columnspan=2,sticky="nswe",pady=(20,20),padx=(20,20))
        self.valid_label = label(self,text=" Valid details:",font=EFB, theme = theme)
        self.valid_label.grid(row=1,column=0,sticky="nswe",padx=(20,10),pady=(0,40))
        if coltype == "Text Detail":
            self.validation_display = entry_with_scrollbar(self, theme = theme)
        else:
            self.validation_display = validation_entry_with_scrollbar(self, coltype, theme = theme)
        if validation:
            self.validation_display.set_my_value(",".join(validation))
            self.validation_display.place_cursor()
        elif not validation:
            try:
                self.validation_display.disable_checking()
            except:
                pass
            x = (
"""Type in values/text separated by commas ',' which will appear in a dropdown box when a user attempts to edit a cell in this column. To set an non-value
either put a comma at the start or a double comma wherever you want a non-value to show in the dropdown box.
For Numerical Detail columns you can set a range of values by using '_' like so: 'start_stop_step' e.g. '5_0_-1' which would create 5,4,3,2,1,0
For Date Detail columns and to limit user entry to only UK working day dates; enter: only uk working days"""
            )
            self.validation_display.set_my_value(" ".join(x.split("\n")))
        self.validation_display.grid(row=1,column=1,sticky="nswe",pady=(0,20),padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=2,column=0,columnspan=2,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Save validation",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(0,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(0,20))
        self.result = False
        self.new_validation = ""
        if not validation:
            self.validation_display.my_entry.bind("<FocusIn>",self.show_remove_help)
            
    def show_remove_help(self,event=None):
        self.validation_display.set_my_value("")
        self.validation_display.my_entry.unbind("<FocusIn>")
        try:
            self.validation_display.enable_checking()
        except:
            pass
        
    def confirm(self,event=None):
        self.result = True
        self.new_validation = self.validation_display.get_my_value()
        if self.new_validation.lower() in ("only uk working days", "only england working days", "only wales working days", "only scotland working days", "only ni working days"):
            self.new_validation = self.new_validation.lower()
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()


class validation_entry_with_scrollbar(tk.Frame):
    def __init__(self,parent,coltype,theme = "dark blue"):
        tk.Frame.__init__(self,parent)
        self.config(bg = theme_bg(theme))
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = validation_normal_entry(self, font = EF, coltype = coltype, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)
        
    def disable_checking(self):
        self.my_entry.disable_checking()
        
    def enable_checking(self):
        self.my_entry.enable_checking()


class validation_normal_entry(tk.Entry):
    def __init__(self, parent, font, coltype, width_ = None, theme = "dark blue"):
        tk.Entry.__init__(self, parent, font = font,
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme))
        if width_:
            self.config(width=width_)
        if coltype == "Numerical Detail":
            self.allowed_chars = {"0","1","2","3","4","5","6","7","8","9",",","-","_","."}
        if coltype == "Date Detail":
            self.allowed_chars = {"0","1","2","3","4","5","6","7","8","9",",","/","-",
                                  "o","O","n","N","l","L","y","Y","u","U","k","K","c","C","t","T","w","W","e","E",
                                  "w","W","o","O","r","R","k","K","i","I","n","N",
                                  "g","G","d","D","a","A","y","Y","s","S"," "}
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def validate_(self,sv):
        self.sv.set("".join([c.lower() for c in self.sv.get() if c in self.allowed_chars]))
        
    def disable_checking(self):
        self.config(textvariable="")
        
    def enable_checking(self):
        self.config(textvariable=self.sv)
        self.sv.set("".join([c.lower() for c in self.sv.get() if c in self.allowed_chars]))
        
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        
    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(get_platform_rc_binding())


class condition_entry_with_scrollbar(tk.Frame):
    def __init__(self,parent,coltype="Text Detail", theme = "dark blue"):
        tk.Frame.__init__(self,parent)
        self.config(bg = theme_bg(theme))
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = condition_normal_entry(self,font=EF,coltype=coltype, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)


class condition_normal_entry(tk.Entry):
    def __init__(self, parent, font, coltype = "Text Detail", width_ = None, theme = "dark blue"):
        tk.Entry.__init__(self, parent, font = font,
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme))
        if width_:
            self.config(width=width_)
        self.coltype = coltype
        if self.coltype not in ("ID", "Parent", "Text Detail"):
            self.validate_text = True
        else:
            self.validate_text = False
        if coltype != "Date Detail":
            self.allowed_chars = {"a","n","d","o","r","A","N","D","O","R","0","1","2","3","4","5","6","7","8","9","!",">","<","="," ","-",".","C","c"}
        else:
            self.allowed_chars = {"a","n","d","o","r","A","N","D","O","R","0","1","2","3","4","5","6","7","8","9","!",">","<","="," ","/","C","c","-"}
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        self.set_my_value(" ")
        
    def validate_(self,sv):
        if self.validate_text:
            self.sv.set("".join([c.lower() for c in self.sv.get().replace("  "," ") if c in self.allowed_chars]))
            
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        
    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(get_platform_rc_binding())


class formula_entry_with_scrollbar(tk.Frame):
    def __init__(self, parent, theme = "dark blue"):
        tk.Frame.__init__(self, parent, bg = theme_bg(theme))
        self.C = parent
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = formula_normal_entry(self, font = EF, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)


class formula_normal_entry(tk.Entry):
    def __init__(self, parent, font, width_ = None, theme = "dark blue"):
        tk.Entry.__init__(self, parent, font = font,
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme))
        if width_:
            self.config(width = width_)
        self.C = parent
        self.allowed_chars = {"d","D","c","C","0","1","2","3","4","5","6","7","8","9","+","(",")","*","-","/","%",".","^"}
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        try:
            if self.C.C.type_ != "Text Detail":
                self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        except:
            self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def validate_(self,sv):
        self.sv.set("".join([c.lower() for c in self.sv.get() if c in self.allowed_chars]))
        
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        
    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(get_platform_rc_binding())


class askconfirm_frame(tk.Frame):
    def __init__(self,parent,action,confirm_text="Confirm",cancel_text="Cancel",bgcolor="green",fgcolor="white", theme = "dark blue"):
        tk.Frame.__init__(self,parent,height=150, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.action_label = label(self,text="Action:",font=EF, theme = theme)
        self.action_label.config(background=bgcolor,foreground=fgcolor)
        self.action_label.grid(row=0,column=0,sticky="nswe",padx=20)
        self.action_display = readonly_entry_with_scrollbar(self, theme = theme)
        self.action_display.my_entry.config(font=ERR_ASK_FNT)
        self.action_display.set_my_value(action)
        self.action_display.grid(row=0,column=1,sticky="nswe",pady=(20,5),padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=1,column=0,columnspan=2,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text=confirm_text,style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=20)
        self.cancel_button = button(self.button_frame,text=cancel_text,style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=20)
        self.boolean = False
        self.action_display.place_cursor()
        
    def confirm(self,event=None):
        self.boolean = True
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()


class error_frame(tk.Frame):
    def __init__(self,parent,msg, theme = "dark blue"):
        tk.Frame.__init__(self,parent,height=150, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = parent
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.errorlabel = label(self,text="Error\nmessage:",font=EF, theme = theme)
        self.errorlabel.config(background="red",foreground="white")
        self.errorlabel.grid(row=0,column=0,sticky="nswe",padx=20)
        self.error_display = readonly_entry_with_scrollbar(self, theme = theme)
        self.error_display.my_entry.config(font=ERR_ASK_FNT)
        self.error_display.set_my_value(msg)
        self.error_display.grid(row=0,column=1,sticky="nswe",pady=(20,5),padx=(0,20))
        self.confirm_button = button(self,text="Okay",style="TF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=1,column=0,columnspan=2,sticky="nswe",padx=20,pady=(10,20))
    def confirm(self,event=None):
        self.destroy()

        
class treeview(ttk.Treeview):
    def __init__(self,parent,selectmode,style):
        ttk.Treeview.__init__(self,
                              parent,
                              selectmode=selectmode,
                              style=style)
        self.new_xview = tk.XView.xview
    def xview(self,*args):
        if "units" in args:
            if args[1] == "1":
                return self.new_xview(self,"scroll","27","units")
            elif args[1] == "-1":
                return self.new_xview(self,"scroll","-27","units")
        else:
            return self.new_xview(self,*args)


class working_text(tk.Text):
    def __init__(self,parent,wrap,font=("Calibri",12), theme = "dark blue", use_entry_bg = True, override_bg = None):
        tk.Text.__init__(self,
                         parent,
                         wrap=wrap,
                         font=font,
                         spacing1=5,spacing2=5)
        self.config(bg = theme_entry_bg(theme) if use_entry_bg else theme_bg(theme),
                    fg = theme_entry_fg(theme) if use_entry_bg else theme_fg(theme),
                    insertbackground = theme_entry_fg(theme) if use_entry_bg else theme_fg(theme))
        if override_bg is not None:
            self.config(bg = override_bg)
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"


class display_text(tk.Frame):
    def __init__(self, parent, text = "", theme = "dark blue"):
        tk.Frame.__init__(self, parent, bg = theme_bg(theme))
        self.C = parent
        self.grid_rowconfigure(0, weight = 1)
        self.grid_columnconfigure(0, weight = 1)
        self.textbox = working_text(self, wrap = "word", theme = theme, use_entry_bg = False)
        self.textbox.grid_propagate(False)
        self.grid_propagate(False)
        self.yscrollb = scrollbar(self,
                                  self.textbox.yview,
                                  "vertical",
                                  self.textbox)
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, text)
        self.textbox.config(state = "disabled", relief = "flat")
        self.textbox.grid(row = 0, column = 0, sticky = "nswe")
        self.yscrollb.grid(row = 0, column = 1, sticky = "ns")
        
    def place_cursor(self, index = None):
        if not index:
            self.textbox.focus_set()
            
    def get_my_value(self):
        return self.textbox.get("1.0", "end")
    
    def set_my_value(self, value):
        self.textbox.config(state = "normal")
        self.textbox.delete(1.0, "end")
        self.textbox.insert(1.0, value)
        self.textbox.config(state = "readonly")
        
    def change_my_state(self, new_state):
        self.current_state = new_state
        self.textbox.config(state = self.current_state)
        
    def change_my_width(self, new_width):
        self.textbox.config(width = new_width)
        
    def change_my_height(self, new_height):
        self.textbox.config(height = new_height)


class wrapped_text_with_find_and_yscroll(tk.Frame):
    def __init__(self,parent,text,current_state,height=None, theme = "dark blue"):
        tk.Frame.__init__(self, parent, bg = theme_bg(theme))
        self.C = parent
        self.theme = theme
        self.grid_rowconfigure(1,weight=1)
        self.grid_columnconfigure(0,weight=1)
        self.current_state = current_state
        self.word = ""
        self.find_results = []
        self.results_number = 0
        self.addon = ""
        self.find_frame = frame(self, theme = theme)
        self.find_frame.grid(row=0,column=0,columnspan=2,sticky="nswe")
        self.save_text_button = button(self.find_frame,text="Save as",
                                       style="Std.TButton",
                                       command=self.save_text)
        self.save_text_button.pack(side="left",fill="x")
        self.find_button = button(self.find_frame,text="Find: ",
                                  style="Std.TButton",
                                  command=self.find)
        self.find_button.pack(side="left",fill="x")
        self.find_window = normal_entry(self.find_frame,font=BF, theme = theme)
        self.find_window.bind("<Return>",self.find)
        self.find_window.pack(side="left",fill="x",expand=True)
        self.find_reset_button = button(self.find_frame,text="X",style="Std.TButton",command=self.find_reset)
        self.find_reset_button.pack(side="left",fill="x")
        self.find_results_label = label(self.find_frame,"0/0",BF, theme = theme)
        self.find_results_label.pack(side="left",fill="x")
        self.find_up_button = button(self.find_frame,text="▲",style="Std.TButton",command=self.find_up)
        self.find_up_button.pack(side="left",fill="x")
        self.find_down_button = button(self.find_frame,text="▼",style="Std.TButton",command=self.find_down)
        self.find_down_button.pack(side="left",fill="x")
        self.textbox = working_text(self,wrap="word", theme = theme)
        if height:
            self.textbox.config(height=height)
        self.yscrollb = scrollbar(self,
                                  self.textbox.yview,
                                  "vertical",
                                  self.textbox)
        self.textbox.delete(1.0,"end")
        self.textbox.insert(1.0,text)
        self.textbox.config(state=self.current_state)
        self.textbox.grid(row=1,column=0,sticky="nswe")
        self.yscrollb.grid(row=1,column=1,sticky="ns")
        
    def place_cursor(self,index=None):
        if not index:
            self.textbox.focus_set()
            
    def get_my_value(self):
        return self.textbox.get("1.0","end")
    
    def set_my_value(self,value):
        self.textbox.config(state="normal")
        self.textbox.delete(1.0,"end")
        self.textbox.insert(1.0,value)
        self.textbox.config(state=self.current_state)
        
    def change_my_state(self,new_state):
        self.current_state = new_state
        self.textbox.config(state=self.current_state)
        
    def change_my_width(self,new_width):
        self.textbox.config(width=new_width)
        
    def change_my_height(self,new_height):
        self.textbox.config(height=new_height)
        
    def save_text(self):
        newfile = filedialog.asksaveasfilename(parent=self,
                                               title="Save text on popup window",
                                               filetypes=[('Text File','.txt'),('CSV File','.csv')],
                                               defaultextension=".txt",
                                               confirmoverwrite=True)
        if not newfile:
            return
        newfile = os.path.normpath(newfile)
        if not newfile.lower().endswith((".csv",".txt")):
            errorpopup = error(self.C,"Can only save .csv/.txt files", theme = self.theme)
            return
        self.save_text_button.change_text("Saving...")
        try:
            with open(newfile,"w") as fh:
                fh.write(self.textbox.get("1.0","end")) #remove last newline? [:-2]
        except:
            errorpopup = error(self.C,"Error saving file", theme = self.theme)
            self.save_text_button.change_text("Save as")
            return
        self.save_text_button.change_text("Save as")
        
    def find(self,event=None):
        self.find_reset(True)
        self.word = self.find_window.get()
        if not self.word:
            return
        self.addon = "+" + str(len(self.word)) + "c"
        start = "1.0"
        while start:
            start = self.textbox.search(self.word,index=start,nocase=1,stopindex="end")
            if start:
                end = start + self.addon
                self.find_results.append(start)
                self.textbox.tag_add("i",start,end)
                start = end
        if self.find_results:
            self.textbox.tag_config("i",background="Yellow")
            self.find_results_label.config(text="1/"+str(len(self.find_results)))
            self.textbox.tag_add("c",self.find_results[self.results_number],self.find_results[self.results_number]+self.addon)
            self.textbox.tag_config("c",background="Orange")
            self.textbox.see(self.find_results[self.results_number])
            
    def find_up(self,event=None):
        if not self.find_results or len(self.find_results) == 1:
            return
        self.textbox.tag_remove("c",self.find_results[self.results_number],self.find_results[self.results_number]+self.addon)
        if self.results_number == 0:
            self.results_number = len(self.find_results) - 1
        else:
            self.results_number -= 1
        self.find_results_label.config(text=str(self.results_number+1)+"/"+str(len(self.find_results)))
        self.textbox.tag_add("c",self.find_results[self.results_number],self.find_results[self.results_number]+self.addon)
        self.textbox.tag_config("c",background="Orange")
        self.textbox.see(self.find_results[self.results_number])
        
    def find_down(self,event=None):
        if not self.find_results or len(self.find_results) == 1:
            return
        self.textbox.tag_remove("c",self.find_results[self.results_number],self.find_results[self.results_number]+self.addon)
        if self.results_number == len(self.find_results) - 1:
            self.results_number = 0
        else:
            self.results_number += 1
        self.find_results_label.config(text=str(self.results_number+1)+"/"+str(len(self.find_results)))
        self.textbox.tag_add("c",self.find_results[self.results_number],self.find_results[self.results_number]+self.addon)
        self.textbox.tag_config("c",background="Orange")
        self.textbox.see(self.find_results[self.results_number])
        
    def find_reset(self,newfind=False):
        self.find_results = []
        self.results_number = 0
        self.addon = ""
        if newfind == False:
            self.find_window.delete(0,"end")
        for tag in self.textbox.tag_names():
            self.textbox.tag_delete(tag)
        self.find_results_label.config(text="0/0")


class scrollbar(ttk.Scrollbar):
    def __init__(self,parent,command,orient,widget):
        ttk.Scrollbar.__init__(self,
                               parent,
                               command=command,
                               orient=orient)
        self.orient = orient
        self.widget = widget
        if self.orient == "vertical":
            self.widget.configure(yscrollcommand=self.set)
        elif self.orient == "horizontal":
            self.widget.configure(xscrollcommand=self.set)


class readonly_entry(tk.Entry):
    def __init__(self, parent, font, width_ = None, theme = "dark blue"):
        tk.Entry.__init__(self, parent, font = font, state = "readonly",
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme))
        if width_:
            self.config(width=width_)
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.config(state="normal")
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        self.config(state="readonly")

    def change_theme(self, theme = "dark blue"):
        self.config(background = theme_entry_bg(theme),
                      foreground = theme_entry_fg(theme),
                      disabledbackground = theme_entry_dbg(theme),
                      disabledforeground = theme_entry_dfg(theme),
                      insertbackground = theme_entry_cursor(theme),
                      readonlybackground = theme_entry_dbg(theme))


class normal_entry(tk.Entry):
    def __init__(self,parent,font,width_=None,relief="sunken",
                 border=1,textvariable=None, theme = "dark blue"):
        tk.Entry.__init__(self,parent,font=font,relief=relief,
                          border=border,textvariable=textvariable,
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme)
                          )
        if width_:
            self.config(width=width_)
        if textvariable:
            self.config(textvariable=textvariable)
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        
    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(get_platform_rc_binding())
        

class readonly_entry_with_scrollbar(tk.Frame):
    def __init__(self,parent,font=EF, theme = "dark blue"):
        tk.Frame.__init__(self,parent, bg = theme_bg(theme)) 
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = readonly_entry(self, font = font, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)

    def change_text(self, text = ""):
        self.my_entry.set_my_value(text)

    def change_theme(self, theme = "dark blue"):
        self.config(bg = theme_bg(theme))
        self.my_entry.change_theme(theme)


class entry_with_scrollbar(tk.Frame):
    def __init__(self, parent, theme = "dark blue"):
        tk.Frame.__init__(self, parent, bg = theme_bg(theme))
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = normal_entry(self, font = EF, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)


class ez_dropdown(ttk.Combobox):
    def __init__(self,parent,font,width_=None):
        self.displayed = tk.StringVar()
        ttk.Combobox.__init__(self,parent,
                              font=font,
                              state="readonly",
                              textvariable=self.displayed)
        if width_:
            self.config(width=width_)
            
    def get_my_value(self,event=None):
        return self.displayed.get()
    
    def set_my_value(self,value,event=None):
        self.displayed.set(value)


class numerical_entry_with_scrollbar(tk.Frame):
    def __init__(self, parent, theme = "dark blue"):
        tk.Frame.__init__(self, parent, background = theme_bg(theme))
        self.grid_columnconfigure(0,weight=1)
        self.grid_rowconfigure(0,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.my_entry = numerical_normal_entry(self, font = EF, theme = theme)
        self.my_entry.grid(row=0,column=0,sticky="nswe")
        self.my_scrollbar = scrollbar(self,self.my_entry.xview,
                                      "horizontal",
                                      self.my_entry)
        self.my_scrollbar.grid(row=1,column=0,sticky="ew")
        
    def change_my_state(self,state,event=None):
        self.my_entry.config(state=state)
        
    def place_cursor(self,event=None):
        self.my_entry.focus_set()
        
    def get_my_value(self,event=None):
        return self.my_entry.get()
    
    def set_my_value(self,val,event=None):
        self.my_entry.set_my_value(val)


class numerical_normal_entry(tk.Entry):
    def __init__(self, parent, font, width_ = None, theme = "dark blue"):
        tk.Entry.__init__(self,
                          parent,
                          font = font,
                          background = theme_entry_bg(theme),
                          foreground = theme_entry_fg(theme),
                          disabledbackground = theme_entry_dbg(theme),
                          disabledforeground = theme_entry_dfg(theme),
                          insertbackground = theme_entry_cursor(theme),
                          readonlybackground = theme_entry_dbg(theme))
        if width_:
            self.config(width=width_)
        self.allowed_chars = {"0","1","2","3","4","5","6","7","8","9","."}
        self.sv = tk.StringVar()
        self.config(textvariable=self.sv)
        self.sv.trace("w", lambda name, index, mode, sv=self.sv: self.validate_(self.sv))
        self.rc_popup_menu = tk.Menu(self,tearoff=0,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Select all",accelerator="Ctrl+A",command=self.select_all,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Cut",accelerator="Ctrl+X",command=self.cut,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Copy",accelerator="Ctrl+C",command=self.copy,**menu_kwargs)
        self.rc_popup_menu.add_command(label="Paste",accelerator="Ctrl+V",command=self.paste,**menu_kwargs)
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def validate_(self,sv):
        x = self.sv.get()
        dotidx = [i for i,c in enumerate(x) if c == "."]
        if dotidx:
            if len(dotidx) > 1:
                x = x[:dotidx[1]]
        if x.startswith("."):
            x = x[1:]
        if x.startswith("-"):
            self.sv.set("-" + "".join([c for c in x if c in self.allowed_chars]))
        else:
            self.sv.set("".join([c for c in x if c in self.allowed_chars]))
            
    def rc(self,event):
        self.focus_set()
        self.rc_popup_menu.tk_popup(event.x_root,event.y_root)
        
    def select_all(self,event=None):
        self.event_generate("<Control-a>")
        return "break"
    
    def cut(self,event=None):
        self.event_generate("<Control-x>")
        return "break"
    
    def copy(self,event=None):
        self.event_generate("<Control-c>")
        return "break"
    
    def paste(self,event=None):
        self.event_generate("<Control-v>")
        return "break"
    
    def set_my_value(self,newvalue):
        self.delete(0,"end")
        self.insert(0,str(newvalue))
        
    def enable_me(self):
        self.config(state="normal")
        self.bind("<1>",lambda event: self.focus_set())
        self.bind(get_platform_rc_binding(), self.rc)
        
    def disable_me(self):
        self.config(state="disabled")
        self.unbind("<1>")
        self.unbind(get_platform_rc_binding())


class date_entry(tk.Frame):
    def __init__(self,parent,DATE_FORM, theme = "dark blue"):
        tk.Frame.__init__(self,
                          parent,
                          relief="flat",
                          bg = theme_bg(theme),
                          highlightbackground = theme_fg(theme),
                          highlightthickness = 2,
                          border=2)
        self.C = parent

        self.allowed_chars = {"0","1","2","3","4","5","6","7","8","9"}
        self.sv_1 = tk.StringVar()
        self.sv_2 = tk.StringVar()
        self.sv_3 = tk.StringVar()
        self.entry_1 = normal_entry(self,font=("Calibri",30,"bold"),width_=4,relief="flat",border=0,
                                    textvariable=self.sv_1, theme = theme)
        self.sep = "/" if "/" in DATE_FORM else "-"
        self.label_1 = tk.Label(self,font=("Calibri",30,"bold"),text=self.sep,
                                background=theme_entry_bg(theme),relief="flat")
        
        self.entry_2 = normal_entry(self,font=("Calibri",30,"bold"),width_=2,relief="flat",border=0,
                                    textvariable=self.sv_2, theme = theme)
        
        self.label_2 = tk.Label(self,font=("Calibri",30,"bold"),text=self.sep,
                                background=theme_entry_bg(theme),relief="flat")
        
        self.entry_3 = normal_entry(self,font=("Calibri",30,"bold"),width_=2,relief="flat",border=0,
                                    textvariable=self.sv_3, theme = theme)
        
        self.sv_1.trace("w", lambda name, index, mode, sv=self.sv_1: self.validate_1(self.sv_1))
        self.sv_2.trace("w", lambda name, index, mode, sv=self.sv_2: self.validate_2(self.sv_2))
        self.sv_3.trace("w", lambda name, index, mode, sv=self.sv_3: self.validate_3(self.sv_3))
        self.entry_1.bind("<BackSpace>",self.e1_back)
        self.entry_2.bind("<BackSpace>",self.e2_back)
        self.entry_3.bind("<BackSpace>",self.e3_back)

        if DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            self.entry_3.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_2.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_1.pack(side="left")
            self.entries = [self.entry_3,self.entry_2,self.entry_1]
        elif DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            self.entry_1.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_2.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_3.pack(side="left")
            self.entries = [self.entry_1,self.entry_2,self.entry_3]
        elif DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            self.entry_2.pack(side="left")
            self.label_2.pack(side="left")
            self.entry_3.pack(side="left")
            self.label_1.pack(side="left")
            self.entry_1.pack(side="left")
            self.entries = [self.entry_2,self.entry_3,self.entry_1]

        self.DATE_FORM = DATE_FORM
        
    def e1_back(self,event):
        x = self.sv_1.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if not x:
                self.entry_3.icursor(2)
                self.entry_3.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_1.set(x)
        return "break"

    def e2_back(self,event):
        x = self.sv_2.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if not x:
                self.entry_3.icursor(2)
                self.entry_3.focus_set()
                return "break"
            elif len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if not x:
                self.entry_1.icursor(4)
                self.entry_1.focus_set()
                return "break"
            elif len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_2.set(x)
        return "break"
            
    def e3_back(self,event):
        x = self.sv_3.get()
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            if not x:
                self.entry_2.icursor(2)
                self.entry_2.focus_set()
                return "break"
            if len(x) == 1:
                x = ""
            else:
                x = x[:-1]
            self.sv_3.set(x)
        return "break"

    def validate_1(self,sv):
        year = []
        for i,c in enumerate(self.sv_1.get()):
            if c in self.allowed_chars:
                year.append(c)
            if i > 3:
                break
        year = "".join(year)
        if len(year) > 4:
            year = year[:4]
        self.entry_1.set_my_value(year)
        if len(year) == 4:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                pass
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                self.entry_2.focus_set()
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                pass
        
    def validate_2(self,sv):
        month = []
        for i,c in enumerate(self.sv_2.get()):
            if c in self.allowed_chars:
                month.append(c)
            if i > 1:
                break
        move_cursor = False
        if len(month) > 2:
            month = month[:2]
        if not month:
            self.entry_2.set_my_value("")
            return
        if len(month) == 1 and int(month[0]) > 1:
            month = ["0",month[0]]
            move_cursor = True
        elif len(month) == 1 and int(month[0]) <= 1:
            return
        e0 = int(month[0])
        e1 = int(month[1])
        if e0 > 1:
            month[0] = "1"
            int(month[0])
        if e1 > 2 and e0 > 0:
            month[0] = "0"
            month[1] = str(e1)
        self.entry_2.set_my_value("".join(month))
        if move_cursor:
            self.entry_2.icursor(2)
        if len(month) == 2:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                self.entry_1.focus_set()
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                self.entry_3.focus_set()
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                self.entry_3.focus_set()
            
    def validate_3(self,sv):
        day = []
        for i,c in enumerate(self.sv_3.get()):
            if c in self.allowed_chars:
                day.append(c)
            if i > 1:
                break
        move_cursor = False
        if len(day) > 2:
            day = day[:2]
        if not day:
            self.entry_3.set_my_value("")
            return
        if len(day) == 1 and int(day[0]) > 3:
            day = ["0",day[0]]
            move_cursor = True
        elif len(day) == 1 and int(day[0]) <= 3:
            return
        e0 = int(day[0])
        e1 = int(day[1])
        if e0 > 3:
            day[0] = "1"
            int(day[0])
        if e0 >= 3 and e1 > 1:
            day[0] = "0"
            day[1] = str(e1)
        self.entry_3.set_my_value("".join(day))
        if move_cursor:
            self.entry_3.icursor(2)
        if len(day) == 2:
            if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
                self.entry_2.focus_set()
            if self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
                pass
            elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
                self.entry_1.focus_set()
        
    def get_my_value(self):
        return self.sep.join([e.get() for e in self.entries])

    def set_my_value(self,date):
        date = re.split('|'.join(map(re.escape, ("/", "-"))), date)
        if self.DATE_FORM in ("%d/%m/%Y", "%d-%m-%Y"):
            try:
                self.sv_1.set(date[2])
            except:
                pass
            try:
                self.sv_2.set(date[1])
            except:
                pass
            try:
                self.sv_3.set(date[0])
            except:
                pass
        elif self.DATE_FORM in ("%Y/%m/%d", "%Y-%m-%d"):
            try:
                self.sv_1.set(date[0])
            except:
                pass
            try:
                self.sv_2.set(date[1])
            except:
                pass
            try:
                self.sv_3.set(date[2])
            except:
                pass
        elif self.DATE_FORM in ("%m/%d/%Y", "%m-%d-%Y"):
            try:
                self.sv_1.set(date[2])
            except:
                pass
            try:
                self.sv_2.set(date[0])
            except:
                pass
            try:
                self.sv_3.set(date[1])
            except:
                pass
        
    def place_cursor(self,index=0):
        self.entries[index].focus()


class rename_column_frame(tk.Frame):
    def __init__(self, C, current_col_name, type_of_col, theme = "dark blue"):
        tk.Frame.__init__(self, C, height = 190, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = C
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(2,weight=1)
        self.col_label = label(self,text="Current column\nname:",font=EF, theme = theme)
        self.col_label.grid(row=0,column=0,sticky="nswe",padx=20)
        self.col_display = readonly_entry_with_scrollbar(self, theme = theme)
        if type_of_col == "detail":
            self.col_display.set_my_value("Detail column: " + current_col_name)
        elif type_of_col == "hierarchy":
            self.col_display.set_my_value("Parent column: " + current_col_name)
        elif type_of_col == "ID":
            self.col_display.set_my_value("ID column: " + current_col_name)
        self.col_display.grid(row=0,column=1,sticky="nswe",pady=(20,5),padx=(0,20))
        self.new_name_label = label(self,text="New column\nname:",font=EF, theme = theme)
        self.new_name_label.grid(row=1,column=0,sticky="nswe",padx=20)
        self.new_name_display = entry_with_scrollbar(self, theme = theme)
        self.new_name_display.grid(row=1,column=1,sticky="nswe",pady=5,padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=2,column=0,columnspan=2,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Confirm",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(10,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(10,20))
        self.result = False
        self.new_name_display.place_cursor()
        
    def confirm(self,event=None):
        self.result = "".join(self.new_name_display.get_my_value().strip().split())
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()


class add_hierarchy_frame(tk.Frame):
    def __init__(self, C, theme = "dark blue"):
        tk.Frame.__init__(self, C, height = 150, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.C = C
        self.grid_columnconfigure(1,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.hier_name_label = label(self,text="New hierarchy\nname:",font=EF, theme = theme)
        self.hier_name_label.grid(row=0,column=0,sticky="nswe",padx=20)
        self.hier_name_display = entry_with_scrollbar(self, theme = theme)
        self.hier_name_display.grid(row=0,column=1,sticky="nswe",pady=(20,5),padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=1,column=0,columnspan=2,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Add Hierarchy Column",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(10,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(10,20))
        self.result = False
        self.hier_name_display.place_cursor()
        
    def confirm(self,event=None):
        self.result = "".join(self.hier_name_display.get_my_value().strip().split())
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()
        

class add_detail_column_frame(tk.Frame):
    def __init__(self,C, theme = "dark blue"):
        tk.Frame.__init__(self, C, height = 150, bg = theme_bg(theme))
        self.grid_propagate(False)
        self.grid_columnconfigure(2,weight=1)
        self.grid_rowconfigure(1,weight=1)
        self.type_display = ez_dropdown(self,EF)
        self.type_display['values'] = ("Text Detail","Numerical Detail","Date Detail")
        self.type_display.set_my_value("Text Detail")
        self.type_display.grid(row=0,column=0,sticky="nswe",padx=(20,0),pady=(20,5))
        self.type_display.bind("<<ComboboxSelected>>",lambda focus: self.detail_name_display.place_cursor())
        self.detail_name_label = label(self,text="New detail\ncolumn name:",font=EF, theme = theme)
        self.detail_name_label.grid(row=0,column=1,sticky="nswe",padx=20,pady=(20,5))
        self.detail_name_display = entry_with_scrollbar(self, theme = theme)
        self.detail_name_display.grid(row=0,column=2,sticky="nswe",pady=(20,5),padx=(0,20))
        self.button_frame = frame(self, theme = theme)
        self.button_frame.grid(row=1,column=0,columnspan=3,sticky="nswe")
        self.button_frame.grid_columnconfigure(0,weight=1,uniform="x")
        self.button_frame.grid_columnconfigure(1,weight=1,uniform="x")
        self.button_frame.grid_rowconfigure(0,weight=1)
        self.confirm_button = button(self.button_frame,text="Add Detail Column",style="EF.Std.TButton",command=self.confirm)
        self.confirm_button.grid(row=0,column=0,sticky="nswe",padx=(20,10),pady=(10,20))
        self.cancel_button = button(self.button_frame,text="Cancel",style="EF.Std.TButton",command=self.cancel)
        self.cancel_button.grid(row=0,column=1,sticky="nswe",padx=(10,20),pady=(10,20))
        self.result = False
        self.type_ = "Text Detail"
        self.formula = ""
        self.detail_name_display.place_cursor()
        
    def confirm(self,event=None):
        self.result = "".join(self.detail_name_display.get_my_value().strip().split())
        self.type_ = self.type_display.get()
        self.destroy()
        
    def cancel(self,event=None):
        self.destroy()


class frame(tk.Frame):
    def __init__(self,parent,background="white",highlightbackground="white",highlightthickness=0,theme = "dark blue"):
        tk.Frame.__init__(self,parent,
                          background=theme_bg(theme),
                          highlightbackground=highlightbackground,
                          highlightthickness=highlightthickness)
        

class button(ttk.Button):
    def __init__(self,parent,style="Std.TButton",text="",command=None,state="normal",underline=-1):
        ttk.Button.__init__(self,
                            parent,
                            style=style,
                            text=text,
                            command=command,
                            state=state,
                            underline=underline)
        
    def change_text(self,text):
        self.config(text=text)
        self.update_idletasks()


class StatusBar(tk.Label):
    def __init__(self, parent, text, theme = "dark blue"):
        tk.Label.__init__(self,
                          parent,
                          text = text,
                          font = ("Calibri",11,"normal"),
                          background = theme_bg(theme),
                          foreground = theme_status_fg(theme),
                          anchor = "w")
        self.text = text
        
    def change_text(self,text=""):
        self.config(text=text)
        self.text = text
        self.update_idletasks()


class label(tk.Label):
    def __init__(self,parent,text,font,theme = "dark blue", anchor = "center"):
        tk.Label.__init__(self,parent,
                          text=text,font=font,
                          background=theme_bg(theme),
                          foreground=theme_fg(theme),
                          anchor = anchor)
        
    def change_text(self,text):
        self.config(text=text)
        self.update_idletasks()


class displabel(tk.Label):
    def __init__(self,parent,text,font, theme = "dark blue"):
        tk.Label.__init__(self,parent,
                          background=theme_bg(theme),
                          text=text,font=font)
        self.config(anchor="w")
        
    def change_text(self,text):
        self.config(text=text)
        self.update_idletasks()




        
