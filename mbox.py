# -*- coding: utf-8 -*-
# @Author: Some Guy Online
# @Date:   2017-09-22 15:48:07
# @Last Modified by:   GalenMoo
# @Last Modified time: 2017-10-20 09:17:04

from tkinter import ttk
from tkinter import *

class Mbox(object):

    root = None

    def __init__(self, msg, dict_key=None):
        """
        msg = <str> the message to be displayed
        dict_key = <sequence> (dictionary, key) to associate with user input
        (providing a sequence for dict_key creates an entry for user input)
        """
        self.top = Toplevel(Mbox.root)
        self.top.grab_set() 

        frm = ttk.Frame(self.top, borderwidth=4, relief='ridge')
        frm.pack(fill='both', expand=True)

        label = ttk.Label(frm, text=msg)
        label.pack(padx=4, pady=4)

        caller_wants_an_entry = dict_key is not None

        if caller_wants_an_entry:
            self.entry = ttk.Entry(frm)
            self.entry.pack(pady=4)

            b_submit = ttk.Button(frm, text='Submit')
            b_submit['command'] = lambda: self.entry_to_dict(dict_key)
            b_submit.pack()

        b_cancel = ttk.Button(frm, text='Cancel')
        b_cancel['command'] = self.top.destroy
        b_cancel.pack(padx=4, pady=4)
        self.top.bind('<Return>', self.destroyTop)
        if caller_wants_an_entry:
            self.dict_key = dict_key
            self.entry.focus_set()
            self.top.bind('<Return>', self.submitBind)
        else:
            b_cancel.focus_set()

    def destroyTop(self, event):
        self.top.destroy()

    def submitBind(self, event):
        self.entry_to_dict(self.dict_key)
        self.top.destroy()

    def entry_to_dict(self, dict_key):
        data = self.entry.get()
        if data:
            d, key = dict_key
            d[key] = data
            self.top.destroy()

class MyDialog(object):
    def __init__(self, parent):
        self.toplevel = Toplevel(parent)
        self.var = StringVar()
        label = ttk.Label(self.toplevel, text="Input Serial Number")
        om = ttk.Entry(self.toplevel, textvariable = self.var)
        button = ttk.Button(self.toplevel, text="OK", command=self.toplevel.destroy)
        label.pack(side="top", fill="x")
        om.pack(side="top", fill="x")
        button.pack()

    def show(self):
        self.toplevel.deiconify()
        self.toplevel.wait_window()
        value = self.var.get()
        return value