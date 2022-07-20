# must install xlwings, pandas, openpyxl, cx_Oracle, dateparser, bs4, xlsxwriter
import tkinter as tk
from tkinter import *
from tkinter import filedialog, simpledialog
import os
from connect_to_db import get_connection
from SQLDataAppMethods import run_diagnostic, run_zb_files, get_pvs_sites
from run_zb_spec_sched import sort_and_print_sched_stops, validation


def get_sites(lines):
    sites_to_print = []
    lines_to_print = []
    for line in lines:
        if line.to_print:
            sites_to_print.append(line.site_name)
            lines_to_print.append(line)
    return sites_to_print, lines_to_print


# ************************
# Scrollable Frame Class
# ************************
class ScrollFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)  # create a frame (self)

        self.canvas = tk.Canvas(self, borderwidth=0, background="#ffffff")  # place canvas on self
        # place a frame on the canvas, this frame will hold the child widgets
        self.viewPort = tk.Frame(self.canvas, background="#ffffff")
        self.vsb = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)  # place a scrollbar on self
        self.canvas.configure(yscrollcommand=self.vsb.set)  # attach scrollbar action to scroll of canvas

        self.vsb.pack(side="right", fill="y")  # pack scrollbar to right of self
        self.canvas.pack(side="left", fill="both", expand=True)  # pack canvas to left of self and expand to fil
        self.canvas_window = self.canvas.create_window((4, 4), window=self.viewPort, anchor="nw",
                                                       # add view port frame to canvas
                                                       tags="self.viewPort")
        # bind an event whenever the size of the viewPort frame changes.
        self.viewPort.bind("<Configure>", self.onFrameConfigure)
        self.canvas.bind("<Configure>", self.onCanvasConfigure)

        # perform an initial stretch on render, otherwise the scroll region has a tiny border until the first resize
        self.onFrameConfigure(None)

    def onFrameConfigure(self, event):
        '''Reset the scroll region to encompass the inner frame'''
        self.canvas.configure(scrollregion=self.canvas.bbox(
            "all"))  # whenever the size of the frame changes, alter the scroll region respectively.

    def onCanvasConfigure(self, event):
        '''Reset the canvas window to encompass inner frame when required'''
        canvas_width = event.width
        # whenever the size of the canvas changes alter the window region respectively.
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)


# Class that creates a frame for each schedule that has a button with the schedule number and
# 3 radio buttons for delete, mvo and tto
class Checkbar(Frame):
    def __init__(self, parent=Frame, picks=[], anchor=W):
        Frame.__init__(self, parent)
        self.vars = []
        self.to_print = False
        self.site_name = picks[0]
        var = IntVar()
        var.set(0)
        chk = Radiobutton(self, text=picks[0], var=var, value=1, command=self.change)
        chk.pack(side=LEFT, anchor=anchor, expand=YES)
        type_var = StringVar(self)
        type_var.set("Schedule Type")
        type_i = OptionMenu(self, type_var, "In-Service", "Draft", "Future", "Discontinued")
        type_i.pack(side=LEFT, anchor=anchor, expand=YES)

        date_var = StringVar(self)
        # this is what shows up originally on the dropdown before selecting anything
        date_var.set("Date")
        date = OptionMenu(self, date_var, "03-03-3333", "02-02-2022", "04-04-4444")
        date.pack(side=LEFT, expand=YES)

        self.date_var = date_var
        self.type_var = type_var
        self.vars = [var]
        self.has_day = False

        Button(self, text='Add Specific Date', command=self.specify_date).pack(side=RIGHT, anchor=anchor)

        # method that adds a drop down with 1-31 so user can pick an exact day of the month to analyze if needed

    def specify_date(self):
        new_date = simpledialog.askstring('Custom Date', 'Please input date in MM-DD-YYYY format:')
        self.vars[-1].set(new_date)

    def change(self):
        vars_list = []
        for var in self.vars:
            vars_list.append(var.get())
        if vars_list[0] == 1 and not self.to_print:
            self.to_print = True
        elif vars_list[0] == 1 and self.to_print:
            self.to_print = False
            self.vars[0].set(0)
        elif vars_list[0] == 0 and not self.to_print:
            self.to_print = True
            self.vars[0].set(1)


# class which creates the main frame that holds all the other frames for the schedules
# must create this frame within a class to be able to get the pop-up scrollable
class MainFrame(tk.Frame):
    def __init__(self, root):
        tk.Frame.__init__(self, root)
        self.scrollFrame = ScrollFrame(self)  # add a new scrollable frame.
        # gets connection to get PVS sites
        pvs_sites_connection = get_connection()
        # returns list of PVS sites from dataframe just retrieved
        pvs_sites = get_pvs_sites(pvs_sites_connection)
        # closes connection in case of time out, connections attached to each method
        pvs_sites_connection.close()

        self.lines = []

        row_counter = 0
        while row_counter < len(pvs_sites):
            line = Checkbar(self.scrollFrame.viewPort, [pvs_sites[row_counter]])
            row_counter += 1
            line.pack(side=TOP, fill=X)
            line.config(relief=GROOVE, bd=10)
            self.lines.append(line)

        def quit():
            root.quit()

        def run_diag():
            diag_sites, lines_to_print = get_sites(self.lines)
            print('Deleting current files in HTML Check')
            for file in os.listdir('HTML Check/'):
                if file.endswith('.xls') or file.endswith('.xlsx') or file.endswith('.XLSX'):
                    os.remove(str('HTML Check/' + file))
            if len(diag_sites) >= 1:
                for line in lines_to_print:
                    print('Printing Diagnostic for: ', diag_sites)
                    run_diagnostic([line.site_name], line.type_var.get(), line.date_var.get())
                    print('Done! Diagnostic printed to DataFiles')
            else:
                print('You forgot to select a site!')

        def run_zb():
            zb_sites, zb_lines_to_print = get_sites(self.lines)
            conn = get_connection()
            if len(zb_sites) >= 1:
                print(f'Printing ZB files for {len(zb_sites)} sites: ', zb_sites)
                # if intvar is 0 that means print each site individually
                if separate.get() == 0:
                    # prints site by line because might have different types/effective dates per site
                    for line in zb_lines_to_print:
                        run_zb_files([line.site_name], line.type_var.get(), line.date_var.get(), True, conn)
                # otherwise, 1 means print together
                elif separate.get() == 1:
                    site_1 = zb_lines_to_print[0]
                    site_2 = zb_lines_to_print[1]
                    if site_1.type_var.get() == site_2.type_var.get() and \
                            site_1.date_var.get() == site_2.date_var.get():
                        run_zb_files([zb_sites], site_1.type_var.get(), site_1.date_var.get(), False, conn)
                    # else:
                    #     run_zb_together_diff_types()
                print('Done! Zero base files printed to DataFiles')
            else:
                print('You forgot to select a site!')
            # closing connection for minimum time connected to sql
            conn.close()
            print('Connection closed')

        def run_zb_sched():
            schedule_numbers_file = filedialog.askopenfilename()
            if schedule_numbers_file != '':
                schedules, stops, site, missing, typedate = sort_and_print_sched_stops(schedule_numbers_file)
                validation(schedules, stops, site, schedule_numbers_file, missing, typedate)
            else:
                print("Did not choose a file!")

        def reset():
            for line in self.lines:
                line.to_print = False
                line.vars[0].set(0)
            print('Selected sites reset')

        separate = IntVar()
        self.separate = separate
        # initializing the buttons for resetting, all mvo/tto and changing the actual htmls
        Button(root, text='Quit', command=quit).pack(side=BOTTOM)
        Button(root, text='Reset', command=reset).pack(side=BOTTOM)
        Button(root, text='Run Diagnostic', command=run_diag).pack(side=BOTTOM)
        Button(root, text='Run ZB Files by Sched #s', command=run_zb_sched).pack(side=BOTTOM)
        Radiobutton(root, text='Together', variable=separate, value=1).pack(side=BOTTOM)
        Radiobutton(root, text='Separate', variable=separate, value=0).pack(side=BOTTOM)
        Button(root, text='Run ZB Files', command=run_zb).pack(side=BOTTOM)

        # when packing the scrollframe, we pack scrollFrame itself (NOT the viewPort)
        self.scrollFrame.pack(side="top", fill="both", expand=True)


def launch_pop_up():
    root = tk.Tk()
    root.title('SQL Data App')
    root.geometry('500x400+400+100')
    # call MainFrame class which will create the scroll frame and Checkbars
    MainFrame(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
    root.quit()


if __name__ == "__main__":
    launch_pop_up()
