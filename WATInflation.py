import xlrd
import xlwt
from tkinter import *
import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


root = os.path.dirname(os.path.abspath(__file__))

# ['employer', 'surname', 'given_name', 'position', 'salary_paid', 'taxable_benefits', 'year']
UWSalaries_File = root + "/Salaries.xlsx"

# ['YEAR', 'CPI']
Inflation_CPI_File = root + "/Inflation_CPI.xls"

out_headers = ['employer', 'surname', 'given_name', 'position', 'salary_paid', 'taxable_benefits', 'year', 'CPI', 'salary_change']
output = root + "/Output/data.xls"

iconpath= root + "\Icon\Iconsmind-Outline-Coins.ico"


class Application:
    def __init__(self, master):
        master.title("UWaterloo Salary VS Inflation")
        master.geometry("500x300")
        master.configure(background='black')
        master.bind('<Return>', self.Lookup)
        master.iconbitmap(default=iconpath)
        self.search = StringVar()
        self.entry = Entry(master, textvariable=self.search, width=60).pack(pady=20)
        self.scrollbar = Scrollbar(master)
        self.listbox = Listbox(master, selectmode=SINGLE, height=12, width=75, yscrollcommand=self.scrollbar.set)
        self.listbox.bind('<<ListboxSelect>>', self.WriteData)
    def Lookup(self, master):
        initial_search = self.search.get()
        search_list = initial_search.split(' ')
        firstval, lastval = search_list[0].upper(), search_list[-1].upper()
        NameDct = {}
        with xlrd.open_workbook(UWSalaries_File) as r_sal:
            rs_sal = r_sal.sheet_by_index(0)
            for row in range(0, rs_sal.nrows):
                row_data = [rs_sal.cell_value(row ,col) for col in range(rs_sal.ncols)]
                if (firstval in row_data[1] or firstval in row_data[2]) and (lastval in row_data[1] or lastval in row_data[2]):
                    NameDct[row_data[2] + " " + row_data[1]] = [row_data[0], row_data[3]]
        self.listbox.pack(side=LEFT, fill=X, padx=10)
        self.scrollbar.pack(side=LEFT, fill=BOTH, padx=10, pady=23)
        self.scrollbar.config(command=self.listbox.yview)
        self.listbox.delete(0, END)
        for name in NameDct:
            name_position = name + ", " + NameDct[name][1]
            self.listbox.insert(END, name_position)
    def ReadData(self, specific_search):
        with xlrd.open_workbook(UWSalaries_File) as r_sal, xlrd.open_workbook(Inflation_CPI_File) as r_CPI:
            # Read UWSalaries_File and create a list of years associated with the search term, and a data_dct dictionary of all the values associated with search term.
            rs_sal = r_sal.sheet_by_index(0)
            year_lst, data_dct = [], {}
            for row in range(0, rs_sal.nrows):
                row_data = [rs_sal.cell_value(row ,col) for col in range(rs_sal.ncols)]
                if specific_search == (row_data[2] + " " + row_data[1]):
                    data_dct[int(row_data[6])] = row_data
                    year_lst.append(int(row_data[6]))
                    print(row_data)
            print(sorted(data_dct))
            # Read Inflation_CPI_File and append the CPI values up with each record in the data_dct dictionary.
            rs_CPI = r_CPI.sheet_by_index(0)
            for row in range(1, rs_CPI.nrows):
                if rs_CPI.cell_value(row, 0) in year_lst:
                    data_dct[rs_CPI.cell_value(row, 0)].append(float(rs_CPI.cell_value(row, 1)))
        return data_dct
    def WriteData(self, master):
        listboxindex = self.listbox.curselection()[0]
        self.specific_search = self.listbox.get(listboxindex).split(",")[0]
        data = self.ReadData(self.specific_search)
        # Add the salary_change
        old_salary = 0
        for rindex, key in enumerate(sorted(data)):
            if old_salary > 0:
                salary_change = ((data[key][4] - old_salary) / old_salary) * 100
            else:
                salary_change = data[key][7]
            old_salary = data[key][4]
            data[key].append(salary_change)
        # Write to xls file
        wb = xlwt.Workbook()
        ws = wb.add_sheet("Data")
        for cindex, val in enumerate(out_headers):
            ws.write(0, cindex, val)
        for rindex, key in enumerate(sorted(data)):
            for cindex, val in enumerate(data[key]):
                ws.write(rindex+1, cindex, val)
        wb.save(output)

        # Plotting using Pandas and Matplotlib below
        xl = pd.ExcelFile(output)
        df = xl.parse("Data")

        fig, axes = plt.subplots(nrows=2, ncols=1, figsize=(10,7), sharex=True)
        year_ticks = df['year'].tolist()
        salary_changes = df['salary_change'].tolist()
        CPIs = df['CPI'].tolist()

        data1 = df[['year', 'CPI', 'salary_change']]
        data1.set_index('year', inplace=True)
        styles1 = ['ro-','go-']
        axes[0].set_ylabel('Percent (%)')

        axes[0].plot(year_ticks, salary_changes, color='blue')
        axes[0].plot(year_ticks, CPIs, color='black')

        salary_changes_array = np.array(salary_changes)
        CPIs_array = np.array(CPIs)

        axes[0].fill_between(year_ticks, salary_changes, CPIs, where=CPIs_array <= salary_changes_array, facecolor='green', interpolate=True)
        axes[0].fill_between(year_ticks, salary_changes, CPIs, where=CPIs_array >= salary_changes_array, facecolor='red', interpolate=True)


        data2 = df[['year', 'salary_paid']]
        data2.set_index('year', inplace=True)
        axes[1].set_ylabel('Salary ($)')
        data2.plot(ax=axes[1], style='bo-', grid=True, xticks=year_ticks)

        axes[0].set_title('{}\'s Salary Change VS Inflation (% Change in CPI)'.format(self.specific_search))

        plt.show()

root = Tk()
GUI = Application(root)
root.mainloop()
