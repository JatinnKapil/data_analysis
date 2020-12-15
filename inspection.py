import subprocess
import win32api
import win32com.client
from pywinauto.application import Application
import pyautogui
import hashlib
import time
import ctypes
import os
import pandas as pd
import numpy as np
import glob
import win32com.client as win32
from pathlib import Path
from datetime import date

pyautogui.alert(text='OverDue Changes Report',
                title='Robotic Process Automation (RPA)', button='OK')

df1 = pd.read_excel(r"Overdue Inspection_previous.xlsx",
                    sheet_name="overdue inspections")
df2 = pd.read_excel(r"Overdue Inspection_current.xlsx",
                    sheet_name="overdue inspections")

df1.rename(columns={'Unit ID': 'Unit_ID',
                    'Equipment ID': 'Equipment_ID', 'Facility ID': 'Facility_ID', 'Region ID': 'Region_ID', "Event Type": "Event_Type"}, inplace=True)
df2.rename(columns={'Unit ID': 'Unit_ID',
                    'Equipment ID': 'Equipment_ID', 'Facility ID': 'Facility_ID', 'Region ID': 'Region_ID', "Event Type": "Event_Type"}, inplace=True)


df1['equip'] = df1['Equipment_ID'].str.cat(
    df1['Unit_ID'], sep="/").str.cat(df1['Event_Type'], sep="/")

df2['equip'] = df2['Equipment_ID'].str.cat(
    df2['Unit_ID'], sep="/").str.cat(df2['Event_Type'], sep="/")

# df2['equip'] = df2['Equipment_ID'].str.cat(df2['Unit_ID'], sep="/")


added = df2[~df2["equip"].isin(df1["equip"])]

removed = df1[~df1["equip"].isin(df2["equip"])]

writer = pd.ExcelWriter(
    'OverdueChanges.xlsx', engine='xlsxwriter')


added.to_excel(
    writer, index=False, sheet_name='Equipment_ID_add-on')
added = added.drop(columns=['equip'])

removed.to_excel(
    writer, index=False, sheet_name='Equipment_ID_removed')
removed = removed.drop(columns=['equip'])

table1 = pd.pivot_table(added, index=['Region_ID', 'Facility_ID', "Event_Type"], values=[
    'Equipment_ID'], aggfunc=len, fill_value=0, margins=True, margins_name='Total')

table1.to_excel(writer, index=True, sheet_name='pivot_data_add-on')

table2 = pd.pivot_table(removed, index=['Region_ID', 'Facility_ID', "Event_Type"], values=[
    'Equipment_ID'], aggfunc=len, fill_value=0, margins=True, margins_name='Total')

table2.to_excel(writer, index=True, sheet_name='pivot_data_removed')


workbook = writer.book
worksheet = writer.sheets['Equipment_ID_add-on']
worksheet1 = writer.sheets['Equipment_ID_removed']
worksheet2 = writer.sheets['pivot_data_add-on']
worksheet3 = writer.sheets['pivot_data_removed']

cell_format = workbook.add_format(
    {'bold': True, 'font_color': '#800000', 'font_size': 12, 'font_script': 'Automated', 'align': 'Justify', 'bg_color': '#f2f2f2'})

# Light red fill with dark red text.
format1 = workbook.add_format({'bg_color':   '#f2f2f2',
                               'font_color': '#9C0006'})

# Green fill with dark green text.
format2 = workbook.add_format({'bg_color':   '#FF6600',
                               'font_color': '#006100'})

worksheet.set_column('A:Z', 27)
worksheet1.set_column('A:Z', 27)
worksheet2.set_column('A:D', 27)
worksheet3.set_column('A:D', 27)
worksheet.set_column('C3:G42', 27, cell_format)
worksheet1.set_column('C3:G42', 27, cell_format)
worksheet2.conditional_format(
    'D2:D11', {'type': 'data_bar', 'bar_axis_color': '#0070C0', 'bar_border_color': '#63C384', 'bar_solid': False})
worksheet2.conditional_format('C2:C11', {
                              'type': 'text', 'criteria': 'containing', 'value': 'External Inspection', 'format': format1})

worksheet2.conditional_format('C2:C11', {
                              'type': 'text', 'criteria': 'containing', 'value': 'Internal Inspection', 'format': format2})

worksheet3.conditional_format(
    'D2:D11', {'type': 'data_bar', 'bar_axis_color': '#0070C0', 'bar_border_color': '#63C384', 'bar_solid': False})

worksheet3.conditional_format('C2:C11', {
                              'type': 'text', 'criteria': 'containing', 'value': 'External Inspection', 'format': format1})

worksheet3.conditional_format('C2:C11', {
                              'type': 'text', 'criteria': 'containing', 'value': 'Internal Inspection', 'format': format2})


chart1 = workbook.add_chart({'type': 'doughnut'})
chart1.add_series({
    'name': 'Overdue Changes data',
    'categories': ['pivot_data_add-on', 1, 1, 10, 1],
    'values': '=pivot_data_add-on!$D$2:$D$11',
    # 'fill':   {'none': True},
    # 'border': {'color': 'black'}
})
chart1.set_style(26)
chart1.set_rotation(36)
chart1.set_hole_size(49)
chart1.set_title({'name': 'Doughnut Chart with user defined colors'})
worksheet2.insert_chart('B21', chart1, {'x_offset': 25, 'y_offset': 10})


chart2 = workbook.add_chart({'type': 'doughnut'})
chart2.add_series({
    'name': 'Overdue Changes data',
    'categories': ['pivot_data_removed', 1, 1, 10, 1],
    'values': '=pivot_data_removed!$D$2:$D$11',
    # 'fill':   {'none': True},
    # 'border': {'color': 'black'}
})
chart2.set_style(26)
chart2.set_rotation(36)
chart2.set_hole_size(49)
chart2.set_title(
    {'name': 'Overdue Data Doughnut Chart with user defined colors'})
worksheet3.insert_chart('B21', chart2, {'x_offset': 25, 'y_offset': 10})


writer.save()
