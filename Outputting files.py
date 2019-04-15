#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Apr 13 16:34:01 2019

@author: pr067127
"""

import pandas as pd
import os

df = pd.read_pickle(os.path.join('Documents','For Python','data_frame.pickle'))

# Smaller object for easier vis
small_df = df.iloc[49980:50019,:].copy()

# Basic excel 
small_df.to_excel("Basic.xlsx")
small_df.to_excel("no_index.xlsx", index=False)
small_df.to_excel("columns.xlsx", columns=["artist","title","year"])

# Multiple worksheets
writer = pd.ExcelWriter('multiple_sheets.xlsx', engine='xlsxwriter')
small_df.to_excel(writer, sheet_name="Preview", index=False)
df.to_excel(writer, sheet_name="Complete", index=False)
writer.save()

# Conditional Formatting
artist_counts = df['artist'].value_counts()
artist_counts.head()
writer = pd.ExcelWriter('colors.xlsx', engine="xlsxwriter")
artist_counts.to_excel(writer, sheet_name="Artist Counts")
sheet = writer.sheets['Artist Counts']
cells_range = 'B2:B{}'.format(len(artist_counts.index))
sheet.conditional_format(cells_range, {'type':'2_color_scale', 'min_value':'10',
                                       'min_type':'percentile', 'max_value':'10',
                                       'max_type':'percentile'})
writer.save()

# SQL
import sqlite3

with sqlite3.connect('my_database.db') as conn:
    small_df.to_sql('Tate', conn)