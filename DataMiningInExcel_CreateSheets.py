import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
from openpyxl import Workbook


result_file="/Users/guqingxian/Desktop/学习/Data Mining in Excel/Results_Exercises/results.xlsx"
book = load_workbook(result_file);

sheet=book.active;
sheet.title='Overall to Total'
book.create_sheet('Total 30 Traces')
book.create_sheet('Total 30 Traces Visualised')
book.create_sheet('Total 30 Traces Sorted')
book.create_sheet('Total HTTP')
book.create_sheet('HTTP Comparision')
book.create_sheet('Total TCP')
book.create_sheet('TCP Comparision')
book.create_sheet('Total UDP')
book.create_sheet('UDP Comparision')
book.create_sheet('Total QUIC')
book.create_sheet('QUIC Comparision')
book.save(result_file)
book.close()