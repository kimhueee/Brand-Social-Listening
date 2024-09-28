import pandas as pd
import vaex as vx
import numpy as np
from datemerge import datemerge
from hyperlink import hyperlink
from to_excel_file import to_excel_file

topic_SP_list = ['THP Trà Xanh Không Độ', 'THP Sữa Đậu Nành Soya', 'THP Nước tăng Lực Number 1', 'THP Trà Thanh Nhiệt Dr. Thanh','THP Trà Sữa Macchiato Không Độ','THP Number 1 Active Chanh Muối', 'THP Number 1 Chanh & Dâu']
channel_list = ['Facebook', 'Fanpage', 'Tiktok', 'News', 'Forum', 'Threads', 'Instagram', 'Youtube', 'E-commerce']
sent_list = ['Positive', 'Neutral', 'Negative']
#labels_list

CORP = pd.read_excel(r"D:\\Data Kompa\\Tân Hiệp Phát\\Monthly\\THPaug.xlsx", sheet_name='CORP')
THP_BLD = pd.read_excel(r"D:\\Data Kompa\\Tân Hiệp Phát\\Monthly\\THPaug.xlsx", sheet_name='THP_BLD')
SP = pd.read_excel(r"D:\\Data Kompa\\Tân Hiệp Phát\\Monthly\\THPaug.xlsx", sheet_name='SP')