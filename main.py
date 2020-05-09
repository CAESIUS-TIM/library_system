import pandas as pd
import numpy as np
import re
import sys

# -- init --
# - file -
# method 1:
src = 'asset/博库3.4-4.2.xls'
dst = 'sample1.xlsx'
df = pd.read_excel(src)
# method 2:
...
# - const -
MIN_SIZE, MAX_SIZE = 17, 31  # not null
LANG_KEYWORDS = ['chi']  # not null
MIN_PAGE = 50  # not null
READER_KEYWORDS = ['幼儿', '中小学', '高职', '高专']  # null
PRESS = ...  # not null
TITLE_KEYWORDS = ['报告', '皮书', '年鉴', '研究', '分析']  # not null
TITLES = ['书名', 'chinesetitle']
columns = list(df)
for TITLE in TITLES:
    if TITLE in columns:
        break


# -- filter_function --
# print 开头部分：\033[显示方式;前景色;背景色m + 结尾部分：\033[0m
def print_color(*args, color=31, sep=' ', end='\n',):
    colors = {
        'black': 30,
        'red': 31,
        'green': 32,
        'yellow': 33,
        'blue': 34,
        'magenta': 35,
        'cyan': 36,
        'white': 37,
    }
    if isinstance(color, str):
        color = colors.get(color.lower(), 30)
    print('\033[%dm' % color+sep.join(map(str, args))+'\033[0m')


def size_filter(item):  # 2.去除小于17cm或大于31cm开本的图书。
    """
    12
    12×12
    12*12
    """
    try:
        matched = re.match(r'(\d+)(?:.+?(\d+))?', item)
        if matched == None:
            return False
        if MIN_SIZE <= int(matched.group(1)) <= MAX_SIZE and\
                (matched.group(2) == None or MIN_SIZE <= int(matched.group(2)) <= MAX_SIZE):
            return True
        else:
            return False
    except:
        print_color('[size_filter]error:', item)
        return False


def lang_filter(item):  # 3.去除影印版，影印本，非中文图书。
    try:
        if item in LANG_KEYWORDS:
            return True
        else:
            return False
    except:
        print_color('[lang_filter]error:', item)
        return False


def reader_filter(item):  # 5.去除读者对象含有幼儿，中小学，甚至高职高专的。
    try:
        for i in READER_KEYWORDS:
            if i in item:
                return False
        else:
            return True
    except:
        if pd.isna(item):  # pd.isna > np.isnan
            return True
        print_color('[reader_filter]error:', item)
        return False


def page_filter(item):  # 9.去除小于50页的图书。

    try:
        matched = re.match(r'(\d+)', item)
        if MIN_PAGE <= int(matched.group(1)):
            return True
        else:
            return False
    except:
        print_color('[page_filter]error:', item)
        return False


def title_filter(item):
    for keyword in TITLE_KEYWORDS:
        if keyword in item:
            return True
    return False

def highlight_max(x):
    return ['background-color: yellow' if title_filter(item) else '' for item in x]
# -- main --


# 12.按分类号排序，A在最前面。
df.sort_values(by=['分类', 'ISBN'], axis=0, ascending=[True, True])  # 12
df = df.loc[df['尺寸'].apply(size_filter)]
df = df.loc[df['语种'].apply(lang_filter)]
df = df.loc[df['读者群'].apply(reader_filter)]
df = df.loc[df['页数'].apply(page_filter)]
print(df.columns)

# -- excel样式 --
# method 2:
# DataFrame.style：便于数据处理
df = df.style.apply(highlight_max,subset=[TITLE])

# -- write --
# 11.excel进去，出来还是个excel.
print(df)
print(type(df))
df.to_excel(dst, index=False, encoding='gbk')

"""
# method 1:
# excel样式：艺术性高，数据处理性低
# Warning: 样式必须在最后做！----------------------------
l_end = len(df.index) + 2  # 表格的行数,便于下面设置格式
yellow = workbook.add_format({'fg_color': '#FFEE99'})

# 7.高亮显示书名中含有“报告”“皮书”“年鉴”“研究”“分析”的图书信息。
TITLE_KEYWORDS = ['报告', '皮书', '年鉴', '研究', '分析']
TITLES = ['书名', 'chinesetitle']
columns = list(df)
col_num = columns.find(title)
for index, value in df[title].items():
    if(title_filter(value)):
"""
