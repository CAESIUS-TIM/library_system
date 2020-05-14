import pandas as pd
import numpy as np
import re
import sys

# -- init --
# - fileInput -
# method 1:
src = r'asset\博库3.4-4.2.xls'
dst = r'sample1.xlsx'  # Styler.to_excel
# Styler只能导出excel且必须是xlsx
df = pd.read_excel(src)
path_PUBLISHER = r'asset\拟删除的出版社.xlsx'
# method 2:
...
# - const -
# [0]
NOT_NULL = ('ISBN', '书名', '著者', '出版社', '定价', '页数', '尺寸', '分类', '出版时间', '语种',)
NULL = ('主题', '副题名', '卷册', '分卷名', '丛书', '版本',
        '附件', '附注', '内容提要', '教材', '读者群', '装帧',)
INF = 1 << 32
RANGE_FILTER = {                            # 保留
    '尺寸': (17, 31),                       # 2
    '页数': (50, INF),                      # 9
}
STR_FILTER = {                              # 去除: 0, 保留: 1
    '版本': (0, ('影印版', '影印本',)),      # 3.1
    '语种': (1, ('chi',)),  # 3.2
    '读者群': (0, ('幼儿', '中小学', '中学', '小学', '高中', '儿童', '少儿', '少年', '高职', '高专',)),  # 5
}
RANGE_HIGHLIGHTER = {
    '定价': (200, INF),                     # 8
    '页数': (3, INF),                       # 10
}
STR_HIGHLIGHTER = {
    '书名': ('报告', '皮书', '年鉴', '研究', '分析',),  # 7
    '分类': lambda x: x[0] == 'I',          # 13
    '卷册': ('辑',),                        # 14
    '装帧': ('线装', '袋装', '函装',),       # 15
}
# [1] 现行
MIN_VOL = 3                     # 10 ...
MIN_SIZE, MAX_SIZE = 17, 31     # 2 # not null
LANG_KEYWORDS = ['chi']         # 3.2 # not null
MIN_PAGE = 50                   # 9 # not null
READER_KEYWORDS = ['幼儿', '中小学', '中学', '小学', '高中',
                   '儿童', '少儿', '少年', '高职', '高专', ]  # 5 #null
PUBLISHER = pd.read_excel(path_PUBLISHER)                   # 4 # not null
PUBLISHER = PUBLISHER.dropna()
PUBLISHER = PUBLISHER.values  # array
BINDING_KEYWORDS = ['线装', '袋装', '函装']
TITLE_KEYWORDS = ['报告', '皮书', '年鉴', '研究', '分析']     # not null


def print_color(*args, color=31, sep=' ', end='\n',):
    # print 开头部分：\033[显示方式;前景色;背景色m + 结尾部分：\033[0m
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

# -- filter --

# function_dict = {}
# def filter():
#     ...
# def highlighter():
#     ...


def size_filter(item):  # 2.去除小于17cm或大于31cm开本的图书。
    """
    a
    a×b
    """
    # not null
    pat1 = r'(\d+)(?=cm|$)'
    pat2 = r'(\d+)×(\d+)(?=cm|$)'
    try:
        matched = re.compile(pat1).match(item)
        if matched != None:
            if MIN_SIZE <= int(matched.group(1)) <= MAX_SIZE:
                return True
        else:
            matched = re.compile(pat2).match(item)
            if MIN_SIZE <= int(matched.group(1)) <= MAX_SIZE and\
                    MIN_SIZE <= int(matched.group(2)) <= MAX_SIZE:
                return True
        return False
    except:
        print_color('Error[size_filter]:', item)
        return False


def lang_filter(item):  # 3.1. 去除影印版，影印本，非中文图书。
    # not null
    if item in LANG_KEYWORDS:
        return True
    else:
        return False


def publisher_filter(item):  # 4. 去除一些小的出版社出版的图书。（详见excel表格）
    # not null
    if item in PUBLISHER:
        return True
    return False


def reader_filter(item):  # 5.去除读者对象含有幼儿，中小学，甚至高职高专的。
    # null
    if pd.isna(item):
        return True
    for i in READER_KEYWORDS:
        if i in item:
            return False
    else:
        return True


def page_filter(item):  # 9.去除小于50页的图书。
    """
    a页
    a-b页
    XXX,XXX页 (计数法)
    (逗号的全半角)
    3册(302;400;308页) r'(\d+)册\((?:(\d+);)*(\d+)页\)'
    10册
    """
    # not null
    # [\u4e00-\u9fa5]汉字，防止页数中有奇怪的注释
    item.replace(',', '，')
    pat1 = r'(\d+)(?=[\u4e00-\u9fa5]|$)'
    pat2 = r'(\d+)-(\d+)(?=[\u4e00-\u9fa5]|$)'
    pat3 = r'\d{1,3}(,\d{3})*(?=[\u4e00-\u9fa5]|$)'
    pat4 = r'(\d+)册'
    try:
        matched = re.compile(pat1).match(item)
        if matched:
            page = int(matched.group(1))
        else:
            matched = re.compile(pat2).match(item)
            if matched:
                page = int(matched.group(2))-int(matched.group(1))
            else:
                matched = re.compile(pat3).match(item)
                if matched:
                    return True
                matched = re.compile(pat4).match(item)
                if matched:
                    return True

        if MIN_PAGE <= page:
            return True
        else:
            return False
    except:
        print_color('[page_filter]error:', item)
        return False

# -- highlighter --


def title_highlighter(item):  # 7. 高亮显示书名中...
    # not null
    for keyword in TITLE_KEYWORDS:
        if keyword in item:
            return True
    return False


def price_highlighter(item):  # 8. 高亮显示价格大于200的图书的整行信息。
    # not null
    if 200 <= item:
        return True
    return False


def volnum_highlighter(item):  # 10. 高亮显示“页数”为3册及以上的图书。
    """
    3册(302;400;308页) r'(\d+)册\((?:(\d+);)*(\d+)页\)'
    10册
    """
    # not null
    pat3 = r'(\d+)册'
    matched = re.compile(pat3).match(item)
    if matched:
        if MIN_VOL <= int(matched.group(1)):
            return True
    return False


def classno_highlighter(item):  # 13. 高亮显示“分类”中头字母为“I”的图书的整行信息。
    # not null
    if item[0] == 'I':
        return True
    return False


def isvol_highlighter(item):  # 14. 高亮显示“卷册”I列中有“辑”字出现的图书的整行信息。
    # null
    if pd.isna(item):
        return False
    if '辑' in item:
        return True
    return False


def binding_highlighter(item):  # 15. 高亮显示“装帧”U列中...
    # null
    if pd.isna(item):
        return False
    for keyword in BINDING_KEYWORDS:
        if keyword in item:
            return True
    return False


def highlight(x, fun):
    return ['background-color: yellow' if fun(item) else '' for item in x]


# -- main --
# - NaN -
df = df.dropna(how='any', subset=NOT_NULL)

# 12.按分类号排序，A在最前面。
df.sort_values(by=['分类', 'ISBN'], axis=0, ascending=[True, True])  # 12
# - filter -
df = df.loc[df['尺寸'].apply(size_filter)]
df = df.loc[df['语种'].apply(lang_filter)]
df = df.loc[df['读者群'].apply(reader_filter)]
df = df.loc[df['页数'].apply(page_filter)]
df = df.loc[df['出版社'].apply(publisher_filter)]
print(df.columns)

# -- excel样式 --
# method*3: 1.Styler 2.ExcelWriter 3.import StyleFrame
# DataFrame.style：便于数据处理
st = df.style
st = st.apply(highlight, subset=['书名'], args=(title_highlighter,))
st = st.apply(highlight, subset=['定价'], args=(price_highlighter,))
st = st.apply(highlight, subset=['页数'], args=(volnum_highlighter,))
st = st.apply(highlight, subset=['分类'], args=(classno_highlighter,))
st = st.apply(highlight, subset=['卷册'], args=(isvol_highlighter,))
st = st.apply(highlight, subset=['装帧'], args=(binding_highlighter,))

# -- write --
# 11.excel进去，出来还是个excel.
# Styler.to_excel
# Styler只能导出excel且必须是xlsx
print(st)
print(type(st))
st.to_excel(dst, index=False, encoding='gbk')
