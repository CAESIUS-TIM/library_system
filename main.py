import pandas as pd
import numpy as np
import re
import sys

# -- init --
# - filePath -
# method 1:
src = r'asset\人天1.9-3.23.xls'
dst = r'sample1.xlsx'  # Styler.to_excel
# Styler只能导出excel且必须是xlsx
path_PUBLISHER = r'asset\拟删除的出版社.xlsx'
# method 2:
...
# - const -
# [0]
# -- isNull --
NOT_NULL = ('ISBN', '书名', '著者', '出版社', '定价', '页数', '尺寸', '分类', '出版时间', '语种',)
NULL = ('主题', '副题名', '卷册', '分卷名', '丛书', '版本',
        '附件', '附注', '内容提要', '教材', '读者群', '装帧',)
# -- type --
OBJECT = ('ISBN', '书名', '著者', '出版社', '页数', '尺寸', '分类', '语种', '主题', '副题名',
          '卷册', '分卷名', '丛书', '版本', '附件', '附注', '内容提要', '教材', '读者群', '装帧',)  # -> str
FLOAT = ('定价',)  # float, int, 单纯数字的字符串 object -> float
DATETIME = ('出版时间',)  # 暂时没有必要
# --- 关于函数合并 ---
# -- float --
# 9 float:纯数字范围
# 2,9,10 string:文本数字(格式,单位)范围
# -- string --
# 3.1,3.2,5,7,14,15 集合中某一字符串在目标字符串中存在
# 13 特别地,在上述基础上出现在特点位置，且为单一字符
# 4 目标字符串在集合中存在
# --- 结论： 差别太大,难以合并
INF = 1 << 32
RANGE_FILTER = {                            # 保留: 闭区间
    '尺寸': (17, 31),                       # 2: 格式
    '页数': (50, INF),                      # 9: 格式
}
STR_FILTER = {                              # 去除: 0, 保留: 1
    '版本': (0, ('影印版', '影印本',)),      # 3.1: 仅可能是影印或非影印
    '语种': (1, ('chi',)),                  # 3.2: 支持混合语言
    '读者群': (0, ('幼儿', '中小学', '中学', '小学', '高中', '儿童', '少儿', '少年', '高职', '高专',)),  # 5
}
RANGE_HIGHLIGHTER = {                       # 闭区间
    '定价': (200, INF),                     # 8: 无格式
    '页数': (3, INF),                       # 10: 格式
}
STR_HIGHLIGHTER = {
    '书名': ('报告', '皮书', '年鉴', '研究', '分析',),  # 7
    '分类': lambda x: x[0] == 'I',          # 13
    '卷册': ('辑',),                        # 14
    '装帧': ('线装', '袋装', '函装',),       # 15
}
# [1] 现行
MIN_VOL = 3                                                 # 10 ...
MIN_SIZE, MAX_SIZE = 17, 31                                 # 2 # not null
EDITION_KEYWORDS = ['影印版', '影印本']                      # 3.1 # null
LANG_KEYWORDS = ['chi']                                     # 3.2 # not null
PUBLISHER = pd.read_excel(path_PUBLISHER)                   # 4 # not null
PUBLISHER = PUBLISHER.dropna()
PUBLISHER = PUBLISHER.values  # array
READER_KEYWORDS = ['幼儿', '中小学', '中学', '小学', '高中',
                   '儿童', '少儿', '少年', '高职', '高专', ]  # 5 #null
TITLE_KEYWORDS = ['报告', '皮书', '年鉴', '研究', '分析']     # 7 #not null
MIN_PAGE = 50                                               # 9 # not null
BINDING_KEYWORDS = ['线装', '袋装', '函装']                  # 15 # null

size_error = 0
page_error = 0


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
    print('\033[%dm' % color+sep.join(map(str, args)) +
          '\033[0m', sep=sep, end=end)

# -- filter --

# function_dict = {}
# def filter():
#     ...
# def highlighter():
#     ...


def size_filter(item):  # 2.去除小于17cm或大于31cm开本的图书。
    """
    a
    a×b(乘 数学符号: ×,英文字母: x,X)
    # 大16开 210×85mm
    # 小16开 185×260mm
    # 20开 206mm×181mm
    """
    global size_error
    # not null
    pat1 = r'(\d+)(?=cm|$)'
    pat2 = r'(\d+)[×xX](\d+)(?=cm|$)'
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
        """
        大16开 210×85mm
        小16开 185×260mm
        20开 206mm×181mm
        """
        unpat1 = r'[大小]?\d+开'
        if re.compile(unpat1).match(item):
            size_error += 1
            return False
        print_color('Error[size_filter]:', item)
        return False


def edition_filter(item):  # 3.1. 去除影印版，影印本。
    # null
    if pd.isna(item):
        return True
    for keyword in EDITION_KEYWORDS:
        if keyword in item:
            return True
    return False


def lang_filter(item):  # 3.2. 去除非中文图书。
    # not null
    for keyword in LANG_KEYWORDS:
        if keyword in item:
            return True
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
    ...全彩 (注释)
    # [a×b]
    # a,b,c,...,n页
    # a页,b页,c页,...,n页
    """
    global page_error
    # not null
    # [\u4e00-\u9fa5]中文字符，防止页数中有奇怪的注释
    # [\uFF00-\uFFFF]全角字符
    item = item.replace('，', ',')  # replace(self, old, new, count=-1, /)
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
        """
        [a×b]
        a,b,c,...,n页
        a页,b页,c页,...,n页
        """
        unpat1 = r'\[\d+×\d+\]'
        unpat2 = r'(\d+,)*(\d+)(?=[\u4e00-\u9fa5]|$)'
        unpat3 = r'(\d+页,)*(\d+页)(?=[\u4e00-\u9fa5]|$)'
        if re.compile(unpat1).match(item) or\
                re.compile(unpat2).match(item) or\
                re.compile(unpat3).match(item):
            page_error += 1
            return False
        print_color('Error[page_filter]:', item)
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
if __name__ == "__main__":
    # -- fileInput --
    print('- read_excel')
    df = pd.read_excel(src)
    # print(df.columns) # 列名检验

    # -- dropNa --
    print('- dropna')
    df = df.dropna(how='any', subset=NOT_NULL)
    print_color('non-null:',len(df),color='blue')
    # -- type --
    for colnum in OBJECT:
        # df[colnum] = df[colnum].astype('str')
        df[colnum] = df[colnum].apply(lambda x: x if pd.isna(x) or
                                      isinstance(x, str) else str(x))

    # -- filter --
    print('- filter')
    df = df.loc[df['尺寸'].apply(size_filter)]          # 2
    df = df.loc[df['版本'].apply(edition_filter)]       # 3.1
    df = df.loc[df['语种'].apply(lang_filter)]          # 3.2
    df = df.loc[df['出版社'].apply(publisher_filter)]   # 4
    df = df.loc[df['读者群'].apply(reader_filter)]      # 5
    df = df.loc[df['页数'].apply(page_filter)]          # 9
    print_color('rest:',len(df),color='blue')
    # -- sort --
    # 12.按分类号排序，A在最前面。
    print('- sort')
    df.sort_values(by=['分类', 'ISBN'], axis=0, ascending=[True, True])  # 12

    # -- excel样式 --
    # method*3: 1.Styler 2.ExcelWriter 3.import StyleFrame
    # DataFrame.style：便于数据处理
    # -- highlighter --
    print('- highlighter')
    st = df.style
    st = st.apply(highlight, subset=['书名'], args=(title_highlighter,))    # 7
    st = st.apply(highlight, subset=['定价'], args=(price_highlighter,))    # 8
    st = st.apply(highlight, subset=['页数'], args=(volnum_highlighter,))   # 10
    st = st.apply(highlight, subset=['分类'], args=(classno_highlighter,))  # 13
    st = st.apply(highlight, subset=['卷册'], args=(isvol_highlighter,))    # 14
    st = st.apply(highlight, subset=['装帧'], args=(binding_highlighter,))  # 15

    # -- write --
    # 11.excel进去，出来还是个excel.
    # Styler.to_excel
    # Styler只能导出excel且必须是xlsx(xls不行)
    print('- to_excel')
    st.to_excel(dst, index=False, encoding='gbk', float_format='%.2f')   # 11
    print_color('- done', color='green')
    print_color('size_error =', size_error,
                ', page_error =', page_error, color='yellow')
