"""
要不要NaN
"""
import pandas as pd
import re

src = "asset/博库3.4-4.2.xls"
dst = "sample1.xls"
data = pd.read_excel(src)


def size_filter(item, inf=17, sup=31):  # 2
    try:
        matched = re.match(r"(\d+)(?:.+?(\d+))?", item)
        if matched == None:
            return False
        if inf <= int(matched.group(1)) <= sup and\
                (matched.group(2) == None or inf <= int(matched.group(2)) <= sup):
            return True
        else:
            return False
    except:
        print("[size_filter]error:",item)
        return False


def lang_filter(item, lang=['chi']):  # 3
    # str or list
    try:
        if item in lang:
            return True
        else:
            return False
    except:
        print("[lang_filter]error:",item)
        return False


def reader_filter(item, reader=["幼儿", "中小学", "高职", "高专"]):  # 5
    try:
        for i in reader:
            if i in item:
                return False
        else:
            return True
    except:
        print("[reader_filter]error:",item)
        return False


def page_filter(item, inf=50):  # 9
    try:
        matched = re.match(r"(\d+)(页)?", item)
        if inf <= int(matched.group(1)):
            return True
        else:
            return False
    except:
        print("[page_filter]error:",item)
        return False


data.sort_values(by=["分类", "ISBN"], axis=0, ascending=[True, True])  # 12
data = data.loc[data["尺寸"].apply(size_filter)]
data = data.loc[data["语种"].apply(lang_filter)]
data = data.loc[data["读者群"].apply(reader_filter)]
data = data.loc[data["页数"].apply(page_filter)]

data.to_excel(dst, index=False, encoding="gbk")

# print(data["定价"][3],type(data["定价"][3])) # 69.0 <class 'numpy.float64'>
# print(data["卷册"][3],type(data["卷册"][3])) # 2018 <class 'str'>
