import re
from copy import copy
#将列的大小转化为字母坐标,如28表示AB
def num_to_scope(c:int)->str:
    ret = ''
    while c>0:
        mod=(c-1)%26
        ret=chr(65+mod)+ret
        c=(c-1)//26
    return ret
#将列的大小转化为数字,如AB表示28
def scope_to_num(col:str)->int:
    num=0
    for c in col:
        num=num*26+(ord(c)-65)+1
    return num
def loader_to_pos(loader:str)->list(list()):
    front, back = loader.split(':')
    index = 0
    for i, v in enumerate(front):
        if v.isdigit():
            index = i
            break
    startC, startR = front[:index], front[index:]
    for i, v in enumerate(back):
        if v.isdigit():
            index = i
            break
    endC,endR=back[:index],back[index:]
    ec=scope_to_num(endC)+1
    sc=scope_to_num(startC)
    sr=int(startR)
    er=int(endR)+1
    return [[sr,er],[sc,ec]]
#复制cell的字体、大小等样式
def copy_sheet_attributes(source_sheet, target_sheet):
    target_sheet.sheet_format = copy(source_sheet.sheet_format)
    target_sheet.sheet_properties = copy(source_sheet.sheet_properties)
    target_sheet.merged_cells = copy(source_sheet.merged_cells)
    target_sheet.page_margins = copy(source_sheet.page_margins)
    target_sheet.freeze_panes = copy(source_sheet.freeze_panes)
    # set row dimensions
    # So you cannot copy the row_dimensions attribute. Does not work (because of meta data in the attribute I think). So we copy every row's row_dimensions. That seems to work.
    for rn in range(len(source_sheet.row_dimensions)):
        target_sheet.row_dimensions[rn] = copy(source_sheet.row_dimensions[rn])

    if source_sheet.sheet_format.defaultColWidth is None:
        print('Unable to copy default column wide')
    else:
        target_sheet.sheet_format.defaultColWidth = copy(source_sheet.sheet_format.defaultColWidth)

    # set specific column width and hidden property
    # we cannot copy the entire column_dimensions attribute so we copy selected attributes
    for key, value in source_sheet.column_dimensions.items():
        target_sheet.column_dimensions[key].min = copy(source_sheet.column_dimensions[key].min)   # Excel actually groups multiple columns under 1 key. Use the min max attribute to also group the columns in the targetSheet
        target_sheet.column_dimensions[key].max = copy(source_sheet.column_dimensions[key].max)  # /sf/ask/2549209491/ discussed the issue. Note that this is also the case for the width, not onl;y the hidden property
        target_sheet.column_dimensions[key].width = copy(source_sheet.column_dimensions[key].width) # set width for every column
        target_sheet.column_dimensions[key].hidden = copy(source_sheet.column_dimensions[key].hidden)


def smaller(tmp,value):
    if value<tmp:
        return True
    return False

def smallerE(tmp,value):
    if value<=tmp:
        return True
    return False

def bigger(tmp,value):
    if value>tmp:
        return True
    return False

def biggerE(tmp,value):
    if value>=tmp:
        return True
    return False

def equal(tmp,value):
    if tmp==value:
        return True
    return False

def unequal(tmp,value):
    if tmp!=value:
        return True
    return False

def contain(tmp,value):
    if tmp in value:
        return True
    return False

def uncontain(tmp,value):
    if tmp not in value:
        return True
    return False