import xlwings as xw
import re

wb = xw.Book('Manifest_zh-CN_de-DE.xlsx')
sheet_names = [sheet.name for sheet in wb.sheets]

def remove_non_german_words(input_string):
    # 使用正则表达式匹配德语单词
    german_words = re.findall(r'\b[a-zA-ZäöÖüß]+[a-zA-ZäöÖüß]*\b', input_string)
    
    # 将匹配到的德语单词连接成新的字符串
    result = ' '.join(german_words)
    
    return result

# 选择名为terminology的工作表
sheet = wb.sheets['terminology']
# 术语列表
terminology = []
# 遍历terminology工作表第一列中的每一行
for row in sheet.range('A1:A78'):
    # 获取单元格原始内容
    raw_value = row.value
    # 去除非德语词汇
    cell_value = remove_non_german_words(raw_value)
    terminology.append(cell_value)
print(terminology)

# 选择raw工作表
raw = wb.sheets['raw']
# 创建一个字典来存储键值对
translation_dict = {}
# 获取raw工作表的行数
rows = raw.range('A1').current_region.rows.count
for i in range(1,rows+1):
    raw_sentence = raw.range('B' + str(i)).value
    translation_sentense = raw.range('A' + str(i)).value
    translation_dict[raw_sentence] = translation_sentense

# 检查是否存在名为res的工作表
if 'res' in sheet_names:
    # 如果存在，则打开res工作表
    res_sheet = wb.sheets['res']
else:
    # 如果不存在，则新建res工作表并打开
    res_sheet = wb.sheets.add('res')

# res工作表的行计数
res_count = 1
# 遍历字典中的每个键值对
for item in terminology:
    for key, value in translation_dict.items():
        if item in key:
            print(res_count)
            res_sheet.range('A' + str(res_count)).value = item
            res_sheet.range('B' + str(res_count)).value = key
            res_sheet.range('C' + str(res_count)).value = value
            res_count += 1

# 保存修改后的Excel文件
wb.save('Manifest_zh-CN_de-DE.xlsx')
wb.close()