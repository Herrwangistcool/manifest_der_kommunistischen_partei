# Manifest_der_Kommunistischen_Partei

#### 介绍
共产党宣言
用xlwings统计共产党宣言中各个术语出现的位置。

#### excel表格内容
1. raw工作表，存放共产党宣言的原文与对应的翻译，A列为翻译，B列为原文。
2. terminology工作表，存放需要统计的术语。
3. res工作表，存放统计结果。

#### 脚本结构

1.  函数定义，定义了一个用于去除术语中出现的中文的函数——remove_non_german_words。
2.  处理术语表，术语表中的术语含有中文，标点符号，不利于后面的查找与统计，故需要删除每个单元格中的无关字符。
3.  创建键值对，将宣言的原文与翻译一一对应，做成键值对。将键值对结果存储到translation_dict，以备后续使用。
4.  遍历术语表，判断该术语是否在translation_dict的键中，如果是则将该术语以及该键值对存入res工作表。
5.  保存并关闭表格


