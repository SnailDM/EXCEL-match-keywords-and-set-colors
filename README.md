# EXCEL-match-keywords-and-set-colors
python匹配文本关键词并设置颜色

思想：主要通过xlsxwriter的write_rich_string模块来标注颜色。需要先将文本匹配到对应关键词，用关键词将句子分割开，然后加上颜色，捆绑关键词，再捆绑后面的句子。
