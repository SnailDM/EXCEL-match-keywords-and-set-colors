#用关键词将句子分割开，然后加上颜色，加上关键词，再拼接后面的句子
#created by xuqingyao
import pandas as pd
import re
from xlsxwriter.workbook import Workbook

def model_set_colors(data, data_result, cate_data):
    workbook = data_result
    #关键词为英文字符时可能会有大小写不一，需要统一转换为小写或大写
    keywords = cate_data['关键词'].str.lower()
    keyword_num = keywords.shape[0]
    #为后面捆绑作准备
    worksheet = workbook.add_worksheet()
    red = workbook.add_format({'color':'red'})
    #读取行数，后面对每行数据循环写入匹配到关键词的富字符串
    nums_data = data.shape[0]
    #设定worksheet的初始写入行列
    work_row = 0
    work_col = 0

    for num in range(nums_data):
        txt = data.loc[num, '文本描述']
        txt = str(txt).lower()
        match_str = ''  #保存匹配到的关键词
        for i in range(keyword_num):
            keyword = keywords[i]
            keyword = str(keyword)

            if txt.find(keyword) >= 0:
                match_str = match_str+keyword+'|'
        match_str = match_str.strip('|')    #虽然关键词已保存，但并未按其在文本中出现的位置顺序
        #通过re.finditer按文本中出现的顺序匹配，所以进行二次匹配
        re_match = re.finditer(match_str,txt)
        re_str = ''
        for m in re_match:
            re_str = re_str + str(m.group()) + '|'
        re_str = re_str.strip('|')
        keyword_split = re_str.split('|')   #这样就按文本中出现关键词的顺序列出了

        keyword_split_num = len(keyword_split)
        keyword_match = keyword_split
        match_words = []
        #keyword_match_num = len(keyword_match)
        #找出能模糊匹配到的字符长度短的关键词
        for j in range(keyword_split_num):
            for k in range(len(keyword_match)):
                if (keyword_split[j].find(keyword_match[k])>= 0) and (len(keyword_split[j])>len(keyword_match[k])):
                    match_words.append(keyword_match[k])

        #原关键词列表删除字符长度短的关键词
        num_match_words = len(match_words)
        for i in range(num_match_words):
            keyword_split.remove(match_words[i])
        keyword_set= keyword_split
        keyword_sep = ''
        for each in keyword_set:
            keyword_sep = keyword_sep + each + '|'
        keyword_sep = keyword_sep.strip('|')
        #print('keyword_sep:', keyword_sep)

        #用所有关键词将整段话分割，再插入富字符串，然后捆绑颜色、关键词和后面的文本，需注意一一对应
        temp_list = re.split(keyword_sep, txt)
        params = []
        temp_list_num = len(temp_list)
        for i in range(temp_list_num):
            if i != 0:
               params.extend((red,keyword_set[i-1],temp_list[i]))
            else:
                params.append(temp_list[i])
        worksheet.write_rich_string(work_row, work_col, *params)
        work_row = work_row+1
    workbook.close()

if __name__ == '__main__':
    data = pd.read_excel(r'E:\DLML\DataHandle\data\data.xlsx', index=None)  #原文本文件
    data_result = Workbook(r'E:\DLML\DataHandle\data\data-result.xlsx') #标注结果
    cate_data = pd.read_excel(r'E:\DLML\DataHandle\data\cate_data.xlsx', index=None) #匹配的文本文件
    model_set_colors(data, data_result, cate_data)