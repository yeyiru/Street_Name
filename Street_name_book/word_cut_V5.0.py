#coding: UTF-8
#開始運行前，請先全選doc，按ctrl + shift + f9移除所有超鏈接並保存
#如果有超鏈接，轉換會出錯
#編譯辭書格式：
#第一章  XXX
#第二章  XXX
#第三章  XX鎮、鄉、區
#   第一節 鎮名緣起
#   第二節 自然環境
#   第三節 區域特色
#   第四節 各里地名釋義
#       第一項  XX里（村）    <-------------------------------抓取起始位置
#           里（村）名由來
#             里（村）的description，若干行
#           地名釋義 
#           （一）具體地名1
#              具體地名1的description，若干行
#           （二）具體地名2
#              具體地名2的description，若干行
#            ......
#           其他
#           （一）具體地名1
#              具體地名1的description，若干行
#           （二）具體地名2
#              具體地名2的description，若干行
#            ......
#       第二項  XX里（村）
#            ......
#       第三項  XX里（村）
#            ......
#第四章  XX鎮、鄉、區
#基本算法架構：step1：根據“第X項”找到里（村）作為開始，以“章”結束一個鎮、鄉、區，據此找到每個里（村）相關內容的上下邊界
#             step2：在每個里（村）中根據“里（村）名由來”找到每個里（村）description的上下邊界
#             step3：在每個里（村）中根據“地名釋義”和“（）”找到每個里（村）的具體地名及其上下邊界                     
#             step4：返回正文，從根據上下邊界抓取具體內文，輸出到csv的固定位置
#***************************************************************************************************************
import numpy as np
import pandas as pd
from docx import Document
import win32com.client as wc
import os
from tqdm import tqdm
#數據處理，doc轉docx轉txt存入本地
class data_processor():
    def __init__(self):
        self.dir_name = 'F:/street_name/book/'

    def doc2docx(self):
        #doc转docx
        word = wc.Dispatch("Word.Application")
        for i in os.listdir(self.dir_name):
            if i.endswith('.doc') and not i .startswith('~$'):
                
                doc_path = os.path.join(self.dir_name, i)
                doc = word.Documents.Open(doc_path)

                rename = os.path.splitext(i)

                save_path = os.path.join(self.dir_name, rename[0] + '.docx')
                doc.SaveAs(save_path, 12)
                doc.Close()
                print(i)
        word.Quit()

    def docx2txt(self):
        #docx转txt，去除所有不必要的格式
        for i in os.listdir(self.dir_name):
            if i.endswith('.docx') and not i.startswith('~$'):
                docx_path = os.path.join(self.dir_name, i)
                document = Document(docx_path)
                txt_path = os.path.join(self.dir_name, str(i).replace('.docx', '.txt'))
                txt_file = open(txt_path, 'w', encoding = 'utf-8')
                for paragraph in tqdm(document.paragraphs):
                    new_paragraph = paragraph.text.strip('/r')
                    new_paragraph = new_paragraph.strip()
                    if new_paragraph != '':
                        txt_file.write(new_paragraph + '\n')
                txt_file.close()
                #删除使用过的docx
                os.remove(docx_path)

#分行
class word_cut():
    def __init__(self):
        #初始化全局變量
        #工作目錄
        self.dir_name = 'F:/street_name/book/'
        #中文字符常量
        self.chinese = ['一', '二', '三', '四', '五', '六', '七', '八', '九', '十']

    def run(self):
        for i in os.listdir(self.dir_name):
            if i.endswith('.txt') and not i .startswith('~$'):
                self.save_name = i.replace('txt', 'csv')
                print('Begin read ' + str(i))
                self.get_txt(i)
                self.get_vil_index_up_down()
                self.cut_vli_by_index()
                self.re_index()  #若註解此行則以村里為單位標記No
                self.save_csv()
                
    def get_txt(self, file_name):
        txt_path = os.path.join(self.dir_name, str(file_name))
        with open(txt_path, 'r', encoding = 'utf-8') as txt_file:
            #获得txt文本存入list
            self.document_list = txt_file.readlines()
        txt_file.close()
        #定義一個df存放村裡對應的行號上下界
        self.vil_df_index = pd.DataFrame(columns = ['No', 'dist_name', 'vil_name', 'vil_index_down', 'vil_index_up'], dtype = int)
        #定義一個df供保存需要的數據
        self.df_save = pd.DataFrame(columns = ['No', 'name_dist', 'name_li', 'name', 'name_eng', 'location', 'description'])

    #獲取各個村里內容的上下限
    def get_vil_index_up_down(self):
        for line_index in tqdm(range(len(self.document_list))):
            #記錄每個里的行號作為index存入df_index
            line = self.document_list[line_index]
            vil_name = self.get_vil_name(line)
            
            if vil_name is not None:
                if 'end' not in vil_name :
                    cache_index = pd.DataFrame({'No': vil_pointer + 1,
                                                'dist_name': dist_name,
                                                'vil_name': vil_name,
                                                'vil_index_down': line_index,
                                                'vil_index_up': 0}, index = [0])
                    vil_df_cache_index = vil_df_cache_index.append(cache_index, ignore_index = True)
                    try:
                        vil_df_cache_index.iloc[vil_pointer - 1, 4] = line_index - 1
                    except:
                        pass
                    vil_pointer += 1
                else:
                    if vil_name[1] != '結論':
                        dist_name = vil_name[1]
                    try:
                        vil_df_cache_index.iloc[vil_pointer - 1, 4] = line_index - 1
                        self.vil_df_index = self.vil_df_index.append(vil_df_cache_index, ignore_index = True)
                    except:
                        pass
                    #重置指針和暫存df
                    vil_pointer = 0
                    vil_df_cache_index = pd.DataFrame(columns = ['No', 'dist_name', 'vil_name', 'vil_index_down', 'vil_index_up'], dtype = int)

    def get_vil_name(self, line):
        #根據內容和長度查找 “第X項 XX村（里）”格式段落
        if '第' in line and '項' in line and len(line) <= 24:
            tmp = line.split('項')
            return tmp[1].strip()
        elif '第' in line and '章' in line and any(s in line for s in self.chinese) and len(line) <= 12:
            tmp = line.split('章')
            return 'end', tmp[1].strip()
        else:
            return None

    def cut_vli_by_index(self):
        #初始化一個df供存放村里名、地名和description的index
        useful_index = pd.DataFrame(columns = ['No', 'dist_name', 'vil_name', 'useful_name', 'useful_index_down', 'useful_index_up'])
        #遍歷vil_df_index，搜索每個村裡名下面的地名和對應的description的行號存入useful_index，
        for i in tqdm(range(len(self.vil_df_index))):
            no = self.vil_df_index.iloc[i, 0]
            dist_name = self.vil_df_index.iloc[i, 1]

            vil_name = self.vil_df_index.iloc[i, 2]
            line_index_down = self.vil_df_index.iloc[i, 3]
            line_index_up = self.vil_df_index.iloc[i, 4]

            cache_index = self.get_name_and_description_index(no, dist_name, vil_name, line_index_down, line_index_up)
            useful_index = useful_index.append(cache_index, ignore_index = True)
        self.get_description_main(useful_index)
    
    #根據上下界，在村里內做進一步細分
    def get_name_and_description_index(self, no, dist_name, vil_name, line_index_down, line_index_up):
        useful_index = pd.DataFrame(columns = ['No', 'dist_name', 'vil_name', 'useful_name', 'useful_index_down', 'useful_index_up'])
        line_pointer = 0
        for i in range(line_index_down, line_index_up + 1):
            line = self.document_list[i].strip()
            line = line.replace('\r', '')
            line = line.replace('\n', '')
            #依次切斷
            #里（村）名由来下面是里（村）的description
            if '名由來' in line and len(line) <= 10:
                try:
                    line = line.split('、')[1]
                except:
                    pass
                cache_index = pd.DataFrame({'No': line_pointer + 1,
                                            'dist_name': dist_name,
                                            'vil_name': vil_name,
                                            'useful_name': line,
                                            'useful_index_down': i,
                                            'useful_index_up': 0}, index = [0])
                useful_index = useful_index.append(cache_index, ignore_index = True)
                line_pointer += 1
            #地名釋義終結了村、里的description，且下面是具體地名
            elif ('地名釋義' in line or '二、其他' in line) and len(line) <=10:
                useful_index.iloc[line_pointer - 1, 5] = i - 1
            #具體地名以括號+中文數字做開頭
            #具體地名緊接著就是description
            elif '（'  in line and '）' in line and any(s in line for s in self.chinese) and (len(line) <= 30 or '【' in line):
                try:
                    #以）分割取出具體地名
                    line = line.split('）', 1)[1]
                except:
                    pass
                cache_index = pd.DataFrame({'No': line_pointer + 1,
                                            'dist_name': dist_name,
                                            'vil_name': vil_name,
                                            'useful_name': line,
                                            'useful_index_down': i,
                                            'useful_index_up': 0}, index = [0])
                useful_index = useful_index.append(cache_index, ignore_index = True)
                try:
                    useful_index.iloc[line_pointer - 1, 5] = i - 1
                except:
                    pass
                line_pointer += 1

        useful_index.iloc[line_pointer - 1, 5] = line_index_up
        return useful_index

    #獲取description
    def get_description_main(self, useful_index):
        for i in tqdm(range(len(useful_index))):
            #初始化save_df內的各項元素
            no = useful_index.iloc[i, 0]
            dist_name = useful_index.iloc[i, 1]
            vil_name = useful_index.iloc[i, 2]
            name = ''
            name_eng = ''
            location = ''
            description = ''
            #獲取內容
            #里（村）名由來下面就是村里的description
            if '名由來' in useful_index.iloc[i, 3] and len(useful_index.iloc[i, 3]) == 4:
                description = self.get_description(useful_index.iloc[i, 4], useful_index.iloc[i, 5])
            #否則對應地名及其下的description
            else:
                name = useful_index.iloc[i, 3]
                description = self.get_description(useful_index.iloc[i, 4], useful_index.iloc[i, 5])
            cache_description = pd.DataFrame({'No': no,
                                                'name_dist': dist_name,
                                                'name_li': vil_name,
                                                'name': name,
                                                'name_eng': name_eng,
                                                'location': location,
                                                'description': description}, index = [0])
            #寫入df
            self.df_save = self.df_save.append(cache_description, ignore_index = True)
    
    def get_description(self, index_down, index_up):
        #根據index上下限獲取description
        description = ''
        #如果只有一行description則直接寫入
        if index_down == index_up - 1:
            description = self.clear_description(self.document_list[index_up])
        else:
            for i in range(index_down + 1, index_up + 1):
                description = description + self.document_list[i]
                description = self.clear_description(description)
        return description
    
    def clear_description(self, description):
        #清理正文內容，去除換行、空格、局末的其他（小地名存在兩個部分，以其他分隔，因只有一行，將其歸上處理）等不必要的字符
        description = description.strip()
        description = description.replace('\r', '')
        description = description.replace('\n', '')
        description = description.strip('三、其他')
        description = description.strip('二、地名釋義')
        return description
    
    def re_index(self):
        #重新編號，以鄉鎮區為編號單位
        count = 0
        for i in range(len(self.df_save)):
            count += 1
            self.df_save.iloc[i, 0] = count
            try:
                if self.df_save.iloc[i, 1] != self.df_save.iloc[i + 1, 1]:
                    count = 0
            except:
                pass

    def save_csv(self):
        self.df_save.to_csv(self.dir_name + self.save_name, header=1, index = False, encoding='utf-8-sig')

if __name__ == '__main__':

    dp = data_processor()
    #doc转txt
    try:
        dp.doc2docx()
        dp.docx2txt()
    except:
        pass
    
    w_c = word_cut()
    w_c.run()