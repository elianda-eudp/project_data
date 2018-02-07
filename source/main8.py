# -*- coding: utf-8 -*-
"""
Created on Fri Jan 12 09:55:52 2018

@author: user
"""

from docx import Document
import json
import table
def fun(document):
    paras = document.paragraphs
    lcount = []  # 存储所有段落的对象 
  
    for i in range(len(paras)):
        if paras[i].style.name == 'Heading 2':
            lcount.append(i)     
    # 提取每一段的数据
    for i in range(len(lcount)):
        if i+1 >= len(lcount):
            break
        ax = lcount[i+7]
        bx = lcount[i+8]
#       print(ax,bx)
        Head4_list = []
        Head3_list = []
        normal_list = []
        # 取出当前一级标题需要的n个段落
        for j in range(int(ax),int(bx),1):
#            print(paras[j].text)
            # 获取并打印二级标题
            if paras[j].style.name == 'Heading 2':
                Head2 = paras[j].text
                print(Head2)
            # 获取并打印三级题
            if paras[j].style.name == 'Heading 3':
                Head3 =  paras[j].text
                print(Head3)
                Head3_list.append(Head3)               
              # 获取并打印四级题
            elif paras[j].style.name == 'Heading 4':
                Head4 =  paras[j].text
                print(Head4)
                Head4_list.append(Head4)
            # 获取并打印正文
            elif paras[j].style.name == 'Normal':
                normal = paras[j].text                
                print(normal)
                new_normal = normal.replace('\u3000', ' ')
                normal_list.append(new_normal)                                     
        break

    data = {Head2:{Head3_list[0]:{Head4_list[0]:normal_list[0],
                                       Head4_list[1]:normal_list[1]+normal_list[2]+normal_list[3],
                                       Head4_list[2]:normal_list[4]+normal_list[5],
                                       Head4_list[3]:table.table(63,document),
                                       Head4_list[4]:table.table(64,document)}       
            }}
    print(data)
    with open('ta_cust_acct_info.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[1]:{Head4_list[5]:normal_list[6],
                                       Head4_list[6]:normal_list[7]+normal_list[8]+normal_list[9],
                                       Head4_list[7]:normal_list[10]+normal_list[11],
                                       Head4_list[8]:'',
                                       Head4_list[9]:table.table(65,document)}
            }}
    print(data)  
    with open('ta_cust_asset_acct.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)        
    data = {Head2:{Head3_list[2]:{Head4_list[10]:normal_list[12],
                                       Head4_list[11]:normal_list[13]+normal_list[14]+normal_list[15],
                                       Head4_list[12]:normal_list[16]+normal_list[17],
                                       Head4_list[13]:'',
                                       Head4_list[14]:table.table(66,document)}       
            }}
    print(data)  
    with open('ta_prod_asset_list.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[3]:{Head4_list[15]:normal_list[18],
                                       Head4_list[16]:normal_list[19]+normal_list[20]+normal_list[21],
                                       Head4_list[17]:normal_list[22]+normal_list[23],
                                       Head4_list[18]:'',
                                       Head4_list[19]:table.table(67,document)}       
            }}
    print(data)  
    with open('ta_retailer_principal_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[4]:{Head4_list[20]:normal_list[24],
                                       Head4_list[21]:normal_list[25]+normal_list[26]+normal_list[27],
                                       Head4_list[22]:normal_list[28]+normal_list[29],
                                       Head4_list[23]:'',
                                       Head4_list[24]:table.table(68,document)}       
            }}
    print(data)  
    with open('ta_retailer_acc_principal_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[5]:{Head4_list[25]:normal_list[30],
                                       Head4_list[26]:normal_list[31]+normal_list[32]+normal_list[33],
                                       Head4_list[27]:normal_list[34]+normal_list[35],
                                       Head4_list[28]:'',
                                       Head4_list[29]:table.table(69,document)}       
            }}
    print(data)  
    with open('ta_huge_redeem_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[6]:{Head4_list[30]:normal_list[36],
                                       Head4_list[31]:normal_list[37]+normal_list[38]+normal_list[39],
                                       Head4_list[32]:normal_list[40]+normal_list[41],
                                       Head4_list[33]:'',
                                       Head4_list[34]:table.table(70,document)}       
            }}
    print(data)  
    with open('ta_principal_trade_list_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[7]:{Head4_list[35]:normal_list[42],
                                       Head4_list[36]:normal_list[43]+normal_list[44]+normal_list[45],
                                       Head4_list[37]:normal_list[46]+normal_list[47],
                                       Head4_list[38]:'',
                                       Head4_list[39]:table.table(71,document)}       
            }}
    print(data)  
    with open('ta_acct_trade_list_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[8]:{Head4_list[40]:normal_list[48],
                                       Head4_list[41]:normal_list[49]+normal_list[50]+normal_list[51],
                                       Head4_list[42]:normal_list[52]+normal_list[53],
                                       Head4_list[43]:'',
                                       Head4_list[44]:table.table(72,document)}       
            }}
    print(data)  
    with open('ta_special_trade_list_note.json','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)
    data = {Head2:{Head3_list[9]:{Head4_list[45]:normal_list[54],
                                       Head4_list[46]:normal_list[55]+normal_list[56]+normal_list[57],
                                       Head4_list[47]:normal_list[58]+normal_list[59],
                                       Head4_list[48]:'',
                                       Head4_list[49]:table.table(73,document)}       
            }}
    print(data)  
    with open('gf_prod_redeem_list_note','w+') as f:
        str_data = json.dumps(data)
        f.write(str_data)

   
   
fun(Document('理财监管新规数据库优化11.9.docx')) 
    
    
    









