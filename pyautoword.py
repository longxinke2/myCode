import re
import wx
import sqlalchemy
import math
import functools



def groupby_data(data,string,other_groupby):
    data1 = data.groupby(string)['xxdm'].count().reset_index(name='num')  # 计算组别数量
    sum_counts = data1['num'].sum()  # 计算所有组别数量的总和
    data1['proportion'] = data1['num'] / sum_counts   # 计算占比并添加到新列中
    if len(other_groupby)!=0:
        data2 = data.groupby([string,other_groupby])['xxdm'].count().reset_index(name='num')  # 计算组别数量
        sum_counts = data2['num'].sum()  # 计算所有组别数量的总和
        data2['proportion'] = data2['num'] / sum_counts   # 计算占比并添加到新列中
        for i in data2[other_groupby].unique():
            data3 = data2[data2[other_groupby]==i]
            data1 = data1.merge(data3,how='left',on=string)
        for i in data1.columns:
            if 'num' in i:
                data1[i] = data1[i].apply(lambda x: not math.isnan(x) and int(x) or 0)
        data1.drop(list(filter(lambda x:other_groupby[1] in x,data1.columns)),axis=1,inplace=True)
        data1.columns=[string]+[j+str(i) for i in range(len(data2[other_groupby].unique())+1) for j in ['num', 'proportion']]
    return data1

def get_yjy_db(db):
    return sqlalchemy.create_engine(f"mysql+pymysql://yjy_user:Yjy123456@am-wz9el267w54i2r7ip131930o.ads.aliyuncs.com:3306/{db}")

def order_table(table,result,group={}):
    a = list(filter(lambda x:group[x]==result,group.keys()))
    if len(a)==0:
        return table
    if a[0]=='xy' :
        xy_order_str = input('请输入排序规则：')
        list1 = []
        try:
            for i in list(table['yx']):
                list1.append(re.search(i, xy_order_str).span(0)[0])
        except:
            return '未在表格中找到学院字段'
        table['order'] = list1
        table.sort_values(by='order',inplace=True)
        table.drop('order',axis=1,inplace=True)
    if a[0]=='num_down':
        table.sort_values(by='num0',inplace=True,ascending=False)
    if a[0]=='num_up':
        table.sort_values(by='num0',inplace=True)
    return table

def sum100(arr):
    # 使用哈希表记录每个元素出现的次数
    freq = {}
    for num in arr:
        if num in freq:
            freq[num] += 1
        else:
            freq[num] = 1
    # 创建一个结果列表
    candidates = []
    # 遍历哈希表，将每个元素和它出现的次数添加到结果列表中
    
    for key, value in freq.items():
        candidates.append([key, int(round(key/sum(arr),5)*10000), key/sum(arr) *10000 - int(round(key/sum(arr),5)*10000), value])
    target = 10000
    for i in candidates:
        target-=i[1]*i[3]
#     print(target,candidates)
    if target == 0:
        return candidates
    result = []

    def dfs(combination, current_sum, start):
        # 如果当前组合的数字总和超过了目标值，则停止递归
        if current_sum > target:
            return
        # 如果当前组合的数字总和等于目标值，则将其添加到结果中
        if current_sum == target:
            result.append(combination)
            return
        # 递归搜索剩余的数字
        for i in range(start, len(candidates)):
            dfs(combination + [candidates[i]], current_sum + candidates[i][3], i + 1)

    dfs([], 0, 0)
    temp_g=[]
    for r in result:
        sum_num = 0
        for r1 in r:
            sum_num+=r1[2]*r1[3]
        temp_g.append(sum_num)
    tgroup = []
#     print(result)
    for i in range(0,len(temp_g)):
        if temp_g[i]==max(temp_g):
            tgroup=result[i]
            break
    if tgroup==[]:
        print('该处百分比相加不为1')
    for j in tgroup:
        candidates[candidates.index(j)][1] += 1
    for i in range(len(arr)):
        for j in candidates:
            if arr[i]==j[0]:
                arr[i]=j[1]
    return arr

class btnFrame(wx.Frame):
    def __init__(self, parent,title, btn_group,callback):
        wx.Frame.__init__(self, parent, title='', size=(10*max(len(v) for v in btn_group.values())+80, len(btn_group)*30+100))
        
        # 回调函数
        self.callback=callback
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, f"{title}", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )

        sizer.Add( self.m_staticText1, 0, wx.ALL, 5 )

        for i in btn_group:
            exec(f"self.{i}=wx.Button(self, label='{btn_group[i]}')") # 创建btn
            exec(f"sizer.Add(self.{i}, 0, wx.ALL, 5)") # 添加进
            exec(f"self.{i}.Bind(wx.EVT_BUTTON, lambda event,self=self, i=btn_group[i]: (self.callback(i),self.Destroy()))")

        self.SetSizer(sizer)
        
        # 初始化窗口
        self.Centre()
        self.Show()
        self.SetWindowStyle(wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
        

def get_order_way(title,btn_group):
    try:
        del app
    except:
        pass
    selection_value = ''
    def getvalue(value):
        nonlocal selection_value
        selection_value = value
    app = wx.App()
    frame = btnFrame(None,title,btn_group,getvalue)
    app.MainLoop()
    return selection_value