import re
import wx
import sqlalchemy
import math
import functools
import numpy
from IPython.display import clear_output

def L(string):
    try:
        return index_dict[string]
    except:
        return index_dict.get(string) or 'sb'

def get_group_order_table(data,string,other_groupby):
    # 获取分组数据
    data1 = groupby_data(data,string,other_groupby)
    print(data1.to_markdown())
    # 排序
    group = {'default':'默认排序','xy':f'按照给定的{L(string)}顺序排序','num_down':'按照人数降序','num_up':'按照人数升序'}
    result = get_order_way('排序规则：',group)
    data1 = order_table(data1,result,string,group)
    if not data1:
        return 
    clear_output()  # 清除输出
    print(data1.to_markdown())
    return data1

def groupby_data(data,string,other_groupby):
    data1 = data.groupby(string)['xxdm'].count().reset_index(name='num')  # 计算组别数量
    sum_counts = data1['num'].sum()  # 计算所有组别数量的总和
    data1['proportion'] = round(data1['num']*100 / sum_counts,2).astype(str)+'%'   # 计算占比并添加到新列中
    if len(other_groupby)!=0:
        data2 = data.groupby([string,other_groupby])['xxdm'].count().reset_index(name='num')  # 计算组别数量
        sum_counts = data2['num'].sum()  # 计算所有组别数量的总和
        data2['proportion'] = round(data2['num']*100 / sum_counts,2).astype(str)+'%'   # 计算占比并添加到新列中
        for i in data2[other_groupby].unique():
            data3 = data2[data2[other_groupby]==i]
            data3.columns = [string,i,f'{i}_num',f'{i}_proportion']
            data1 = data1.merge(data3,how='left',on=string)
        for i in data1.columns:
            if 'num' in i:
                data1[i] = data1[i].apply(lambda x: not math.isnan(x) and int(x) or 0)
        data1.drop(data2[other_groupby].unique(),axis=1,inplace=True)
#         data1.columns=[string]+[j+str(i) for i in range(len(data2[other_groupby].unique())+1) for j in ['num', 'proportion']]
    data1.replace(numpy.nan,'-',inplace=True)#nan值替换成'-'
    data1.replace(0,'-',inplace=True)
    return data1.reindex(columns=data1.columns.tolist()[:1]+data1.columns.tolist()[3:]+data1.columns.tolist()[1:3])

def get_yjy_db(db):
    return sqlalchemy.create_engine(f"mysql+pymysql://yjy_user:Yjy123456@am-wz9el267w54i2r7ip131930o.ads.aliyuncs.com:3306/{db}")

def order_table(table,result,string,group={}):
    a = list(filter(lambda x:group[x]==result,group.keys()))
    if len(a)==0:
        return table
    if a[0]=='xy' :
        xy_order_str = input('请输入排序规则：')
        list1 = []
        try:
            for i in list(table[string]):
                list1.append(re.search(i, xy_order_str).span(0)[0])
        except:
            print(f'未在表格中找到{L(string)}字段')
            return
        table['order'] = list1
        table.sort_values(by='order',inplace=True)
        table.drop('order',axis=1,inplace=True)
    if a[0]=='num_down':
        table.sort_values(by='num',inplace=True,ascending=False)
    if a[0]=='num_up':
        table.sort_values(by='num',inplace=True)
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

# 字典
index_dict = {
            #基础指标
            'myd':'满意','ppd':'匹配','xgd':'相关','tjd':'愿意','gzd':'关注','xb':'性别','yx':'院系','zy':'专业','sysf':'生源省份',
            'mz':'民族','jyq':'乐观','zqj':'乐观','zsp':'满意','jjx':'满意','lzt':'乐观','sjy':'满意','yjy':'满意','ljy':'满意',
            #派遣数据
            'sxgxlx':'升学院校类型','sxgxbq':'升学院校层次',
            #分数指标
            '5':'非常高','4':'高','3':'比较高','2':'不太高','1':'不高',
            #教育教学调研
            'dy_cyzyxgd':'创业专业相关度','dy_cyyy':'创业原因','dy_cylx':'创业类型','dy_cyfs':'创业方式',
              'dy_cykn':'创业困难','dy_cyfwjy':'创业服务建议','dy_mxmyd':'母校满意度','dy_mxtjd':'母校推荐度','dy_jyjx':'教育教学',
              'dy_mxtjd':'母校推荐度','dy_jyjx_tsjy':'通识教育满意度','dy_tsjyjy':'通识教育建议','dy_jyjx_zyjy':'专业教育满意度','dy_zykcjy':'专业教育建议',
             'dy_jyjx_sjjx':'实践教学满意度','dy_sjjxjy':'实践教学建议','dy_zynljy':'职业能力教育建议','dy_zzdgjjy':'师资水平建议',
             'dy_jyjx_zynljy':'职业教育满意度','dy_jyjx_szsp':'师资水平','dy_jyjx_xsjz':'学术讲座','dy_jyjx_jxss':'教学设施',
            'dy_jyjx_khfs':'考核方式','dy_sjjxmyd':'实践教学满意度','dy_jyjx_zhmyd':'综合满意度',
            'dy_sjjx_kcsy':'课程实验','dy_sjjx_kcsyx':'课程实用性','dy_sjjx_sjkbz':'实践课比重','dy_sjjx_bylw':'毕业论文',
             'dy_sjjx_sys':'实验室的使用与管理','dy_kcszjy':'课程设置建议','dy_szsp':'师资水平满意度','dy_szsp_jxtd':'教学态度满意度',
             'dy_szsp_jxnr':'教学内容满意度','dy_szsp_jxff':'教学方法、方式满意度','dy_qzwt':'求职问题','dy_zxjl':'在校经历',
             'dy_szsp_sshd':'师生互动满意度','dy_szsp_zynl':'专业知识能力满意度',
             'dy_jycyfwmyd':'就业创业服务满意度','dy_jycyfwjy':'就业创业服务建议','dy_jycy':'就业创业',
             'dy_jycy_zcxc':'政策宣传与讲解','dy_jycy_xxtg':'信息提供与发布','dy_jycy_zyzx':'职业选择咨询/辅导','dy_jycy_jnpx':'就业/创业技能培训','dy_jycy_xlts':'求职心理调适','dy_jycy_mszd':'面试指导与训练',
             'dy_jycy_sxzd':'升学指导','dy_jycy_lxzd':'留学指导','dy_jycy_kc':'就业/创业课程','dy_jycy_xssxh':'线上双选会','dy_jycy_xyzp':'校园招聘活动','dy_jycy_jysx':'就业手续办理','dy_jycy_knbf':'就业困难群体帮扶',
              #就业调研
              '自主创业':'zzcy','dwhy':'单位行业','zzcy':'自主创业','sfmc':'所在省份','dy_lzcs':'离职次数','dy_jyqjyq':'专业前景预期',
              'dy_jymyd':'就业满意度','dy_jyzyxgd':'就业专业相关度','dy_jyjxmyd':'教育教学满意度','dy_jyqs':'就业歧视',
              'dy_jyzybxgyy':'就业专业不相关原因','dy_zwppd':'职业期待匹配度','dy_xzmyd':'薪资满意度','dy_jyyx':'月薪',
              'dy_jypjxz':'平均月薪','dy_shbz':'社会保障','dy_gwfzqj':'发展前景预期','dy_qwyz':'期望月薪',
              'dy_shbzmyd':'社会保障满意度','dy_dwpj':'岗位评价','dy_dwzmd':'单位知名度','dy_pxjh':'培训机会','dy_jskj':'晋升空间',
              'dy_gzhj':'工作环境','dy_gzzzx':'工作自主性','dy_gzyl':'工作压力','dy_gzwdx':'工作稳定性','dy_gwfzqj':'岗位发展前景',
              'dy_gzdd':'工作地点','dy_qywh':'企业文化','dy_gzfw':'工作氛围','dy_dwzhmyd':'单位综合满意度',
              'dy_lzyy':'离职原因','dy_gzys':'求职关注因素','dy_gzys_xcsp':'薪酬水平','dy_gzys_shbz':'社会保障',
              'dy_gzys_gzwdx':'工作稳定度','dy_gzys_gzcs':'工作城市','dy_gzys_dwsw':'单位社会声望','dy_gzys_gzhj':'工作环境',
              'dy_gzys_fzkj':'发展空间','dy_gzys_jrqw':'父母家人期望','dy_gzys_rjgx':'人际关系','dy_gzys_zydk':'专业对口',
              'dy_gzys_grxq':'个人兴趣','dy_gzys_rczc':'人才政策','dy_qztj':'求职途径','dy_offer':'offer数','dy_qzsc':'求职时长',
              'dy_wjyyy':'未就业原因','dy_qwdwlx':'期望单位类型','dy_qwjyss':'期望就业省市','dy_cgcjyy':'出国出境原因',
              'dy_qwbf':'期望帮扶','dy_nltsxq':'期望求职帮助','dy_jxszzyxgd':'继续深造与专业相关度','dy_gnsxyy':'国内升学原因',
              'dy_yqyx_zygh':'职业规划影响','dy_yqyx_xlzt':'就业乐观度',
              #单位调研
              'cdy_dwgm':'单位规模','cdy_dwxz':'单位性质','cdy_dwhy':'单位行业','cdy_dwszd':'单位所在地','cdy_zpqd':'招聘渠道',
              'cdy_kzys':'关注因素','cdy_zpxz':'招聘薪酬范围','cdy_zygzd':'专业关注度','cdy_rcpymyd':'人才培养认同度',
              'cdy_bysmyd':'毕业生满意度','cdy_byspj_sxpz':'思想品质','cdy_rcpyjy':'人才培养建议','cdy_lxzpyy':'来校招聘原因',
              'cdy_byspj_gztd':'工作态度','cdy_byspj_zysp':'专业水平','cdy_byspj_zynl':'职业能力','cdy_byspj_fzql':'发展潜力',
              'cdy_byspj':'毕业生评价','cdy_gwsysj':'岗位适应时间','cdy_lzqk':'离职情况','cdy_lzyy':'离职原因','cdy_zpkn':'招聘困难',
              'cdy_jyfwmyd':'就业服务满意度','cdy_zmyjyfw':'就业服务','cdy_jyfwjy':'就业服务建议'}