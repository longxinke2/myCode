import requests
from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives import hashes
from IPython.display import clear_output
from IPython.display import HTML
import re
import pyperclip
import pyautogui
import keyboard
from bs4 import BeautifulSoup
import chardet
from PIL import ImageGrab
from PIL import Image
from IPython.display import clear_output
import win32com.client
import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import time
import datetime
from email.utils import formataddr
from email.header import Header
import xlwings as xw
import pyperclip
import keyboard
import pyautogui

def f2_shuru(i):
    keyboard.wait('f2')
    pyperclip.copy(i)
    pyautogui.hotkey('ctrl','v')
    
def open_zb(path='lib/研究院工作周报-龙辛柯.xlsx'):
    # 主要用于查看改动的效果
    # 创建Word应用程序对象
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = True
    path=os.path.abspath(path)# 不知道为什么当前路径不行，要完整路径
    workbook = excel.Workbooks.Open(path)
    workbook.Activate() # 置于所有应用之前

def click_image(image_path, confidence=0.95, lag_x=0,lag_y=0,wait=0.2, double=False):
    # 获取屏幕分辨率
    screen_width, screen_height = pyautogui.size()

    # 全屏截图
    screen_image = ImageGrab.grab()

    # 在屏幕截图中查找匹配
    image=Image.open(image_path)
    location = pyautogui.locate(image, screen_image, confidence=confidence)
    if location:
        time.sleep(wait)
        location = pyautogui.locate(image, screen_image, confidence=confidence)
        # 匹配成功，返回位置信息
        x, y, a, b = location
        if double==False:
            pyautogui.click(x+a//2+lag_x, y+b//2+lag_y)
        else:
            pyautogui.doubleClick(x+a//2+lag_x, y+b//2+lag_y)

def get_soup(url):
    headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
    response = requests.get(url, headers=headers)
    encoding = chardet.detect(response.content)['encoding']
    response.encoding = encoding
    return BeautifulSoup(response.text,'html.parser')

def ctrl_a():
    pyautogui.hotkey('ctrl','a')
    pyautogui.hotkey('ctrl','c')
    return pyperclip.paste()

def ctrl_v(x):
    pyperclip.copy(x)
    pyautogui.hotkey('ctrl','v')
        
def translate():
    s = input(":")
    dat = {
        "kw":s
    }
    resp = requests.post("https://fanyi.baidu.com/sug", data=dat)
    print(resp.json()['data'][0]['v']) #将返回的内容直接返回为json

def is_English():
    string = input(":")
    for ch in string:
        if ord(ch) >= 128:
            return ch
    return True

def progress_bar(now,end,text='',scare=2):
    percent = now*100//end
    line = '-'*(percent//scare)+' '*((100-percent)//scare)
    return f'{text}{line}{percent}%'#print的时候end='\r'可以把光标挪到最前面
    
def encode(clear=False):
    public_key_bytes = b'-----BEGIN PUBLIC KEY-----\nMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAzVsQAFaz1kTZXPY3Og9x\nlsspWAKXJBqv81Q4DKeL8YP1OOpihA9OYHO2um2GlVVGhlgRxbkIAEWjzelJJb8+\nSvCUm2CDwFuSYNfx7k/i5UeEK1Y7gbCjBalN2PzBuTqcvAeSfexX85Lm+lW3iQw0\nbv1k4ifKabrooq2ILJTNH3NnGjcERO8QiFa6jY5/H3eWo3POknahAX26rhpYMl1X\na/r0cSbI/c4DqzMpd6sGFQkL2DHsVCF2saY+HMPMqQHm+oT903GVAIpw0x0u6p6o\nAf1QIy8uGvZH0ee4YReLJbc80JvyHysZst3lHQu3E/0UV7nxQOPfumNfphFcwE5S\nFwIDAQAB\n-----END PUBLIC KEY-----\n'
    if clear:
        with open("lib/encode.txt", "wb") as file:
            file.write(b'')
        return
    # 读取公钥
    public_key = serialization.load_pem_public_key(
        public_key_bytes
    )
    temp=''
    while True:
        with open("lib/encode.txt", "ab") as file:
            x=input()
            if x=='show':
                print(temp)
                continue
            temp=x
            plaintext =x.encode()
            # 使用公钥加密
            ciphertext = public_key.encrypt(
                plaintext,
                padding.OAEP(
                    mgf=padding.MGF1(algorithm=hashes.SHA256()),
                    algorithm=hashes.SHA256(),
                    label=None
                )
            )
            file.write(ciphertext)
            clear_output()

def decode():
    # 读取私钥
    with open("C:/Users/龙辛柯/Desktop/code/private_key_bytes.txt", "rb") as file:
        private_key_bytes = file.read()  # 读取二进制字符串
    private_key = serialization.load_pem_private_key(
        private_key_bytes,
        password=None
    )
    with open("myCode/lib/encode.txt", "rb") as file:
        ciphertext = file.read()  # 读取二进制字符串
        text=''
        for i in range(int(len(ciphertext)/256)):
            decrypted_plaintext = private_key.decrypt(
                ciphertext[i*256:(i+1)*256],
                padding.OAEP(
                    mgf=padding.MGF1(algorithm=hashes.SHA256()),
                    algorithm=hashes.SHA256(),
                    label=None
                )
            )
            text1 = decrypted_plaintext.decode()
            if text1!='end':
                text +=text1
            else:
                text +='\n\t'
        print(text)
        
def get_week_of_month(date):
    first_day = date.replace(day=1)
    dom = date.day
    adjusted_dom = dom + first_day.weekday()
    if adjusted_dom <= 7:
        return 1
    elif adjusted_dom <= 14:
        return 2
    elif adjusted_dom <= 21:
        return 3
    elif adjusted_dom <= 28:
        return 4
    else:
        return 5

# date: "2022-08-09"
def get_current_week():
    monday, sunday = datetime.date.today(), datetime.date.today()
    month = datetime.date.today().month
    one_day = datetime.timedelta(days=1)
    while monday.weekday() != 0:
        monday -= one_day
    while sunday.weekday() != 6:
        sunday += one_day
    today = datetime.date.today()
    week_num= get_week_of_month(today)
    
    # return monday, sunday
    # 返回时间字符串
    return '龙辛柯-周报('+ datetime.datetime.strftime(monday, "%Y年%m月%d日") + '-' + datetime.datetime.strftime(sunday, "%Y年%m月%d日")\
    +f'){month}月第{week_num}周'

def send_myemail(timet=0):
    # 配置邮箱及密码
    from_mail_name = formataddr((Header('龙辛柯','utf-8').encode(), 'longxk@bibibi.net'))
    to_mail_name = '陈静 <chenj@bibibi.net>; 高玉 <gaoyu@bibibi.net>; 史册 <shice@bibibi.net>; 夏超群 <xiacq@bibibi.net>; 曾诚睿 <zengcr@bibibi.net>; 朱黎 <zhul@bibibi.net>; 龙辛柯 <longxk@bibibi.net>'
    from_mail = 'longxk@bibibi.net'
    from_mail_password = '740926Lxk'
#     to_mail = ['longxk@bibibi.net','2245247439@qq.com']
    to_mail =['sunb@bibibi.net','hr@bibibi.net','chenj@bibibi.net','longxk@bibibi.net','zengcr@bibibi.net','gaoyu@bibibi.net','shice@bibibi.net','xiacq@bibibi.net','zhul@bibibi.net']
    # 截图路径
    pic_file = r'lib/周报截图.png'
    
    app = xw.App(visible=True, add_book=False)
    wb = app.books.open('lib/研究院工作周报-龙辛柯.xlsx')
    sheet = wb.sheets[0]
    sheet['A1'].value = get_current_week()

    # 写完截图
    all = sheet.used_range    # 获取有内容的range
    all.api.CopyPicture()    # 复制图片区域
    sheet.api.Paste()    # 粘贴
    pic = sheet.pictures[0]    # 当前图片
    pic.api.Copy()    # 复制图片
    time.sleep(3)
    img = ImageGrab.grabclipboard()    # 获取剪贴板的图片数据
    # 有一个缓存的时间
    time.sleep(5)
    img.save(pic_file)    # 保存图片
    pic.delete()    # 删除sheet上的图片

    wb.save() # 保存文件
    wb.close() # 关闭文件
    app.quit() # 关闭程序
    
    time.sleep(timet)
    
    curr_time = datetime.datetime.now()
    time_str = curr_time.strftime('%m%d')
    
    
    # 设置总的邮件体对象，对象类型为mixed
    msg = MIMEMultipart('mixed')

    # 邮件的发件人及收件人信息
    msg['From'] = from_mail_name
    msg['To'] = to_mail_name
    msg['Cc'] = '孙柏 <sunb@bibibi.net>; hr <hr@bibibi.net>'
    # 邮件的主题
    msg['Subject'] = f'研究院工作周报-龙辛柯-{time_str}'

    # 构造文本内容
#     text_info = 'hello world'
#     text_sub = MIMEText(text_info, 'plain', 'utf-8')
#     msg.attach(text_sub)

    # 构造超文本
    html_info = """
    <div>各位同事：</div><div>&nbsp; &nbsp; &nbsp; 大家好！这是我本周的工作总结，请查收！</div>
    <div>
    <sign signid="1">
    <div>
    <font>祝：身体健康！<br></font>
    <font><br>
    <img src="cid:image1"><br>
    </font>
    <div style="color:#909090;font-family:Arial Narrow;font-size:12px"></div>
    </div>
    <div style="font-size:14px;font-family:Verdana;color:#000;">
    <div>
    <img src="https://exmail.qq.com/cgi-bin/viewfile?type=signature&amp;picid=ZX0228-TpCGSnFBZkAbzAtofQjgIcm&amp;uin=3341171678" onerror=""><font size="4">长沙市云研网络科技有限公司</font></div><div><table id="QMEditorArea" class="bd" cellpadding="0" cellspacing="0" style="border: 1px solid rgb(220, 224, 225); color: rgb(0, 0, 0); font-family: &quot;Helvetica Neue&quot;, Arial, &quot;PingFang SC&quot;, &quot;Hiragino Sans GB&quot;, STHeiti, &quot;Microsoft YaHei&quot;, sans-serif; font-size: 12px; background-color: rgb(245, 245, 245); min-width: 450px;"><tbody><tr><td class="bizCardContainer_td" style="-webkit-font-smoothing: antialiased; background: rgb(255, 255, 255); padding: 5px; line-height: 1.5;"><div id="signPreview" style="height: auto; min-height: 80px; padding: 5px;"><div><b><font size="4">龙辛柯</font><font size="2"> 研究院</font></b></div><div>地址：湖南省长沙市天心区芙蓉南路二段390号湖南长沙人力资源产业园B座6层</div><div>电话：17620085592</div><div>E-mail：longxk@bibibi.net</div></div></td></tr></tbody></table></div>
    </div></sign></div><div>&nbsp;</div><div><tincludetail></tincludetail></div><!--<![endif]-->
    """
    html_sub = MIMEText(html_info, 'html', 'utf-8')
    # 如果不加下边这行代码的话，上边的文本是不会正常显示的，会把超文本的内容当做文本显示——法男女搭配，加了就作为附近上传了，别加
   #html_sub["Content-Disposition"] = 'attachment; filename="csdn.html"'
    # 把构造的内容写到邮件体中
    msg.attach(html_sub)
    
    #添加图片到正文
    with open(pic_file,'rb') as image:
        image_info = MIMEImage(image.read())
        image_info.add_header('Content-Id','<image1>')
        msg.attach(image_info)

    # 构造附件
    txt_file = open(r'lib/研究院工作周报-龙辛柯.xlsx', 'rb').read()
    txt = MIMEText(txt_file, 'base64', 'utf-8')
    txt["Content-Type"] = 'application/octet-stream'
    txt.add_header('Content-Disposition', 'attachment', filename=f'研究院工作周报-龙辛柯-{time_str}.xlsx')
    msg.attach(txt)
    server = smtplib.SMTP('smtp.exmail.qq.com')
    server.login(from_mail,from_mail_password)
    server.sendmail(from_mail,to_mail,msg.as_string())
    server.quit()
    return