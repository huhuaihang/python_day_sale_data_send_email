# -*- coding: utf-8 -*-

import xlsxwriter
import pymysql
import time,datetime
import sys
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email import encoders

reload(sys)
sys.setdefaultencoding('utf-8')

data_file_name = time.strftime("%Y-%m-%d",time.localtime(time.time()))

wb = xlsxwriter.Workbook(data_file_name + '.xlsx')

sheet = wb.add_worksheet('newsheet')
merge_format = wb.add_format({
    'bold':     True,
    'border':   6,
    'align':    'center',#水平居中
    'valign':   'vcenter',#垂直居中
    'fg_color': '#D7E4BC',#颜色填充
})
sheet.merge_range(0,0,0,6, '周销售数据',merge_format)

conn = pymysql.connect(
        host = '127.0.0.1',
        port = 3306,
        user = 'root',
        password = 'ytb@123',
        db = 'ytb',
        charset = 'utf8'
)

cursor = conn.cursor()
old_time = 1556640000
new_time = time.time()
sql = 'select * from day_data where day_time > '+ str(new_time - 7 * 86400)

row = cursor.execute(sql)
result = cursor.fetchall()
cursor.close()

conn.close()

sheet.write_row(1,0,['日期','注册数','激活数','销量','销售额','非激活销售量','非激活销售额'])

i = 2
for item in result:
        sheet.write_row(i,0,[time.strftime("%Y-%m-%d",time.localtime(item[7])),item[1],item[2],item[3],item[5],item[4],item[6]])
        i+= 1

sheet.write_row(i,0,["7日合计：","=sum(B3:B8)","=sum(C3:C8)","=sum(D3:D8)","=sum(E3:E8)","=sum(F3:F8)","=sum(G3:G8)"])
chart = wb.add_chart({"type":"line"})
chart.set_title({'name':'周统计'})
chart.set_x_axis({'name':'日期'})
chart.set_y_axis({'name':'数量'})
chart.add_series({"name":"注册量","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,1,7,1],'data_labels':{'value':True}})
chart.add_series({"name":"激活量","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,2,7,2],'data_labels':{'value':True}})
chart.add_series({"name":"销量","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,3,7,3],'data_labels':{'value':True}})
chart.add_series({"name":"普通商品销量","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,5,7,5],'data_labels':{'value':True}})

sheet.insert_chart('A11',chart)

chart2 = wb.add_chart({"type":"line"})
chart2.set_title({"name":"销售量周统计折线图"})
chart2.set_x_axis({"name":"日期"})
chart2.set_y_axis({"name":"销售额"})
chart2.add_series({"name":"销售额","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,4,7,4],'data_labels':{'value':True}})
chart2.add_series({"name":"普通商品销售额","categories":'=newsheet!$A$3:$A$8',"values":['newsheet',2,6,7,6],'data_labels':{'value':True}})

sheet.insert_chart('A25',chart2)
wb.close()


from_addr = '550134146@qq.com'
password = '****'

#to_addr = 'huhuaihangf@163.com'
to_addr = '550134146@qq.com'
smtp_server = 'smtp.qq.com'

content = '周运营报告请查看附件！'
#msg = MIMEText(content, 'plain', 'utf-8')
msg = MIMEMultipart()

msg['From'] = from_addr
msg['To'] = to_addr
msg['Subject'] = unicode('云淘帮上周运营数据')

part = MIMEApplication(open('/python/' + data_file_name +'.xlsx', 'rb').read())
part.add_header('Content-Disposition','attachment', filename= data_file_name +".xlsx")
msg.attach(part)

server = smtplib.SMTP_SSL(smtp_server)
server.connect(smtp_server, 465)
server.login(from_addr, password)

server.sendmail(from_addr, to_addr.split(','), msg.as_string())
server.quit()

                        
