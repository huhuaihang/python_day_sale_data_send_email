# -*- codecoding: utf-8 -*-
import pymysql

import time,datetime
import smtplib
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

conn = pymysql.connect(
        host = '127.0.0.1',
        port = 3306,
        user = 'root',
        password = '***',
        db = '***',
        charset = 'utf8'
)

cursor = conn.cursor()
old_time = 1556640000
new_time = time.time()
day_time = time.mktime(datetime.date.today().timetuple())
yesterday_time = int(day_time) - 86400

content = '注册人数: '
cursor = conn.cursor()
sql = 'select * from `day_data` where day_time >= '+ str(yesterday_time) + ' and day_time <' + str(day_time)



row2 = cursor.execute(sql)
result = cursor.fetchall()
cursor.close()
print(result)

conn.close()


for item in result:
	content += str(item[1]) + ' '
        content += ' 激活人数：' + str(item[2]) + ' '
	content += ' 总销量：' + str(item[3]) + ' '
	content += ' 普通商品销量：' + str(item[4]) + ' '
	content += ' 销售额：' + str(item[5]) + ' '
	content += ' 普通商品销售额：' + str(item[6]) + ' '

print(content)
from email.mime.text import MIMEText

from email.header import Header

from_addr = '550134146@qq.com'
password = '****'

#to_addr = 'huhuaihangf@163.com'
to_addr = '550134146@qq.com'
smtp_server = 'smtp.qq.com'

msg = MIMEText(content, 'plain', 'utf-8')

msg['From'] = Header(from_addr)
msg['To'] = Header(to_addr)
msg['Subject'] = Header(unicode('云淘帮昨日运营数据'))

server = smtplib.SMTP_SSL(smtp_server)
server.connect(smtp_server, 465)
server.login(from_addr, password)

server.sendmail(from_addr, to_addr.split(','), msg.as_string())
server.quit()

