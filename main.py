import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
import csv
my_sender = 'INPUT MAIL ADDRESS HERE'  # 发件人邮箱账号,已脱敏
my_pass = 'INPUT PASSWORD HERE'  # 发件人邮箱密码,已脱敏
# 如果邮箱使用授权码登录，请填写授权码

def getReceiver(x):
    with open('mail.csv',encoding='UTF-8-sig') as file:
        f_csv=csv.DictReader(file)
        for i, rows in enumerate(f_csv):
            if i == x:
                row = rows
                return(row)


def mail(receiverNumber,receiverName,receiverMail,receiverAccount,receiverPassword):


    ret = True
    try:
        msg = MIMEText('你好！这是来自CIC的礼物，一份office 365订阅，附带2TB的OneDrive云存储空间。\n请在office.com登录。\n你的账号是:'+receiverAccount+'\n你的密码是:'+receiverPassword+'\n有问题请发送邮件到cicpublic@outlook.com。', 'plain', 'utf-8')
        msg['From'] = formataddr(['Liz', my_sender])  # 括号里的对应发件人邮箱昵称、发件人邮箱账号
        msg['To'] = formataddr([receiverName, receiverMail])  # 括号里的对应收件人邮箱昵称、收件人邮箱账号
        msg['Subject'] = "来自CIC的一份礼物:office365订阅-"+receiverNumber  # 邮件的主题，也可以说是标题

        server = smtplib.SMTP("smtp-mail.outlook.com", 587)  # 发件人邮箱中的SMTP服务器，端口是25
        ####根据自己的邮件服务修改
        server.starttls()
        server.login(my_sender, my_pass)  # 括号中对应的是发件人邮箱账号、邮箱密码
        server.sendmail(my_sender, [receiverMail], msg.as_string())  # 括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.quit()  # 关闭连接
    except Exception:  # 如果 try 中的语句没有执行，则会执行下面的 ret=False
        ret = False

    return ret

for i in range(0,233):
    receiver=getReceiver(i)

    ret = mail(receiver['序号'],receiver['姓名'],receiver['邮箱'],receiver['账号'],receiver['密码'])
    if ret:
        print("邮件发送成功,收件人:"+receiver['姓名']+";收件人邮箱:"+receiver['邮箱']+";序号:"+receiver['序号'])
    else:
        print("邮件发送失败")
