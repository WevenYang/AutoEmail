# -*- coding: utf-8 -*-
import re
import smtplib
from email.mime.text import MIMEText


class EmailUtil:

    def __init__(self, arr, title):
        self.xuhao = arr[0]
        self.bumen = arr[1]
        self.gonghao = arr[2]
        self.xingming = arr[3]
        self.yuexin = arr[4]
        self.jibengongzi = arr[7]
        self.tongxunbuzhu = arr[8]
        self.jiaotongbuzhu = arr[9]
        self.gangweigongzi = arr[10]
        self.jixiaogongzi = arr[11]
        self.quanqinjiang = arr[12]
        self.butie = arr[13]
        self.diaozengjiam = arr[15]
        self.yingfaheji = arr[16]
        self.shebaoyibao = arr[17]
        self.zhufanggongjijin = arr[18]
        self.gerensuodesui = arr[20]
        self.shifaheji = arr[22]
        self.beizhu = arr[23]
        self.email = arr[24]
        self.title = title
        # self._user = "594771590@qq.com"     #发送者email
        # self._pwd = "ajvkbaysjtxbbejd"      #发送者email密码
        self._user = "604060689@qq.com"  # 发送者email
        self._pwd = "gjfzgmuvgcuqbejj"  # 发送者email密码

    def send_by_email(self):
        #由于该字符串的拼接方式会对使百分号%变得敏感，所以在html的宽度属性需要加两个百分号区分
        htmlcontent = '<!DOCTYPE html><html><head><style type="text/css">td{text-align: center;}body{' \
                      'padding:20px;}</style></head><body><h1 style="text-align:center">广东普惠智能教育技术有限公司</h1><h1 ' \
                      'style="text-align:center">' \
                      + self.title + '</h1><p style="float:right;">单位：RMB元</p><table border="1" cellspacing="0" ' \
                                     'width="100%%"><tr><td rowspan="2">序号</td><td rowspan="2">部门</td><td ' \
                                     'rowspan="2">工号</td><td rowspan="2">姓名</td><td rowspan="2">月薪</td><th ' \
                                     'colspan="8">应发工资</th><td rowspan="2">应发合计</td><th colspan="3">扣除项目</th><td ' \
                                     'rowspan="2">实发合计</td><td ' \
                                     'rowspan="2">备注</td></tr><tr><td>基本工资</td><td>通讯补助</td><td>交通补助</td><td>岗位工资</td' \
                                     '><td>绩效工资</td><td>全勤奖</td><td>补贴</td><td>调增减</td><td>社保/医保</td><td>住房公积金</td' \
                                     '><td>个人所得税</td></tr><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td' \
                                     '>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td' \
                                     '>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td><td>%s</td' \
                                     '></table></body></html>' % (str(self.xuhao), str(self.bumen), str(self.gonghao),
                                                                  str(self.xingming), str(self.yuexin),
                                                                  str(self.jibengongzi),
                                                                  str(self.tongxunbuzhu), str(self.jiaotongbuzhu),
                                                                  str(self.gangweigongzi), str(self.jixiaogongzi),
                                                                  str(self.quanqinjiang), str(self.butie),
                                                                  str(self.diaozengjiam),
                                                                  str(self.yingfaheji), str(self.shebaoyibao),
                                                                  str(self.zhufanggongjijin), str(self.gerensuodesui),
                                                                  str(self.shifaheji), str(self.beizhu))

        msg = MIMEText(htmlcontent, 'html', 'utf-8')
        msg["Subject"] = self.title  # 发送的标题
        msg["From"] = self._user  # 发送人
        msg["To"] = self.email  # 发送对象

        try:
            #正则匹配是否是QQ邮箱
            if re.match(r'[0-9a-zA-Z_]{0,19}@qq.com', self.email):
                s = smtplib.SMTP_SSL("smtp.qq.com", 465)
                s.login(self._user, self._pwd)
                s.sendmail(self._user, self.email, msg.as_string())
                s.quit()
                print("Success!")
                return True
            else:
                print("识别到不是邮箱格式")
                return False
        except smtplib.SMTPException as e:
            print("Falied,%s" % e)
            return False
