#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import poplib,email,telnetlib
import datetime,time,sys,traceback
from datetime import datetime as mydatetime
from time import mktime
import json
import os
import pandas as pd
import xlwt
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

addresser_dict = {}
user = ""
password = ""
email_server = ""
start_time = ""
end_time = ""
output_file = ""
write_res_excel = xlwt.Workbook(encoding = 'utf-8') 
worksheet = write_res_excel.add_sheet('sheet1',cell_overwrite_ok=True)
line_num = 1

class init_config():
    def __init__(self):
        with open('config.json', 'r') as f:
            content = f.read()
            config = json.loads(content)
            print(config)
            global addresser_dict, start_time, end_time, output_file, user, password, email_server
            user = config['user']
            password = config['password']
            email_server = config['email_server']
            addr_list = config['addresser_list']
            
            for addr_name in addr_list:
                addresser_dict[addr_name]= 1
                print(addr_name)
            
            time_range_str = config['time_range'].split(',') 
            if len(time_range_str)>1:
                start_time=time_range_str[0]
                end_time=time_range_str[1]
                if start_time>end_time:
                    print("start_time is bigger than end_time")
                    exit(0)
            else:
                start_time=time.strftime("%Y-%m-%d", time.localtime()) 
                end_time=start_time
                
            print(start_time,end_time)
            output_file = config['output_xls']

            worksheet.write(0,1, label = '今日单位净值')
            worksheet.write(0,2, label = '应交税费')
            write_res_excel.save(output_file)    

class down_email():

    def __init__(self):
        # 输入邮件地址, 口令和POP3服务器地址:
        global user, password, email_server
        self.user = user
        self.password = password
        self.pop3_server = email_server 

    def read_excel_line(self,sheets_list,f_path,f_name):
        global line_num
        for sheet in sheets_list:
            filled_num = 0
            df=pd.read_excel(f_path, sheet,na_values='#N') #,names="今日单位净值："
            print(df)
            for line in df.values:
                if filled_num == 2:
                    line_num = line_num + 1
                    break
                print(line)
                if str(line[1]) == "应交税费" or  str(line[1]) == "应付税费" :
                    f_sheet = f_name+sheet
                    worksheet.write(line_num,0, label = f_sheet)
                    if str(line[4]) == 'nan':  
                        worksheet.write(line_num,2, label = line[7]) 
                        print(line[7])
                    else:
                        worksheet.write(line_num,2, label = line[4]) 
                        print(line[4])
                    filled_num = filled_num + 1

                if str(line[0]).find("今日单位净值") != -1:
                    worksheet.write(line_num,1, label = line[1]) 
                    filled_num = filled_num + 1
                    print(line[1])

    def get_excel_info(self):
        print(os.listdir('./file_save'))
        file_list=os.listdir('./file_save')
        for i in range (len(file_list)) :
            if (file_list[i].find(".xlsx") != -1): 
                f_name=str(file_list[i]).split("_")[0]
                # print("openpyxl")
                f_path = "./file_save/%s" % file_list[i]
                print(f_path)
                xlsx = pd.ExcelFile(f_path,"openpyxl")
                sheets_list = xlsx.sheet_names
                print(sheets_list)
                self.read_excel_line(sheets_list,f_path,f_name)

            elif (file_list[i].find(".xls") != -1) : 
                f_name=str(file_list[i]).split("_")[0]
                # print("xlrd")
                f_path = "./file_save/%s" % file_list[i]
                print(f_path)
                xls = pd.ExcelFile(f_path,"xlrd")
                sheets_list = xls.sheet_names
                print(sheets_list)
                self.read_excel_line(sheets_list,f_path,f_name)
            else:
                continue
        return 
    
    # 获得msg的编码
    def guess_charset(self,msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        return charset

    #获取邮件内容
    def get_content(self,msg):
        content=''
        content_type = msg.get_content_type()
        # print('content_type:',content_type)
        if content_type == 'text/plain': # or content_type == 'text/html'
            content = msg.get_payload(decode=True)
            charset = self.guess_charset(msg)
            if charset:
                content = content.decode(charset)
        return content

    # 字符编码转换
    def decode_str(self,str_in):
        value, charset = decode_header(str_in)[0]
        if charset:
            value = value.decode(charset)
        return value

    # 解析邮件,获取附件
    def get_att(self,msg_in):
        print("enter get attached files")
        attachment_files = []
        for part in msg_in.walk():
            # 获取附件名称类型
            file_name = part.get_param("name")  # 如果是附件，这里就会取出附件的文件名
            contType = part.get_content_type()
            if file_name=='noneType':
                print("is nonetype")
                print(self.get_content(part))
            elif file_name:
                print("is a file")
                h = email.header.Header(file_name)
                # 对附件名称进行解码
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # 将附件名称可读化
                    filename = self.decode_str(str(filename, dh[0][1]))
                    
                # 下载附件
                data = part.get_payload(decode=True)
                # 在指定目录下创建文件，二进制文件用wb模式打开
                path = "./file_save/"
                isExists=os.path.exists(path)
                if not isExists:
                    os.makedirs(path) 
                    print("创建目录",path,"成功")
                att_file = open(path + filename, 'wb')
                att_file.write(data)  # 保存附件
                att_file.close()
                attachment_files.append(filename)
            else:
                print("is a text, ",self.get_content(part))

        return attachment_files

    def run_ing(self):
        # 连接到POP3服务器,有些邮箱服务器需要ssl加密，可以使用poplib.POP3_SSL
        try:
            telnetlib.Telnet(self.pop3_server, 995)
            server = poplib.POP3_SSL(self.pop3_server, 995, timeout=10)
        except:
            time.sleep(5)
            print("not ssl connect")
            server = poplib.POP3(self.pop3_server, 110, timeout=10)
            server.set_debuglevel(1)

        # server.set_debuglevel(1) # 打开或关闭调试信息
        # 打印POP3服务器的欢迎文字:
        print(server.getwelcome().decode('utf-8'))
        # 身份认证:
        server.user(self.user)
        server.pass_(self.password)
        # 返回邮件数量和占用空间:
        print('Messages: %s. Size: %s' % server.stat())
        # list()返回所有邮件的编号:
        resp, mails, octets = server.list()
        print(mails)
        index = len(mails)
        global addresser_dict
        for i in range(index, 0, -1):
            resp, lines, octets = server.retr(i)
            # lines存储了邮件的原始文本的每一行
            msg_content = b'\r\n'.join(lines).decode('utf-8','ignore') 
            # 解析邮件:
            msg = Parser().parsestr(msg_content)
            #获取邮件的发件人，收件人， 抄送人,主题
            From = parseaddr(msg.get('from'))[1]
            To = parseaddr(msg.get('To'))[1]
            Cc=parseaddr(msg.get_all('Cc'))[1]# 抄送人
            date_str = msg.get('Date')
            dt=datetime.datetime.strptime(date_str, '%a, %d %b %Y %H:%M:%S %z').strftime('%Y-%m-%d') 
          
            Subject = self.decode_str(msg.get('Subject'))
            if From not in addresser_dict :
                continue
            print('from:%s,to:%s,Cc:%s,subject:%s,date_str:%s'%(From,To,Cc,Subject,dt))
            global start_time,end_time
            ddd=str(dt)
            if ddd < start_time or ddd > end_time:
                continue
            
            # 获取附件
            attach_file=self.get_att(msg)
            print(attach_file)

        server.quit()


if __name__ == '__main__':

    try:
        # config内输入邮件地址, 口令和POP3服务器地址:
        init_config()
        email_class=down_email()
        email_class.run_ing()
        email_class.get_excel_info()
        write_res_excel.save(output_file)
    except Exception as e:
        import traceback
        ex_msg = '{exception}'.format(exception=traceback.format_exc())
        print(ex_msg)
       
