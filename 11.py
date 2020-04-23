import os
import ftplib
import re
import csv
from xlwt import Workbook
import time
import pandas as pd
import zipfile

class Ftp_Load:
    def __init__(self):
        self.id = [2051, 3424, 2819, 2724, 2725, 108, 111, 194, 195, 395,
                   396, 1031, 2035, 3428, 3445, 886, 887,4260, 4261, 4262, 4263, 4264, 4265, 4266, 4267, 4268,
                   4269, 4270, 4271, 4272, 4273, 4274, 4275, 4276, 4277, 4278, 4279, 4280]
        self.ftp = ftplib.FTP('')
        self.ftp.login('')
        self.dir_id = [x for x in self.id if str(x) in self.ftp.nlst()]

    def csv_to_xls_conv(self, csv_file, excel=False):
        encodings = ['utf-8', 'windows-1251', 'windows-1252']
        for code in encodings:
            try:
                df = pd.read_excel(csv_file) if excel else pd.read_csv(csv_file,  sep=';', encoding=code)
                if 'Price' in df.columns:
                    df.rename(columns={'Price': 'amount'}, inplace=True)
                    df = df[['Name', 'amount']]
                writer = pd.ExcelWriter(csv_file.split('.')[0] + '.xls')
                df.to_excel(writer, 'Sheet1', index=False, encoding=code)
                writer.save()
            except Exception as e:
                print(e)

    def unzip_file(self, f_name):
        with zipfile.ZipFile(f_name, 'r') as zip_f:
            for name in zip_f.namelist():
                new_name = f_name[:-4] + name[name.rfind('.'):]

                with zip_f.open(name) as file:
                    content = file.read()
                    full_path = os.path.join(os.getcwd(), new_name)
                    with open(full_path, 'wb') as file_w:
                        file_w.write(content)
                    print(new_name)
                    self.csv_to_xls_conv(new_name, excel=True)

    def ftp_uploads(self, id, file_ftp):
        fin = open('%s' % file_ftp, 'rb')
        self.ftp.cwd('/%s' % id)
        self.ftp.storbinary('STOR %s' % file_ftp, fin)

    def uploads_conv_xls(self):

        for id in self.dir_id:

            print(id)
            self.ftp.cwd('/%s' % id)
            print(self.ftp.nlst())
            if len(self.ftp.nlst()) > 2:
                for fn in self.ftp.nlst():
                    if len(fn) > 3 and fn[fn.rfind('.csv'):].lower() == '.csv':
                        # print(fn)
                        print('Processing file {}'.format(fn))

                        self.ftp.voidcmd('TYPE I')
                        if self.ftp.size(fn) > 100:
                            self.ftp.retrbinary('RETR ' + fn, open(fn, 'wb').write)
                            self.csv_to_xls_conv(fn)
                            self.ftp.delete(fn)
                            self.ftp_uploads(id, fn[:-4] + '.xls')
                        else:
                            self.ftp.delete(fn)
                    elif len(fn) > 3 and fn[fn.rfind('.zip'):].lower() == '.zip':
                        # print(fn)
                        print('Processing file {}'.format(fn))

                        self.ftp.voidcmd('TYPE I')
                        self.ftp.retrbinary('RETR ' + fn, open(fn, 'wb').write)
                        self.unzip_file(fn)
                        self.csv_to_xls_conv(fn)
                        self.ftp.delete(fn)
                        self.ftp_uploads(id, fn[:-4] + '.xls')
        self.ftp.quit()


text = open('zvit.txt', 'a')
try:
    Ftp_Load().uploads_conv_xls()
except Exception as e:
    text.write('\n' + '\n' + time.strftime("%b_%d_%Y_%H_%M_") + '  ' + "ERROR:" + str(e) + '\n' + '\n')
text.close()

