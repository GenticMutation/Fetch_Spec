# -*- coding: utf-8 -*-
"""
Created on 3/15/2024
@author: jdeng4/Quzhi
"""
import datetime, os, re,shutil, pathlib, sys,csv
import subprocess
import logging, getpass, requests,yaml
import json

import requests_ntlm
from requests_kerberos import HTTPKerberosAuth, REQUIRED, OPTIONAL
from prettytable import PrettyTable
from cryptography.fernet import Fernet
#import win32com.client as win32


working_path=os.getcwd()  #r'C:\Python\Fething Spec\dist\FetchSPEC' #
File_path = pathlib.Path(os.path.join(working_path, "Doc"))
cert_url=r"IntelSHA256RootCA-base64.crt"
os.chdir(working_path)


class SPECDownloader:
    def __init__(self, cert_url:str=False) -> None:
        self.url = 'http://cdpspwsfclb1.cd.intel.com:95'
        self.session = requests.Session()
        self.session.auth = requests_ntlm.HttpNtlmAuth('CCR\Jdeng4', 'Gentic@2122')
        self.headers = {
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0)'}
        self.session.headers = self.headers

    def login(self, domain:str=None, username:str=None, password:str=None,sub_url=None, method=None):

        form_dict = {}
        form_dict.setdefault('Domain', os.getenv('USERDOMAIN') if domain is None else domain)
        form_dict.setdefault('userName', os.getenv('USERNAME') if username is None else username)
        form_dict.setdefault('Password','Gentic@2122') #getpass.getpass())  #userPass
        response = self.session.post(url=self.url+sub_url, data=form_dict)
        try:
            response_Dict = json.loads(response.text)
            response_result = response_Dict['result']
            response_Sub_Url = response_Dict['url']
            response_load = self.session.get(url=self.url+response_Sub_Url,data=form_dict)

            search_head={'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; Touch; .NET4.0C; .NET4.0E; Tablet PC 2.0; InfoPath.3)'}
            self.session.headers=search_head
            search_data={"documentNumber":"","operation_equip":"","checklistId":"","description":"","cabinet":"y","vline":"","loggedUser":r"CCR\JDENG4"}
            response_Search=self.session.post(url=self.url+'/RnUUI/PostRnU/SearchData',data=search_data)

            Spec_Down=self.session.get(self.url+'/RnUUI/PostRnUSearch/DisplayPDF?specId=121-0124&rev=148&pdfStatus=')
        except Exception as e:
            print(f'login_response is not an dict')

        return response.text

    def login_Vista(self, domain:str=None, username:str=None, password:str=None):
        #self.url = 'https://s46-apps-vista.s46prod.mfg.intel.com/Vista/Specs/'
        self.url = 'http://cdpspwsfclb1.cd.intel.com:95/RnUUI/'
        self.session = requests.Session()
        self.session.auth = HTTPKerberosAuth(mutual_authentication=REQUIRED, force_preemptive=True,)
        self.session.auth=requests_ntlm.HttpNtlmAuth('CCR\Jdeng4','Gentic@2122')
        self.headers = { 'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0)'}
        self.session.headers = self.headers
        self.session.verify= cert_url
        self.spec_list_headers = ['Spec ID', 'Rev #', 'Description', 'Effective Date', 'Owner', 'RnU State', 'Post Date']
        response = self.session.get(url=self.url + '1dms.asp')
        form_dict = {}
        form_dict.setdefault('userDomain', os.getenv('USERDOMAIN') if domain is None else domain)
        form_dict.setdefault('userName', os.getenv('USERNAME') if username is None else username)
        if password is None:
            form_dict.setdefault('Password',getpass.getpass())  #userPass
        else:
            form_dict.setdefault('Password', password)
        response = self.session.post(url=self.url+'1dms.asp', data=form_dict)
        psw_Patten='The user name or password is incorrect'
        if re.search(psw_Patten,response.text):
            return False
        else:
            return True

    def download_spec_by_id(self, spec_id:str, dst_url:str=None):
        latest_spec_list = self.get_latest_spec_list()
        spec_info = None
        for item in latest_spec_list:
            if item[self.spec_list_headers[0]].strip() == spec_id.strip():
                spec_info = item
                break
        if spec_info is None:
            #logger.error(f"SPEC {spec_id} not found!")
            return
        spec_rev = spec_info[self.spec_list_headers[1]]
        self.print_spec_list([spec_info])

        target_url = f'frameset.asp?docID={spec_id}&version={spec_info[self.spec_list_headers[1]]}&spectype=doc'
       #logger.info(f"Fetching spec from URL: {self.url+target_url}")

        response = self.session.get(url=self.url + target_url)
        #logger.info(f'Fetching status code: {response.status_code}')


        headers = self.headers
        headers.setdefault('Accept', 'image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*')
        headers.setdefault('Accept-Encoding', 'gzip, deflate')
        headers.setdefault('Accept-Language', 'en-US, en; q=0.5')
        headers.setdefault('Connection', 'Keep-Alive')
        headers.setdefault('Host', 's46-apps-vista.s46prod.mfg.intel.com' )
        headers.setdefault('Referer', self.url + target_url)

        #nav_url = f'navigateFrame.asp?docID={spec_id}&version={spec_rev}'
        #response = self.session.get(url=self.url + nav_url, headers=headers)
        #logger.info(f'NavigateFrame status code: {response.status_code}')


        docview_url = f'docview.asp?docID={spec_id}&version={spec_rev}&fontbase=4'
        self.response = self.session.get(url=self.url + docview_url, headers=headers)
        #logger.info(self.response.headers)
        #logger.info(self.response.apparent_encoding)
        #logger.info(f'Spec ID : ({spec_id}.{spec_rev}) ; Encoding:{self.response.apparent_encoding}')

        #logger.info(f'DocViewFrame status code: {response.status_code}')
        os.chdir(working_path)
        if dst_url is None:
            dst_url = f'{spec_id}_{spec_rev}'

        folder_path = pathlib.Path(dst_url)
        shutil.rmtree(folder_path, ignore_errors=True)
        try:
            folder_path.mkdir()
            #print(os.chdir(os.path.join(working_path,folder_path)))
            #print(folder_path)
        except:
            pass
        #logger.info(self.response.encoding)
        self.response.encoding=self.response.apparent_encoding  # 'windows-1250'
        #logger.info(self.response.encoding)

        encoding_Dict={'windows-1250':'cp1250','windows-1252':'cp1252'}

        f = open(folder_path.joinpath(folder_path.name + '.html'), 'w', encoding=self.response.apparent_encoding)#'cp1250')
        f.write(self.response.text)  #.replace('','')
        f.close()

        pattern = '(?:\w|\s|\d)+\.(?:jpg|jpeg|png|gif|bmp|jif)'
        #href = "temp_image019.jpg"
        image_names = set()

        for item in re.findall(pattern, self.response.text, re.MULTILINE):
            image_names.add(item)
        spec_root_url = os.path.dirname(self.response.url)
        #logger.info(f'{len(image_names)} images found in SPEC ROOT {spec_root_url}')
        #out_put_log(Spec_Id=spec_id,Spec_Rev=spec_rev,Comments=f'{len(image_names)} images found in SPEC ROOT {spec_root_url}')
        QTY=0
        for each_image in image_names:
            r = self.session.get(url=f'{spec_root_url}/{each_image}', stream=True)
            # print(r.text)
            # print(r.headers)
            # print(r.request.headers)
            if r.status_code == 200:
                data = b''
                with open(folder_path.joinpath(each_image), 'wb') as f:
                    # r.raw.decode_content = True
                    for chunk in r.iter_content(chunk_size=1024):
                        if (chunk):
                            data += chunk
                    f.write(data)
                QTY +=1
                Finish_Rate='{:.2%}'.format(QTY/len(image_names))
                #click.echo('This is %d img file,total %d . Comp Rate: %s' % (QTY,len(image_names),Finish_Rate), nl=True)
            # break
        #click.echo()


if __name__ == '__main__':
    domain=None
    username=None
    #password='Gentic@2122'
    app =SPECDownloader(cert_url=cert_url)
    login_response=app.login(domain=domain, username=username,sub_url='/RnUUI/Login/AuthenticateAuthorizeUser',method='post')

