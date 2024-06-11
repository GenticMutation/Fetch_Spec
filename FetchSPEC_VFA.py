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
Config_Path = pathlib.Path(os.path.join(working_path, "script_config.yaml"))
config=yaml.load(open(Config_Path),yaml.FullLoader)
key=b'4r8UdYQkO-brheWle6xF7CXHmw7KJX8nulSwY7Vza1Y='
cipher_suite=Fernet(key)
encrypted_password=config['password']
decrypted_password=cipher_suite.decrypt(encrypted_password).decode()
Spec_Id_Source=config['SpecID_List_Vista']


class SPECDownloader:
    def __init__(self, cert_url:str=False) -> None:
        self.url = 'http://cdpspwsfclb1.cd.intel.com:95'
        self.session = requests.Session()
        self.session.auth = requests_ntlm.HttpNtlmAuth('CCR\Jdeng4', 'Gentic@2122')
        self.headers = {
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0)'}
        self.search_head = {
            'User-Agent': 'Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; Touch; .NET4.0C; .NET4.0E; Tablet PC 2.0; InfoPath.3)'}
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
            if response_result=='Redirect':
                response_load = self.session.get(url=self.url+response_Sub_Url,data=form_dict)

        except Exception as e:
            print(f'login_response is not an dict')

        return response.text
    def get_latest_spec_list(self):
        '''
        Returns: List of Spec ID
        '''
        self.session.headers = self.search_head
        search_data = {"documentNumber": "", "operation_equip": "", "checklistId": "", "description": "",
                       "cabinet": "y", "vline": "", "loggedUser": r"CCR\JDENG4"}
        self.response_Search = self.session.post(url=self.url + '/RnUUI/PostRnU/SearchData', data=search_data)
        self.response_Search_dict=json.loads(self.response_Search.text)['responsestring']['Table']
        Spec_List=[]
        for i in self.response_Search_dict:
            Spec_List.append([i['SPEC_ID'],i['REV'],i['DESCRIPTION'],i['OWNER'],i['EFFECTIVE_DATE']])

        print(Spec_List)
        return Spec_List

    def download_spec_by_id(self, spec_id:str,spec_Rev, dst_url:str=None):
        '''

        Args: Download Spec Base on provided Spec Id and Spec rev
            spec_id: Str Data
            spec_Rev: Str or int
            dst_url: Str

        Returns: NA

        '''
        target_url = f'/RnUUI/PostRnUSearch/DisplayPDF?specId={spec_id}&rev={spec_Rev}&pdfStatus='
       #logger.info(f"Fetching spec from URL: {self.url+target_url}")

        self.response = self.session.get(url=self.url + target_url)
        #logger.info(f'Fetching status code: {response.status_code}')
        folder_path=f'{spec_id}.{spec_Rev}.pdf'

        f = open(os.path.join(File_path,folder_path), 'wb')#, encoding='utf-8')#'cp1250')
        f.write(self.response.content)  #.replace('','')
        f.close()




if __name__ == '__main__':
    domain=None
    username=None
    #password='Gentic@2122'
    app =SPECDownloader(cert_url=cert_url)
    login_response=app.login(domain=domain, username=username,sub_url='/RnUUI/Login/AuthenticateAuthorizeUser',method='post')
    Spec_List=app.get_latest_spec_list()
    for i in Spec_List:
        spec_Id=i[0]
        spec_rev=i[1]
        if spec_Id in Spec_Id_Source:
            app.download_spec_by_id(spec_id=spec_Id,spec_Rev=spec_rev)
            print(f'spec {spec_Id}.{spec_rev}download successfully')
        else:
            print(f'spec {spec_Id}.{spec_rev} not in the defined list')


