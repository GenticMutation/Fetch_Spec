# -*- coding: utf-8 -*-
"""
Created on 3/15/2024
@author: jdeng4/Quzhi
"""
import datetime, os, re,shutil, pathlib, sys,csv
import subprocess
import click, logging, getpass, requests,yaml
from requests_kerberos import HTTPKerberosAuth, REQUIRED, OPTIONAL
from prettytable import PrettyTable
from cryptography.fernet import Fernet
import win32com.client as win32
#import Mail_Sender
import shutil

#Mail_Sender.send_mail()

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)
logging.getLogger('requests_kerberos').setLevel(logging.WARNING)
working_path=r'E:\Fetch_Spec\FetchSpec6.3.2024'#'#os.getcwd()  #r'C:\Python\Fething Spec\dist\FetchSPEC' #
File_path = pathlib.Path(os.path.join(working_path, "Doc"))
File_History_path = pathlib.Path(os.path.join(working_path, "Doc_History"))
cert_url=pathlib.Path(os.path.join(working_path, "IntelSHA256RootCA-base64.crt"))
os.chdir(working_path)

try:
    File_path.mkdir()
except:
    pass
try:
    File_History_path.mkdir()
except:
    pass


# Decode Password
Config_Path = pathlib.Path(os.path.join(working_path, "script_config.yaml"))
config=yaml.load(open(Config_Path),yaml.FullLoader)
key=b'4r8UdYQkO-brheWle6xF7CXHmw7KJX8nulSwY7Vza1Y='
cipher_suite=Fernet(key)
encrypted_password=config['password']
decrypted_password=cipher_suite.decrypt(encrypted_password).decode()
Spec_Id_Source=config['SpecID_List']


def out_put_log( Spec_Id,Spec_Rev, Comments):
    '''

    Args: wirte log into log file
        Spec_Id: str
        Spec_Rev: str
        Comments: str

    Returns: NA

    '''
    writer.writerow([datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),Spec_Id,Spec_Rev, Comments])
class SPECDownloader:
    def __init__(self, cert_url:str=False) -> None:
        self.url = 'https://s46-apps-vista.s46prod.mfg.intel.com/Vista/Specs/'
        self.session = requests.Session()
        self.session.auth = HTTPKerberosAuth(mutual_authentication=OPTIONAL, force_preemptive=True)
        self.headers = { 'User-Agent': 'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.1; Trident/4.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0)'}
        self.session.headers = self.headers
        self.session.verify= cert_url
        self.spec_list_headers = ['Spec ID', 'Rev #', 'Description', 'Effective Date', 'Owner', 'RnU State', 'Post Date']

    def login(self, domain:str=None, username:str=None, password:str=None):
        '''

        Args: Login to sever
            domain: str
            username: str
            password: str

        Returns: True/False

        '''
        response = self.session.get(url=self.url + '1dms.asp')
        form_dict = {}
        form_dict.setdefault('userDomain', os.getenv('USERDOMAIN') if domain is None else domain)
        form_dict.setdefault('userName', os.getenv('USERNAME') if username is None else username)
        if password is None:
            click.echo(f'Logging in as {form_dict["userDomain"]}/{form_dict["userName"]}:')
            form_dict.setdefault('userPass', getpass.getpass())
        else:
            form_dict.setdefault('userPass', password)
        logger.info(f'Logging in as {form_dict.get("userDomain")}/{form_dict.get("userName")}')
        response = self.session.post(url=self.url + '1dms.asp', data=form_dict)
        logger.info(f"Logging in status: {response.status_code}")

        psw_Patten='The user name or password is incorrect'
        if re.search(psw_Patten,response.text):
            return False
        else:
            return True

    def save_as_html(self, text:str, dst_url:str):
        f = open(dst_url, 'w', encoding='utf-8')
        f.write(text)
        f.close()

    def get_latest_spec_list(self, text_url:str=None, print_to_console:bool=False):
        '''

        Args:
            text_url: str
            print_to_console: bool

        Returns: List

        '''

        logger.info("Querying latest SPEC list ...")

        form_dict = {'btnExit': 'Un-Register', 'btnSendBehind': 'Send Behind', 'btnHelp': 'Help',
                    'btnSearch': 'Search', 'btnClear': 'btnClear', 'sSeries': '',  'sFuncArea': '',
                    'sSerialNum': '', 'sSpecType': '', 'sEquipmentID': '', 'sChecklist': '',
                    'sDescription': '', 'sOwner': '', 'sVersion': 'C'}
        if text_url is None:
            response = self.session.post(url=self.url + '1dms.asp', data=form_dict)
            logger.info(f"Querying status: {response.status_code}")
            content = response.text
        else:
            content = open(text_url, 'r', encoding='utf8').read()
        pattern = 'new Array\("(?P<id>[0-9A-Za-z\-]+)",\s+"(?P<ver>\d+)",\s+"(?P<desc>.+)",\s+"(?P<date>.+)",\s+"(?P<owner>.+)",\s+"(?P<rnu_state>.*)",\s+"(?P<post_date>.*)"\),'
        spec_list = list()
        for item in re.findall(pattern, content, re.MULTILINE):
            spec_info = dict()
            for idx, name in enumerate(self.spec_list_headers):
                spec_info.setdefault(name, item[idx])
            spec_list.append(spec_info)

        if print_to_console:
            self.print_spec_list(spec_list)

        return spec_list

    def print_spec_list(self, spec_list:list):
        p = PrettyTable()
        p.field_names = self.spec_list_headers
        p.align = 'l'
        for each_row in spec_list:
            p.add_row([ str(each_row[x]) for x in each_row])
        print(p)

    def html2word(self,spec_id,Spec_Rev):
        os.chdir(os.path.join(working_path, r"%s_%s" % (spec_id, str(Spec_Rev))))
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')

            targetfile=os.path.join(os.getcwd(), f"{spec_id}_{Spec_Rev}.html" )
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) get target file Successfully')
            logger.info(f'Target_file is: {targetfile}')
            doc = word.Documents.Open(targetfile)#f'{spec_id}_{Spec_Rev}.html')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) word open html Successfully \n'
                                 f'Target_File_Path: {targetfile}')

            Save_To_Path = os.path.join(working_path, r"Doc\%s_%s.Docx" % (spec_id, str(Spec_Rev)))
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments= f'save to path: {Save_To_Path}')
            doc.SaveAs(Save_To_Path, FileFormat=16)
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) html save as word Successfully,\n'
                                 f'save to path: {Save_To_Path}')

            #filename=os.path.join(working_path, r"Doc\%s_%s.Docx" % (spec_id, str(Spec_Rev)))
            doc.Close(True)#,filename)
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) word close Successfully')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) app quit Successfully')
            word.Quit()
        except Exception  as e :
            print(f'word app load fail, failed msg: {e}')
        '''
        #print(f'Spec Converter working Path={os.getcwd()}')
        html=open("%s_%s.html" % (spec_id,str(Spec_Rev))).read()
        pypandoc.convert_text(html, 'docx',outputfile=os.path.join(working_path,r"Doc\%s_%s.Docx" % (spec_id,str(Spec_Rev))),format='html')
        '''
        os.chdir(working_path)

    def download_spec_by_id(self, spec_id:str, dst_url:str=None):
        latest_spec_list = self.get_latest_spec_list()
        spec_info = None
        for item in latest_spec_list:
            if item[self.spec_list_headers[0]].strip() == spec_id.strip():
                spec_info = item
                break
        if spec_info is None:
            logger.error(f"SPEC {spec_id} not found!")
            return
        spec_rev = spec_info[self.spec_list_headers[1]]
        self.print_spec_list([spec_info])

        target_url = f'frameset.asp?docID={spec_id}&version={spec_info[self.spec_list_headers[1]]}&spectype=doc'
        logger.info(f"Fetching spec from URL: {self.url+target_url}")

        response = self.session.get(url=self.url + target_url)
        logger.info(f'Fetching status code: {response.status_code}')


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
        logger.info(self.response.headers)
        logger.info(self.response.apparent_encoding)
        #logger.info(f'Spec ID : ({spec_id}.{spec_rev}) ; Encoding:{self.response.apparent_encoding}')

        logger.info(f'DocViewFrame status code: {response.status_code}')
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
        logger.info(self.response.encoding)
        self.response.encoding=self.response.apparent_encoding  # 'windows-1250'
        logger.info(self.response.encoding)

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
        logger.info(f'{len(image_names)} images found in SPEC ROOT {spec_root_url}')
        out_put_log(Spec_Id=spec_id,Spec_Rev=spec_rev,Comments=f'{len(image_names)} images found in SPEC ROOT {spec_root_url}')
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
                click.echo('This is %d img file,total %d . Comp Rate: %s' % (QTY,len(image_names),Finish_Rate), nl=True)
            # break
        click.echo()

def Spec_exist(Spec_ID,Spec_rev):
    path=[]
    SpecID_Rev=Spec_ID+"_"+str(Spec_rev)
    for i in os.listdir(File_path):
        if os.path.isfile(os.path.join(File_path,i)):
            path.append(i.split('.')[0])

    if SpecID_Rev in path:
        return True
    else:
        return False
def Store_Spec():
    Spec_List_Vista_list =[]
    for item, value in enumerate(Spec_List_Vista):
        Spec_List_Vista_list.append('_'.join([value['Spec ID'],value['Rev #']]))

    for i in os.listdir(File_path):
        specId_rev=i.split('.')[:-1]
        if os.path.isfile(os.path.join(File_path, i)) and not (specId_rev[0] in Spec_List_Vista_list) :
            shutil.move(File_path.joinpath(i),File_History_path)
            print(f"File {i} moved to {File_History_path}")


if __name__ == '__main__':
    domain=None
    username=None
    password=decrypted_password
    app =SPECDownloader(cert_url=cert_url)
    fp = open(os.path.join(working_path, 'log-' + str(datetime.date.today().strftime('%Y-%m-%d')) + '.txt'), mode='a+',
              encoding='utf-8-sig', newline='\n')
    writer = csv.writer(fp)
    writer.writerow(['Date', 'Spec_Id', 'Spec_Rev', 'Comments'])
    # try to Login to Vista, Feedback if Login Fails
    if  app.login(domain=domain, username=username, password=password) ==False:
        logger.info('Login Fail, The user name or password is incorrect... Check and Retry')
        out_put_log(Spec_Id="", Spec_Rev="",
                    Comments='Login Fail, The user name or password is incorrect... Check and Retry')
        sys.exit()
    # Get the Lastest Spec list
    Spec_List_Vista=app.get_latest_spec_list(text_url=None,print_to_console=False)

    for i in Spec_List_Vista:
        SpecId=i['Spec ID']
        SpecID_Patten = '[A-Z]{1,3}[0-9]{2}-[0-9]{2}-[0-9]{4}-[0-9]{3}'
        try:
            SpecID_to_Match=re.search(SpecID_Patten, SpecId).group()
        except:
            SpecID_to_Match=''
        print(f'Match Spec is: {SpecID_to_Match}')
        Spec_Rev=i['Rev #']
        Spec_Effective_Date = i['Effective Date']
        Spec_des=i['Description']
        Spec_Owner=i['Owner']

        if   SpecID_to_Match in Spec_Id_Source and Spec_exist(SpecId,Spec_Rev)==True:
            print(f'{datetime.datetime.now()} : Spec NO.( {SpecId}. {Spec_Rev} ) Already in the local driver,  Pass..')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) Already in the local driver,  Pass..')

        elif SpecID_to_Match in Spec_Id_Source and Spec_exist(SpecId,Spec_Rev)==False:
            app.download_spec_by_id(spec_id=SpecId)
            print(f'{datetime.datetime.now()} : Spec NO.( {SpecId}. {Spec_Rev} ) Download Successfully')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) Download Successfully')
            app.html2word(spec_id=SpecId,Spec_Rev=Spec_Rev)
            print(f'{datetime.datetime.now()} : Spec NO.( {SpecId}. {Spec_Rev} ) Converate to Docx Successfully')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) Converate to Docx Successfully')
            History_Log=open(os.path.join(working_path,"Spec_History.txt"),mode='a+',newline='\n',encoding='utf-8')
            History_Log.write(f'\n {SpecId},{Spec_Rev},{Spec_des},{Spec_Effective_Date},{Spec_Owner},{datetime.datetime.now()}')
            History_Log.close()
        else:
            print(f'{datetime.datetime.now()} : Spec NO.( {SpecId}. {Spec_Rev} ) is not in the target list ,  Pass..')
            out_put_log(Spec_Id=SpecId, Spec_Rev=Spec_Rev,
                        Comments=f'Spec NO.( {SpecId}. {Spec_Rev} ) is not in the target list ,  Pass..')
    Store_Spec()
    logger.info('Spec Fetch Succeed!')
    out_put_log(Spec_Id="", Spec_Rev="",
                Comments='Spec Fetch Succeed!')
    fp.close()

'''
Versions:
Build Main Frame (vista Spec Downloader to Html)
Build Confog File
Build Htmlto Word Function, reply on ' pypandoc
Password Function (password generator,decorder..)
Update setting function to yaml
Bug Fix from local html Encoding
BUild logger Function


'''
