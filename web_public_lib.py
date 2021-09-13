import re,os,json,xlrd
from django.http import HttpResponse,StreamingHttpResponse
from django.core.exceptions import ObjectDoesNotExist, MultipleObjectsReturned
from django.utils.encoding import escape_uri_path
import datetime,math
import requests,os,re
import subprocess, copy,glob
from pathlib import Path
import urllib
import pandas as pd
import urllib.request
class File:
    def __init__(self,*args, **kwargs):
        pass
    @staticmethod
    def Webupload(self,srcf,outdir):
        '''uploadfile and save'''
        outpath = os.path.join(outDir, srcf.name)
        outpath = re.sub("[()£¨£©]", "_", outpath)
        srcf.name = re.sub("[()£¨£©]", "_", srcf.name)
        try:
            with open(outpath, 'wb+') as file:
                for chunk in srcf.chunks():
                    file.write(chunk)
        except Exception as e:
            printErr(e.message)
            return False,outpath
        return True,outpath
    @staticmethod
    def Webdownload(self,filepath,filetype=""):
        '''download file'''
        def file_iter(filepath, blockSize = 4096):
            with open(filepath,"rb") as iFile:
                while True:
                    buf = iFile.read(blockSize)
                    if buf:
                        yield buf
                    else:
                        break
        if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
            fname = os.path.basename(file_path)
            response = StreamingHttpResponse(file_iter(file_path))
            ContentType = 'application/octet-stream'
    #         response['Content-Type'] = 'application/octet-stream'
            if filetype == "word":
                ContentType = 'application/msword;charset=UTF-8'           
            response['Content-Type'] = ContentType
            response["Content-Disposition"] = "attachment;filename*=UTF-8''{}".format(escape_uri_path(fname))
            return response
        else:
            return HttpResponse('Error! The file is not existed or can not be read.'+file_path)
    @staticmethod
    def Findfilepath(self,filename,Dirlist):
        '''find abs filepath'''
        filepath = ""
        filepathArr = []
        for Dir in Dirlist:
            filepath = Path(Dir).rglob(f'*{filename}*')  #
            for path in filepath:
                filepathArr.append(path.__str__())
        return filepathArr
    @staticmethod
    def Excel2df(self,filepath,sheet_num=0,start_index=0,end_index=None):
        '''excel to df'''
        data = xlrd.open_workbook(filepath)  
        table = data.sheets()[sheet_num]  
        if not end_index:
            end_index = table.nrows
        data = {}
        head = table.row_values(start_index)
        # data.keys()= head
        for i in range(start_index+1, end_index):
            row = table.row_values(i)
            for n, m in enumerate(head):
                col = []
                if m in data.keys():
                    col = data.get(m)
                col.append(row[n])
                data[m] = col
        df = pd.DataFrame(data)
        return df

    def __str__(self):
        return f"relation file"
class Web:
    def __init__(self,*args, **kwargs):
        pass
    @staticmethod
    def requests2obj(self,request,obj):
        data_dict=request.POST.dict()
        for k,v in data_dict.items():
            v=v.strip()
            if not v:
                v=None
            elif v.upper()=="YES":
                v=True
            elif v.upper()=="NO":
                v=False
            setattr(obj,k,v)
        return obj
        
    def exeCmd(cmdArgs, bshell=False):
        '''execute analysis program
        cmdArgs: type:list.
        return value: tuple, (exitcode, stdout, stderr)'''
        #child = subprocess.Popen(cmdArgs, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=bshell)
        shellEnv = os.environ.copy()
        print(' '.join(cmdArgs))
        if bshell:
            child = subprocess.Popen(' '.join(cmdArgs), stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True, env=shellEnv)
        else:
            child = subprocess.Popen(cmdArgs, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=False, env=shellEnv)
        child.wait()
        return (child.returncode, child.stdout.read(), child.stderr.read())
    @staticmethod
    def opRecord(cmd,argstest):
    '''decorator for recording user operation.
    arguments: Array, ['paraName1', 'paraName2']
    '''
        
    #arguments=copy.deepcopy(argstest)
        def make_wrapper(func):
            def recordLog(fParams,request, op, fname):
                userId = request.session.get('user_id', 1)
                userName = request.session.get('username', '')
                #print(request.GET, type(request.GET))
                paramInfo = {}
                for k,v in request.POST.items():
                    vs = str(v)
                    if len(vs) > FIELD_MAX_LEN:
                        # vs = vs[:FIELD_MAX_LEN:]
                        vs = 'ignore...'
                    paramInfo[k] = vs
                post_params = json.dumps(paramInfo)[:MAX_LEN:]
                paramInfo = {}
                for k,v in request.GET.items():
                    vs = str(v)
                    if len(vs) > FIELD_MAX_LEN:
                        vs = vs[:FIELD_MAX_LEN:]
                    paramInfo[k] = vs
                get_params = json.dumps(paramInfo)[:MAX_LEN:]
                obj = operate_log.objects.create(user_id=userId)
                obj.user_id=userId 
                obj.user_name=userName
                obj.operation=op
                obj.post_params=post_params
                obj.get_params=get_params
                obj.func_name=fname 
                obj.func_params=fParams
                obj.operate_time = datetime.now()
                obj.save()
            def wrapper(*args, **kwargs):
                fcode = func.__code__
                #print(fcode)
                names = list(fcode.co_varnames[:fcode.co_argcount])
                argsArr = list(args)
                #print(names, arguments, argsArr, kwargs)
                paradic = {}
                delArr = []
                arguments=copy.deepcopy(argstest)       #very important
                for argument in arguments:
                    if argument in kwargs:
                        paradic[argument] = kwargs[argument]
                        delArr.append(argument)
                for d in delArr:
                    arguments.remove(d)
                for i in range(0, min(len(arguments), len(argsArr))):
                    paradic[arguments[i]] = argsArr[i]
                rq = paradic.get('request', None)
                if rq:
                    print('============')
                    print("userId: %s\toperation: %s" % (str(rq.session['username']), cmd))
                    print('============')
                if 'request' in paradic:
                    del paradic['request']
                recordLog(json.dumps(paradic),rq,cmd,func.__name__)
                return func(*args, **kwargs)
            return wrapper
        return make_wrapper
    def __str__(self):
        return f"relation Web"
