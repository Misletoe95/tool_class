from configparser import ConfigParser
from jpype import *
import jpype
import os
def getconfig(config_path):
    parse = ConfigParser()
    parse.read(config_path,encoding="utf-8")
    
    ip  = parse.get("server","ip")
    port  = parse.get("server","port")
    text = parse.get("content","text")
    font_size = parse.get("content","font_size")
    style_num = parse.get("content","style_num")
    row_id = parse.get("content","row_id")
    return ip,int(port),text,int(font_size),int(style_num),row_id
def sendmessage(ip,port,text,font_size,style_num,row_id,data_path):
    jvm_path = getDefaultJVMPath()
    jar_lab_path = os.path.abspath('C://Users//Administrator//Downloads//BX_05_06_SDK_20210910//BX_05_06_SDK_20210910//JAVA SDK//lib//lib_06')
    abs_jar_lab_path = []
    jar_lab_list = os.listdir(jar_lab_path)
    for i in jar_lab_list:
        i = os.path.join(jar_lab_path,i)
        abs_jar_lab_path.append(i)
    clastr=";".join(abs_jar_lab_path)

    jpype.startJVM(jvm_path,'ea','-Djava.class.path=%s'%(clastr))
    for ip in ips.split(","):
        M6cls = jpype.JClass('onbon.bx06.Bx6GEnv')#导入类
        M6cls.initial(30000)#初始化
        bx6m = jpype.JClass('onbon.bx06.series.Bx6M')
        screencls = jpype.JClass('onbon.bx06.Bx6GScreenClient')
        screen = screencls('myscreen',bx6m())

        screen.connect(ip, port) #连接屏幕
        procls = jpype.JClass('onbon.bx06.file.ProgramBxFile')#导入节目类
        pf = procls('P000',screen.getProfile())
        bxtexta = jpype.JClass('onbon.bx06.area.TextCaptionBxArea')#导入文本区域类
        textp= jpype.JClass('onbon.bx06.area.page.TextBxPage')#导入文本类
        stylescls = jpype.JClass('onbon.bx06.utils.DisplayStyleFactory')#导入显示风格仓库类
        fontc = jpype.JClass('java.awt.Font')#导入字体类
        color = jpype.JClass('java.awt.Color')#
        area1 = part1content(screen,bxtexta,textp,fontc,stylescls,color,text,font_size,style_num,0,0,256,32)#第一个区域
        pf.addArea(area1)
        text2 = part2content(data_path,row_id)
        area2 = part1content(screen,bxtexta,textp,fontc,stylescls,color,text2,14,2,0,32,256,32)
        pf.addArea(area2)
        screen.writeProgram(pf)
        screen.disconnect()
def part2content(data_path,row_id):
    dataf = open(data_path,"r",encoding="utf-8")
    lines = dataf.read().split("~")
    ids = {}
    for line in lines:
        row = line.split(";")
        unit = ""
        value = ""
        title = ""
        dataId = row[0]
        tmp = {}
        try:
            unit = row[5]
            value = row[6]
            title = row[18]
        except:
            print(row)
            
        tmp["unit"] = unit
        tmp["value"] = value
        tmp["title"] = title
        ids[dataId] = tmp
    dict_d = ids.get(row_id)
    text = f"{dict_d.get('title')}：{dict_d.get('value')}  {dict_d.get('unit')}"
    dataf.close()
    return text
def part1content(screen,bxtexta,textp,fontc,stylescls,color,text,font_size,style_num,x,y,width,hight):
    area1 = bxtexta(x,y,width, hight,screen.getProfile())
    page = textp(text)#实例化一个文本类
    styles = stylescls.getStyle(style_num)
    fonti = fontc("宋体", fontc.PLAIN, font_size)#实例化一个字体类并传入参数
    page.setFont(fonti)
    page.setDisplayStyle(styles)
    area1.addPage(page)
    return area1
if __name__ == '__main__':
    try:
        config_path = "E://LedConfig//ledconfig.ini"
        ips,port,text,font_size,style_num,row_id = getconfig(config_path)
        data_path = "E://LedConfig//_CDDY_20220115110424.txt"
        
        sendmessage(ips,port,text,font_size,style_num,row_id,data_path)
        print("消息发布成功")
        os.system("pause")
    except Exception as e:
        print(str(e))
    