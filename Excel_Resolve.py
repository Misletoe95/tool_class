import base64
import re
import xml.dom.minidom as xmldom
import os
import zipfile
import shutil
import xlrd


# �ж��Ƿ����ļ����ж��ļ��Ƿ����
def isfile_exist(file_path):
    if not os.path.isfile(file_path):
        print("It's not a file or no such file exist ! %s" % file_path)
        return False
    else:
        return True


# ���Ʋ��޸�ָ��Ŀ¼�µ��ļ�����������excel��׺���޸�Ϊ.zip
def copy_change_file_name(file_path, new_type='.zip'):
    if not isfile_exist(file_path):
        return ''

    extend = os.path.splitext(file_path)[1]  # ��ȡ�ļ���չ��
    if extend != '.xlsx' and extend != '.xls':
        print("It's not a excel file! %s" % file_path)
        return False

    file_name = os.path.basename(file_path)  # ��ȡ�ļ���
    new_name = str(file_name.split('.')[0]) + new_type  # �µ��ļ���������Ϊ��xxx.zip

    dir_path = os.path.dirname(file_path)  # ��ȡ�ļ�����Ŀ¼
    new_path = os.path.join(dir_path, new_name)  # �µ��ļ�·��
    if os.path.exists(new_path):
        os.remove(new_path)
    shutil.copyfile(file_path, new_path)
    return new_path  # �����µ��ļ�·����ѹ����


# ��ѹ�ļ�
def unzip_filez(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False

    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # ��ȡ�ļ���
    zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # ��ȡ�ļ�����Ŀ¼
    print("zipdir:{zipdir}")
    for files in file_zip.namelist():
#         unzipdir = os.path.join(zipfile_path, zipdir)
#         if not os.path.exists(unzipdir):
#             os.makedirs(unzipdir)

#         if not unzipdir.exists:
#             os.makdir(unzipdir)
        file_zip.extract(files, os.path.join(zipfile_path, zipdir))  # ��ѹ��ָ���ļ�Ŀ¼

    file_zip.close()
    return True


# ��ȡ��ѹ����ļ��У���ӡͼƬ·��
def read_img(zipfile_path):
    img_dict = dict()
    if not isfile_exist(zipfile_path):
        return False

    dir_path = os.path.dirname(zipfile_path)  # ��ȡ�ļ�����Ŀ¼
    file_name = os.path.basename(zipfile_path)  # ��ȡ�ļ���
    pic_dir = 'xl' + os.sep + 'media'  # excel���ѹ�������ٽ�ѹ��ͼƬ��mediaĿ¼
    pic_path = os.path.join(dir_path, str(file_name.split('.')[0]), pic_dir)
    if not os.path.exists(pic_path):
        img_dict[1] =dict(img_index=1,img_path=None,img_base64=None)
        return img_dict
    file_list = os.listdir(pic_path)
    for file in file_list:
        filepath = os.path.join(pic_path, file)
        print(filepath)
        img_index = int(re.findall(r'image(\d+)\.', filepath)[0])
        img_base64 = get_img_base64(img_path=filepath)
        img_dict[img_index] = dict(img_index=img_index, img_path=filepath, img_base64=img_base64)
    return img_dict


# ��ȡimg_base64
def get_img_base64(img_path):
    if not isfile_exist(img_path):
        return ""
    with open(img_path, 'rb') as f:
        base64_data = base64.b64encode(f.read())
        s = base64_data.decode()
        return 'data:image/jpeg;base64,%s' % s


def get_img_pos_info(zip_file_path, img_dict, img_feature):
    """����xml �ļ�����ȡͼƬ��excel����е�����λ����Ϣ"""
    os.path.dirname(zip_file_path)
    dir_path = os.path.dirname(zip_file_path)  # ��ȡ�ļ�����Ŀ¼
    file_name = os.path.basename(zip_file_path)  # ��ȡ�ļ���
    xml_dir = 'xl' + os.sep + 'drawings' + os.sep + 'drawing1.xml'
    xml_path = os.path.join(dir_path, str(file_name.split('.')[0]), xml_dir)
    image_info_dict ={(1,10):None}
    if not os.path.exists(xml_path):
        return image_info_dict
    image_info_dict = parse_xml(xml_path, img_dict, img_feature=img_feature)  # ����xml �ļ��� ����ͼƬ����λ����Ϣ
    return image_info_dict


# ��������ѹ��ȡͼƬλ�ã���ͼƬ���������Ϣ
def get_img_info(excel_file_path, img_feature):
    if img_feature not in ["img_index", "img_path", "img_base64"]:
        raise Exception('ͼƬ���ز������� ["img_index", "img_path", "img_base64"]')
    zip_file_path = copy_change_file_name(excel_file_path)
    if zip_file_path != '':
        if unzip_file(zip_file_path):
            img_dict = read_img(zip_file_path)  # ��ȡͼƬ�������ֵ䣬ͼƬ img_index�� img_index�� img_path�� img_base64
            image_info_dict = get_img_pos_info(zip_file_path, img_dict, img_feature)
            return image_info_dict
    return dict()


# ����xml�ļ�����ȡ��ӦͼƬλ��
def parse_xml(file_name, img_dict, img_feature='img_path'):
    # �õ��ĵ�����
    image_info = dict()
    dom_obj = xmldom.parse(file_name)
    # �õ�Ԫ�ض���
    element = dom_obj.documentElement

    def _f(subElementObj):
        for anchor in subElementObj:
            xdr_from = anchor.getElementsByTagName('xdr:from')[0]
            col = xdr_from.childNodes[0].firstChild.data  # ��ȡ��ǩ�������
            row = xdr_from.childNodes[2].firstChild.data
            embed = \
            anchor.getElementsByTagName('xdr:pic')[0].getElementsByTagName('xdr:blipFill')[0].getElementsByTagName(
                'a:blip')[0].getAttribute('r:embed')  # ��ȡ����
            image_info[(int(row), int(col))] = img_dict.get(int(embed.replace('rId', '')), {}).get(img_feature)

    sub_twoCellAnchor = element.getElementsByTagName("xdr:twoCellAnchor")
    sub_oneCellAnchor = element.getElementsByTagName("xdr:oneCellAnchor")
    _f(sub_twoCellAnchor)
    _f(sub_oneCellAnchor)
    return image_info
def unzip_file(zipfile_path):
    if not isfile_exist(zipfile_path):
        return False

    if os.path.splitext(zipfile_path)[1] != '.zip':
        print("It's not a zip file! %s" % zipfile_path)
        return False

    file_zip = zipfile.ZipFile(zipfile_path, 'r')
    file_name = os.path.basename(zipfile_path)  # ��ȡ�ļ���
    zipdir = os.path.join(os.path.dirname(zipfile_path), str(file_name.split('.')[0]))  # ��ȡ�ļ�����Ŀ¼
    if not os.path.exists(zipdir):
        os.makedirs(zipdir)
    for files in file_zip.namelist():
        file_zip.extract(files, zipdir)  # ��ѹ��ָ���ļ�Ŀ¼

    file_zip.close()
    return True

def read_excel_info(file_path, img_col_index, img_feature='img_path'):
    """
    ��ȡ����ͼƬ��excel���ݣ� �������б�
    :param file_path:
    :param img_col_index: ͼƬ�����У�list
    :param img_feature: ָ��ͼƬ������ʽ img_index", "img_path", "img_base64"
    :return:
    """
    img_info_dict = get_img_info(file_path, img_feature)
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)
    head = dict()
    for i, v in enumerate(sheet.row(0)):
        head[i] = v.value
    info_list = []
    for row_num in range(sheet.nrows):
        d = dict()
        for col_num in range(sheet.ncols):
            if row_num ==0:
                continue
            #if 'empty:' in sheet.cell(row_num, col_num).__str__():
            cell_value = sheet.cell_value(row_num, col_num)
            if not cell_value:
                if col_num in img_col_index: #ͼƬ���ڵ���
                    d[head[col_num]] = img_info_dict.get((row_num, col_num))
                else:
                    d[head[col_num]] = sheet.cell(row_num, col_num).value
            else:
                d[head[col_num]] = sheet.cell(row_num, col_num).value
        if d:
            info_list.append(d)
    return info_list


if __name__ == '__main__':
    i = read_excel_info(r'147512880.xlsx', img_col_index=[10])
    from pprint import pprint