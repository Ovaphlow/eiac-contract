# -*- coding=UTF-8 -*-

from __future__ import division

import os
import sys
import datetime

import xlwt

from sqlalchemy import create_engine, text
from xlrd import open_workbook
from xlutils.copy import copy

# 生产环境
cnx = {
    'user': 'cmtech',
    'password': 'cmtech.1123',
    'host': '125.211.221.215',
    'database': 'emsdatabase',
}

# 开发环境
cnx_dev = {
    'user': 'root',
    'password': '',
    'host': '127.0.0.1',
    'database': 'emsdatabase',
}

TEMPLATE_FILE_PATH = 'static\\contract'
TEMPLATE_FILE_NAME = 'template.xls'
SOFFICE_PATH = 'C:\\"Program Files (x86)"\\"LibreOffice 4"\\program\\soffice.exe'

db_engine = create_engine('mysql+pymysql://%s:%s@%s/%s' % \
    (cnx['user'], cnx['password'], cnx['host'],cnx['database']),
    pool_recycle=900, pool_size=3)


def generate_pdf(project_id):
    generate_contract_1(project_id)
    xls_path = os.path.join(os.getcwd(), TEMPLATE_FILE_PATH, '%s.xls' % \
        (project_id))
    # print(xls_path)
    pdf_path = os.path.join(os.getcwd(), TEMPLATE_FILE_PATH)
    # print(pdf_path)
    cmd_str = '%s --headless --invisible -convert-to pdf %s --outdir %s ' % \
        (SOFFICE_PATH, xls_path, pdf_path)
    # print(cmd_str)
    os.system(cmd_str)
    os.system('taskkill /f /im soffice.bin')
    os.system('taskkill /f /im soffice.exe')
    return True


# def generate_contract(project_id):
#     reload(sys)
#     sys.setdefaultencoding("utf-8")
#     sql = '''
#         select *
#         from tbs001_developprojectbasicinfo
#         where ID=:project_id
#     '''
#     param = {'project_id': project_id}
#     res = db_engine.execute(text(' '.join(sql.split())), param)
#     row_project = res.fetchone()
#     res.close()
#     sql = '''
#         select *
#         from tbs001_constructionunit
#         where ID=:cu_id
#     '''
#     param = {'cu_id': row_project.js}
#     res = db_engine.execute(text(' '.join(sql.split())), param)
#     row_cu = res.fetchone()
#     res.close()
#     sql = '''
#         select *
#         from tbs001_user
#         where UserName=:user_name
#     '''
#     param = {'user_name': row_project.pingshenxiangmufuzeren}
#     res = db_engine.execute(text(' '.join(sql.split())), param)
#     row_charge = res.fetchone()
#     res.close()
#     if not row_project:
#         print('invalid project id')
#         return False
#     rb = open_workbook(os.path.join(os.getcwd(), TEMPLATE_FILE_PATH,
#         TEMPLATE_FILE_NAME), formatting_info=True)
#     rs = rb.sheet_by_index(0)
#     wb = copy(rb)
#     ws = wb.get_sheet(0)
#     style_code_140 = (u'font: name 仿宋_GB2312, height 0x140, underline true;'
#         'align:wrap true, vertical center;')
#     style_code_118 = (u'font: name 仿宋_GB2312, height 0x118, underline true;'
#         'align:wrap true, vertical center;')
#     style_1 = xlwt.easyxf(style_code_140)
#     style_2 = xlwt.easyxf(style_code_118)
#     dt = datetime.datetime.now().strftime('%Y年%m月%d日')
#     dt1 = datetime.datetime.now() + datetime.timedelta(days=365)
#     dt1 = dt1.strftime(u'%Y年%m月%d日')
#     ws.write(4, 10, row_project.ProjectName.decode('utf-8'), style_1)
#     ws.write(5, 10, row_cu.Name, style_1)
#     ws.write(7, 10, dt.decode('utf-8'), style_1)
#     ws.write(9, 10,
#         dt.decode('utf-8') + u'至'.decode('utf-8') + dt1.decode('utf-8'),
#         style_1)
#     ws.write(23, 10, row_cu.Name.decode('utf-8'), style_2)
#     ws.write(24, 8, row_cu.Address.decode('utf-8'), style_2)
#     ws.write(25, 8, row_cu.ren.decode('utf-8'), style_2)
#     ws.write(26, 8, row_project.ProjectPerson.decode('utf-8'), style_2)
#     ws.write(27, 7, row_project.Mobil, style_2)
#     ws.write(28, 7, row_cu.Address.decode('utf-8'), style_2)
#     ws.write(29, 6, row_cu.Phone, style_2)
#     ws.write(29, 17, row_cu.chuanzhen, style_2)
#     ws.write(30, 7, row_cu.youxiang, style_2)
#     ws.write(34, 8, row_project.pingshenxiangmufuzeren.decode('utf-8'), style_2)
#     ws.write(35, 7, row_charge.GeRenShouJi, style_2)
#     ws.write(37, 6, row_charge.GuDingDianHua, style_2)
#     ws.write(37, 17, row_charge.ChuanZhen, style_2)
#     ws.write(38, 7, row_charge.DianZiXinXiang, style_2)
#     ws.write(39, 12, u'《'.decode('utf-8') + \
#         row_project.ProjectName.decode('utf-8') + u'》'.decode('utf-8'),
#         style_2)
#     ws.write(42, 8, u'对《'.decode('utf-8') + \
#         row_project.ProjectName.decode('utf-8') + u'》'.decode('utf-8'),
#         style_2)
#     ws.write(67, 12,
#         num2chn(float(row_project.hedinghetongjine) * 10000).decode('utf-8'),
#         style_2)
#     ws.write(113, 22, row_project.ProjectPerson.decode('utf-8'), style_2)
#     ws.write(114, 11,
#         row_project.pingshenxiangmufuzeren.decode('utf-8'), style_2)
#     for i in range(30):
#         ws.col(i).width = 850
#     ws.left_margin = 0.79
#     ws.right_margin = 0.79
#     ws.header_str = ''
#     ws.footer_str = ''
#     wb.save(os.path.join(os.getcwd(), TEMPLATE_FILE_PATH,
#         '%s.xls' % (project_id)))
#     return True


def generate_contract_1(project_id):
    reload(sys)
    sys.setdefaultencoding("utf-8")
    sql = '''
        select *
        from tbs001_developprojectbasicinfo
        where ID=:project_id
    '''
    param = {'project_id': project_id}
    res = db_engine.execute(text(' '.join(sql.split())), param)
    row_project = res.fetchone()
    res.close()
    sql = '''
        select *
        from tbs001_constructionunit
        where ID=:cu_id
    '''
    param = {'cu_id': row_project.js}
    res = db_engine.execute(text(' '.join(sql.split())), param)
    row_cu = res.fetchone()
    res.close()
    sql = '''
        select *
        from tbs001_user
        where UserName=:user_name
    '''
    param = {'user_name': row_project.pingshenxiangmufuzeren}
    res = db_engine.execute(text(' '.join(sql.split())), param)
    row_charge = res.fetchone()
    res.close()
    if not row_project:
        print('invalid project id')
        return False
    style_code_140 = (u'font: name 仿宋_GB2312, height 0x140, underline true;'
        'align:wrap true, vertical center;')
    style_code_118 = (u'font: name 仿宋_GB2312, height 0x118, underline true;'
        'align:wrap true, vertical center;')
    style_1 = xlwt.easyxf(style_code_140)
    style_2 = xlwt.easyxf(style_code_118)
    dt = datetime.datetime.now().strftime('%Y年%m月%d日')
    dt1 = datetime.datetime.now() + datetime.timedelta(days=365)
    dt1 = dt1.strftime(u'%Y年%m月%d日')
    rb = open_workbook(os.path.join(os.getcwd(), TEMPLATE_FILE_PATH,
        'template.xls'), formatting_info=True)
    # rs = rb.sheet_by_index(0)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ws.write(4, 10, row_project.ProjectName.decode('utf-8'), style_1)
    ws.write(5, 10, row_cu.Name.decode('utf-8'), style_1)
    ws.write(7, 10, dt.decode('utf-8'), style_1)
    ws.write(9, 10,
        dt.decode('utf-8') + u'至'.decode('utf-8') + dt1.decode('utf-8'),
        style_1)
    ws = wb.get_sheet(2)
    ws.write(3, 9, row_cu.Name.decode('utf-8'), style_2)
    ws.write(4, 7, row_cu.Address.decode('utf-8'), style_2)
    ws.write(5, 7, row_cu.ren.decode('utf-8'), style_2)
    ws.write(6, 7, row_project.ProjectPerson.decode('utf-8'), style_2)
    ws.write(7, 6, row_project.Mobil.decode('utf-8'), style_2)
    ws.write(8, 6, row_cu.Address.decode('utf-8'), style_2)
    ws.write(9, 5, row_cu.Phone.decode('utf-8'), style_2)
    ws.write(9, 15, row_cu.chuanzhen.decode('utf-8'), style_2)
    ws.write(10, 6, row_cu.youxiang.decode('utf-8'), style_2)
    ws.write(14, 7, row_project.pingshenxiangmufuzeren.decode('utf-8'), style_2)
    ws.write(15, 6, row_charge.GeRenShouJi.decode('utf-8'), style_2)
    ws.write(17, 5, row_charge.GuDingDianHua.decode('utf-8'), style_2)
    ws.write(17, 15, row_charge.ChuanZhen.decode('utf-8'), style_2)
    ws.write(18, 6, row_charge.DianZiXinXiang.decode('utf-8'), style_2)
    if len(row_project.ProjectName) > 19:
        ws.write(19, 10, u'《'.decode('utf-8') + \
            row_project.ProjectName.decode('utf-8')[:19], style_2)
        ws.write(20, 0, row_project.ProjectName.decode('utf-8')[19:] + \
            u'》'.decode('utf-8'), style_2)
    else:
        ws.write(19, 10, u'《'.decode('utf-8') + \
            row_project.ProjectName.decode('utf-8') + \
            u'》'.decode('utf-8'), style_2)
    if len(row_project.ProjectName) > 22:
        ws.write(23, 7, u'对《'.decode('utf-8') + \
            row_project.ProjectName.decode('utf-8')[:22], style_2)
        ws.write(24, 0, row_project.ProjectName.decode('utf-8')[22:] + \
            u'》'.decode('utf-8'), style_2)
    else:
        ws.write(23, 7, u'《'.decode('utf-8') + \
            row_project.ProjectName.decode('utf-8') + \
            u'》'.decode('utf-8'), style_2)
    ws.write(49, 10,
        num2chn(float(row_project.hedinghetongjine) * 10000).decode('utf-8'),
        style_2)
    ws.write(92, 20, row_project.ProjectPerson.decode('utf-8'), style_2)
    ws.write(93, 8,
        row_project.pingshenxiangmufuzeren.decode('utf-8'), style_2)
    for i in range(4):
        ws = wb.get_sheet(i)
        for i in range(30):
            ws.col(i).width = 850
        ws.left_margin = 0.79
        ws.right_margin = 0.79
        ws.header_str = ''
        ws.footer_str = ''
    wb.save(os.path.join(os.getcwd(), TEMPLATE_FILE_PATH,
        '%s.xls' % (project_id)))
    return True


def IIf(b, s1, s2):
    if b:
        return s1
    else:
        return s2


def num2chn(nin=None):
    cs = ('零','壹','贰','叁','肆','伍','陆','柒','捌','玖','◇','分','角','元','拾',
        '佰','仟','万','拾','佰','仟','亿','拾','佰','仟','万')
    st = st1 = ''
    s = '%0.2f' % (nin)
    sln =len(s)
    if sln > 15: return None
    fg = (nin<1)
    for i in range(0, sln-3):
    		ns = ord(s[sln-i-4]) - ord('0')
    		st = IIf((ns==0)and(fg or (i==8)or(i==4)or(i==0)), '', cs[ns]) + \
            IIf((ns==0)and((i<>8)and(i<>4)and(i<>0)or fg
            and(i==0)),'', cs[i+13]) + st
    		fg = (ns==0)
    fg = False
    for i in [1,2]:
    		ns = ord(s[sln-i]) - ord('0')
    		st1 = IIf((ns==0)and((i==1)or(i==2)and(fg or (nin<1))), '', cs[ns]) + \
            IIf((ns>0), cs[i+10], IIf((i==2) or fg, '', '整')) + st1
    		fg = (ns==0)
    st.replace('亿万','万')
    return IIf( nin==0, '零', st + st1)


if __name__ == '__main__':
    # print(calculate_fee(u'报告表', 260000))
    # print(num2chn(1.64 * 10000).decode('utf-8'))
    generate_pdf(1047)
