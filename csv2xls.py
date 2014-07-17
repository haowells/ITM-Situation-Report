#!/usr/bin/env python
# coding=utf-8

import re
import sys
from xlrd import open_workbook
from xlwt import Workbook, easyxf, Formula, Utils
#from xlutils.copy import copy
import logging
import logging.handlers
import sys
from operator import itemgetter
from collections import defaultdict
from itertools import chain

logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s %(name)-12s %(levelname)-8s %(message)s',
    datefmt='%m-%d %H:%M',
    filename='csv2xls.log',
    filemode='w')

console = logging.StreamHandler()
#console.setLevel(logging.INFO)
console.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)

log = logging.getLogger('main')

#rb = open_workbook('sitchndsc.xls',formatting_info=True, encoding_override='utf-8')
#rb = open_workbook('sitchndsc.xls')
#book = copy(rb)

with open(sys.argv[1], 'r') as csvf:
    lines = [x.strip() for x in csvf.readlines()]

csv_cont_len = len(lines[0][1:].split(','))

book = Workbook(encoding='utf-8')
ws_sit = book.add_sheet('sitlist')
ws_sitchn = book.add_sheet('sit_chn_dsc')

app_index = {}
##### {appname: [(host,ip), (host,ip)]
app_seq_index = {}
##### {appname: (str_index, end_index), appname1: (str_index, end_index)}
pcindex = {}
pc_pos_index = []
type_pos_index = []

def set_style(fontname='Arial', fontheight=200, colour='gray', rotation=None, wrap=False):
    return easyxf(
        'font: height %s, name %s;' % (fontheight, fontname) +
        'pattern: pattern solid, fore_colour %s;' % colour +
        'borders: left thin, right thin, top thin, bottom thin;'
        'align: vertical center, horizontal center, rotation %s, wrap %s;' % (rotation, wrap)
    )


def write_stdline(str_row=0, str_col=0, cross_col_no=2, csvlineno=0, height=None, style_head=None, style=None):
    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    cont_list = lines[csvlineno].split(',')
    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, cont_list[0], style_head)
    for (i, content) in enumerate(cont_list[1:]):
        ws_sit.write(str_row, i + cross_col_no + 1, content, style)


def write_levelline(str_row=None, str_col=0, cross_col_no=2, csvlineno=None, height=None, style_head=None,
                    style_level2=None, style_level3=None):
    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    cont_list = lines[csvlineno].split(',')
    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, cont_list[0], style_head)
    for (i, content) in enumerate(cont_list[1:]):
        if content == '2':
            ws_sit.write(str_row, i + cross_col_no + 1, content, style_level2)
        elif content == '3':
            ws_sit.write(str_row, i + cross_col_no + 1, content, style_level3)
        else:
            ws_sit.write(str_row, i + cross_col_no + 1, content, style_head)


def write_notiline(str_row=None, str_col=0, cross_col_no=2, csvlineno=None, height=None, style_head=None,
                   style_ok=None, style_ko=None):
    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    ws_sit.row(str_row + 1).height_mismatch = 1
    ws_sit.row(str_row + 1).height = height

    cont_list = lines[csvlineno].split(',')
    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, '邮件通知', style_head)
    ws_sit.write_merge(str_row + 1, str_row + 1, str_col, str_col + cross_col_no, '短信通知', style_head)
    for (i, content) in enumerate(cont_list[1:]):
        if re.match(r'邮件通知', content):
            ws_sit.write(str_row, i + cross_col_no + 1, 'Y', style_ok)
            ws_sit.write(str_row + 1, i + cross_col_no + 1, 'N', style_ko)
        elif re.match(r'邮件和短信通知', content):
            ws_sit.write(str_row, i + cross_col_no + 1, 'Y', style_ok)
            ws_sit.write(str_row + 1, i + cross_col_no + 1, 'Y', style_ok)
        else:
            ws_sit.write(str_row, i + cross_col_no + 1, 'N', style_ko)
            ws_sit.write(str_row + 1, i + cross_col_no + 1, 'N', style_ko)

def write_pcline(str_row=None, str_col=0, cross_col_no=2, csvlineno=None, height=None, style_head=None, style=None):
    global pc_pos_index
    pc_chndsc = {
        'UX': 'Unix操作系统',
        'UL': 'Unix日志',
        'PX': 'AIX LPAR',
        'UD': 'DB2数据库',
        'MQ': 'MQ消息队列',
        'LZ': 'Linux操作系统',
        'LO': '日志监控LO',
        'C1': 'DB2日志空间',
        'C2': '系统交换空间',
        'C3': '网络端口',
        'C4': 'SNA',
        'C5': 'UAEdge',
        'C6': 'Linux虚拟内存',
        'C8': 'Ping',
        'C9': 'GPFS',
        }

    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    pclist = lines[csvlineno].split(',')

    ### count merge count
    ### [{ux:(strindex, endindex)}, {ul:(strindex, endindex)} .. ]
    seq_pc = []
    tmp = ''
    for pc in pclist[1:]:
        if pc != tmp:
            tmp = pc
            pc_count = pclist.count(pc)
            stri = pclist[1:].index(pc)
            endi = stri + pc_count - 1
            pcindex[pc] = (stri, endi)
            seq_pc.append(pc)
    log.debug(pcindex)

    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, pclist[0], style_head)
    for pc in seq_pc:
        col_str = pcindex[pc][0] + cross_col_no + 1
        cross_len = pcindex[pc][1] - pcindex[pc][0]
        ws_sit.write_merge(str_row, str_row, col_str, col_str + cross_len,
                           "%s(%s)" % (pc_chndsc[pc], pc), style)

    pc_pos_index = [x[0] for x in pcindex.values()]
    pc_pos_index.sort()
    log.debug(pc_pos_index)

def write_typeline(str_row=None, str_col=0, cross_col_no=2, csvlineno=None, height=None, style_head=None, style_type=None):
    global type_pos_index
    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    cont_list = lines[csvlineno].split(',')
    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, cont_list[0], style_head)

    type_list = cont_list[1:]
    seq_typelist = list(set(type_list))
    seq_typelist.sort(key=type_list.index)

    type_index = defaultdict(list)
    seq_type_index = defaultdict(list)
    for k, va in [(v, i) for i, v in enumerate(type_list)]:
        seq_type_index[k].append(va)

    log.debug(seq_type_index)

    for type in seq_typelist:
        tv = seq_type_index[type]
        if len(tv) > 1:
            ntv = tv[:1]
            if tv[1] - tv[0] > 1:
                ntv.append(tv[0])
            for i in range(len(tv))[1:-1]:
                if tv[i-1] +1 == tv[i] == tv[i+1] -1:
                    continue
                elif tv[i-1] +1 == tv[i] <= tv[i+1] -1:
                    ntv.append(tv[i])
                elif tv[i-1] +1 <= tv[i] == tv[i+1] -1:
                    ntv.append(tv[i])
                elif tv[i-1] +1 <= tv[i] <= tv[i+1] -1:
                    ntv.extend([tv[i], tv[i]])
            if tv[-1] - tv[-2] > 1:
                ntv.extend([tv[-1], tv[-1]])
            else:
                ntv.append(tv[-1])

            str = ntv[::2]
            end = ntv[1::2]
            type_index[type] = zip(str, end)
        else:
            type_index[type] = [(tv[0],) * 2]


    log.debug(type_index)

    for type in seq_typelist:
        for (str, end) in type_index[type]:
            col_str = str + cross_col_no + 1
            cross_len = end - str
            if type == 'Non':
                type_cont = None
            else:
                type_cont = type
            ws_sit.write_merge(str_row, str_row, col_str, col_str + cross_len, type_cont, style_type )


    pos_in_type=[]
    for li in type_index.values():
        pos_in_type.extend(map(itemgetter(0), li))
    pos_in_type.sort()
    log.debug(pos_in_type)

    type_pos_index = [i for i in pos_in_type if i not in pc_pos_index]
    log.debug(type_pos_index)


def write_sitdsc(sitlineno=0, str_row=None, str_col=0, cross_col_no=2, csvlineno=None, height=None, style_head=None, style=None):
    ws_sit.row(str_row).height_mismatch = 1
    ws_sit.row(str_row).height = height
    cont_list = lines[csvlineno].split(',')
    ws_sit.write_merge(str_row, str_row, str_col, str_col + cross_col_no, cont_list[0], style_head)
    for (i, content) in enumerate(cont_list[1:]):
        if content == 'Non':
            cellname = Utils.rowcol_to_cell(sitlineno, i + cross_col_no + 1)
            ws_sit.write(str_row, i + cross_col_no + 1,
                         Formula("VLOOKUP(%s,sit_chn_dsc!$A$1:$B$1000,2,FALSE)" % cellname), style)
        else:
            ws_sit.write(str_row, i + cross_col_no + 1, content, style)


def write_app_ip_host(str_row=None, csvlineno=None, style_app=None, style_ip=None, style_host=None):

    cmdb = open_workbook('CMDB.xlsx')
    cmdbws = cmdb.sheet_by_index(0)

    hostlist = [l.split(",")[0] for l in lines[csvlineno:]]

    for (i, host) in enumerate(hostlist):
        ip = ''
        appname = ''
        for row_i in range(cmdbws.nrows):
            if host == cmdbws.cell(row_i, 0).value:
                ip = cmdbws.cell(row_i, 1).value
                appname = cmdbws.cell(row_i, 2).value
                break

        if app_index.has_key(appname):
            app_index[appname].append((ip, host))
        else:
            app_index[appname] = [(ip, host)]

    app_seq_list = app_index.keys()
    app_seq_list.sort()
    i = 0
    for app in app_seq_list:
        num_of_hosts = len(app_index[app])
        j = i + num_of_hosts - 1
        app_seq_index[app] = (i, j)
        i = j + 1

    log.debug(app_seq_index)

    for app in app_seq_list:
        ws_sit.write_merge(str_row + app_seq_index[app][0],
                           str_row + app_seq_index[app][0] + len(app_index[app]) - 1,
                           0, 0, app, style_app)
        for i in range(len(app_index[app])):
            ws_sit.write(str_row + app_seq_index[app][0] + i, 1, app_index[app][i][0], style_ip)
            ws_sit.write(str_row + app_seq_index[app][0] + i, 2, app_index[app][i][1], style_host)

    app_start_index = [v[0] for k, v in app_seq_index.items()]
    app_end_index = [v[1] for k, v in app_seq_index.items()]

    for i in range(max(app_end_index)):
        if i in app_start_index:
            ws_sit.row(i + str_row).level = 1
        else:
            ws_sit.row(i + str_row).level = 2


def write_content(str_row=None, str_col=3, csvlineno=None, style_ok=None, style_ko=None, style_not_exist=None):
    app_seq_list = app_index.keys()
    app_seq_list.sort()
    seq_hostlist = []
    for hosts in [[x[1]for x in app_index[app]] for app in app_seq_list]:
        seq_hostlist.extend(hosts)

    log.debug(seq_hostlist)

    for (i, host) in enumerate(seq_hostlist):
        for li in lines[csvlineno:]:
            if host == li.split(',')[0]:
                statuslist = li.split(',')[1:]
                for (j, status) in enumerate(statuslist):
                    if re.search(r'Non', status):
                        ws_sit.write(str_row + i, str_col + j, None, style_not_exist)
                    elif status == "Stopped":
                        ws_sit.write(str_row + i, str_col + j, status, style_ko)
                    elif status == "Started" or status == "Open" or status == "Closed":
                        ws_sit.write(str_row + i, str_col + j, None, style_ok)
                    elif re.search(r'|', status):
                        if re.search(r'Stopped', status):
                            status = status.replace("|", chr(10))
                            ws_sit.write(str_row + i, str_col + j, status, style_ko)
                        else:
                            status1 = status.split("|")
                            status2 = [s.split('->')[0] for s in status1]
                            status = "|".join(status2)
                            ws_sit.write(str_row + i, str_col + j, status, style_ok)
                    elif re.search(r'->', status):
                        if re.search(r'Stopped', status):
                            status = status.split('->')[0]
                            ws_sit.write(str_row + i, str_col + j, status, style_ko)
                        else:
                            status = status.split('->')[0]
                            ws_sit.write(str_row + i, str_col + j, status, style_ok)
                break

### main write excel
ws_sit.col(0).width = 6000
ws_sit.col(1).width = 8000
ws_sit.col(2).width = 4000

sit_head = set_style(colour='gray40')
sit_style = set_style(colour='olive_ega', rotation='-90')
write_stdline(str_row=0, csvlineno=0, height=1800, style_head=sit_head, style=sit_style)

level_head = set_style(colour='gray40')
level2_style = set_style(colour='yellow')
level3_style = set_style(colour='orange')
write_levelline(str_row=1, csvlineno=1, height=300, style_head=level_head, style_level2=level2_style, style_level3=level3_style)

noti_head = set_style(colour='gray40')
noti_ok = set_style(colour='bright_green')
noti_ko = set_style(colour='tan')
write_notiline(str_row=2, csvlineno=2, height=300, style_head=noti_head, style_ok=noti_ok, style_ko=noti_ko)

pc_head = set_style(colour='gray40')
pc_style = set_style(colour='gold')
write_pcline(str_row=4, csvlineno=4, height=600, style_head=pc_head, style=pc_style)

type_head = set_style(colour='gray40')
type_style = set_style(fontheight=180, colour='ocean_blue', wrap=True)
write_typeline(str_row=5, csvlineno=3, height=600, style_head=type_head, style_type=type_style)

sitdsc_head = set_style(colour='gray40')
sitdsc_style = set_style(fontheight=180, colour='pale_blue', wrap=True)
write_sitdsc(str_row=6, csvlineno=5, height=2400, style_head=sitdsc_head, style=sitdsc_style)

app_head = set_style(colour='lavender')
ip_head = set_style(colour='teal')
host_head = set_style(colour='lime')
write_app_ip_host(str_row=7, csvlineno=6, style_app=app_head, style_ip=ip_head, style_host=host_head)

ok_style = set_style(colour='sea_green')
ko_style = set_style(colour='light_yellow')
not_exist_style = set_style(colour='coral')
write_content(str_row=7, csvlineno=6, style_ok=ok_style, style_ko=ko_style, style_not_exist=not_exist_style)

### set outlines
def set_outline(l1=None, l2=None, cross_col_no=2):
    for i in range(csv_cont_len):
        if i in l1:
            ws_sit.col(i + cross_col_no + 1).level = 1
        elif i in l2:
            ws_sit.col(i + cross_col_no + 1).level = 2
        else:
            ws_sit.col(i + cross_col_no + 1).level = 3

set_outline(l1=pc_pos_index, l2=type_pos_index)

### freeze row and col
def set_freeze(vert=None, horz=None):
    ws_sit.panes_frozen = True
    ws_sit.remove_splits = True
    ws_sit.vert_split_pos = vert
    ws_sit.horz_split_pos = horz
    ws_sit.vert_split_first_visible = vert
    ws_sit.horz_split_first_visible = horz

set_freeze(vert=3, horz=7)

### set scaling
ws_sit.normal_magn = 70

xlsfname = sys.argv[1].split(".")[0] + ".xls"
book.save(xlsfname)