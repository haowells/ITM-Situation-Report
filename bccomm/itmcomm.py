#!/usr/bin/env python
__author__ = 'ibm_linh'
import logging
import re
import os
import sys


def login(temsip, user, password):
    log = logging.getLogger('login')
    #temslogin = 'tacmd login -s ' + ' 182.248.56.60 ' + ' -u ' + ' sysadmin ' + ' -p ' + 'bcdctiv1'
    temslogin = 'tacmd login -s ' + temsip + ' -u ' + user + ' -p ' + password
    log.debug(temslogin)
    #tepslogin = 'tacmd tepslogin -s ' + cfg.get('teps','localhost') + ' -u ' + cfg.get('user') + ' -p ' + cfg.get('pw')
    pat = re.compile('.* logged into server on http.*')
    try:
        log.info('start login to tems...')
        if pat.match([l for l in os.popen(temslogin)][3]):
            log.info('tems login done')
        else:
            log.error('Can not log in tems while running \"tacmd login\" ! Program exiting...\n')
            sys.exit()
    except IndexError:
        log.error('Can not log in tems while running \"tacmd login\" ! Program exiting...\n')
        sys.exit()

#def listsystems(pc):
#    log = logging.getLogger('listsystem')
#    cmd = 'tacmd listsystems -t ' + pc
#    li = [l for l in os.popen(cmd)]
#    msnl = []
#    for l in li[1:]:
#        (msn, pc, ver, status) = re.split(r'\s+', l.strip())
#        if status == 'Y' or status == 'N':
#            msnl.append(msn)
#    log.debug(msnl)
#    return msnl


def viewnode(host, pcfilter=None):
    log = logging.getLogger('viewnode')
    managed_system_dict = {}

    managed_os_ux = host + ':KUX'
    managed_os_lz = host + ':LZ'
    cmd_ux = 'tacmd viewnode -n ' + managed_os_ux
    cmd_lz = 'tacmd viewnode -n ' + managed_os_lz
    log.debug(cmd_ux)
    log.debug(cmd_lz)
    li_ux = [l for l in os.popen(cmd_ux)]
    li = li_ux
    #log.debug(li)
    if re.match(r'KUICVN002E', li_ux[1]):
        log.info(''.join(li_ux))
        li_lz = [l for l in os.popen(cmd_lz)]
        li = li_lz
        if re.match(r'KUICVN002E', li_lz[1]):
            log.info(''.join(li_lz))
            return None
    for l in li[1:]:
        log.info(l)
        (managed_system, pc, ver, kax, kgl) = re.split(r'\s+', l.strip())
        if pc in pcfilter:
            managed_system_dict[managed_system] = pc
    log.debug(managed_system_dict)
    return managed_system_dict


def listsit(managed_system, pc, host):
    log = logging.getLogger('listsit')
    cmd = 'tacmd listsit -l -d "#" -m ' + managed_system
    log.debug(cmd)
    li = [l for l in os.popen(cmd)]

    if len(li) == 0:
        log.info("not found situatoin for " + managed_system)
        return None

    single_key = ("Type", "Name", "Status", "FullName", "InstName")
    sit_number = len(li)/5
    sit_list = []
    for i in range(sit_number):
        try:
            single_value = map(lambda x: x.split('#')[1].strip(), [l for l in li[i*5:i*5+4]])
            if pc in ['UD', 'MQ', 'LO']:
                single_value.append(managed_system.split(":")[0])
            else:
                single_value.append(host)
        except IndexError:
            log.info("not found situation for " + managed_system)
            return None
        single_dict = dict(zip(single_key, single_value))
        sit_list.append(single_dict)
    return sit_list

def getipfile():
    log = logging.getLogger('getipfile')
    cmd = 'scp ITMHMS1:/opt/itm6/UA/C08/ipfile .'
    log.debug(cmd)
    li = [l for l in os.popen(cmd)]
    log.debug(li)


