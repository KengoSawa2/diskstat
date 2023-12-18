import pathlib
import socket
import wmi
import pywintypes
import win32wnet
import win32netcon
import win32api
import copy
import datetime
import csv
import os
import re
import natsort

from pprint import pprint

import smtplib
import sys
from email import encoders, utils
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes

__version__ = "0.0.1"

ipv4addr = '192.168.'
TB_BASE = 1024 * 1024 * 1024 * 1024

#denied_Desc = ["Default share","Remote IPC"]
denied_Desc = ["Remote IPC"]
denied_Name = ["ADMIN$","print$","IPC$"]

exedirpath = os.path.dirname(os.path.abspath(__file__))
exedirname = os.path.split(exedirpath)[-1]
logdirname = "Log"

logdirpath = exedirpath + "\\" + logdirname
configpath = exedirpath + "\\" + "config.txt"


# E$ などのIPC内部共有の特殊表現マッチ用
ipcdmatchpattern = re.compile('[A-Z]\$')

csv_headers = ["ホスト名","共有ポイント名","ドライブレター","ボリューム名","ボリューム名2",
               "ディスク総容量","ディスク空き容量","使用率","IPアドレス","UNCフルパス","デバッグ情報"]

dict_keylist = ['hostname','type','sharename','drive','VolumeName','TotalDisk (TB)','RemainDisk (TB)','Usage (%)','IPAddress','UNCPath','OSVer','debuginfo']

# checkzumi = 171.118
# iplist = [45,118,171]

# 送信中継先の社内サーバやポート番号変えたかったら、ここを直接いじってね
smtpserver = '192.168.0.25'
smtpport =   25
# 宛先アドレス固定のナゴリ
# toaddr = "sawatsu@lespace.co.jp"
# fromに入るアドレス
fromaddr = "manage@lespace.co.jp"

# wmiで問い合わせするときに使うユーザ名とパスワード
username = "lespace"
password = "lespace"



resultlist = []
# 出力最終結果リスト value=結果辞書
#      'hostname'  : str(ホスト名. 解決失敗時はIPアドレス)
#      'type'      : str(共有名目印 [hidden] = IPCによるドライブ共有="E$"みたいな, [share] = それ以外
#      'sharename' : str(共有ポイント名 Cache_A_MUSIC)
#      'drive'     : str(A:\ とか) sambaは空
#      'IPAddress' : str IPアドレスの文字列
#      'VolumeName': str(180T-01とか) sambaは空
#      'TotalDisk (TB)' : int ディスク総容量(TiBだよ)
#      'RemainDisk (TB)': int ディスク空き容量(TiBだよ)
#      'Usage (%)'     : float ディスク総容量に対する空き容量の% 小数点1桁まで
#      'UNCPath'   : Windowsで絶対パスとして有効なUNCパス
#      'OSVer'     : WindowsのVer
#      'debuginfo' : str debug用エラー情報(なにもないと空)

def getdiskinfo(mylist,myaddr):

    # recdict = {}
    # recdict['IPAddress'] = myaddr

    try:
        hostname = socket.gethostbyaddr(myaddr)[0]
        # recdict['hostname'] = hostname
        print(hostname + "(" + myaddr + ")")

    except OSError as oserr:
        hostname = myaddr
        # recdict['hostname'] = hostname
        print("(" + hostname + ")" + str(oserr))

    # まずはWMIで投げてみるか
    if getshareinfo_wmi(myaddr,hostname,mylist):
        # WMIで取れてるならそれでいいや
        pass
    else:
        # 取れないならしゃあないからwnetAPIで取れる分だけ取ろっと
        getshareinfo_wnet(myaddr,hostname,mylist)

def getshareinfo_wmi(myaddr,hostname,mylist):

    recdict = {}
    recdict['hostname'] = hostname
    recdict['type'] = None
    recdict['sharename'] = None
    recdict['drive'] = None
    recdict['IPAddress'] = myaddr
    recdict['VolumeName'] = None
    recdict['TotalDisk (TB)'] = None
    recdict['RemainDisk (TB)'] = None
    recdict['Usage (%)'] =  None
    recdict['UNCPath'] = None
    recdict['OSVer'] = None
    recdict['debuginfo'] = None

    osstr = None

    try:
        c = wmi.WMI(myaddr, user=username, password=password)
        for oswmi in c.Win32_OperatingSystem():
            osstr = oswmi.Caption

        drdict = {} # key=ドライブレター文字, value=tuple(ディスク総容量,ディスク残容量,ボリューム名)

        # ドライブレターをkeyにした総容量、空き容量辞書を作るよ
        for logic in c.Win32_LogicalDisk():

            # print(logic)
            # 共有の設定や特殊な共有次第では中身Noneなことあるので、Noneのはスキップする
            if logic.Size and logic.FreeSpace and logic.VolumeName:
                drdict[logic.Name] = (int(logic.Size),int(logic.FreeSpace),logic.VolumeName)
            else:
                pass
                # print("None Logical?:")
                # print(logic.Name)

        #pprint(drdict)

        ipclist = []
        sharelist = []

        for share in c.Win32_Share():


            # print(share)
            # if (share.Description == "Default share" or share.Description == "Remote IPC" or share.Name == "ADMIN$") and share.Name != "C$":
            # if (share.Description in denied_Desc or share.Name in denied_Name) and share.Name != "C$":
            if (share.Description in denied_Desc or share.Name in denied_Name):
                # print("skip list {0}".format(share.Name))
                continue

            dretter = str(share.Path).split('\\')[0]

            # if dretter == 'None':
            #     dretter = share.Name[0] + ':'
            #
            # try:

            recdict['sharename'] = share.Name
            recdict['drive'] = dretter
            recdict['VolumeName'] = drdict[dretter][2]
            recdict['TotalDisk (TB)'] = round((drdict[dretter][0] / TB_BASE), 2)
            recdict['RemainDisk (TB)'] = round((drdict[dretter][1] / TB_BASE), 2)
            recdict['Usage (%)'] = round(100 - (((drdict[dretter][1] * 100) / drdict[dretter][0])) , 1)
            recdict['UNCPath'] = os.path.join('\\\\',hostname,share.Name)
            recdict['OSVer'] = osstr
            recdict['debuginfo'] = None
            if ipcdmatchpattern.match(share.Name):
                recdict['type'] = "[hidden]"
                ipclist.append(copy.deepcopy(recdict))
            else:
                recdict['type'] = "[share]"
                sharelist.append(copy.deepcopy(recdict))

            # except Exception as err:
            #     recdict['debuginfo'] = "wmi:" + str(err)
            #     mylist.append(copy.deepcopy(recdict))
            #     continue

        for ipcrec in ipclist:
            mylist.append(copy.deepcopy(ipcrec))

#       sharelisted = natsort.natsorted(sharelist,key='VolumeName')

        for sharerec in sharelist:
            mylist.append(copy.deepcopy(sharerec))

        # print(recdict['sharename'] + recdict['TotalDisk (TB)'])

    except Exception as err:
        recdict['debuginfo'] = "wmi:" + str(err)
        # mylist.append(copy.deepcopy(recdict))
        return False

    return True

def getshareinfo_wnet(myaddr,hostname,mylist):

    recdict = {}
    recdict['hostname'] = hostname
    recdict['type'] = None
    recdict['sharename'] = None
    recdict['drive'] = None
    recdict['IPAddress'] = myaddr
    recdict['VolumeName'] = None
    recdict['TotalDisk (TB)'] = None
    recdict['RemainDisk (TB)'] = None
    recdict['Usage (%)'] = None
    recdict['UNCPath'] = None
    recdict['OSVer'] = None
    recdict['debuginfo'] = None

    netresouce = win32wnet.NETRESOURCE()
    netresouce.dwScope = win32netcon.RESOURCE_GLOBALNET
    netresouce.lpProvider = 'Microsoft Windows Network'
    netresouce.dwType = win32netcon.RESOURCETYPE_DISK
    netresouce.dwDisplayType = win32netcon.RESOURCEDISPLAYTYPE_SHARE
    netresouce.lpRemoteName = '\\\\' + myaddr

    try:
        hnd = win32wnet.WNetOpenEnum(win32netcon.RESOURCE_GLOBALNET,
                                     win32netcon.RESOURCETYPE_DISK,
                                     0,
                                     netresouce)

    except win32wnet.error as wneterr:
        # print(netresouce.lpRemoteName + "error isn't exist?")
        recdict['debuginfo'] = str(wneterr)
        mylist.append(copy.deepcopy(recdict))
        return False

    if hnd:

        shname = ""

        # sharelist = []
        retlist = win32wnet.WNetEnumResource(hnd)
        for retres in retlist:
            try:
                shname = str(retres.lpRemoteName).split('\\')[-1]
                totals = win32api.GetDiskFreeSpaceEx(retres.lpRemoteName)
                totaltb = totals[1] / TB_BASE
                remaintb = totals[2] / TB_BASE
                # print("{0} {1:.2f}TiB/{2:.2f}TiB".format(myaddr + '\\' + shname, remaintb, totaltb))
                # print(retres.lpLocalName)
                recdict['IPAddress'] = myaddr
                recdict['sharename'] = shname
                recdict['drive'] = None
                recdict['VolumeName'] = None
                recdict['TotalDisk (TB)'] = round((totals[1] / TB_BASE), 2)
                recdict['RemainDisk (TB)'] = round((totals[2] / TB_BASE), 2)
                recdict['Usage (%)'] = round(100 - (((totals[2] * 100) / totals[1])), 1)
                recdict['UNCPath'] = os.path.join('\\\\',hostname,shname)
                recdict['OSVer'] = None
                recdict['debuginfo'] = None

                mylist.append(copy.deepcopy(recdict))

            except win32wnet.error as wneterr:
                recdict['debuginfo'] = str(wneterr)
                mylist.append(copy.deepcopy(recdict))
                # print(retres.lpRemoteName)
        win32wnet.WNetCloseEnum(hnd)

        # なんかnatsortできないけどlinuxだしどうでもいっか。。
        # sharelisted = natsort.natsorted(sharelist, key='sharename')

        # for sharerec in sharelisted:
        #     mylist.append(copy.deepcopy(sharerec))

    return True

def csvout(resultlist, sendmail = True, addrlist = None):

    starttime = datetime.datetime.now()
    filename = exedirpath + "\\" + logdirname + "\\" + starttime.strftime("%Y%m%d_%H%M%S") + ".csv"
    ofd = open(filename, 'wt', encoding='utf-8',newline='')
    outwriter = csv.DictWriter(ofd,fieldnames=dict_keylist,extrasaction='ignore',dialect='excel')

    outwriter.writeheader()
    for recdict in resultlist:
        outwriter.writerow(recdict)
    ofd.close()

    if sendmail:
        for addr in addrlist:

            subjctstr = 'DiskStat Report' + "(" + starttime.strftime("%Y/%m") + ") " + '[' + exedirname + ']'

            bodystr = ""
            bodystr += "batch name = " + exedirname + "\n"
            bodystr += "execute time = " + starttime.strftime("%Y/%m/%d %H:%M:%S") + "\n"
            bodystr += "version = " + __version__ + "\n"

            message = create_message(fromaddr, addr,subjctstr,bodystr,filename)
            send(fromaddr, addr, message)

def attachment(filename):
    fd = open(filename, 'rb')
    mimetype, mimeencoding = mimetypes.guess_type(filename)
    if mimeencoding or (mimetype is None):
        mimetype = 'application/octet-stream'
    maintype, subtype = mimetype.split('/')
    if maintype == 'text':
        retval = MIMEText(fd.read(),subtype)
    else:
        retval = MIMEBase(maintype, subtype)
        retval.set_payload(fd.read())
        encoders.encode_base64(retval)
    # gmailの添付データにC:\\とかのパス名含まれちゃうの微妙なので取っ払って送信する
    header_filename = os.path.basename(filename)
    retval.add_header('Content-Disposition', 'attachment', filename=header_filename)
    fd.close()
    return retval

def create_message(fromaddr, toaddr, subject, message, file):
    msg = MIMEMultipart()
    msg['To'] = toaddr
    msg['From'] = fromaddr

    msg['Subject'] = subject
    msg['Date'] = utils.formatdate(localtime=True)
    msg['Message-ID'] = utils.make_msgid()

    body = MIMEText(message, 'plain')
    msg.attach(body)

    msg.attach(attachment(file))
    return msg.as_string()

def send(fromaddr, toaddr, message):
    s = smtplib.SMTP(host=smtpserver, port=smtpport)
    s.sendmail(fromaddr, [toaddr], message)
    s.close()

if __name__ == '__main__':

    if not os.path.exists(logdirpath):
        os.makedirs(logdirpath,exist_ok=True)
    if not os.path.exists(configpath):
        usage = "↑1行目に送信したい相手のメールアドレスをcsvで入れて2行目以降に調査したいIPアドレスのクラスB.Cの0-255の任意の数字を一つずつ入れてね！この行は削除して保存してね！\n"
        ofd = open(configpath,"wt",encoding='utf-8')
        ofd.write(usage)
        ofd.close()
        print('config.txt is not exist. generate config.txt plz open and write config.')
        print('Press any key to exit()...')
        input('>> ')
        sys.exit(-1)

    fd = open(configpath,"r",encoding='utf-8')

    toaddrlist = fd.readline().split(",")

    iplist = fd.readlines()

    fd.close()

    #空ファイルじゃない場合は中身に従う
    if len(iplist):
        pass
    #空ファイルの場合は暗黙的に全IP調査する
    else:
        iplist.clear()
        iplist = range(0,255)

    for i in iplist:

        if isinstance(i,int):
            i = str(i)
        # 空行スキップ
        if not i.strip():
            continue
        wkipaddr = ipv4addr + i.strip()
        getdiskinfo(resultlist, wkipaddr)
        # serverごとの改行用に空の辞書入れとく
        lfdict = {}
        resultlist.append(lfdict)

    # print("ResultDict:")
    csvout(resultlist,True,toaddrlist)

    print("exit")


