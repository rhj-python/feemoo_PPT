#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/13'

import requests
import os,time

url='https://jxup.fmapp.com/rc_upload3.php'

file_path='F:/资源/PPT资源/2019/9月/淡雅灰低三角形背景时尚简约扁平风年度工作总结汇报ppt模板.pptx'.replace('\u202a','')


headers = {
        'User-Agent' : 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:46.0) Gecko/20100101 Firefox/46.0',

}

feemoo_cookie={
    'Cookie':'acw_tc=781bad3115665436677108240e31d00f27d88832303d038336c4f5343fba2e; feemootoken=MDAwMDAwMDAwMJbfe2eakXHNtMuAro-Xs8-DdX3JmKqZnYPOq4eEqXatf7qKoJuKbaE.MDAwMDAwMDAwMJbfe6eYbH3NtMuAlp2ZxduRn4nNmLpvqoDSr6mTt3qnetOdn5l8eZqytouVh9C3loSsjJaEpq2djr3JqnrOoKx_uopsgo19lbPMoZOFqayXmImEyoXMf62BucysgKiYaHrUYXQ.fadeaffd89e45f8ff4bef3361e5e8aa53776a465116bf36b0ff90e258e4b0e37; view_stat=1; PHPSESSID=s20e3sek2mj885qiqbj30ob9o7'

}


'''
chunkNumber: 1
chunkSize: 2097152
currentChunkSize: 793296
totalSize: 793296
identifier: 793296-pptpptx
filename: 财务会计数据分析年终工作总结报告ppt模板.pptx
relativePath: 财务会计数据分析年终工作总结报告ppt模板.pptx
totalChunks: 1
ptype: 13
subType: 14
name: 财务会计数据分析年终工作总结报告ppt模板.pptx
time: 1568320213
uid: 2192888
token: 252c685058b5c1083ad9f3e1e7f0c8a0
pic: /Public/imgquality/attached/image/20190913/20190913043006_97525.jpg,/Public/imgquality/attached/image/20190913/20190913043006_17966.jpg,/Public/imgquality/attached/image/20190913/20190913043007_82977.jpg,/Public/imgquality/attached/image/20190913/20190913043007_53394.jpg,/Public/imgquality/attached/image/20190913/20190913043007_50868.jpg
file: (binary)
'''

'''
chunkNumber: 3
chunkSize: 2097152
currentChunkSize: 3998092
totalSize: 8192396
identifier: 8192396-pptpptx
filename: 黑板背景可爱卡通风小学教师说课通用ppt模板.pptx
relativePath: 黑板背景可爱卡通风小学教师说课通用ppt模板.pptx
totalChunks: 3
ptype: 13
subType: 18
name: 黑板背景可爱卡通风小学教师说课通用ppt模板.pptx
time: 1568320530
uid: 2192888
token: 252c685058b5c1083ad9f3e1e7f0c8a0
pic: /Public/imgquality/attached/image/20190913/20190913043513_51986.jpg,/Public/imgquality/attached/image/20190913/20190913043513_20639.jpg,/Public/imgquality/attached/image/20190913/20190913043513_50981.jpg,/Public/imgquality/attached/image/20190913/20190913043513_11670.jpg,/Public/imgquality/attached/image/20190913/20190913043515_72366.jpg
file: (binary)
'''

# ptype
PT={
    'PPT':13,
}

# subType
ST={
    "商务报告":14,
    "企业宣传":15,
    "求职简历":16,
    "公关策划":16,
    "教学课件":16,

}


def send_ppt_req(file_path):
    now_time=int(time.time())
    totalSize=os.path.getsize(file_path)
    _,file_name=os.path.split(file_path)

    data={
        'chunkNumber': 1,
        'chunkSize': 40971522,
        'currentChunkSize': totalSize,
        'totalSize':totalSize,
        'identifier': '{}-pptpptx'.format(totalSize),
        'totalSize':os.path.getsize(file_path),
        'filename': file_name,
        'relativePath': file_name,
        'totalChunks': 1,
        # ptype: 13
        # subType: 18
        # name: 黑板背景可爱卡通风小学教师说课通用ppt模板.pptx
        # time: 1568320530
        # uid: 2192888
        # token: 252c685058b5c1083ad9f3e1e7f0c8a0
        # pic: /Public/imgquality/attached/image/20190913/20190913043513_51986.jpg,/Public/imgquality/attached/image/20190913/20190913043513_20639.jpg,/Public/imgquality/attached/image/20190913/20190913043513_50981.jpg,/Public/imgquality/attached/image/20190913/20190913043513_11670.jpg,/Public/imgquality/attached/image/20190913/20190913043515_72366.jpg
        # file: (binary)

    }

    print(totalSize,file_name)

    # requests.post(headers=headers,cookies=feemoo_cookie)


if __name__=='__main__':
    send_ppt_req(file_path)


# 测试中功能有待完善
