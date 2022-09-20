#!/usr/bin/env python
# -*- coding: utf-8 -*-

__mtime__ = '2019/9/11'

import random
import constant as cst

YEAR='2021'
MONTH='4月'
DATE='2021-4-19'

save_path='g:/资源/PPT资源/{}/{}/{}'.format(YEAR,MONTH,DATE)
pic_path='e:/图片保存/{}_PPT'.format(DATE)


save_path_word='g:/资源/WORD资源/{}/{}/{}'.format(YEAR,MONTH,DATE)
word_test_path='g:/资源/WORD资源/{}/{}/测试'.format(YEAR,MONTH)
ppt_test_path='g:/资源/PPT资源/{}/{}/测试'.format(YEAR,MONTH)
ppt_upload_path='g:/资源/PPT资源/{}/{}/upload/'.format(YEAR,MONTH)
pic_path_word='e:/图片保存/{}_WORD'.format(DATE)

can_use_path='g:/资源/PPT资源/2019/10月/可用9'

ppt_word_filter=cst.PPT_WORD_FILTER
ppt_word_filter=sorted(ppt_word_filter,key=lambda x:len(x),reverse=True)




replace_small='X'
replace_medium='添加文本'
replace_large='点击输入简要文字内容文字内容需概括精炼不用多余的文字修饰言简意赅的说明分项内容'



# JOB_TYPE=['求职','简历','个人简介','职业规划']
# PROMOT_TYPE=['企业','宣传','简介','员工风采']
# PLAN_TYPE=['节','计划书','策划','颁奖','纪念册','活动','公祭']
# TEACH_TYPE=['答辩','教师','教育','教学','党','培训']
# BUSINESS_TYPE=['商务','工作','报告','总结']

ppt_type=dict(
HOMEWORK_TYPE_1=(['作业','寒假','暑假'],('作业汇报','教育')),
JOB_TYPE_1=(['求职','简历','个人简介','职业规划'],('求职招聘','教育')),
COMMUNITY_TYPE_1=(['社团','活动'],('社团活动','教育')),
REPLY_TYPE_1=(['答辩','毕业','开题'],('毕业答辩','教育')),
TEACH_TYPE_1=(['幼儿','教师','教育','教学'],('教学课件','教育')),
TALK_TYPE_1=(['班会','家长会'],('主题班会','教育')),
COMMEMORATE_TYPE_1=(['毕业纪念','毕业相册'],('毕业纪念','教育')),

REPORT_TYPE_2=(['汇报','报告','述职','旅游','旅行','游记'],('工作汇报','商业')),
PRODUCT_TYPE_2=(['发布'],('产品发布','商业')),
PLAN_TYPE_2=(['计划书','策划','颁奖','计划','相册'],('计划总结','商业')),
TEAM_TYPE_2=(['团建','团队建设'],('企业团建','商业')),
LEARN_TYPE_2=(['培训','入职'],('入职培训','商业')),
FESTIVAL_TYPE_2=(['节'],('节日庆典','商业')),
PROMOT_TYPE_2=(['宣传','画册'],('企业宣传','商业')),

LEARNING_TYPE_3=(['入党','入党培训'],('入党培训','政府')),
WORK_TYPE_3=(['党建'],('党建工作','政府')),
LESSION_TYPE_3=(['党','党课'],('主题党课','政府')),
FESTIVAL_TYPE_3=(['国庆','建军','阅兵'],('节日庆典','政府')),
PROMOT_TYPE_3=(['党史'],('党史宣传','政府')),
)


ppt_type_overainbow=dict(
    HOMEWORK_TYPE=(['作业汇报'],('作业汇报',0)),
    JOB_TYPE=(['求职招聘'],('求职招聘',1)),
    COMMUNITY_TYPE=(['社团活动'],('社团活动',2)),
    REPORT_TYPE=(['毕业答辩'],('毕业答辩',3)),
    TEACH_TYPE=(['教学课件'],('教学课件',4)),
    TALK_TYPE=(['主题班会'],('主题班会',5)),
    COMMEMORATE_TYPE=(['毕业纪念'],('毕业纪念',65)),
    BUSINESS_REPORT_TYPE=(['工作汇报'],('工作汇报',7)),
    PRODUCT_TYPE=(['产品发布'],('产品发布',8)),
    PLAN_TYPE=(['计划总结'],('计划总结',9)),
    TEAM_TYPE=(['企业团建'],('企业团建',10)),
    LEARN_TYPE=(['入职培训'],('入职培训',11)),
    FESTIVAL_TYPE=(['节日庆典'],('节日庆典',12)),
    PROMOT_TYPE=(['企业宣传'],('企业宣传',13)),
    PARTY_TYPE=(['入党培训','党建工作','主题党课','党史宣传'],('党政军警',14)),
)

headers = {
        'User-Agent' : 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:46.0) Gecko/20100101 Firefox/{}.0'.format(random.randint(100,6500)),

}

common_headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.97 Safari/537.36'
}


fm_headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/77.0.3865.90 Safari/537.36',
    'Cookie':'showstl=yes; feemootoken=MDAwMDAwMDAwMJbfe2eakXHNtMuAro-Xs8-DdX3JmKqZnYPOq4eEqXatf7qKoJuKbaE.MDAwMDAwMDAwMJbfe6eYbH3NtMuAlp2ZxduRn4nNmLpvqoDSr6mTt3qnetOdn5l8eZqyto_ciLqvmoScnpaEzK2djr3JqnrOoKx_0H6vhI2TlLS2oZWFqayXmImEyoXMf62BucysgKiYaHrUYXQ.8fa46f107309472e2bda80ba88cb102384d7dc7022e36eac5e661a85a6a806dd; PHPSESSID=rrcgitgr50i32tqet3e8p0u2cu'
}



first_day='UM_distinctid=16d4d440f5cb0a-027b76fc7be451-67e1b3f-1fa400-16d4d440f5de1f; advanced-frontend=f5onfhro92tqoedd49nqd938j5; _identity-frontend=dfac7f9f7007d14b7dcdc9c6925e4c3e20bcd85cc91541b063b4a4f67d3c8516a%3A2%3A%7Bi%3A0%3Bs%3A18%3A%22_identity-frontend%22%3Bi%3A1%3Bs%3A50%3A%22%5B1452813%2C%224pQgksysgbY-4ZZsEyqZSxvlKwoc_o9p%22%2C86400%5D%22%3B%7D; CNZZDATA1272242338=698697745-1569189993-https%253A%252F%252Fppt.lq-sf.com%252F%7C1569189993; CNZZDATA1271446102=86754677-1568958692-%7C1569194217'



# 10.3
'UM_distinctid=16d86bfe892210-0e7b13040a8f0c-67e1b3f-1fa400-16d86bfe8933ba; _identity-frontend=dfac7f9f7007d14b7dcdc9c6925e4c3e20bcd85cc91541b063b4a4f67d3c8516a%3A2%3A%7Bi%3A0%3Bs%3A18%3A%22_identity-frontend%22%3Bi%3A1%3Bs%3A50%3A%22%5B1452813%2C%224pQgksysgbY-4ZZsEyqZSxvlKwoc_o9p%22%2C86400%5D%22%3B%7D; advanced-frontend=bkkqh33ocrkd7b2uv8d3divv46; CNZZDATA1271446102=1204787383-1569919726-%7C1570057952'

# jj_word
# https://www.jjppt.com/vip/download?id=259&type=0

jp_cookie={
    'Cookie':'UM_distinctid=16d86bfe892210-0e7b13040a8f0c-67e1b3f-1fa400-16d86bfe8933ba; _identity-frontend=dfac7f9f7007d14b7dcdc9c6925e4c3e20bcd85cc91541b063b4a4f67d3c8516a%3A2%3A%7Bi%3A0%3Bs%3A18%3A%22_identity-frontend%22%3Bi%3A1%3Bs%3A50%3A%22%5B1452813%2C%224pQgksysgbY-4ZZsEyqZSxvlKwoc_o9p%22%2C86400%5D%22%3B%7D; advanced-frontend=bkkqh33ocrkd7b2uv8d3divv46; CNZZDATA1271446102=1204787383-1569919726-%7C1570057952'

}

# cookie1
'bt_guid=%2222ae2423cd69152ff739d5d817ae5b54%22; adIssem=0; UM_distinctid=16d8f72c8aec6a-008f45aae18026-67e1b3f-1fa400-16d8f72c8afba1; 588ku_login_refer_url=%22https%3A%5C%2F%5C%2F588ku.com%5C%2F%22; FIRSTVISITED=1570069731.081; _ga=GA1.2.1372574000.1570069736; auth_id=%2228239214%7C%5Cu98ce%7C1570933765%7C9700c194e8ffb0e66756101b874999b3%22; success_target_path=%22%5C%2F%5C%2F588ku.com%5C%2F%22; sns=%7B%22token%22%3A%7B%22access_token%22%3A%220AF54E20290459493522BD19A098ED23%22%2C%22expires_in%22%3A%227776000%22%2C%22refresh_token%22%3A%22B3D38826973DC630A69451690CAF8955%22%2C%22openid%22%3A%229BD2E0779DCDFDFB6E4480B197AE9342%22%7D%2C%22type%22%3A%22qq%22%7D; ISREQUEST=1; source_url=588ku.com; _k_iprec_1=58.35.179.8; _k_iprec_wd_1=%7C-; _k_ut_v1=%7B%22ip%22%3A%2258.35.179.8%22%2C%22d%22%3A%222019-10-06%22%2C%22h%22%3A%22%22%2C%22hd%22%3A%22%22%2C%22sem%22%3A%22%22%2C%22host%22%3A%22588ku.com%22%2C%22bm%22%3A2%2C%22kw%22%3A%22%22%7D; _gid=GA1.2.2060465024.1570360084; host=588ku.com; back_search_words=%5B%22keji%22%5D; vipGather=1; backConditionFilterField=8_2; mid_autumn_show_pv=1; mid_autumn_list_pv=8; stat_str=%u5145%u503C%3A%u4FA7%u8FB9%u680F%u6309%u94AE%3A%u5168%u7AD9; WEBPARAMS=is_pay=1; PHPSESSID=cmrduq74cobf61m9euu6jef710; CNZZDATA1277846593=279948192-1570066301-https%253A%252F%252Fwww.google.com%252F%7C1570372006; Hm_lvt_8226f7457e3273fa68c31fdc4ebf62ff=1570069737,1570084937,1570360084,1570373384; Hm_lvt_3e90322e8debb1d06c9c463f41ea984b=1570069737,1570084937,1570360084,1570373384; 588KUSSID=10udmmucicmeqnd9eibfgk40r2; location=4; e9e446b10c4934194c555fcbd6eb6994=%220a35a8a8059a99fb4f9570dd8c0a57c7%22; write_phone_pv=80; referer=%22%5C%2F%5C%2F588ku.com%5C%2Foffice%5C%2F11493.html%22; Hm_lpvt_8226f7457e3273fa68c31fdc4ebf62ff=1570377385; Hm_lpvt_3e90322e8debb1d06c9c463f41ea984b=1570377385'



'千库之前'
'bt_guid=%2222ae2423cd69152ff739d5d817ae5b54%22; adIssem=0; UM_distinctid=16d8f72c8aec6a-008f45aae18026-67e1b3f-1fa400-16d8f72c8afba1; FIRSTVISITED=1570069731.081; _ga=GA1.2.1372574000.1570069736; ISREQUEST=1; WEBPARAMS=is_pay=1; back_search_words=%5B%22dangjian%22%2C%22keji%22%5D; illus_search_words=%5B%22zhongyangjie%22%5D; _k_iprec_wd_1=%7C-; _k_iprec_1=58.35.4.235; _gid=GA1.2.635038395.1571824100; _k_ut_v1=%7B%22ip%22%3A%2258.35.4.235%22%2C%22d%22%3A%222019-10-24%22%2C%22h%22%3A%22%22%2C%22hd%22%3A%22%22%2C%22sem%22%3A%22%22%2C%22host%22%3A%22www.baidu.com%22%2C%22bm%22%3A16642%2C%22kw%22%3A%22%22%7D; PHPSESSID=aocfq77d8b44grc7l9ig4ru024; 464a92bec59e6fb530b6d1b7d42e9f82=%220a35a8a8059a99fb4f9570dd8c0a57c7%22; source_url=588ku.com; no_login_pv=9; location=4; referer=%22%5C%2F%5C%2F588ku.com%5C%2Fresume%5C%2F0-default-0-0-0-0-0-0-1%5C%2F%22; CNZZDATA1277846593=279948192-1570066301-https%253A%252F%252Fwww.google.com%252F%7C1571928452; 588KUSSID=mch8netokhqt0ndjem5cmpm583; _gat_gtag_UA_139867171_1=1; Hm_lvt_3e90322e8debb1d06c9c463f41ea984b=1571824100,1571846298,1571870632,1571931002; Hm_lvt_8226f7457e3273fa68c31fdc4ebf62ff=1571824100,1571846298,1571870632,1571931002; host=588ku.com; 588ku_login_refer_url=%22https%3A%5C%2F%5C%2F588ku.com%5C%2Fresume%5C%2F0-default-0-0-0-0-0-0-1%5C%2F%22; phoneold28239214=1; temp_login_uid=28239214; temp_login_avator=%22http%3A%5C%2F%5C%2Fthirdqq.qlogo.cn%5C%2Fg%3Fb%3Doidb%26k%3D36R8B8uLaTGxyJ52eficV7Q%26s%3D100%26t%3D1555697451%22; temp_login_flag2=1; auth_id=%2228239214%7C%5Cu98ce%7C1572795020%7C113ff6745d5d149769718a1f23c574f3%22; success_target_path=%22%5C%2F%5C%2F588ku.com%5C%2Fresume%5C%2F0-default-0-0-0-0-0-0-1%5C%2F%22; sns=%7B%22token%22%3A%7B%22access_token%22%3A%220AF54E20290459493522BD19A098ED23%22%2C%22expires_in%22%3A%227776000%22%2C%22refresh_token%22%3A%22B3D38826973DC630A69451690CAF8955%22%2C%22openid%22%3A%229BD2E0779DCDFDFB6E4480B197AE9342%22%7D%2C%22type%22%3A%22qq%22%7D; write_phone_pv=17; Hm_lpvt_8226f7457e3273fa68c31fdc4ebf62ff=1571931021; Hm_lpvt_3e90322e8debb1d06c9c463f41ea984b=1571931021'


qk_cookie={
    'Cookie':'bt_guid=%221330aa01c5d614fa3e88b685f73e4c4f%22; adIssem=0; UM_distinctid=178c93a2203692-05531b8e01b29-c3f3568-1fa400-178c93a2204dba; FIRSTVISITED=1618283733.549; _ga=GA1.2.2119399197.1618283734; 588ku_login_refer_url=%22https%3A%5C%2F%5C%2F588ku.com%5C%2F%22; auth_id=%2228239214%7C%5Cu98ce%7C1619147742%7Ca088af00044ac93b12d98043456f2380%22; success_target_path=%22%5C%2F%5C%2F588ku.com%5C%2F%22; sns=%7B%22token%22%3A%7B%22access_token%22%3A%2268E081387D0FD18AA9046A44A088A3AA%22%2C%22expires_in%22%3A%227776000%22%2C%22refresh_token%22%3A%22B3D38826973DC630A69451690CAF8955%22%2C%22openid%22%3A%229BD2E0779DCDFDFB6E4480B197AE9342%22%7D%2C%22type%22%3A%22qq%22%7D; ISREQUEST=1; WEBPARAMS=is_pay=1; ui_588ku=dWlkPTAmdWM9JnVzPSZ0PTZkMDg0MDQxZjNkZGE0NDc3ZTk5YWQyNGRkNWFmNGU4MTYxODI4Mzc0Mi43Mjg4NzI1MCZncj0xJnVycz0%3D; IPSSESSION=mrfi1pdcevgjr05bsr553ksvu5; e6f327d84cf03d9ac9576cdc25f9bc1b=%220a35a8a8059a99fb4f9570dd8c0a57c7%22; _k_ut_v1=%7B%22ip%22%3A%22124.79.102.248%22%2C%22d%22%3A%222021-04-19%22%2C%22h%22%3A%22%22%2C%22hd%22%3A%22%22%2C%22sem%22%3A%22%22%2C%22host%22%3A%22588ku.com%22%2C%22bm%22%3A5%2C%22kw%22%3A%22%22%7D; location=4; editorCenterAd210315=1; 588KUSSID=8mqfn6bje5mmv1q2voh3cveee4; Hm_lvt_8226f7457e3273fa68c31fdc4ebf62ff=1618283734,1618778806; Hm_lvt_3e90322e8debb1d06c9c463f41ea984b=1618283734,1618778806; _gid=GA1.2.2073293865.1618778806; _gat_gtag_UA_139867171_1=1; now_page=1; no_login_pv_page=1; write_phone_pv=4; referer=%22%5C%2F%5C%2F588ku.com%5C%2Fppt%5C%2F0-new-0-0-0-0-0-0-2%5C%2F%22; Hm_lpvt_3e90322e8debb1d06c9c463f41ea984b=1618778828; Hm_lpvt_8226f7457e3273fa68c31fdc4ebf62ff=1618778828',
}

bt_cookie={
    'Cookie':'_uab_collina=156967461386562177476064; sensorsdata2015jssdkcross=%7B%22distinct_id%22%3A%2216d77e5d4d7cb0-0642ab2f551fe3-67e1b3f-2073600-16d77e5d4d84bf%22%2C%22%24device_id%22%3A%2216d77e5d4d7cb0-0642ab2f551fe3-67e1b3f-2073600-16d77e5d4d84bf%22%2C%22props%22%3A%7B%22%24latest_traffic_source_type%22%3A%22%E8%87%AA%E7%84%B6%E6%90%9C%E7%B4%A2%E6%B5%81%E9%87%8F%22%2C%22%24latest_referrer%22%3A%22https%3A%2F%2Fwww.google.com%2F%22%2C%22%24latest_referrer_host%22%3A%22www.google.com%22%2C%22%24latest_search_keyword%22%3A%22%E6%9C%AA%E5%8F%96%E5%88%B0%E5%80%BC%22%7D%2C%22first_id%22%3A%22%22%7D; FIRSTVISITED=1573180470.157; bt_guid=4d6a5ff29d7014205142c2ffed609e50; sign=QQ; login_type=QQ; head_27378999=%2F%2Fqzapp.qlogo.cn%2Fqzapp%2F101334430%2F72A35C25CA546888868698A43DE84931%2F100; ISREQUEST=1; WEBPARAMS=is_pay=1; md_session_id=20191116001533874; user_refers=a%3A2%3A%7Bi%3A0%3Bs%3A6%3A%22ibaotu%22%3Bi%3A1%3Bs%3A6%3A%22ibaotu%22%3B%7D; sign=null; islogout=1; referer=http%3A%2F%2Fibaotu.com%2F%3Fm%3Dstats%26callback%3DjQuery1910030768695653239364_1573906256726%26lx%3D1%26pagelx%3D11%26exectime%3D0.0002%26loadtime%3D1.112%26_%3D1573906256728; auth_id=27378999%7C%E9%A3%8E%7C1575202273%7C9c98e305c2135053cd0014f64edde9a4; sns=%7B%22type%22%3A%22qq%22%2C%22token%22%3A%7B%22access_token%22%3A%22231E95916B83B06E83D02C76952E2E50%22%2C%22expires_in%22%3A%227776000%22%2C%22refresh_token%22%3A%2288F7FB07ED1DFD81C9AB37CBF66C8B69%22%2C%22openid%22%3A%2272A35C25CA546888868698A43DE84931%22%7D%7D; Hm_lvt_2b0a2664b82723809b19b4de393dde93=1573655930,1573811460,1573902807,1573906274; Hm_lvt_4df399c02bb6b34a5681f739d57787ee=1573655930,1573811465,1573902808,1573906274; answer_key=a_62; Hm_lpvt_2b0a2664b82723809b19b4de393dde93=1573906299; Hm_lpvt_4df399c02bb6b34a5681f739d57787ee=1573906299'}


st_cookie={
    'Cookie':"PHPSESSID=0a30fa7b08388d3ce1fd75ba9180c443; uniqid=60750a8eb2c42; from_data=YTo3OntzOjQ6Imhvc3QiO3M6MTM6Ind3dy5iYWlkdS5jb20iO3M6Mzoic2VtIjtiOjE7czoxMDoic291cmNlZnJvbSI7aTowO3M6NDoid29yZCI7czowOiIiO3M6Mzoia2lkIjtpOjMwMjA7czo4OiJzZW1fdHlwZSI7aToxO3M6NDoiZnJvbSI7aTowO30%3D; channel_info_reg=d3d3LmJhaWR1LmNvbQ%3D%3D; channel_info=d3d3LmJhaWR1LmNvbQ%3D%3D; sem_from=1; bd_vid=NzYwNDA5NjEzMzU5MDA1ODA2MixodHRwcyUzQSUyRiUyRjY5OXBpYy5jb20lMkYlM0ZzZW0lM0QxJTI2c2VtX2tpZCUzRDMwMjAlMjZzZW1fdHlwZSUzRDElMjZiZF92aWQlM0Q3NjA0MDk2MTMzNTkwMDU4MDYy; act_layer_1=A; act_layer_2=B; act_layer_3=B; act_layer_4=A; act_layer_5=A; act_layer_6=B; act_layer_7=B; search_notice=1; user_uniqid=275E23F9D4DCEE39; Qs_lvt_375926=1618283151; mediav=%7B%22eid%22%3A%22278616%22%2C%22ep%22%3A%22%22%2C%22vid%22%3A%22!fmf*EC'Fl8XRo%2B%23YpLm%22%2C%22ctn%22%3A%22%22%2C%22vvid%22%3A%22!fmf*EC'Fl8XRo%2B%23YpLm%22%2C%22_mvnf%22%3A1%2C%22_mvctn%22%3A0%2C%22_mvck%22%3A0%2C%22_refnf%22%3A0%7D; FIRSTVISITED=1618283151.801; Hm_lvt_e37e21a48e66c7bbb5d74ea6f717a49c=1618283152; Hm_lvt_ddcd8445645e86f06e172516cac60b6a=1618283152; Hm_lvt_1154154465e0978ab181e2fd9a9b9057=1618283152; s_token=a5cc7749aaec13f3e497590d94eddb2f; index-collectHInt=1; lastTime=contact-qq; session_data=YTo1OntzOjM6InVpZCI7czo4OiIyNzEwMzI0MSI7czo1OiJ0b2tlbiI7czozMjoiYWNlODliZWZhMTg0YjFhYjk4YjhmNzk2MDVhY2E4M2UiO3M6MzoidXV0IjtzOjMyOiI5Mzg5YzJkOWNkMmJiMmE4YTI4YzNjMzU0OGIyMWYzZCI7czo0OiJkYXRhIjthOjE6e3M6ODoidXNlcm5hbWUiO3M6Mzoi6aOOIjt9czo2OiJleHRpbWUiO2k6MTYxODg4Nzk2Mzt9; loginToken_27103241=7ebb14e4cc5cd5d794de7563c1fb962a276c1acd; uid=27103241; username=%E9%A3%8E; head_pic=%2F%2Fthirdqq.qlogo.cn%2Fg%3Fb%3Doidb%26k%3D36R8B8uLaTGxyJ52eficV7Q%26s%3D40%26t%3D1555697451; isdiyici=1; login_user=1; recommend-rukou=1; is_qy_vip=1; Hm_lpvt_e37e21a48e66c7bbb5d74ea6f717a49c=1618283164; active90=1; isSearch=1; searchKeyword=PPT; search_Kw=%22PPT%22; referer_page=search:index_5; ISREQUEST=1; WEBPARAMS=is_pay=1; redirect=https%3A%2F%2F699pic.com%2Fppt-246992-263-new-all-0-all-all-1-0-0-0-0-0-0-all-all.html; Qs_pv_375926=44917706411005736%2C1034005694481494700%2C3695722706837983000%2C145147728210995680; Hm_lpvt_ddcd8445645e86f06e172516cac60b6a=1618283190; Hm_lpvt_1154154465e0978ab181e2fd9a9b9057=1618283190",

}

# st_headers={
#     'Cookie':'PHPSESSID=b86a5db19bff1ad2bfbfb6e16401c8a7; uniqid=6037e7788af8d; from_data=YTo3OntzOjQ6Imhvc3QiO3M6MTM6Ind3dy5iYWlkdS5jb20iO3M6Mzoic2VtIjtiOjE7czoxMDoic291cmNlZnJvbSI7aTowO3M6NDoid29yZCI7czowOiIiO3M6Mzoia2lkIjtpOjMwMjA7czo4OiJzZW1fdHlwZSI7aToxO3M6NDoiZnJvbSI7aTowO30%3D; channel_info_reg=d3d3LmJhaWR1LmNvbQ%3D%3D; channel_info=d3d3LmJhaWR1LmNvbQ%3D%3D; sem_from=1; act_layer_1=B; act_layer_2=B; act_layer_3=A; act_layer_4=A; act_layer_5=D; act_layer_6=A; search_notice=1; user_uniqid=34DC23853438B5B2; Qs_lvt_375926=1614276474; mediav=%7B%22eid%22%3A%22841650%22%2C%22ep%22%3A%22%22%2C%22vid%22%3A%22%22%2C%22ctn%22%3A%22%22%2C%22vvid%22%3A%22%22%2C%22_mvnf%22%3A1%2C%22_mvctn%22%3A0%2C%22_mvck%22%3A1%2C%22_refnf%22%3A0%7D; FIRSTVISITED=1614276475.393; Hm_lvt_e37e21a48e66c7bbb5d74ea6f717a49c=1614276476; Hm_lvt_ddcd8445645e86f06e172516cac60b6a=1614276476; Hm_lvt_1154154465e0978ab181e2fd9a9b9057=1614276476; s_token=8fe9ff059e5e6b7af2c7fa2c48146a3b; index-collectHInt=1; lastTime=contact-qq; session_data=YTo1OntzOjM6InVpZCI7czo4OiIyNzEwMzI0MSI7czo1OiJ0b2tlbiI7czozMjoiYzljYjRlZjZkNjAzNmViZTRhOGE1NmNlZmIxNTg0NjgiO3M6MzoidXV0IjtzOjMyOiI1NGVmYjFlZDE0NDEyYWNmMDYyOGY5YWI5YmRjYjcwYiI7czo0OiJkYXRhIjthOjE6e3M6ODoidXNlcm5hbWUiO3M6Mzoi6aOOIjt9czo2OiJleHRpbWUiO2k6MTYxNDg4MTMwNDt9; loginToken_27103241=d25a873c5167202f26b31076b36bfed0682d0bd5; uid=27103241; username=%E9%A3%8E; head_pic=%2F%2Fthirdqq.qlogo.cn%2Fg%3Fb%3Doidb%26k%3D36R8B8uLaTGxyJ52eficV7Q%26s%3D40%26t%3D1555697451; isdiyici=1; recommend-rukou=1; is_qy_vip=1; Hm_lpvt_e37e21a48e66c7bbb5d74ea6f717a49c=1614276521; ISREQUEST=1; WEBPARAMS=is_pay=1; login_view=1; search_mode=PPT; referer_page=search:index_5; redirect=https%3A%2F%2F699pic.com%2Fppt-0-263-new-all-0-all-all-4-0-0-0-0-0-0-all-all.html; Qs_pv_375926=3492933931971196400%2C1096138123728091300%2C704233357561831200%2C4113960020562972000%2C3795933667573122000; Hm_lpvt_1154154465e0978ab181e2fd9a9b9057=1614278820; Hm_lpvt_ddcd8445645e86f06e172516cac60b6a=1614278820; prevClickAjax={"index":1,"sid":0,"page":4}'
#
# }


def get_headers():
    headers = {
        'User-Agent' : 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:46.0) Gecko/20100101 Firefox/{}.0'.format(random.randint(200,650)),

    }
    return headers



if __name__=='__main__':
    print(ppt_word_filter)

