#祈愿-「赤团开时」_ver2.2
#项目完成时间2021年11月1日 13点03分
import xlsxwriter as xw #请先安装XlsxWriter模块
import time 
import random
import os
def clear():
    os.system('cls')
wishRecordDocument = xw.Workbook("wishRecord.xlsx") # 创建文件
style_blue = wishRecordDocument.add_format({
    'fg_color':'#00ccff' ,
    'align' : 'center' ,
})
style_purple_wuQi = wishRecordDocument.add_format({
    'fg_color':'#cc99FF',
    'align' : 'center' ,
})
style_purple_jueSe = wishRecordDocument.add_format({
    'fg_color':'#FF99CC',
    'bold': True ,
    'align': 'center' ,
})
style_golden = wishRecordDocument.add_format({
    'fg_color':'#FFCC00' ,
    'bold': True ,
    'align': 'center' ,
})
style_Time_Number = wishRecordDocument.add_format({
    'align': 'center' ,
})
wishRecord1 = wishRecordDocument.add_worksheet("wishREcord1") # 创建表
wishRecord1.activate() # 激活表
title = ['序号','记录','时间'] # 设置表头
wishRecord1.set_column(1,1,30)
wishRecord1.set_column(2,2,20)
wishRecord1.write_row('A1',title,style_Time_Number) # 从A1单元格开始写入表头
blue = ["以理服人", "沐浴龙血的剑", "飞天御剑","黑缨枪","神射手之誓","讨龙英杰谭","魔导绪论","铁影阔剑","反曲弓","翠玉法球","冷刃","弹弓","鸦羽弓"]
purple_JueSe = ["忍里之貉.早柚(风) up","猫尾特调.迪奥娜(冰) up","渡来介者.托马(火) up","掩月天权.凝光(岩)", "智明无邪.烟绯(火)","雪融有踪.重云(冰)","棘冠恩典.罗莎莉亚(冰)","无冕的龙王.北斗(雷)","少年春衫薄.行秋(水)","断罪皇女.菲谢尔(雷)","闪耀的偶像.芭芭拉(水)","燥热旋律.辛焱(火)","未授勋之花.诺艾尔(岩)","万民百味.香菱(火)","无害甜度.砂糖(风)","命运试金石.班尼特(火)","狼少年.雷泽(雷)"]
purple_WuQi =["弓藏", "祭礼弓", "绝弦", "西风猎弓", "祭礼残章", "昭心", "流浪乐章", "西风秘典", "西风长枪", "匣里灭辰","雨裁","祭礼大剑","钟剑","西风大剑","匣里龙吟", "祭礼剑", "笛剑", "西风剑"]
Golden = ["雪霁梅香.胡桃(火)", "霆霓快雨.刻晴(雷)", "星天水镜.莫娜(水)", "晨曦的暗面.迪卢克(火)", "蒲公英骑士.琴(风)", "冰冻回魂夜.七七(冰)"]
purple_baoDi = 0
purple_up = 0
Golden_up = 0
which_wish = 0
last_Golden = 0
Golden_baoDi = 0
JiLu = {}
def qiYuan_1_jueSe():
    global which_wish,  purple_baoDi, JiLu, last_Golden , Golden_baoDi 
    which_wish = which_wish + 1
    if which_wish - last_Golden < 60 :
        Golden_baoDi += 1
    elif which_wish - last_Golden < 70 :
        Golden_baoDi += 20
    elif which_wish - last_Golden < 80 :
        Golden_baoDi += 100
    elif which_wish - last_Golden < 90 :
        Golden_baoDi += 1064.5
    this_wish = random.randint(1,10000)
    if 60 + Golden_baoDi >= this_wish :
        wu_xing()
    elif purple_baoDi  >= 9 :
        si_xing()
    else:
        qi_1 = random.randint(1,1000)
        if qi_1 <= 51:
            si_xing()
        else :
            BM = blue[random.randint(0,12)]
            print('\033[0;34m',BM.center(38,' '),'\033[0m')
            JiLu[which_wish] = "(三星)"+BM+"(武器)"
            wishRecord1.write(which_wish, 0 , which_wish , style_Time_Number)
            wishRecord1.write(which_wish, 1 , BM , style_blue)
            wishRecord1.write(which_wish , 2 ,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),style_Time_Number)
            purple_baoDi = purple_baoDi +1
    return which_wish , purple_baoDi , JiLu  
def si_xing():
    global which_wish ,  purple_up , purple_baoDi , JiLu
    wishRecord1.write(which_wish , 0 , which_wish , style_Time_Number)
    if purple_up == 1 :
        PMa = purple_JueSe[random.randint(0,2)]
        print('\033[1;35m',PMa.center(38,' '),'\033[0m')  # 四星up
        JiLu[which_wish] = "(四星)"+PMa
        wishRecord1.write(which_wish , 1 , PMa , style_purple_jueSe)
        purple_up = 0
    else:    
        get_purple = random.randint(0,1)
        if get_purple == 1:
            PMa = purple_JueSe[random.randint(0,2)]
            print('\033[1;35m',PMa.center(38,' '),'\033[0m')  # 四星up
            JiLu[which_wish] = "(四星)"+PMa
            wishRecord1.write(which_wish , 1 , PMa , style_purple_jueSe)
            purple_up = 0
        else:
            what_purple = random.randint(0,9)%2
            purple_up = 1
            if what_purple == 0 :
                PMa = purple_JueSe[random.randint(3,16)]
                print('\033[1;35m',PMa.center(38,' '),'\033[0m')
                JiLu[which_wish] = "(四星)"+PMa
                wishRecord1.write(which_wish , 1 , PMa , style_purple_jueSe)
            else :
                PMa = purple_WuQi[random.randint(0,17)]
                print('\033[1;35m',PMa.center(38,' '),'\033[0m')
                JiLu[which_wish] = "(四星)"+PMa+"(武器)"
                wishRecord1.write(which_wish , 1 , PMa , style_purple_wuQi)
    purple_baoDi = 0
    wishRecord1.write(which_wish , 2 ,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),style_Time_Number)
    return which_wish , purple_baoDi , JiLu , purple_up
def wu_xing():
    global Golden_up , which_wish , purple_baoDi , last_Golden , Golden_baoDi , JiLu
    if Golden_up == 1 :
        GMa = Golden[0]
        print('\033[22;33m',GMa.center(38,' '),"up",'\033[0m')     #五星up
        JiLu[which_wish] = "(五星)"+GMa+"up"
        Golden_up = 0  
    else :
        may_5 = random.randint(0,9)%2
        if may_5!=0 :
            GMa = Golden[0]
            print('\033[22;33m',GMa.center(38,' '),"up",'\033[0m')
            JiLu[which_wish] = "(五星)"+GMa+"up"
            Golden_up = 0
        else:
            GMa = Golden[random.randint(1,5)]
            print('\033[22;33m',GMa.center(38,' '),'\033[0m')
            JiLu[which_wish] = "(五星)"+GMa
            Golden_up = 1
    wishRecord1.write(which_wish , 0 , which_wish , style_Time_Number)
    wishRecord1.write(which_wish , 1 , GMa , style_golden)
    wishRecord1.write(which_wish , 2 ,time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()),style_Time_Number)
    last_Golden = which_wish
    purple_baoDi = purple_baoDi +1
    Golden_baoDi = 0
    return purple_baoDi , last_Golden , Golden_up , which_wish , JiLu , Golden_baoDi 
def qiYuan_10_jueSe():
    for i in range(0,10,1):      
        qiYuan_1_jueSe()      
def logs_wish() :
    print('\033[33m','祈愿记录'.center(45,'#'),'\033[0m')
    for key in sorted(JiLu.keys()):
        if JiLu[key].find("五星") >-1 :
            print(key,'\033[22;33m',JiLu[key],'\033[0m')
        elif JiLu[key].find("四星") >-1 :
            print(key,'\033[1;35m',JiLu[key],'\033[0m')
        elif JiLu[key].find("三星") >-1 :
            print(key,JiLu[key])
def logs_many_5() :
    global which_wish
    many_5 = 0
    many_4 = 0
    print('\033[33m',"祈愿结果统计".center(30,'-'),'\033[0m')
    print('\033[1;36m',"共计祈愿",which_wish,"次",'\033[0m')
    for key in sorted(JiLu.keys()):
        if JiLu[key].find("五星") >-1 :
            many_5 = many_5 +1
            print('\033[22;33m',JiLu[key],"(",key,")",'\033[0m')
    print('\033[1;36m',"共获得五星",many_5,"个",'\033[0m')
    for key in sorted(JiLu.keys()):
        if JiLu[key].find("四星") >-1 :
            many_4 = many_4 +1
            print('\033[1;35m',JiLu[key],"(",key,")",'\033[0m')
    print('\033[1;36m',"共获得四星",many_4,"个",'\033[0m')
def maker():
    print('\033[33m','开发者简介'.center(40,'*'),'\033[0m')
    print('\033[22;33m',"作者:顾博荣".center(38,' '),'\033[0m')
    print('\033[1;36m',"个人网站:https://www.luotianyi.press".center(38,' '),'\033[0m')
    print('\033[22;33m',"Copyright ©2020-2021 © 顾博荣 版权所有".center(38,' '),'\033[0m')
    print('\033[22;31m',"可交流学习使用，未经许可，禁止商用".center(26,' '),'\033[0m')
    print('\033[33m','*'.center(45,'*'),'\033[0m')
def wish_profile():
    wish_tell = '''
※ 以上角色中，限定角色不会进入「奔行世间」常驻祈愿。

※ 更多祈愿信息可在祈愿界面输入0进行查询。
    '''
    wish_talk = "祈愿介绍"
    print('\033[33m',wish_talk.center(70,'='))
    print('●活动期间，限定五星角色','\033[22;31m','「雪霁梅香.胡桃(火)」','\033[33m','的祈愿获取概率将大幅提升！\n')
    print('\033[33m','●活动期间，四星角色','\033[1;31m','「渡来介者.托马(火)」','\033[0m',end='')
    print('\033[1;36m','「忍里之貉.早柚(风)」\n','\033[1;34m','\n  「猫尾特调.迪奥娜(冰)」','\033[33m','的祈愿获取概率将大幅提升！\n')
    print('\033[33m','●活动结束后，四星角色','\033[1;31m','「渡来介者.托马(火)」','\033[33m','将在2.3版本进入「奔行世间」常驻祈愿。')
    print('\033[33m',wish_tell,'\033[0m')
    
def wish_rules():
    rules_name = '祈愿规则'
    rules1 = '''
【5星物品】
在本期「赤团开时」活动祈愿中，5星角色祈愿的基础概率为0.600%，
综合概率（含保底）为1.600%，最多90次祈愿内必定能通过保底获取5星角色。
当祈愿获取到5星角色时，有50.000%的概率为本期UP角色'''
    rules2 = '''
如果本次祈愿获取的5星角色非本期UP角色，下次祈愿获取的5星角色必定为本期5星UP角色。
【4星物品】
在本期「赤团开时」活动祈愿中，4星物品祈愿的基础概率为5.100%，
4星角色祈愿的基础概率为2.550%，4星武器祈愿的基础概率为2.550%，
4星物品祈愿的综合概率（含保底）为13.000%。
最多10次祈愿必定能通过保底获取4星或以上物品，通过保底获取4星物品的概率为99.400%，
获取5星物品的概率为0.600%。
当祈愿获取到4星物品时，有50.000%的概率为本期4星UP角色'''
    rules3 = '''如果本次祈愿获取的4星物品非本期4星UP角色，下次祈愿获取的4星物品必定为本期4星UP角色。
当祈愿获取到4星UP物品时，每个本期4星UP角色的获取概率均等。'''
    print('\033[33m',rules_name.center(70,'='),'\033[0m')
    print('\033[33m',rules1,'\033[0m',end='')
    print('\033[22;31m','「雪霁梅香.胡桃(火)」。\n','\033[0m',end='')
    print('\033[33m',rules2,'\033[0m',end='')
    print('\033[1;36m','\n「忍里之貉.早柚(风)」','\033[0m',end='')
    print('\033[33m','、','\033[0m',end='')
    print('\033[1;34m','「猫尾特调.迪奥娜(冰)」','\033[0m',end='')
    print('\033[33m','、','\033[0m',end='')
    print('\033[1;31m','「渡来介者.托马(火)」','\033[0m',end='')
    print('\033[33m','中的一个。\n',rules3,'\033[0m')
def wish():
    wish_name = "祈愿池: 「赤团开时」"
    print('\033[1;36m',wish_name.center(30,'-'))
    print("0.查看祈愿规则\n1.祈愿一次\n2.祈愿十次\n3.查看祈愿记录\n4.统计祈愿结果\n5.查看制作者信息\n6.导出记录并退出(需安装XlsxWriter模块)",'\033[0m')
    x = str(input("请选择\t"))
    if x == '0' :
        clear()
        wish_rules()
        wish()
    elif x == "1" :
        clear()
        print('\033[33m',"恭喜获得".center(38,' '),'\033[0m')
        qiYuan_1_jueSe()
        wish()
    elif x == "2" :
        clear()
        print('\033[33m',"恭喜获得".center(38,' '),'\033[0m')
        qiYuan_10_jueSe()
        wish()
    elif x== "3" :
        clear()
        logs_wish()
        wish()
    elif x== "4" :
        clear()
        logs_many_5()
        wish()
    elif x== "5" :
        clear()
        maker()
        wish()
    elif x== "6" :
        clear()
        wishRecordDocument.close() # 关闭表
        print('\033[33m','祈愿记录已保存到当前文件路径下，文件名为wishRecord.xlsx  请及时查看')
        print('\033[1;31m','(如需保存，请修改xlsx文件名称，否则下次启动程序时将自动覆盖上次的记录)','\033[0m')
    else:
        clear()
        wish()    
wish_profile()
wish()


        