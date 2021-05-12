# coding=utf-8

import os
import sys
import xlrd
import xlwt
from inspect import getsourcefile
from os.path import abspath
# from numpy.core import unicode


class read_log:
    def __init__(self, fileaddress):
        x = 0
        y = 0
        # fileaddress = input("please input log_file address")
        (path, filename) = os.path.split(fileaddress)
        self.txt_addr = fileaddress
        self.txt_name = filename[0:-4]
        self.filepath = path
        self.xlsname = "Log_Check_" + self.txt_name + ".xls"
        self.xlspath = os.path.join(self.filepath, self.xlsname)
        print('*' * 20)
        print("初始化" + self.txt_name + "中")
        if not os.path.exists(self.xlspath):
            file = open(self.txt_addr, 'r')
            xls = xlwt.Workbook()
            sheet = xls.add_sheet('sheet1', cell_overwrite_ok=True)
            while True:
                lines = file.readline()
                if not lines:
                    break
                for i in lines.split(','):
                    item = i.strip().encode('utf8').decode('utf8')
                    sheet.write(x, y, item)
                    # if x == 0 and y == 0:
                    #     self.start_time = float(item)
                    y += 1
                x += 1
                y = 0
            file.close()
            os.chdir(self.filepath)
            xls.save(self.xlsname)
        print(self.txt_name + '初始化完成')

    def read_data(self, testnum):
        dict_data = {}
        workbook = xlrd.open_workbook(self.xlspath)
        sheet_new = workbook.sheet_by_name('sheet1')
        new_time = []
        self.col_count = sheet_new.ncols
        for item in sheet_new.col_values(0):
            item = item.strip().encode('raw_unicode_escape')
            new_time.append(item)
        time2 = []
        for item in new_time[0:-2]:
            item = float(item)
            time2.append(item)
        test_item = []
        for item in sheet_new.col_values(testnum):
            item = item.strip().encode('raw_unicode_escape')
            test_item.append(item)
        test_item_new = []
        for item in test_item[0:-2]:
            item = float(item)
            test_item_new.append(item)
        if len(test_item_new) == len(time2):
            dict_data = dict(zip(time2, test_item_new))
        self.start_time = float(time2[0])
        return dict_data

    def check_distract(self):
        print("分心检查... ... ")
        distract_dict = self.read_data(8)
        region_dict = self.read_data(9)
        distract_appear = []
        distract_key = []
        a = 0
        b = 0
        for key, value in distract_dict.items():
            if int(value) >= 2:
                if a == 0:
                    if b == 0:
                        distract_appear.append(key)
                        # distract_key.append(key)
                        key_last = key
                        b += 1
                        a += 1
                        continue
                    elif (key - key_last) > 100:
                        distract_appear.append(key)
                        # distract_key.append(key)
                        key_last = key
                        a += 1
                        continue
                if key - key_last < 100:
                    key_last = key
                    distract_key.append(key)
                    continue
            else:
                a = 0
        for key in distract_key:
            del distract_dict[key]
            del region_dict[key]
        key_del = []
        for key, value in distract_dict.items():
            for item in distract_appear:
                if key < item and (item - key) < 4000:
                    key_del.append(key)
        for key in key_del:
            if key in distract_dict.keys():
                del distract_dict[key]
                del region_dict[key]
        for key in distract_appear:
            if key in distract_dict.keys():
                del distract_dict[key]
                del region_dict[key]
        c = 0
        d = 0
        start_time = []
        end_time = []
        for key, value in region_dict.items():
            if int(value) >= 2:
                if c == 0:
                    start_time.append(key)
                    c += 1
                key_last = key
            else:
                if c == 1:
                    end_time.append(key_last)
                c = 0
        stt = []
        edt = []
        if len(start_time) == len(end_time):
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 2000:
                    stt.append(start_time[i])
                    edt.append(end_time[i])
        elif len(start_time) - len(end_time) == 1:
            del start_time[-1]
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 2000:
                    stt.append(start_time[i])
                    edt.append(end_time[i])
        key_record = []
        if len(stt) != 0:
            print('正在记录分心异常时间点前后三秒数据')
            xls = xlwt.Workbook()
            for stts in stt:
                file = open(self.txt_addr, 'r')
                sheet = xls.add_sheet(str(stts), cell_overwrite_ok=True)
                for key, value in region_dict.items():
                    if abs(key - stts) <= 3000:
                        key_record.append(key)
                x = 0
                y = 0
                while True:
                    lines = file.readline()
                    if not lines:
                        break
                    line = []
                    for i in lines.split(','):
                        item = i.strip().encode('utf8').decode('utf8')
                        line.append(item)
                    # line[0].find('\x00')
                    if '\x00' in line[0]:
                        break
                    if float(line[0]) in key_record:
                        for i in lines.split(','):
                            item = i.strip().encode('utf8').decode('utf8')
                            sheet.write(x, y, item)
                            y += 1
                        x += 1
                    y = 0
                key_record.clear()
                file.close()
            report_distraction_filepath = self.filepath + "\\distrction_log_extract"
            if not os.path.exists(report_distraction_filepath):
                os.mkdir(report_distraction_filepath)
            os.chdir(report_distraction_filepath)
            xls.save(self.txt_name + "_distration_check.xls")
        else:
            print('文件' + self.txt_name + '无分心异常')
        return stt, edt

    def check_fatigue(self):
        print("疲劳检查... ... ")
        fitigue_dict = self.read_data(159)
        event_dict = self.read_data(1)
        fitigue_appear = []
        fitigue_key = []
        a = 0
        b = 0
        for key, value in fitigue_dict.items():
            if int(value) >= 2:
                if a == 0:
                    if b == 0:
                        fitigue_appear.append(key)
                        key_last = key
                        b += 1
                        a += 1
                        continue
                    elif (key - key_last) > 100:
                        fitigue_appear.append(key)
                        # distract_key.append(key)
                        key_last = key
                        a += 1
                        continue
                if key - key_last < 100:
                    key_last = key
                    fitigue_key.append(key)
                    continue
            else:
                a = 0
        for key in fitigue_key:
            del fitigue_dict[key]
            del event_dict[key]
        key_del = []
        for key, value in fitigue_dict.items():
            for item in fitigue_appear:
                if key < item and (item - key) < 5000:
                    key_del.append(key)
        for key in key_del:
            if key in fitigue_dict.keys():
                del fitigue_dict[key]
                del event_dict[key]
        for key in fitigue_appear:
            if key in fitigue_dict.keys():
                del fitigue_dict[key]
                del event_dict[key]
        yawn_list = []
        eyeclose_list = []
        xls = xlwt.Workbook()
        sheet_yawn = xls.add_sheet('yawn_list', cell_overwrite_ok=True)
        sheet_eyeclose = xls.add_sheet('eyeclose_list', cell_overwrite_ok=True)
        for key, value in event_dict.items():
            if int(value) == 5:
                yawn_list.append(key)
            elif int(value) == 6:
                eyeclose_list.append(key)
        if len(yawn_list) != 0 and len(eyeclose_list) != 0:
            print('正在记录疲劳异常时间点')
            xls = xlwt.Workbook()
            sheet_yawn = xls.add_sheet('yawn_list', cell_overwrite_ok=True)
            sheet_eyeclose = xls.add_sheet('eyeclose_list', cell_overwrite_ok=True)
            c = 0
            d = 0
            for key, value in event_dict.items():
                if int(value) == 5:
                    sheet_yawn.write(c, 0, key)
                    c += 1
                elif int(value) == 6:
                    sheet_eyeclose.write(c, 0, key)
                    d += 1
            report_fatigue_filepath = self.filepath + "\\fatigue_log_extract"
            if not os.path.exists(report_fatigue_filepath):
                os.mkdir(report_fatigue_filepath)
            os.chdir(report_fatigue_filepath)
            xls.save(self.txt_name + "_fitigue_check.xls")
        else:
            print('文件' + self.txt_name + '无疲劳异常')
        return yawn_list, eyeclose_list

    def check_noface(self):
        print("人脸检查... ... ")
        face_result_dict = self.read_data(240)
        dace_detect_dict = self.read_data(42)
        face_lost_appear = []
        face_lost_key = []
        a = 0
        b = 0
        for key, value in face_result_dict.items():
            if int(value) == 2:
                if a == 0:
                    if b == 0:
                        face_lost_appear.append(key)
                        # distract_key.append(key)
                        key_last = key
                        b += 1
                        a += 1
                        continue
                    elif (key - key_last) > 100:
                        face_lost_appear.append(key)
                        # distract_key.append(key)
                        key_last = key
                        a += 1
                        continue
                if key - key_last < 100:
                    key_last = key
                    face_lost_key.append(key)
                    continue
            else:
                a = 0
        for key in face_lost_key:
            del face_result_dict[key]
            del dace_detect_dict[key]
        key_del = []
        for key, value in face_result_dict.items():
            for item in face_lost_appear:
                if key < item and (item - key) < 5000:
                    key_del.append(key)
        for key in key_del:
            if key in face_result_dict.keys():
                del face_result_dict[key]
                del dace_detect_dict[key]
        for key in face_lost_appear:
            if key in face_result_dict.keys():
                del face_result_dict[key]
                del dace_detect_dict[key]
        c = 0
        d = 0
        start_time = []
        end_time = []
        for key, value in dace_detect_dict.items():
            if int(value) == 0:
                if c == 0:
                    start_time.append(key)
                    c += 1
                key_last = key
            else:
                if c == 1:
                    end_time.append(key_last)
                c = 0
        stt = []
        edt = []
        if len(start_time) == len(end_time):
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 3000:
                    stt.append(start_time[i])
                    edt.append(end_time[i])
        elif len(start_time) - len(end_time) == 1:
            del start_time[-1]
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 3000:
                    stt.append(start_time[i])
                    edt.append(end_time[i])

        key_record = []

        if len(stt) != 0:
            print('正在写入人脸丢失异常时间点')
            xls = xlwt.Workbook()
            sheet = xls.add_sheet('noface_check', cell_overwrite_ok=True)
            c = 0
            for item in stt:
                sheet.write(c, 0, item)
                c += 1
            report_noface_filepath = self.filepath + "\\noface_log_extract"
            if not os.path.exists(report_noface_filepath):
                os.mkdir(report_noface_filepath)
            os.chdir(report_noface_filepath)
            xls.save(self.txt_name + "_noface_check.xls")
        else:
            print('文件' + self.txt_name + '无人脸丢失异常')
        return stt, edt


if __name__ == '__main__':

    # folderaddress = input("请输入文件夹路径: ")
    # if not os.path.exists(folderaddress):
    #     print("请输入正确路径")
    #     continue

    folderaddress, filepath = os.path.split(abspath(getsourcefile(lambda:0)))
    # while True:
    #     print(folderaddress)
    # C:\Users\Admin\Desktop\2021_5_7-dms\2021_5_7_13_51_4_dms.txt
    filelist = os.listdir(folderaddress)
    # test_item = input("请输入测试项目：\n 1.分心 2.疲惫 3.人脸丢失 4.全部（不包括3）\n")
    xls = xlwt.Workbook()
    sheet1 = xls.add_sheet('distract_check', cell_overwrite_ok=True)
    sheet2 = xls.add_sheet('fitigue_check', cell_overwrite_ok=True)
    sheet3 = xls.add_sheet('noface_check', cell_overwrite_ok=True)
    x1 = 0
    x2 = 0
    x3 = 0
    sheet1.write(x1, 0, '编号')
    sheet1.write(x1, 1, '绝对时间')
    sheet1.write(x1, 2, '起始时间戳')
    sheet1.write(x1, 3, '结束时间戳')
    sheet1.write(x1, 4, '所属数据')

    sheet2.write(x2, 0, '编号')
    sheet2.write(x2, 1, '绝对时间')
    sheet2.write(x2, 2, '时间点')
    sheet2.write(x2, 3, '所属数据')

    sheet3.write(x3, 0, '编号')
    sheet3.write(x3, 1, '绝对时间')
    sheet3.write(x3, 2, '起始时间戳')
    sheet3.write(x3, 3, '结束时间戳')
    sheet3.write(x3, 4, '所属数据')

    for file in filelist:
        if os.path.splitext(file)[1] == ".txt":
            fileaddress = os.path.join(folderaddress, file)
            test = read_log(fileaddress)
            test.read_data(0)
            start_time = test.start_time
            cols = test.col_count
            # if int(test_item) == 1:
            str1 = test.txt_name
            position = []
            for i in range(0, len(str1)):
                if str1[i] == '_':
                    position.append(i)
            year = str1[0: int(position[0])]
            monate = str1[(int(position[0]) + 1): int(position[1])]
            day = str1[(int(position[1]) + 1): int(position[2])]
            year_mon_day = str1[0: int(position[2]) + 1]
            hour = int(str1[int(position[2]) + 1: int(position[3])])
            minute = int(str1[int(position[3]) + 1: int(position[4])])
            second = int(str1[int(position[4]) + 1: int(position[5])])
            if cols >= 240:
                stt_noface, edt_noface = test.check_noface()
                for i in range(0, len(stt_noface)):
                    x3 += 1
                    time_diff_noface = (stt_noface[i] - start_time) // 1000
                    ms = (stt_noface[i] - start_time) % 1000
                    if time_diff_noface < 3600:
                        hr_noface = 0
                        ses_noface = time_diff_noface % 60
                        mnt_noface = time_diff_noface // 60
                    else:
                        hr_noface = time_diff_noface // 3600
                        a_noface = time_diff_noface % 3600
                        ses_noface = a_noface % 60
                        mnt_noface = a_noface // 60
                    second_noface = second + ses_noface
                    if second_noface >= 60:
                        second_noface -= 60
                        minute_noface = minute + mnt_noface + 1
                    else:
                        minute_noface = minute + mnt_noface
                    if minute_noface >= 60:
                        minute_noface -= 60
                        hour_noface = hour + hr_noface + 1
                    else:
                        hour_noface = hour + hr_noface
                    video_time_noface = year_mon_day + str(hour_noface) + '_' + str(minute_noface) + '_' \
                                        + str(second_noface) + '_' + str(ms)
                    sheet3.write(x3, 0, x1)
                    sheet3.write(x3, 1, video_time_noface)
                    sheet3.write(x3, 2, stt_noface[i])
                    sheet3.write(x3, 3, edt_noface[i])
                    sheet3.write(x3, 4, test.txt_name)

            stt, edt = test.check_distract()
            for i in range(0, len(stt)):
                x1 += 1
                time_diff_dis = int((stt[i] - start_time) // 1000)
                ms_dis = int((stt[i] - start_time) % 1000)
                if time_diff_dis < 3600:
                    hr_dis = 0
                    ses_dis = int(time_diff_dis % 60)
                    mnt_dis = int(time_diff_dis // 60)
                else:
                    hr_dis = int(time_diff_dis // 3600)
                    a_dis = int(time_diff_dis % 3600)
                    ses_dis = int(a_dis % 60)
                    mnt_dis = int(a_dis // 60)
                second_dis = second + ses_dis
                if second_dis >= 60:
                    second_dis -= 60
                    minute_dis = minute + mnt_dis + 1
                else:
                    minute_dis = minute + mnt_dis
                if minute_dis >= 60:
                    minute_dis -= 60
                    hour_dis = hour + hr_dis + 1
                else:
                    hour_dis = hour + hr_dis
                video_time_dis = year_mon_day + str(hour_dis) + '_' + str(minute_dis) + '_' \
                                    + str(second_dis) + '_' + str(ms_dis)
                sheet1.write(x1, 0, x1)
                sheet1.write(x1, 1, video_time_dis)
                sheet1.write(x1, 2, stt[i])
                sheet1.write(x1, 3, edt[i])
                sheet1.write(x1, 4, test.txt_name)
            yawn_list, eyeclose_list = test.check_fatigue()
            for i in range(0, len(yawn_list)):
                x2 += 1
                time_diff_yawn = (yawn_list[i] - start_time) // 1000
                ms_yawn = (yawn_list[i] - start_time) % 1000
                if time_diff_yawn < 3600:
                    hr_yawn = 0
                    ses_yawn = time_diff_yawn % 60
                    mnt_yawn = time_diff_yawn // 60
                else:
                    hr_yawn = time_diff_yawn // 3600
                    a_yawn = time_diff_yawn % 3600
                    ses_yawn = a_yawn % 60
                    mnt_yawn = a_yawn // 60
                second_yawn = second + ses_yawn
                if second_yawn >= 60:
                    second_yawn -= 60
                    minute_yawn = minute + mnt_yawn + 1
                else:
                    minute_yawn = minute + mnt_yawn
                if minute_yawn >= 60:
                    minute_yawn -= 60
                    hour_yawn = hour + hr_yawn + 1
                else:
                    hour_yawn = hour + hr_yawn
                video_time_yawn = year_mon_day + str(hour_yawn) + '_' + str(minute_yawn) + '_' \
                                 + str(second_yawn) + '_' + str(ms_yawn)
                sheet2.write(x2, 0, x1)
                sheet2.write(x2, 1, video_time_yawn)
                sheet2.write(x2, 2, yawn_list[i])
                sheet2.write(x2, 3, test.txt_name)
            for i in range(0, len(eyeclose_list)):
                x2 += 1
                time_diff_eyeclose = (eyeclose_list[i] - start_time) // 1000
                ms_eyeclose = (eyeclose_list[i] - start_time) % 1000
                if time_diff_eyeclose < 3600:
                    hr_eyeclose = 0
                    ses_eyeclose = time_diff_eyeclose % 60
                    mnt_eyeclose = time_diff_eyeclose // 60
                else:
                    hr_eyeclose = time_diff_eyeclose // 3600
                    a_eyeclose = time_diff_eyeclose % 3600
                    ses_eyeclose = a_eyeclose % 60
                    mnt_eyeclose = a_eyeclose // 60
                second_eyeclose = second + ses_eyeclose
                if second_eyeclose >= 60:
                    second_eyeclose -= 60
                    minute_eyeclose = minute + mnt_eyeclose + 1
                else:
                    minute_eyeclose = minute + mnt_eyeclose
                if minute_eyeclose >= 60:
                    minute_eyeclose -= 60
                    hour_eyeclose = hour + hr_eyeclose + 1
                else:
                    hour_eyeclose = hour + hr_eyeclose
                video_time_eyeclose = year_mon_day + str(hour_eyeclose) + '_' + str(minute_eyeclose) + '_' \
                                  + str(second_eyeclose) + '_' + str(ms_eyeclose)
                sheet2.write(x2, 0, x1)
                sheet2.write(x2, 1, video_time_eyeclose)
                sheet2.write(x2, 2, eyeclose_list[i])
                sheet2.write(x2, 3, test.txt_name)
        print("\n")
    print("正在生成报告")
    os.chdir(folderaddress)
    xls.save("log_check_report_all.xls")
    print("检查结束。\n")

    # test = read_log(r'G:\Hirain\WR\W16\2021_5_8-dms\2021_5_8_10_15_31_dms.txt')
    # test_dict = test.check_distract()
