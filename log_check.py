# coding=utf-8

import os

import xlrd
import xlwt


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
        if len(start_time) == len(end_time):
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 2000:
                    stt.append(start_time[i])
        elif len(start_time) - len(end_time) == 1:
            del start_time[-1]
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 2000:
                    stt.append(start_time[i])
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
            os.chdir(self.filepath)
            xls.save(self.txt_name + "_distration_check.xls")
        else:
            print('文件' + self.txt_name + '无分心异常')

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
            os.chdir(self.filepath)
            xls.save(self.txt_name + "_fitigue_check.xls")
        else:
            print('文件' + self.txt_name + '无疲劳异常')

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
        if len(start_time) == len(end_time):
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 3000:
                    stt.append(start_time[i])
        elif len(start_time) - len(end_time) == 1:
            del start_time[-1]
            for i in range(0, len(start_time)):
                if end_time[i] - start_time[i] >= 3000:
                    stt.append(start_time[i])
        key_record = []

        if len(stt) != 0:
            print('正在写入人脸丢失异常时间点')
            xls = xlwt.Workbook()
            sheet = xls.add_sheet('noface_check', cell_overwrite_ok=True)
            c = 0
            for item in stt:
                sheet.write(c, 0, item)
                c += 1
            os.chdir(self.filepath)
            xls.save(self.txt_name + "_noface_check.xls")
        else:
            print('文件' + self.txt_name + '无人脸丢失异常')

if __name__ == '__main__':
    while True:
        folderaddress = input("请输入文件夹路径: ")
        if not os.path.exists(folderaddress):
            print("请输入正确路径")
            continue
        # C:\Users\Admin\Desktop\2021_5_7-dms\2021_5_7_13_51_4_dms.txt
        filelist = os.listdir(folderaddress)
        test_item = input("请输入测试项目：\n 1.分心 2.疲惫 3.人脸丢失 4.全部（不包括3）\n")
        for file in filelist:
            if os.path.splitext(file)[1] == ".txt":
                fileaddress = os.path.join(folderaddress, file)
                test = read_log(fileaddress)
                if int(test_item) == 1:
                    test.check_distract()
                if int(test_item) == 2:
                    test.check_fatigue()
                if int(test_item) == 3:
                    test.check_noface()
                if int(test_item) == 4:
                    test.check_distract()
                    test.check_fatigue()
                    # test.check_noface()
            print("\n")
        print("检查结束。\n")




    # test = read_log(r'G:\Hirain\WR\W16\2021_5_8-dms\2021_5_8_10_15_31_dms.txt')
    # test_dict = test.check_distract()
