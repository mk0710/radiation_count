# -*- coding: utf-8 -*-
import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtGui import QIcon
import numpy as np
from r_c import Ui_Form
import time
import xlsxwriter as xw


class rad_c(QtWidgets.QWidget, Ui_Form):
    def __init__(self):
        super(rad_c, self).__init__()
        self.setupUi(self)
        self.init()
        self.setWindowTitle("辐射表实验数据处理助手")
        self.setWindowIcon(QIcon("26.ico"))

    def init(self):
        self.open.clicked.connect(self.openfile)
        self.use.clicked.connect(self.para_used)
        self.fac.clicked.connect(self.para_new)
        self.clear.clicked.connect(self.cle)
        self.rec.clicked.connect(self.save_excel)

    def openfile(self):

        openfile_name = QFileDialog.getOpenFileName(self, '选择文件', '', '(*.dat )')
        # print(openfile_name[0])
        data_file = open(openfile_name[0])
        data = data_file.readlines()

        data_file.close()

        # data = openfile_name.read_table(openfile_name[0], header=None)
        #
        # data_file = open('C:/Users/Administrator/Desktop/F190327.dat')
        # data = data_file.readlines()
        # data_file.close()

        # 字符串时间转秒
        def t2s(t):
            h, m, s = t.strip().split(":")
            return int(h) * 3600 + int(m) * 60 + int(s)
        time_list = list()
        for x in data:
            pair = x.split()
            if len(pair) == 0:
                break
            time_list.append(t2s(pair[0]))
        result_index = list()
        result_index.append(0)
        index = 0
        u = len(time_list)
        for x in range(0, len(time_list)):
            x = index
            time_diff = time_list[x]
            while x != len(time_list) and time_list[x] - time_diff < 60:
                x += 1
            if x == len(time_list):
                break
            index = x
            result_index.append(x)
        global y
        y = []
        for x in result_index:
            y.append(data[x])
        # print('y', y)
        # 列表元素个数=矩阵行数=测量次数

        row = len(y)
        if row < 60:
            self.stat.setText('测量次数不足')
        # 去无数据list
        list_arr = []
        for i in range(row):
            if len(y[i]) > 15:
                list_arr.append(y[i])
        # 去尾部空list
        for i in range(len(list_arr)):
            if len(list_arr[i].split()) == 0:
                list_arr.pop(i)
        ll = len(list_arr)
        for i in range(ll):
            list_arr[i] = list_arr[i].split()
        # print("列表中嵌套的每一个列表代表同一时间所有仪器的测量数据:", list_arr)
        # 把二维list转数组，以list中每个嵌套的list为一行
        a = np.array(list_arr)
        a = a[0:60, :]

        # 删除无意义数据列，获取有意义数据的列坐标#############################

        cols = a.shape[1]
        # print(a.shape[0])
        true_date = []
        for c in range(1, cols-1):
            eval("true_date.append(self.lineEdit"+str(c)+".text())")
        # print(true_date)
        # 被检表所在通道数 - 1
        index_true_date = []
        for x in true_date:
            if len(x) == 0:
                index_true_date.append(99)
            else:
                index_1 = true_date.index(x)
                index_true_date.append(index_1)

        # print("被检表的坐标", index_true_date)

        # 删除第一列（时间数据）,axis=1代表列
        a = np.delete(a, 0, axis=1)
        # 转字符串为浮点
        # list_arr 行代表不同测量时间，列代表不同的设备
        a = a.astype(float)

        # 生成被检表所在通道的list
        global true_index
        true_index = [index_true_date.index(i)+1 for i in index_true_date if i != 99]
        # print(true_index)

        # 加入标准器所在的0通道，形成新的list
        index_2 = [0]
        for i in true_index:
            index_2.append(i)
        # 以index_2 删除矩阵没有意义（没有实际被检表）的列##############################a
        a = a[:, index_2]
        # print('del cols', a)

        # 削减矩阵行数，使其可以平均拆分为3组
        par = a.shape[0] // 3
        # print('par', par)
        yu = a.shape[0]-par*3
        if yu == 1:
            a = np.delete(a, 0, axis=0)
        elif yu == 2:
            a = np.delete(a, [0, 1], axis=0)
        global m1, m2, m3
        m1, m2, m3 = np.split(a, 3, axis=0)
        # 将np矩阵 四舍五入

        def np_round(matric, d):
            b = []
            for aa in matric:
                b.append(np.around(aa, decimals=d))
            return b
        m1 = np.array(np_round(m1, 4))
        m2 = np.array(np_round(m2, 4))
        m3 = np.array(np_round(m3, 4))

        # print("m1", m1)
        # 数据格式整理完毕，行代表不同时间，列代表不同设备。
        ################################################
        # 开始依据检定规程算法处理数据，第一列是标准器的数据，第二列往后从1通道顺序排列
        # fij_list:
        # 1. shape[1] means the quantity of matrix column

        # 2. 按照测量次数i，对第j被检表进行计算，结果存放为list类型

        def f_j(h):

            fij_list = []
            for j in range(1, h.shape[1]):
                for k in range(h.shape[0]):
                    # 每次测量比值 = 被检的第i个输出 / 标准的第i个输出
                    fij_list.append(h[k][j] / (h[k][0]))

            # 把测量比值的list转矩阵fij，行代表每一个被检表，列代表不同测量次数
            fij = np.array(fij_list).reshape(h.shape[1] - 1, h.shape[0]).round(3)
            # 行数/组数 = 每一组占有的行数
            # 计算样本方差，得到每一个设备的
            s = np.std(fij, axis=1, ddof=1)
            # print("样本标准偏差s:", s)
            # 每组比值的平均值fj,即每个被检表的平均比值
            fj = fij.mean(axis=1)

            #  fij-fj,fij的列分别与Fj做减法
            diff = []
            for v in range(fij.shape[1]):
                diff.append(fij[:, v] - fj)
            diff = np.array(diff)
            #  差值大于3s的位置显示1，否则为0，
            kick = []
            for v in range(diff.shape[0]):
                kick.append(diff[v, :] >= 3 * s)
            # 行代表不同测量时间，列代表不同被检表
            kick = np.array(kick)
            # 布尔转1、0
            k_b = kick + 0
            # # 算法测试##############
            # k_b[1][0] += 1
            # #######################

            k_b = k_b.T
            # kk代表k_b中1元素的坐标，行代表不同测量时间，列代表不同被检表
            kk = np.argwhere(k_b == 1)
            kkl = kk.tolist()
            if len(kk) != 0:
                # 不合格的元素替换为0
                def dell(aa, z):
                    for zz in z:
                        aa[zz[0]][zz[1]] = 0
                    return aa
                fij = dell(fij, kkl)

                fij = fij.tolist()
                index_00 = 0
                mm = []
                for ii in fij:
                    k = len(ii)
                    for i0 in ii:
                        if int(i0 * 1000000) == 0:
                            index_00 += 1
                    mm.append(sum(ii) / (k - index_00))
                    index_00 = 0
                return [round(m, 3) for m in mm]

            else:
                fjf = []
                for ff in fj:
                    fjf.append(ff)
                return [round(m, 3) for m in fjf]
        try:
            x = f_j(m1)
            y = f_j(m2)
            z = f_j(m3)
            fin = []
            fin.append(x)
            fin.append(y)
            fin.append(z)
            fin_array = np.array(fin)
            # print(fin_array)
            global mean_fin
            mean_fin = np.mean(fin_array, 0)
            # print('最终的F', mean_fin[0])

        except ZeroDivisionError:
            pass
        # 提取灵敏度################################
        sen = []
        for i in range(1, 20):
            eval("sen.append(self.lineEdit" + str(i+40) + ".text())")
        # print('被检灵敏度:', sen)
        int_sen = []
        for i in sen:
            if len(i) == 0:
                int_sen.append(0)
            else:
                int_sen.append(float(i))
        # print('被检灵敏度str2flo:', int_sen)
        global fin_sen
        fin_sen = []
        for i in int_sen:
            if i != 0:
                fin_sen.append(i)
        # print('输入的灵敏度:', fin_sen)
        # 提取完毕 fin_sen

    def para_used(self):
        lim = 0.08
        k_0 = float(self.lineEdit40.text())
        self.new_sen = []
        self.stability = []
        for i in range(len(true_index)):
            self.new_sen.append(round(k_0 * mean_fin[i], 3))
            self.stability.append(abs(1 - (self.new_sen[i] / fin_sen[i])))
            if abs(self.stability[i]) <= lim:
                eval("self.lineEdit" + str(true_index[i]+20) + ".setText('合格')")
            else:
                eval("self.lineEdit" + str(true_index[i]+20) + ".setText('不合格')")
        print(self.stability)
        # 把稳定性数值转为小数点后3位有效
        self.stability = np.array(self.stability)
        self.stab = [float('{:.3f}'.format(i)) for i in self.stability]
        # self.stab.append(np.around(self.stability, decimals=3))

    def para_new(self):
        lim = 0.05
        k_0 = float(self.lineEdit40.text())

        self.new_sen = []
        self.stability = []
        for i in range(len(true_index)):
            self.new_sen.append(round(k_0 * mean_fin[i], 3))
            self.stability.append(abs(1 - (self.new_sen[i] / fin_sen[i])))

            if abs(self.stability[i]) <= lim:
                eval("self.lineEdit" + str(true_index[i] + 20) + ".setText('合格')")
            else:
                eval("self.lineEdit" + str(true_index[i] + 20) + ".setText('不合格')")

    def cle(self):

        for i in range(1, 20):
            eval("self.lineEdit" + str(i + 20) + ".setText('')")

    def save_excel(self):
        #  检定结果装入列表
        results = []
        old_sense = []
        sn = []
        m11 = []
        m22 = []
        m33 = []
        index = 1
        for i in range(len(true_index)):
            eval("results.append(self.lineEdit" + str(true_index[i]+20) + ".text())")
            eval("old_sense.append(self.lineEdit" + str(true_index[i]+40) + ".text())")
            eval("sn.append(self.lineEdit" + str(true_index[i]) + ".text())")
            m11.append(m1[:, [0, index]])
            m22.append(m2[:, [0, index]])
            m33.append(m3[:, [0, index]])
            index += 1
        # print('m11:', m11)
        # print('m22:', m22)
        # print('m33:', m33)
        #
        # print('serial_no：', sn)
        # print('sense：', old_sense)
        # print('new：', self.new_sen)
        print('stab', self.stab)
        # print('result：', results)

        def excel(serial_no, sense, new, stab, result, data1, data2, data3):

                workbook = xw.Workbook(serial_no + '.xlsx')  # 创建一个名为Dome2.xlsx的表格

                worksheet1 = workbook.add_worksheet()  # 添加第一个表单，默认为sheet1

                merge1_format = workbook.add_format({
                    'font_name': '黑体',
                    'size': '16',
                    'bold': True,
                    'align': 'center',  # 水平居中
                    'valign': 'vcenter',  # 垂直居中

                })
                # merge1_format.set_num_format('0.0000')
                worksheet1.set_column('A:A', 7)

                worksheet1.merge_range('A1:J4', ' 总辐射表检定记录表', merge1_format)

                merge2_format = workbook.add_format({
                    'border': 1,
                    'font_name': '黑体',
                    'size': '12',
                    'align': 'light',  # 水平居中
                    'valign': 'vcenter',  # 垂直居中
                })
                # merge2_format.set_num_format('0.0000')
                worksheet1.set_row(4, 12)
                worksheet1.set_row(5, 12)
                worksheet1.merge_range('A5:J6', ' 温度:     （℃）               湿度:       （RH%）             '
                                                '风速:     （m/s）', merge2_format)

                worksheet1.set_row(6, 24)
                worksheet1.merge_range('A7:B7', '被检仪器序号:', merge2_format)
                worksheet1.merge_range('C7:J7', serial_no, merge2_format)

                worksheet1.set_row(7, 24)
                worksheet1.merge_range('A8:B8', '被检器灵敏度:          ', merge2_format)
                worksheet1.merge_range('C8:J8', sense + '（μV/W/m²）       ', merge2_format)

                worksheet1.set_row(8, 24)
                worksheet1.merge_range('A9:B9', '标准表灵敏度:          ', merge2_format)
                worksheet1.merge_range('C9:J9', self.lineEdit40.text() + '（μV/W/m²）       ', merge2_format)

                for i in range(9, 30):
                    worksheet1.set_row(i, 20)

                merge3_format = workbook.add_format({
                    'border': 1,
                    'font_name': '黑体',
                    'size': '12',
                    'align': 'center',  # 水平居中
                    'valign': 'vcenter',  # 垂直居中
                    'text_wrap': True
                })

                merge4_format = workbook.add_format({
                    'border': 1,
                    'bold': True,
                    'font_name': 'Times New Roman',
                    'size': '12',
                    'align': 'center',  # 水平居中
                    'valign': 'vcenter',  # 垂直居中
                    'text_wrap': True
                })

                worksheet1.merge_range('A10:A30', '检\n\n定\n\n数\n\n据', merge3_format)
                worksheet1.write('B10', '序号', merge3_format)
                worksheet1.write('C10', '标准器', merge3_format)
                worksheet1.write('D10', '被检表', merge3_format)
                worksheet1.write('E10', '序号', merge3_format)
                worksheet1.write('F10', '标准器', merge3_format)
                worksheet1.write('G10', '被检表', merge3_format)
                worksheet1.write('H10', '序号', merge3_format)
                worksheet1.write('I10', '标准器', merge3_format)
                worksheet1.write('J10', '被检表', merge3_format)

                index1 = 1
                index2 = 21
                index3 = 41

                for i in range(10, 30):
                    worksheet1.write(i, 1, index1, merge4_format)
                    worksheet1.write(i, 2, str(data1[i-10, 0]), merge3_format)
                    worksheet1.write(i, 3, str(data1[i-10, 1]), merge3_format)
                    worksheet1.write(i, 4, index2, merge4_format)
                    worksheet1.write(i, 5, str(data2[i-10, 0]), merge3_format)
                    worksheet1.write(i, 6, str(data2[i-10, 1]), merge3_format)
                    worksheet1.write(i, 7, index3, merge4_format)
                    worksheet1.write(i, 8, str(data3[i-10, 0]), merge3_format)
                    worksheet1.write(i, 9, str(data3[i-10, 1]), merge3_format)

                    index1 += 1
                    index2 += 1
                    index3 += 1

                worksheet1.set_row(30, 24)

                worksheet1.merge_range('A31:B31', '被检表新灵敏度:', merge2_format)
                worksheet1.merge_range('C31:E31', '{:.2f}'.format(new) + '(μV/W/m²)', merge2_format)
                worksheet1.merge_range('F31:G31', '稳定性:', merge2_format)
                print("stab2", stab)
                worksheet1.merge_range('H31:J31', '{:.1f}'.format(stab*100)+"%", merge2_format)

                worksheet1.set_row(31, 24)
                worksheet1.merge_range('A32:B32', '检定结果:', merge2_format)
                worksheet1.merge_range('C32:J32', result, merge3_format)

                worksheet1.set_row(32, 24)
                worksheet1.merge_range('A33:J33', '检定员:                                  核验员:                 ',
                                       merge2_format)
                worksheet1.set_row(33, 24)

                merge4_format = workbook.add_format({

                    'font_name': '黑体',
                    'size': '13',
                    'align': 'center',  # 水平居中
                    'valign': 'vcenter',  # 垂直居中
                })
                dt = time.localtime(time.time())
                dtt = time.strftime('%Y-%m-%d', dt)
                worksheet1.merge_range('I34:J34', dtt, merge4_format)
                workbook.close()
        for i in range(len(fin_sen)):
            excel(sn[i], old_sense[i], self.new_sen[i], self.stab[i], results[i], m11[i], m22[i], m33[i])


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    myshow = rad_c()
    myshow.show()
    sys.exit(app.exec_())
