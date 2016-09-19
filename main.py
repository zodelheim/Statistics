# -*- coding: utf-8 -*-
import xlrd, xlwt
import os
import sys
import numpy as np
from numpy import mean, var, std, sqrt
import seaborn as sns
# import matplotlib.pyplot as plt
def canal_apparatus():
    file_name=np.array('')
    for a in xrange(0,len(os.listdir('.'))):
        if os.listdir('.')[a].find('xlsx') != -1:
            # print os.listdir('.')[a]
            file_name=np.append(file_name, (os.listdir('.')[a]))
    file_name=np.delete(file_name,0)
    sred = np.zeros(len(file_name))

    for j in xrange(0, len(file_name)):
        rb = xlrd.open_workbook(file_name[j])
        sheet = rb.sheet_by_index(0)
        cytoplasm = np.zeros(sheet.nrows/2)
        core = np.zeros(sheet.nrows/2)
        print j
        for i in xrange(1,sheet.nrows/2+1,1):
            # print sheet.cell_value(2*i,-1)
            cytoplasm[i-1] = sheet.cell_value(2*i,-1)

        for i in xrange(0,sheet.nrows/2,1):
            core[i] = sheet.cell_value(2*i+1,-1)
        sred[j] = mean(core)/mean(cytoplasm)
    print sred
    print 'среднее = ', mean(sred), 'среднеквадратичное =', sqrt(std(sred)**2*len(sred)/(len(sred)-1)), 'ошибка среднего =', sqrt(std(sred)**2*len(sred)/(len(sred)-1))/sqrt(len(sred))
def square_ur():
    file_name = np.array('')
    for a in xrange(0, len(os.listdir('.'))):
        if os.listdir('.')[a].find('xlsx') != -1:
            # print os.listdir('.')[a]
            file_name = np.append(file_name, (os.listdir('.')[a]))
    file_name = np.delete(file_name, 0)
    sred = np.zeros(len(file_name))
    for j in xrange(0, len(file_name)):
        rb = xlrd.open_workbook(file_name[j])
        sheet = rb.sheet_by_index(0)
        capsula = np.zeros(sheet.nrows/2)
        clubok = np.zeros(sheet.nrows/2)
        S = np.zeros(sheet.nrows/2)
        for i in xrange(1,sheet.nrows/2+1,1):
            capsula[i-1] = sheet.cell_value(2*i,1)
        for i in xrange(0,sheet.nrows/2,1):
            clubok[i] = sheet.cell_value(2*i+1,1)
        S = capsula - clubok
        sred[j] = mean(S)
    print 'среднее = ', mean(sred), 'среднеквадратичное =', std(sred), 'ошибка среднего =', std(sred)/sqrt(len(sred))
def gui():
    if __name__ == '__main__':
        app = QApplication(sys.argv)
        w = QWidget()
        w.resize(250, 150)
        w.move(300, 300)
        w.setWindowTitle('Simple')
        w.show()

        sys.exit(app.exec_())
canal_apparatus()
# square_ur()
# from bokeh.plotting import figure, output_file, show
# x = [1, 2, 3, 4, 5]
# y = [6, 7, 2, 4, 5]
# output_file("lines.html")
# p = figure(title="simple line example", x_axis_label='x', y_axis_label='y')
# p.line(x, y, legend="Temp.", line_width=2)
# show(p)
# square_ur()
# gui()