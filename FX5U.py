"""
author: alai
"""

import socket
import time
import xlwt,xlrd,struct
from xlutils.copy import copy
import sys
import re,os,sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from FX5U_UI import Ui_MainWindow
from PyQt5.QtCore import *
from PyQt5.QtSql import QSqlDatabase,QSqlQuery,QSqlTableModel,QSqlQueryModel
# from threading import Thread
data="ok"

class SqlThread(QThread,QWidget):
    trigger2=pyqtSignal(str)

    def __int__(self):
        super(SqlThread, self).__init__()
    def run(self):
        #必须从run开始启动程序

        # global data
        # pass
        # self.pushButton5.clicked.connect(self.test)
        # pass
        self.createDB()
    def createView(title, model):
        view = QTableView()
        view.setModel(model)
        view.setWindowTitle(title)
        return view

    # def initializeModel(self,model):
    #     model.setTable('people')
    #     model.setEditStrategy(QSqlTableModel.OnFieldChange)
    #     model.select()
    #     model.setHeaderData(0, Qt.Horizontal, "ID")
    #     model.setHeaderData(1, Qt.Horizontal, "name")
    #     model.setHeaderData(2, Qt.Horizontal, "address")
    #     view1 = self.createView("Table Model (View 1)", model)
    #     # view1.clicked.connect(findrow)


    def createDB(self):
        global glb_query_tag
        global glb_str1
        glb_query_tag=0
        glb_str1=""

        db = QSqlDatabase.addDatabase('QSQLITE')
        db.setDatabaseName('./database/R18003.db')
        # model=QSqlTableModel()
        # self.initializeModel(model)
        # model.setEditStrategy(QSqlTableModel.OnManualSubmit)


        # model.select()
        while True:
            if not db.open():
                QMessageBox.critical(None, ("无法打开数据库"),
                                     ("无法建立到数据库的连接,这个例子需要SQLite 支持，请检查数据库配置。\n\n"
                                      "点击取消按钮退出应用。"),
                                     QMessageBox.Cancel)
                return False

            if glb_query_tag == 1:
                scanfinished = glb_str1[22:26]
                barcode = glb_str1[26:90]
                wenjian = glb_str1[90:130]
                pressure = glb_str1[130:134]
                result = glb_str1[134:138]
                # print(wenjian)
                query=QSqlQuery()



                time1 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
                query.exec_(

                    "insert into R04 values('{}', '{}', '{}')".format(time1, barcode, wenjian))
                glb_query_tag = 0

                db.close()
            if glb_query_tag == 2:
                # print("button5 is clicked")
                self.model = QSqlTableModel(self)
                self.model.setTable("R04")
                self.model.setHeaderData(0, Qt.Horizontal, QVariant("xuhao"))
                self.model.setHeaderData(1, Qt.Horizontal, QVariant("content1"))
                self.model.setHeaderData(2, Qt.Horizontal, QVariant("content2"))
                # self.model.select()
                #
                # self.view = QTableView(self)
                # self.view.resize(300, 300)
                # self.view.setModel(self.model)
                # self.view.resizeColumnsToContents()
                # self.view = QTableView(self.tab_3)
                # 基类是qmainwindow是用这个方法添加
                # self.setCentralWidget(self.view)

                query = QSqlQuery()
                # self.queryModel=QSqlQueryModel(self)
                # # self.recordQuery(0)
                self.tableview=QTableView(self.tab_3)
                # self.tableView.setModel(self.queryModel)

                query.exec_("select * from R04")
                if query.next():
                    print(query.value(1),query.value(2))
                    # print("query search")
                glb_query_tag == 0










class TcpThread(QThread):

    trigger = pyqtSignal(str)
    trigger1=pyqtSignal(str)

    def __int__(self):
        super(TcpThread, self).__init__()







    # 分离出一个函数用于循环tcp 重连

    def doConnect(self):
        with open('IP_Address.txt', 'r')as fn:
            s1 = fn.read()
        host = ''.join(re.findall(r'IP:(.*)', s1))  # 服务器IP地址取出来，并转成字符串
        port = int(''.join(re.findall(r'PORT:(.*)', s1)))  # 从文本中提取端口号转成整数
        BUFFSIZE = int(''.join(re.findall(r'BUFFSIZE:(.*)', s1)))  # 从文本中提取字符长度转整数
        ADDRESS = (''.join(re.findall(r'ADDRESS:(.*)', s1)))  # 从文本中提取字符长度转整数
        LENGTH = (''.join(re.findall(r'LENGTH:(.*)', s1)))  # 从文本中提取字符长度转整数
        print(host,port,ADDRESS,LENGTH)

        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            sock.connect((host, port))
            print("connect is ok")
        except:
            print('fail to setup socket connection')
        return sock, BUFFSIZE,ADDRESS,LENGTH

    # 发送数据

    def tcpClient(self):
        global glb_tag
        glb_tag=0
        global glb_data

        sockLocal, BUFFSIZE,ADDRESS,LENGTH = self.doConnect()


        while True:

            # 等于0，那么一直查询
            if glb_tag==0:
                self.data1 = '500000FF03FF000018000004010000D*00'

                self.data = self.data1+ADDRESS+LENGTH
                # print(self.data)


                try:
                    sockLocal.send(self.data.encode())
                except socket.error:
                    print("\r\nsocket error,do reconnect ")
                    sockLocal.close()
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()
                except:
                    sockLocal.close()
                    print("other error")
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()

                time.sleep(0.5)
                str1 = sockLocal.recv(BUFFSIZE).decode()


                if str1[22:26]=="0001":

                    glb_tag=1
            # 等于1，写入并等待结果

            elif glb_tag==1:
                # 发射信号，连接到datadisplay

                self.trigger.emit(str1)



                self.data1 = '500000FF03FF00001C000014010000D*00499900010001'
                try:
                    sockLocal.send(self.data1.encode())

                except socket.error:
                    print("\r\nsocket error,do reconnect ")
                    sockLocal.close()
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()
                except:
                    sockLocal.close()
                    print("other error")
                    time.sleep(1)

                    sockLocal, BUFFSIZE = self.doConnect()

                time.sleep(0.3)
                str1 = sockLocal.recv(BUFFSIZE).decode()
                # print(glb_str1[16:18])
                # 如果返回的数据长度是对应长度，那么不再写
                if str1[16:18]=="04":
                    glb_tag = 0
            # 等于2，PLC 数据查询
            elif glb_tag==2:
                self.data1='500000FF03FF000018000004010000D*00'
                # 先转成字符串，然后填满四位
                self.data2=("".join(re.findall("[D|d](\d+)",glb_data)))
                self.data3=self.data2.zfill(4)
                # print(glb_data+str(12))
                # print(self.data2)
                self.data=self.data1+self.data3+"0010"
                # print(self.data)

                try:
                    sockLocal.send(self.data.encode())

                except socket.error:
                    print("\r\nsocket error,do reconnect ")
                    sockLocal.close()
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()
                except:
                    sockLocal.close()
                    print("other error")
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()


                time.sleep(0.3)
                glb_tag = 0
                str1 = sockLocal.recv(BUFFSIZE).decode()
                # print(glb_str1)
                self.trigger1.emit(str1)
            #     PLC数据写入
            elif glb_tag==3:
                # print("li")
                self.data1='500000FF03FF00001C000014010000D*00'
                # 先转成字符串，然后填满四位
                self.data2=("".join(re.findall("(\d+)=",glb_data))).zfill(4)
                self.data3=("".join(re.findall("=(\d+)",glb_data))).zfill(4)
                # print(hex(int(self.data3)))
                # 先转成16进制，然后匹配出来，舍掉0x，然后转成字符串，然后填满0
                self.data4=("".join(re.findall("0x(\w+)",hex(int(self.data3))))).zfill(4)

                self.data=self.data1+self.data2+"0001"+self.data4
                # print(self.data)

                try:
                    sockLocal.send(self.data.encode())

                except socket.error:
                    print("\r\nsocket error,do reconnect ")
                    sockLocal.close()
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()
                except:
                    sockLocal.close()
                    print("other error")
                    time.sleep(1)
                    sockLocal, BUFFSIZE = self.doConnect()


                time.sleep(0.3)
                str1 = sockLocal.recv(BUFFSIZE).decode()
                glb_tag = 2
                # glb_str1 = sockLocal.recv(BUFFSIZE).decode()








    def run(self):
        #必须从run开始启动程序

        # global data


        self.tcpClient()





class MainWindow(QMainWindow,Ui_MainWindow):


    chaxunSignal=pyqtSignal(str)
    txtSignal=pyqtSignal(str)
    excelSignal=pyqtSignal(list)

    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)

        self.list=[]
        self.setupUi(self)
        self.initUI()


    def initUI(self):
        global glb_data
        global glb_query_tag
        glb_query_tag=0
        global glb_row1
        glb_row1=0

        self.pushButton1.clicked.connect(self.dataread)
        self.pushButton2_2.clicked.connect(self.datawrite)
        self.pushButton4.clicked.connect(self.sqlquery)
        # self.pushButton5.clicked.connect(lambda glb_query_tag )

        self.thread1 = TcpThread()

        self.thread1.trigger.connect(self.firstdata)
        self.thread1.trigger1.connect(self.secondedata)

        self.thread1.start()
        self.thread2 = SqlThread()
        self.thread2.start()




    def sqlquery(self):
        global glb_query_tag
        global glb_query_data
        glb_query_data=self.lineEdit2.text()
        glb_query_tag=2

    def dataread(self):
        global glb_tag
        global glb_data
        glb_data = self.lineEdit1.text()
        glb_tag=2
    def datawrite(self):
        global glb_tag
        global glb_data

        glb_data = self.lineEdit1.text()
        glb_tag=3
    def secondedata(self,str):

        # print(isinstance(str1,str))
        # 16进制字符串转成10进制
        str1=int(str[22:26],16)
        str2=" ".join(re.findall("[D|d](\d+)",glb_data))
        # print(str2,glb_data)

        str3="D{0}={1}".format(str2,str1)

        self.listWidget1.addItem(str3)

        # print("hello")



    def firstdata(self,str):
        global glb_row1
        global glb_str1
        global glb_query_tag
        glb_query_tag=1
        glb_str1=str

        time1 = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        # print(str)


        # print(data)
        item0=QTableWidgetItem(str[22:26])
        item1 = QTableWidgetItem(str[26:90])
        item2 = QTableWidgetItem(str[90:130])
        item3 = QTableWidgetItem( str[130:134])
        item4 = QTableWidgetItem(str[134:138])
        self.tableWidget1.setItem(glb_row1,1,item1)
        self.tableWidget1.setItem(glb_row1, 2, item2)
        self.tableWidget1.setItem(glb_row1, 3, item3)
        self.tableWidget1.setItem(glb_row1, 4, item4)


        # self.tabWidget1.setEditTriggers(QAbstractItemView.NoEditTriggers)
        glb_row1 += 1

        if glb_row1>=10:
            glb_row1=0
        # for i in range(10):
        #     # a="item{}".format(i)
        #     # print (a)
        #     item1.setBackground(QBrush(QColor(128,128,128)))
        # print(1)



    # 操作Excel

    # def excel():
    #     global local_day1, lglb_data1, ldata2, ltime1
    #     local_day = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    #
    #     local_time = time.strftime('%H-%M-%S', time.localtime(time.time()))
    #     name = u"马夸特" + local_day + ".xls"
    #     style = xlwt.XFStyle()
    #     # 创建字体
    #     font = xlwt.Font()
    #     # 指定字体名字
    #     font.name = 'Times New Roman'
    #     # 字体加粗
    #     font.bold = False
    #     # 将该font设定为style的字体
    #     style.font = font
    #
    #     try:
    #         rbook = xlrd.open_workbook(name)
    #         # table = rbook.sheets()[0]
    #         # 如果日期变化，说明过了一天，那么清空列表，新建表格
    #         table = rbook.sheets()[0]
    #         nrows = table.nrows  # 获取表的行数
    #         wbook = copy(rbook)
    #         # 获取sheet对象，通过sheet_by_index()获取的sheet对象没有write()方法
    #         table = wbook.get_sheet(0)
    #         # print(lglb_data1)
    #         # print(ltime1)
    #         if len(lglb_data1) > 0:
    #             table.write(nrows, 0, nrows, style)
    #             table.write(nrows, 1, lglb_data1[0], style)
    #             table.write(nrows, 2, ldata2[0], style)
    #             table.write(nrows, 3, ltime1[0], style)
    #
    #             try:
    #                 wbook.save(name)
    #                 del lglb_data1[0]
    #                 del ldata2[0]
    #                 del ltime1[0]
    #                 return excel()
    #             except Exception as e:
    #                 print("save failure")
    #         # print (lglb_data1)
    #
    #
    #     except Exception as e:
    #         book = xlwt.Workbook()
    #         # 新建工作表,可对同一个单元格重复操作
    #         table = book.add_sheet(u'塌陷值', cell_overwrite_ok=True)
    #         table.write(0, 0, u"序号", style)
    #         table.write(0, 1, u"塌陷值1", style)
    #         table.write(0, 2, u"塌陷值2", style)
    #         table.write(0, 3, u'时间', style)
    #
    #         try:
    #             book.save(name)
    #         except Exception as e:
    #             print("NG")
    #         time.sleep(1)
    #         return excel()
    #
    #         # 保存文件,不支持xlsx格式
    #         # 初始化样式
    #
    #         # 写入到文件时使用该样式
    #         # data=tcpserver()
    #         # table.write(0, 0, "序号", style)
    #
    #     # 添加sheet页
    #     # wbook.add_sheet('sheetnnn2', cell_overwrite_ok=True)
    #     # 利用保存时同名覆盖达到修改excel文件的目的,注意未被修改的内容保持不变



        # while True:
        #     try:
        #         msg = str(time.time())
        #         sockLocal.send(msg)
        #         print
        #         "send msg ok : ", msg
        #         print
        #         "recv data :", sockLocal.recv(BUFFSIZE)
        #     except socket.error:
        #         print
        #         "\r\nsocket error,do reconnect "
        #         time.sleep(3)
        #         sockLocal = self.doConnect()
        #     except:
        #         print
        #         '\r\nother error occur '
        #         time.sleep(3)
        #     time.sleep(1)

    # def tcpclient(self):
    #     with open('IP_Address.txt', 'r')as fn:
    #         s1 = fn.read()
    #     HOST = ''.join(re.findall(r'IP:(.*)', s1))  # 服务器IP地址取出来，并转成字符串
    #
    #     PORT = int(''.join(re.findall(r'PORT:(.*)', s1)))  # 从文本中提取端口号转成整数
    #     BUFFSIZE=int(''.join(re.findall(r'BUFFSIZE:(.*)', s1)) ) # 从文本中提取字符长度转整数
    #     # BUFFSIZE = 2048
    #     # print(HOST,PORT,BUFFSIZE)
    #
    #     tctimeClient = socket(AF_INET, SOCK_STREAM)
    #
    #     try:
    #         tctimeClient.connect((HOST, PORT))
    #
    #     except socket.error:
    #         print ('fail to setup socket connection')
    #     sock = tctimeClient
    #     # tctimeClient.close()
    #
    #     while True:
    #         data = input(">")
    #         if not data:
    #             break
    #         tctimeClient.send(data.encode())
    #         data = tctimeClient.recv(BUFFSIZE).decode()
    #         if not data:
    #             break
    #         print(data)
    #     tctimeClient.close()

if __name__ == "__main__":



    app = QApplication(sys.argv)
    mainWindow = MainWindow()
    time1 = time.strftime("%Y/%m/%d/%H/%M/%S", time.localtime())
    # print(time1)
    # print(isinstance(time1,str))


    mainWindow.setWindowTitle("三菱上位软件")
    mainWindow.show()


    sys.exit(app.exec_())


