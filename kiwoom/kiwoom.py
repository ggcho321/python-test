from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorcode import *

class Kiwoom(QAxWidget):
    OnEventConnect = pyqtSignal(int, name='OnEventConnect')

    def __init__(self):
        super().__init__()
        print("Kiwoom class 입니다")

        ######## Event loop 모음
        self.login_event_loop = None
        ############################

        ######## 변수모음
        self.account_num = None
        #######################

        self.login_event_loop = 0

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.OnEventConnect.connect(self.login_slot)
        self.get_account_info()


    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        self.OnEventConnect.emit(0)

    def login_slot(self, errCode):
        print(errors(errCode))

        self.login_event_loop.exit()

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLogininfo(String)", "ACCNO")

        self.account_num = account_list.split(';')[0]

        print("나의 보유 계좌번호 %s" % self.account_num)

'''from PyQt5.QAxContainer import *
from PyQt5.QtCore import *

class Kiwoom(QObject):
    def __init__(self):
        super().__init__()
        self.ax = QAxWidget()
        self.ax.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        self.OnEventConnect = pyqtSignal(int)
        self.OnEventConnect.connect(self.login_slot)
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()

        # Other logic here

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
            # self.OnEventConnect(self.login_slot)
            self.OnEventConnect.emit(0)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")

    def login_slot(self, errCode):
        print(errCode)

    # Other methods here
class Kiwoom(QObject, QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom class 입니다")

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.OnEventConnect = pyqtSignal(int)
        self.OnEventConnect.connect(self.login_slot)

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_slots(self):
        #self.OnEventConnect(self.login_slot)
        self.OnEventConnect.emit(0)

    def login_slot(self, errCode):
        print(errCode)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")'''
