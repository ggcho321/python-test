from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorcode import *
from PyQt5.QtTest import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom class 입니다")

        ######## Event loop 모음
        self.login_event_loop = None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        ###############################

        ####### 계좌관련변수
        self.use_money = 0
        self.use_money_percent = 0.5
        ###########################

        ######## 스크린번호 모음
        self.screen_my_info = "2000"
        self.screen_calculation_stock = "4000"
        self.screen_real_stock = "5000"  # 종목별 할당할 스크린 번호
        self.screen_meme_stock = "6000"  # 종목별 할당할 주문용 스크린 번호
        self.screen_start_stop_real = "1000"  # 실시간 주문 스크린 번호

        ######## 변수모음
        self.account_num = None
        self.account_stock_dict = {}      #예수금 가져오는 것
        self.not_account_stock_dict = {}  #계좌평가 잔고 내역 요청
        self.portfolio_stock_dict = {}    #미체결 요청
        self.jango_dict = {}
        ###############################

        ###### 종목분석용
        self.calcul_data = []
        ##############################

        #########함수실행
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()      # 예수금 가져오는 곳
        self.detail_account_mystock()   # 계좌평가 잔액 요청
        self.not_concluded_account()    # 미체결 요청
        self.calculator_fnc()           # 종목 분석용, 임시용으로 실행

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  #응용프로그램 제어, 경로지정 (레지스트리에 등록)

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  #API의 login_slots을 만듦
        self.OnReceiveTrData.connect(self.trdata_slot) #Tr data를 받는 slot, 예수금 등

    def login_slot(self, errCode):   #event_slots login
        print(errors(errCode))       # API error 처리
        self.login_event_loop.exit()    #login 실행 후 신호 끊음

    def signal_login_commConnect(self):    #시그날 받음 (연결), login_event_loop 시도
            self.dynamicCall("CommConnect()")

            self.login_event_loop = QEventLoop()
            self.login_event_loop.exec()     #login 실행

    def get_account_info(self):  #계좌번호 가져오기
        account_list = self.dynamicCall("GetLogininfo(String)", "ACCNO") #계좌번호

        self.account_num = account_list.split(';')[0] # ['2020202':'']--> ['2020202']

        print("나의 보유 계좌번호 %s" % self.account_num)

    def detail_account_info(self):  #TR 목록에서 인적 사항요청
        print("계좌평가현황을 요청하는 부분")

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "상장폐지조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가현황요청", "OPW00004",  "0",  self.screen_my_info)

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec()

    def detail_account_mystock(self, sPrevNext = "0"):
        print("계좌평가 잔고내역 요청하기 연속조회 %s" % sPrevNext)

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내용요청", "OPW00018", sPrevNext, self.screen_my_info)

        #self.detail_account_info_event_loop2 = QEventLoop()
        self.detail_account_info_event_loop.exec()

    def not_concluded_account(self, sPrevNext = "0"):
        print("미체결 요청")

        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString,QString, int, QString)", "실시간미체결요청", "opt10075", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr 요청을 받은 구역 및 슬롯
        :param sScrNo: 스크린 번호
        :param sRQName: 요청 이름
        :param sTrCode: 요청 id tr 코드
        :param sRecordName: 사용 안함
        :param sPrevNext: 다음페이지
        :return:
        '''

        if sRQName == "계좌평가현황요청":
            deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예수금")
            print("예수금 %s" % int(deposit))

            self.use_money = int(deposit) * self.use_money_percent  #50% 투자
            self.use_money = self.use_money / 4     # 12.5% 투자

            ok_deposit = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "예탁자산평가액")
            print("예탁자산평가액 %s" % int(ok_deposit))

            self.detail_account_info_event_loop.exit()

        elif sRQName == "계좌평가잔고내용요청":

            total_buy_money = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)

            print("총매입금액: %s" % total_buy_money_result)

            total_profit_loss_rate = self.dynamicCall("GetCommData(String, String, int, String)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)

            print("총수익률(%s): %s" % ("%", total_profit_loss_rate_result))

            self.detail_account_info_event_loop.exit()
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)  #보유 종목 수, multi-data 조회
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                profit_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")

                #print("현재가: %s" % current_price)

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code: {}})

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                profit_rate = float(profit_rate.strip())
                current_price = int(current_price.strip())
                total_buy_price = int(total_buy_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": profit_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매매금액": total_buy_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt += 1
            print("계좌에 있는 종목 %s" % self.account_stock_dict)
            print("계좌에 보유종목 카운트 %s" % cnt)

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태")
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                self.not_account_stock_dict[order_no].update({"종목코드": code})
                self.not_account_stock_dict[order_no].update({"종목명": code_nm})
                self.not_account_stock_dict[order_no].update({"주문번호": order_no})
                self.not_account_stock_dict[order_no].update({"주문상태": order_status})
                self.not_account_stock_dict[order_no].update({"주문수량": order_quantity})
                self.not_account_stock_dict[order_no].update({"주문가격": order_price})
                self.not_account_stock_dict[order_no].update({"주문구분": order_gubun})
                self.not_account_stock_dict[order_no].update({"미체결수량": not_quantity})
                self.not_account_stock_dict[order_no].update({"체결량": ok_quantity})

                print("미체결 종목 : %s" % self.not_account_stock_dict[order_no])

            self.detail_account_info_event_loop.exit()

        if sRQName == "주식일봉차트조회":
             code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "종목코드")
             code = code.strip()

             print("%s 일봉데이터 요청" % code)

             cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
             print("데이터 일수 %s" % cnt)

             # data = self.dynamicCall("GetCommDataEx(QString, QString", sTrCode, sRQName)
             # [['', '현재가', '거래량', '거래대금', '날짜', '시가', '고가', '저가', ''] ['', '현재가', '거래량', '거래대금', '']
             # 한번 조회하면 600일 데이터를 받음
             # 600일치를 가져오는 함수 GetComEx []

             for i in range(cnt):  # i = 0 - 599, cnt = 599
                 data = []

                 current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "현재가")
                 value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                          "거래량")
                 trading_value = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "거래대금")
                 date = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                         "일자")
                 start_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                "시가")
                 high_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "고가")
                 low_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "저가")

                 data.append("")
                 data.append(current_price.strip())
                 data.append(value.strip())
                 data.append(trading_value.strip())
                 data.append(date.strip())
                 data.append(start_price.strip())
                 data.append(high_price.strip())
                 data.append(low_price.strip())
                 data.append("")

                 self.calcul_data.append(data.copy())

             #print(self.calcul_data) #데이터 값

             if sPrevNext == "2":
                 self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)

             else:
                 self.calculator_event_loop.exit()



        
    def get_code_list_by_market(self, market_code):  #종목코드를 받음

        code_list = self.dynamicCall("GetCodeListByMarket(QString)", market_code) #"11111"; "22222"; 코스닥:10
        code_list = code_list.split(";")[:-1]
        return code_list

    def calculator_fnc(self):
        '''
        종목분석 실행용 함수
        :param self:
        :return:
        '''

        code_list = self.get_code_list_by_market("10")
        print("코스닥 갯수 %s" % len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)
            print("%s / %s : KOSDAQ Stock Code : %s is updating" % (idx+1, len(code_list), code))
            self.day_kiwoom_db(code=code)

    def day_kiwoom_db(self, code=None,date = None, sPrevNext = "0"):
        QTest.qWait(3600)

        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString,QString, int, QString)", "주식일봉차트조회", "opt10081", sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec()

