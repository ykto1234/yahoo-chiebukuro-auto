import PySimpleGUI as sg
import os
import threading
import sys
import time
import traceback

import setting_read
import spread_sheet
from search_item import SearchItem
import scraip_yahoo

import mylogger
import datetime

# ログの定義
logger = mylogger.setup_logger(__name__)

# 監視フラグ
RUNNING_FLG = False


class MainForm:
    def __init__(self):

        global RUNNING_FLG

        # デザインテーマの設定
        sg.theme('BlueMono')

        # ウィンドウの部品とレイアウト
        main_layout = [
            [sg.Text('Yahoo知恵袋のURLを取得し、スプレッドシートに自動出力します')],
            [sg.Text('処理状態：', size=(11, 1)), sg.Text('停止中', key='process_status', text_color='#191970', size=(11, 1))],
            [sg.Text(size=(40, 1), justification='center', text_color='#191970', key='message_text1'), sg.Button('実行開始', key='execute_button')]
        ]

        # ウィンドウの生成
        self.window = sg.Window('Yahoo知恵袋検索', main_layout)

        # イベントループ
        while True:
            event, values = self.window.read(timeout=100, timeout_key='-TIMEOUT-')

            if event == sg.WIN_CLOSED:
                # ウィンドウのXボタンを押したときの処理
                break

            elif event == 'execute_button':
                if not RUNNING_FLG:
                    RUNNING_FLG = True
                    self.window['execute_button'].update(disabled=True)
                    self.update_text('process_status', '実行中．．．')
                    t1 = threading.Thread(target=yahoo_chiebukuro_worker)
                    # スレッドをデーモン化する
                    t1.setDaemon(True)
                    t1.start()
                else:
                    RUNNING_FLG = False
                    self.update_text('process_status', '停止中')

            elif event == '-TIMEOUT-':
                if RUNNING_FLG:
                    if not self.window['execute_button'].Disabled:
                        self.window['execute_button'].update(disabled=True)
                        self.update_text('process_status', '実行中．．．')
                else:
                    if self.window['execute_button'].Disabled:
                        self.window['execute_button'].update(disabled=False)
                        self.update_text('process_status', '停止中')

            elif event == 'ERROR':
                sg.popup_error('実行に失敗しました', title='エラー', button_color=('#f00', '#ccc'))

            elif event == 'SUCCESS':
                sg.Popup('実行が完了しました', title='実行結果')

        self.window.close()

    def enable_button(self, key):
        self.window[key].update(disabled=False)

    def disable_button(self, key):
        self.window[key].update(disabled=True)

    def update_text(self, key, message):
        self.window[key].update(message)


def yahoo_chiebukuro_worker():
    global RUNNING_FLG

    try:
        # INIファイル読み込み
        gspread_info_dic = setting_read.read_config('GSPREAD_SHEET')
        AUTH_KEY_PATH = gspread_info_dic['AUTH_KEY_PATH']
        SPREAD_SHEET_KEY = gspread_info_dic['SPREAD_SHEET_KEY']
        INPUT_SHEETNAME = gspread_info_dic['INPUT_SHEETNAME']
        OUTPUT_SHEETNAME = gspread_info_dic['OUTPUT_SHEETNAME']

        # スプレッドシートから検索キーワードを取得
        logger.debug('スプレッドシートから検索キーワードを取得開始')
        searchinfo_workbook = None
        searchinfo_workbook = spread_sheet.connect_gspread_workbook(AUTH_KEY_PATH, SPREAD_SHEET_KEY)
        search_keyword_list = []
        search_keyword_list = spread_sheet.read_gspread_sheet_from_workbook(searchinfo_workbook, INPUT_SHEETNAME)
        search_keyword_list = search_keyword_list[1:]

        logger.debug('スプレッドシートから検索キーワードを取得完了')

        if not len(search_keyword_list):
            logger.info('検索キーワードが0件のため、終了します')
            return

        # Yahoo知恵袋検索結果取得
        logger.debug('Yahoo知恵袋検索開始')
        TARGET_URL = 'https://search.yahoo.co.jp/'
        driver = scraip_yahoo.create_driver()

        item_list = []
        for search_keyword in search_keyword_list:
            _item = SearchItem()
            _item.search_keyword = search_keyword[0]
            _item = scraip_yahoo.get_chiebukuro_url(driver, TARGET_URL, _item)
            item_list.append(_item)
        driver.quit()
        logger.debug('Yahoo知恵袋検索完了')

        OUT_DATA_COL = 3
        HEADER_COL = 1
        RANKING_COL = 2

        # シート名存在チェック
        now = datetime.datetime.now()
        this_month_str = now.strftime('%Y/%m')
        TARGET_SHEET_NAME = OUTPUT_SHEETNAME + "_" + this_month_str
        if not spread_sheet.is_exist_worksheet(searchinfo_workbook, TARGET_SHEET_NAME):
            # シートが存在しない場合、新規作成
            logger.debug('月のワークシートが存在しないため、新規作成')
            spread_sheet.create_new_worksheet(searchinfo_workbook, TARGET_SHEET_NAME)
            OUT_DATA_COL = 3

        logger.debug('スプレッドシートに検索結果書き込み開始')

        chiebukuro_sheet_list = []
        chiebukuro_sheet_list = spread_sheet.read_gspread_sheet_from_workbook(searchinfo_workbook, TARGET_SHEET_NAME)

        if len(chiebukuro_sheet_list) > 0:
            OUT_DATA_COL = len(chiebukuro_sheet_list[0]) + 1

        # 日付の記入
        today_str = now.strftime('%Y/%m/%d')
        output_worskheet = searchinfo_workbook.worksheet(TARGET_SHEET_NAME)
        spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=1, cell_col=OUT_DATA_COL, update_value=today_str)

        reload_flg = False
        for _item in item_list:
            ROW_NUM = 2
            is_exist_flg = False

            if reload_flg:
                chiebukuro_sheet_list = spread_sheet.read_gspread_sheet_from_workbook(searchinfo_workbook, TARGET_SHEET_NAME)
                reload_flg = False

            for data_row in chiebukuro_sheet_list[1:]:
                if not data_row[0]:
                    ROW_NUM += 1
                    continue

                if _item.search_keyword == data_row[0]:
                    spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM, cell_col=OUT_DATA_COL, update_value=_item.search_url_no1)
                    spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 1, cell_col=OUT_DATA_COL, update_value=_item.search_url_no2)
                    spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 2, cell_col=OUT_DATA_COL, update_value=_item.search_url_no3)
                    is_exist_flg = True
                    break
                else:
                    ROW_NUM += 1

            if not is_exist_flg:
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM, cell_col=HEADER_COL, update_value=_item.search_keyword)
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM, cell_col=RANKING_COL, update_value='１位')
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 1, cell_col=RANKING_COL, update_value='２位')
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 2, cell_col=RANKING_COL, update_value='３位')
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM, cell_col=OUT_DATA_COL, update_value=_item.search_url_no1)
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 1, cell_col=OUT_DATA_COL, update_value=_item.search_url_no2)
                spread_sheet.update_gspread_sheet(worksheet=output_worskheet, cell_row=ROW_NUM + 2, cell_col=OUT_DATA_COL, update_value=_item.search_url_no3)
                reload_flg = True

        logger.debug('スプレッドシートに検索結果書き込み完了')

    except Exception as err:
        logger.error(err)
        logger.error(traceback.format_exc())

    finally:
        RUNNING_FLG = False
        if driver:
            driver.quit()


def expexpiration_date_check():
    import datetime
    now = datetime.datetime.now()
    expexpiration_datetime = now.replace(month=4, day=15, hour=12, minute=0, second=0, microsecond=0)
    logger.info("有効期限：" + str(expexpiration_datetime))
    if now < expexpiration_datetime:
        return True
    else:
        return False


# プログラム実行部分
if __name__ == "__main__":
    logger.info('プログラム起動開始')

    # # 有効期限チェック
    # if not (expexpiration_date_check()):
    #     logger.info("有効期限切れため、プログラム起動終了")
    #     sys.exit(0)

    args = sys.argv
    if len(args) > 1:
        logger.info('引数あり。引数：' + args[1])
        if args[1] == "-s":
            # 引数がある場合、サイレントでクローリングを実行
            logger.info('バックグラウンドで実行開始')
            yahoo_chiebukuro_worker()
            logger.info('バックグラウンドで実行完了')
            pass
        else:
            logger.info('引数が正しくありません。')
    else:
        # 引数なしの場合
        app = MainForm()