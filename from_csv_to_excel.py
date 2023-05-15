import logging
import time, pythoncom, re, joblib, configparser, win32com.client, os, traceback
from csv_read import scan_data

# Create Child_logger
logger = logging.getLogger('logging_settings').getChild('from_csv_to_excel')

# Read iniFile
inifile = configparser.ConfigParser()
inifile.read('settings.ini','UTF-8')


# Function For Search
def search_rectangle(rectangle, keyword):
    rows = len(rectangle)
    cols = len(rectangle[0])
    res = [[row_num,col_num] \
            for row_num in range(1,rows+1) \
            for col_num in range(1,cols+1) \
            if re.sub("[\u3000 \t \u200b]","",str(rectangle[row_num-1][col_num-1])) == keyword]
    return res

  
def from_csv_to_excel(self):
    
    # Read iniFile_Option
    excel_setting = joblib.load('excel_file_path.txt')
    FILE_PATH = excel_setting[0]
    FILE_NAME = excel_setting[1]
    TEMPLATE_SHEET = inifile['EXCEL']['TEMPLATE_SHEET']
    PRINT_SHEET = inifile['EXCEL']['PRINT_SHEET']
    SEARCH_START_CELL = inifile['EXCEL']['SEARCH_START_CELL']
    ADD_ROW = int(inifile['PYTHON']['ADD_ROW'])
    SEARCH_END_CELL = inifile['EXCEL']['SEARCH_END_CELL']
    SEARCH_KEYWORD, WRITE_WORD = scan_data()
    COLUMN_WIDTH = float(inifile['EXCEL']['COLUMN_WIDTH'])
    ROW_HEIGHT = float(inifile['EXCEL']['ROW_HEIGHT'])
    KEY_NAME_CSV = inifile['CSV']['KEY_NAME_CSV']
    KEY_NAME_EXCEL = inifile['EXCEL']['KEY_NAME_EXCEL']
    NAME_DIFFERENCE_ROW = int(inifile['EXCEL']['NAME_DIFFERENCE_ROW'])
    NAME_DIFFERENCE_COLUMN = int(inifile['EXCEL']['NAME_DIFFERENCE_COLUMN'])
    KEY_DATE_CSV = inifile['CSV']['KEY_DATE_CSV']
    KEY_YEAR_EXCEL = inifile['EXCEL']['KEY_YEAR_EXCEL']
    KEY_MONTH_EXCEL = inifile['EXCEL']['KEY_MONTH_EXCEL']
    YEAR_DIFFERENCE_ROW = int(inifile['EXCEL']['YEAR_DIFFERENCE_ROW'])
    YEAR_DIFFERENCE_COLUMN = int(inifile['EXCEL']['YEAR_DIFFERENCE_COLUMN'])
    MONTH_DIFFERENCE_ROW = int(inifile['EXCEL']['MONTH_DIFFERENCE_ROW'])
    MONTH_DIFFERENCE_COLUMN = int(inifile['EXCEL']['MONTH_DIFFERENCE_COLUMN'])
    EXCEL_MACRO_NAME = eval(inifile['EXCEL']['EXCEL_MACRO_NAME'])
    ROWS_DATA = int(inifile['EXCEL']['ROWS_DATA'])
    REPLASE_WORD = eval(inifile['PYTHON']['REPLASE_WORD'])
    
    # Get Now Time
    time1 = time.time()
    # Start_Button Off
    self.strat_btn['state'] = 'disabled'
    # COM Start
    pythoncom.CoInitialize()
    logger.info("Start CoInitialaze")
    try:
        self.prg_text.set("Excel起動中...")
        start_time = time.time()
        time1 = time.time()
        data_count = 1
        # Open Excel With Win32
        Excel = win32com.client.Dispatch('Excel.Application')
        Excel.Visible = False
        Excel.DisplayAlerts = False
        fullpath = os.path.join(FILE_PATH)
        wb = Excel.Workbooks.Open(Filename=fullpath)
        logger.info("Excel Launched")
        # Print_Sheet Delete
        wb.Worksheets(PRINT_SHEET).Delete()
        ws = wb.worksheets(TEMPLATE_SHEET)
        wb.Sheets.Add(Before=None, After=ws).Name = PRINT_SHEET
        ws2 = wb.worksheets(PRINT_SHEET)
        # Activate The Template_Sheet
        ws.Activate()
        # Set The Row_Height And The Column_Width of A Template_Sheet
        ws.Range(f"{SEARCH_START_CELL}:{SEARCH_END_CELL}").ColumnWidth = COLUMN_WIDTH
        ws.Range(f"1:{str(ROWS_DATA)}").RowHeight = ROW_HEIGHT
        # Activate The Print_Sheet
        ws2.Activate()
        # Set The Row_Height And The Column_Width of A Print_Sheet
        ws2.Range(f"{SEARCH_START_CELL}:{SEARCH_END_CELL}").ColumnWidth = COLUMN_WIDTH
        ws2.Range(f"1:5000").RowHeight = ROW_HEIGHT
        # List Search Scope Value
        rectangle = [[cell for cell in rows] for rows in ws.Range(f"{SEARCH_START_CELL}:{SEARCH_END_CELL}").Value]
        logger.info("Load Complete Setting Before Postting Processing")
        # Search Processing
        # Loop Processing For The Number Of Data Records
        for write_data in WRITE_WORD:
            i = 0
            # Activate The Template_Sheet
            ws.Activate()
            # Clear Values Of The Template_Sheet
            rows = len(rectangle)
            cols = len(rectangle[0])
            for row_num in range(1,rows+1):
                for col_num in range(1,cols+1):
                    if type(rectangle[row_num-1][col_num-1]) is int \
                        or type(rectangle[row_num-1][col_num-1]) is float:
                        ws.Cells(row_num,col_num).Value = ""
            # Loop Processing For The Number Of Kyewords
            for keydata in SEARCH_KEYWORD:
                try:
                    if keydata == KEY_NAME_CSV:
                        # Processing Name
                        keydata = KEY_NAME_EXCEL
                        # Search And Write
                        result = search_rectangle(rectangle, keydata)
                        write_cell = ws.Cells(int(result[0][0])+NAME_DIFFERENCE_ROW,\
                                            int(result[0][1])+NAME_DIFFERENCE_COLUMN)
                        write_cell.Value = write_data[i]

                        i += 1

                    elif keydata == KEY_DATE_CSV:
                        # Processing Year And Month
                        year = str(write_data[i])[2:4]
                        month = str(write_data[i])[4:6]
                        key_year = KEY_YEAR_EXCEL
                        key_month = KEY_MONTH_EXCEL
                        # Search And Write
                        # Year
                        result = search_rectangle(rectangle, key_year)
                        write_cell = ws.Cells(int(result[0][0])+int(YEAR_DIFFERENCE_ROW),\
                                            int(result[0][1])+YEAR_DIFFERENCE_COLUMN)
                        write_cell.Value = int(year)
                        # Month
                        result = search_rectangle(rectangle, key_month)
                        write_cell = ws.Cells(int(result[0][0])+MONTH_DIFFERENCE_ROW,\
                                            int(result[0][1])+MONTH_DIFFERENCE_COLUMN)
                        write_cell.Value = int(month)
                        i += 1
                    else:
                        # Actions Other Than The Above
                        # Processing Blank
                        keydata = re.sub("[\u3000 \t \u200b]","",keydata)
                        for rep in REPLASE_WORD:
                            keydata = keydata.replace(rep,'')
                        # Search And Write
                        result = search_rectangle(rectangle,keydata)
                        write_cell = ws.Cells(int(result[0][0])+1,int(result[0][1]))
                        # Processing Parentheses
                        if write_cell.Value == '(':
                            write_cell = ws.Cells(int(result[0][0]+1),int(result[0][1])+1)
                        write_cell.Value = float(write_data[i])
                        i += 1
                except:
                    i += 1
                    continue
            # Processing Copy
            # Default Scope Copy
            ws.Range(f"{SEARCH_START_CELL}:{SEARCH_END_CELL}").Copy()
            # Activate The Print_Sheet
            ws2.Activate()
            # Paste Cell Selection
            # Calculating Which Rows To Paste
            cells_num = 1 + ADD_ROW
            cells_pos = str(cells_num)
            ws2.Range(f"A{cells_pos}").Select()
            # Processing Paste
            ws2.Paste()
            # Incremental Calculation Of The Next Row To Paste
            ADD_ROW += int(SEARCH_END_CELL[len(SEARCH_END_CELL)-2:len(SEARCH_END_CELL)])
            # ProgressBar Updates
            add_prg = 390/len(WRITE_WORD)
            self.prgVal.set(self.prgVal.get()+add_prg)
            # ProgressLabel Updates
            time2 = time.time()
            elapsed_time = "{:.3g}".format(time2 - time1)
            self.prg_text.set(f"{data_count}/{len(WRITE_WORD)}件目処理時間: {elapsed_time}秒")
            time1 = time2
            data_count += 1
        logger.info("Finished Copy Loop")
        # Run Macro
        self.prg_text.set("行数、改ページ設定,置換処理中(約1分)...\n完了ボタンがオンになるまでプログラムを終了しないで下さい。")
        for macro in EXCEL_MACRO_NAME:
            Excel.Application.Run(FILE_NAME + '!' + macro)
        logger.info("Executed Excel Macro")
        wb.Save()
        #wb.Close()
        Excel.DisplayAlerts = True
        Excel.Visible = True
        #Excel.Application.Quit()
        logger.info("Excel Task Finished")
        self.close_btn['state'] = 'normal'
        self.close_btn['default'] = 'active'
        # ProgressBar Updates
        add_prg = 400
        self.prgVal.set(add_prg)
        # ProgressLabel Updates
        end_time = time.time()
        total_time = "{:.3g}".format(end_time - start_time)
        self.prg_text.set(f"全ての処理が完了しました。(所要時間合計：{total_time}秒)")

        # Releasing Coinitialize
        pythoncom.CoUninitialize()
        logger.info("Removed Coinitialize")
        logger.info("Finished Posting Processing")

    except FileNotFoundError:
        logger.error(traceback.format_exc())
        self.strat_btn['state'] = 'normal'
        self.text.set("設定したファイルパスにファイルが存在していません。")
        pythoncom.CoUninitialize()
        logger.info("Removed CoInitialixze")

    except KeyError:
        logger.error(traceback.format_exc())
        self.strat_btn['state'] = 'normal'
        self.text.set("指定したxlsm又はcsvファイルでのエラー\n開かれた状態になっているか、存在しません。")
        pythoncom.CoUninitialize()
        logger.info("Removed Coinitialize")

    except PermissionError:
        logger.error(traceback.format_exc())
        self.strat_btn['state'] = 'normal'
        self.text.set("指定したxlsm又はcsvファイルが開かれている為,\n処理が保存できませんでした。")
        pythoncom.CoUninitialize()
        logger.info("Removed Coinitialize")

    except Exception as e:
        logger.error(traceback.format_exc())
        try:
            wb.Close()
            Excel.Application.Quit()
        except Exception as ep:
            logger.error(ep)
        pythoncom.CoUninitialize()
        logger.info("Removed Coinitialize")
        self.strat_btn['state'] = 'normal'
        self.text.set(str(repr(e)))