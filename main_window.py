import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import joblib, threading, logging, configparser
from from_csv_to_excel import from_csv_to_excel

logger = logging.getLogger('logging_settings').getChild('main_window')
ch = logging.FileHandler(filename="logging.log")

# Read iniFile
inifile = configparser.ConfigParser()
inifile.read('settings.ini','UTF-8')

class Application(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master)
        # Main Frame Settings
        self.config(width=500,height=400,bg='#FFDEAD')
        # Main Frame Set
        self.pack(expand=1,fill=tk.BOTH)
        # Icon_image Path
        iconfile = '.\\rogo1.ico'
        # Icon_image Set
        self.master.iconbitmap(default=iconfile)
        # WindowSize Setting
        self.master.geometry("500x400+200+100")
        # WindowTitle Setting
        self.master.title("給与明細データ転記ツール")
        # Widgets Set
        self.create_widgets()
      
    
    def create_widgets(self):
        # Font_Style Settings
        self.style = ttk.Style()
        self.style.theme_use('vista')
        self.style.configure("office.TButton",font=("Sans","10","bold"),background="blue")
        # Create Warning_Label
        self.msg_lbl = tk.Label(self,text="※注意 : 全てのExcelファイル、CSVファイルを閉じてから開始を押してください。", \
                          background='#FFDEAD', \
                          foreground='red', \
                          font=('MSゴシック',"9",'bold'))
        self.msg_lbl.place(x=5,y=10)
        # Create CSVFile_Label
        self.csv_lbl = tk.StringVar()
        try:
            csv_path = joblib.load('csv_file_path.txt')
            self.csv_lbl.set(csv_path[0])
        except:
            pass
        self.csv_file_name_label = tk.Label(self,\
                                            textvariable=self.csv_lbl,\
                                            width=59,\
                                            anchor=tk.W,\
                                            font=("MSゴシック","10","bold"),\
                                            fg='#2E8B57',\
                                            background='#D3D3D3')
        self.csv_file_name_label.place(x=10,y=50)
        # Create Select_CSVFile_Dialog
        self.file_button = ttk.Button(self,\
                                      text='CSVファイル選択',\
                                      width=20,\
                                      command=self.csv_button_clicked,\
                                      default="active",\
                                      style="office.TButton")
        self.file_button.place(x=340,y=80)
        # Create ExcelFile_Label
        self.excel_lbl = tk.StringVar()
        try:
            excel_path = joblib.load('excel_file_path.txt')
            self.excel_lbl.set(excel_path[0])
        except:
            pass
        self.csv_file_name_label = tk.Label(self,\
                                            textvariable=self.excel_lbl,\
                                            width=59,\
                                            anchor=tk.W,\
                                            font=("MSゴシック","10","bold"),\
                                            fg='#2E8B57',\
                                            background='#D3D3D3')
        self.csv_file_name_label.place(x=10,y=120)
        # Create Radio_Button
        self.radio_value = tk.IntVar(value=1)
        self.radio_button_1 = tk.Radiobutton(self,\
                                             text='部門コード順',\
                                            command=self.radio_button_clicked,\
                                            variable=self.radio_value,\
                                            value=1)
        self.radio_button_1.place(x=50,y=150)

        self.radio_button_2 = tk.Radiobutton(self,\
                                             text='社員コード順',\
                                            command=self.radio_button_clicked,\
                                            variable=self.radio_value,\
                                            value=2)
        self.radio_button_2.place(x=50,y=170)
        joblib.dump([1],'radio_value.txt',compress=3)
        # Create Select_ExcelFile_Dialog
        self.file_button = ttk.Button(self,\
                                      text='Excelファイル選択',\
                                      width=20,\
                                      command=self.excel_button_clicked,\
                                      default="active",\
                                      style="office.TButton")
        self.file_button.place(x=340,y=150)
        # Create Setting_Button
        self.setting_btn = ttk.Button(self,\
                                      text="設定",\
                                        width=10,\
                                        command=self.setting_button_clicked,\
                                        default='active',\
                                        style="office.TButton")
        self.setting_btn.place(x=410,y=200)
        # Create Start_Button
        self.strat_btn = ttk.Button(self,\
                                    text="開始",\
                                    width=20,\
                                    command=self.button_clicked,\
                                    default='active',\
                                    style="office.TButton")
        self.strat_btn.place(x=180,y=220)
        # Create Close_Button
        self.close_btn = ttk.Button(self,\
                                    text="完了",\
                                    width=20,\
                                    command=self.close,\
                                    state='disable',\
                                    style="office.TButton")
        self.close_btn.place(x=180,y=260) # disabled
        # Create ProgressBar
        self.prgVal = tk.IntVar(value=0)
        self.prgbar = ttk.Progressbar(self, maximum=400 , length=400, variable=self.prgVal,mode="determinate")
        self.prgbar.place(x=55,y=300)
        # Create ProgressBar_Label
        self.prg_text = tk.StringVar()
        self.prg_lbl = tk.Label(self,\
                                textvariable=self.prg_text,\
                                width=68,\
                                fg='black',\
                                bg='#FFDEAD',\
                                font=("Sans","9",'bold'),\
                                anchor=tk.E)
        self.prg_lbl.place(x=10,y=330)
        # Create Error_Message_Label
        self.text = tk.StringVar()
        self.Err_lbl = tk.Label(self,\
                                textvariable=self.text,\
                                width=59,fg='#ff0000',\
                                bg='#FFDEAD',\
                                font=("Sans","10",'bold'),\
                                anchor=tk.CENTER)
        self.Err_lbl.place(x=10,y=350)
    

    def csv_button_clicked(self):
        logger.info("Change_CSV_File_Button Clicked")
        file_type = [("CSVファイル","*.csv;")]
        dir = inifile['DIALOG']['CSVPATH']
        fld = filedialog.askopenfilename(filetypes=file_type, initialdir=dir)
        slippath = fld.split('/')[:-1]
        fldpath = ''
        for i in slippath:
            fldpath += i + '/'

        if fld != "":
            logger.info("File Selected")
            file_path = [fld]
            joblib.dump(file_path,'csv_file_path.txt',compress=3)
            self.csv_lbl.set(fld)
            inifile['DIALOG']['CSVPATH'] = fldpath
            # Write Configration
            with open("settings.ini", "w", encoding="utf-8") as configfile:
                inifile.write(configfile, True)
        else:
            logger.info("Selected Cancel")

    def excel_button_clicked(self):
        logger.info("ExcelFileDialog_Button Clicked")
        file_type = [("Excelファイル","*.xlsm;")]
        dir = inifile['DIALOG']['EXCELPATH']
        fld = filedialog.askopenfilename(filetypes=file_type, initialdir=dir)
        flname = [item for item in fld.split(r'/') if item.endswith('.xlsm') is True]
        fld = fld.replace(r"/",r"\\")

        slippath = fld.split('\\')[:-1]
        fldpath = ''
        for i in slippath:
            fldpath += i + '/'

        if fld != "":
            logger.info("File Selected")
            file_path = [fld,flname[0]]
            joblib.dump(file_path,'excel_file_path.txt',compress=3)
            self.excel_lbl.set(fld)
            inifile['DIALOG']['EXCELPATH'] = fldpath
            # Write Configration
            with open("settings.ini", "w", encoding="utf-8") as configfile:
                inifile.write(configfile, True)
        else:
            logger.info("Selected Calcel")

    def radio_button_clicked(self):
        radiovalue = [self.radio_value.get()]
        joblib.dump(radiovalue,'radio_value.txt',compress=3)

    def button_clicked(self):
        logger.info("Start_Button Clicked")
        self.YesNoBox = messagebox.askquestion('csv転記処理の開始',\
                                          'csv転記処理を開始します。\nCSV・Excelファイルの指定は間違いありませんか？',\
                                            icon='warning')
        if self.YesNoBox == 'yes':
            logger.info("Yes Selected")
            # New Thread Create
            self.th = threading.Thread(target=from_csv_to_excel,args=(self,))
            # New Thread Start
            self.th.start()
        else:
            logger.info("No Selected")
    
    def close(self):
        # MainWindow Close
        self.master.quit()
        # Output LOG
        logger.info("Program Closed")
        # LogHandler Remove and LogFile Close 
        logger.removeHandler(ch)
        ch.close()

    
    def setting_button_clicked(self):
        # Setting_Button Disable
        self.setting_btn['state'] = 'disabled'

        # Inifile Read
        inifile = configparser.ConfigParser(comment_prefixes='/', allow_no_value=True)
        inifile.read('settings.ini','UTF-8')

        # Function With Updating
        def conf_update_click():
            
            logger.info("Update_Button Clicked")
            # Confirmation Message
            Start_YesNoBox = messagebox.askquestion('設定の更新',\
                                            '設定を更新します。\n続行しますか？',\
                                            icon='warning')
            if Start_YesNoBox == 'yes':
                logger.info("Yes Selected")
                inifile['CSV']['KEY_NAME_CSV'] = csv_name_text.get()
                inifile['CSV']['KEY_DATE_CSV'] = csv_ym_text.get()
                inifile['EXCEL']['TEMPLATE_SHEET'] = excel_tmp_name_text.get()
                inifile['EXCEL']['PRINT_SHEET'] = excel_pr_name_text.get()
                inifile['EXCEL']['SEARCH_START_CELL'] = excel_tmp_start_text.get()
                inifile['EXCEL']['SEARCH_END_CELL'] = excel_tmp_end_text.get()
                inifile['EXCEL']['COLUMN_WIDTH'] = excel_width_text.get()
                inifile['EXCEL']['ROW_HEIGHT'] = excel_height_text.get()
                inifile['EXCEL']['KEY_NAME_EXCEL'] = excel_name_text.get()
                inifile['EXCEL']['NAME_DIFFERENCE_ROW'] = excel_name_cell_row_text.get()
                inifile['EXCEL']['NAME_DIFFERENCE_COLUMN'] = excel_name_cell_col_text.get()
                inifile['EXCEL']['KEY_YEAR_EXCEL'] = excel_year_text.get()
                inifile['EXCEL']['KEY_MONTH_EXCEL'] = excel_month_text.get()
                inifile['EXCEL']['YEAR_DIFFERENCE_ROW'] = excel_year_cell_row_text.get()
                inifile['EXCEL']['YEAR_DIFFERENCE_COLUMN'] = excel_year_cell_col_text.get()
                inifile['EXCEL']['MONTH_DIFFERENCE_ROW'] = excel_month_cell_row_text.get()
                inifile['EXCEL']['MONTH_DIFFERENCE_COLUMN'] = excel_month_cell_col_text.get()
                inifile['EXCEL']['EXCEL_MACRO_NAME'] = macro_text.get()
                inifile['EXCEL']['ROWS_DATA'] = excel_rows_text.get()
                inifile['PYTHON']['ADD_ROW'] = add_rows_text.get()
                inifile['PYTHON']['REPLASE_WORD'] = rep_word_text.get()
                # Write Configration
                with open("settings.ini", "w", encoding="utf-8") as configfile:
                    inifile.write(configfile, True)
                # Confirmation Message
                End_YesNoBox = messagebox.askquestion('設定の変更','設定の変更が完了しました。\n設定画面を閉じますか？')
                if End_YesNoBox == 'yes':
                    self.strat_btn['state'] = 'disabled'
                    self.text.set("設定が変更されました。アプリを再起動して下さい。")
                    conf_window_close()
            else:
                logger.info("No Selected")

        # Function On Close
        def conf_window_close():
            # Setting_Button Enabled
            self.setting_btn['state'] = 'nomal'
            # Conf_Window Close
            conf_window.destroy()


        logger.info("Configration_Button Clicked")
        # Child Window With Main_window As Parent Create
        conf_window = tk.Toplevel(self.master)
        # Window Size
        conf_window.geometry("400x550+710+50")
        # Window Title
        conf_window.title("設定")
        # Canvas Create
        conf_canvas = tk.Canvas(conf_window)
        conf_canvas.pack(expand=1,fill=tk.BOTH)

        # Create CSV_Range
        conf_canvas.create_rectangle(10,10,390,80,outline="#68D868",fill="#8DF68D")
        csv_conf_lbl = tk.Label(conf_canvas,\
                                   text="CSV",\
                                    font=("MSゴシック","10","bold"),\
                                    fg="#000000",\
                                    bg="#D3D3D3")
        csv_conf_lbl.place(x=15,y=5)
        # CSV_Name_Field Setting
        csv_name_lbl = tk.Label(conf_canvas,text="・氏名のフィールド名",bg="#8DF68D")
        csv_name_lbl.place(x=15,y=30)
        csv_name_text = tk.StringVar()
        if inifile.has_option('CSV','KEY_NAME_CSV'):
            csv_name_text.set(inifile['CSV']['KEY_NAME_CSV'])
        csv_name_entry = tk.Entry(conf_canvas,textvariable=csv_name_text,width=30)
        csv_name_entry.place(x=130,y=30)
        # CSV_YearMonth_Field Setting
        csv_ym_lbl = tk.Label(conf_canvas,text="・年月のフィールド名",bg="#8DF68D")
        csv_ym_lbl.place(x=15,y=50)
        csv_ym_text = tk.StringVar()
        if inifile.has_option('CSV','KEY_DATE_CSV'):
            csv_ym_text.set(inifile['CSV']['KEY_DATE_CSV'])
        csv_ym_entry = tk.Entry(conf_canvas,textvariable=csv_ym_text,width=30)
        csv_ym_entry.place(x=130,y=50)

        # Create Excel_Range
        conf_canvas.create_rectangle(10,100,390,390,outline="#87CEFA",fill="#87CEFA")
        excel_setting_lbl = tk.Label(conf_canvas,\
                                     text="EXCEL",\
                                    font=("MSゴシック","10","bold"),\
                                    fg="#000000",\
                                    background="#D3D3D3")
        excel_setting_lbl.place(x=15,y=95)
        # Excel_Template_Sheet_Name Setting
        excel_tmp_name_lbl = tk.Label(conf_canvas,text="・テンプレートシート名",bg="#87CEFA")
        excel_tmp_name_lbl.place(x=15,y=120)
        excel_tmp_name_text = tk.StringVar()
        if inifile.has_option('EXCEL','TEMPLATE_SHEET'):
            excel_tmp_name_text.set(inifile['EXCEL']['TEMPLATE_SHEET'])
        excel_tmp_name_entry = tk.Entry(conf_canvas,textvariable=excel_tmp_name_text,width=30)
        excel_tmp_name_entry.place(x=130,y=120)
        # Excel_Print_Sheet_Name Setting
        excel_pr_name_lbl = tk.Label(conf_canvas,text="・印刷用シート名",bg="#87CEFA")
        excel_pr_name_lbl.place(x=15,y=140)
        excel_pr_name_text = tk.StringVar()
        if inifile.has_option('EXCEL','PRINT_SHEET'):
            excel_pr_name_text.set(inifile['EXCEL']['PRINT_SHEET'])
        excel_pr_name_entry = tk.Entry(conf_canvas,textvariable=excel_pr_name_text,width=30)
        excel_pr_name_entry.place(x=130,y=140)
        # Template_Sheet's Start_Cell and End_Cell Setting
        excel_tmp_start_lbl = tk.Label(conf_canvas,text="・テンプレートシートの     開始セル:",bg="#87CEFA")
        excel_tmp_start_lbl.place(x=15,y=160)
        excel_tmp_start_text = tk.StringVar()
        if inifile.has_option('EXCEL','SEARCH_START_CELL'):
            excel_tmp_start_text.set(inifile['EXCEL']['SEARCH_START_CELL'])
        excel_tmp_start_entry = tk.Entry(conf_canvas,textvariable=excel_tmp_start_text,width=8,justify='center')
        excel_tmp_start_entry.place(x=175,y=160)
        excel_tmp_end_lbl = tk.Label(conf_canvas,text="~終了セル:",bg="#87CEFA")
        excel_tmp_end_lbl.place(x=230,y=160)
        excel_tmp_end_text = tk.StringVar()
        if inifile.has_option('EXCEL','SEARCH_END_CELL'):
            excel_tmp_end_text.set(inifile['EXCEL']['SEARCH_END_CELL'])
        excel_tmp_end_entry = tk.Entry(conf_canvas,textvariable=excel_tmp_end_text,width=8,justify='center')
        excel_tmp_end_entry.place(x=290,y=160)
        # Excel_Template_Sheet_Rows Setting
        excel_rows_lbl = tk.Label(conf_canvas,text="・テンプレートの行数",bg="#87CEFA")
        excel_rows_lbl.place(x=15,y=180)
        excel_rows_text = tk.StringVar()
        if inifile.has_option('EXCEL','ROWS_DATA'):
            excel_rows_text.set(inifile['EXCEL']['ROWS_DATA'])
        excel_rows_entry = tk.Entry(conf_canvas,textvariable=excel_rows_text,width=5,justify='center')
        excel_rows_entry.place(x=130,y=180)
        # Cell's Width and Height Setting
        excel_width_lbl = tk.Label(conf_canvas,text="・セルの                                   幅:",bg="#87CEFA")
        excel_width_lbl.place(x=15,y=200)
        excel_width_text = tk.StringVar()
        if inifile.has_option('EXCEL','COLUMN_WIDTH'):
            excel_width_text.set(inifile['EXCEL']['COLUMN_WIDTH'])
        excel_width_entry = tk.Entry(conf_canvas,textvariable=excel_width_text,width=8,justify='center')
        excel_width_entry.place(x=175,y=200)
        excel_height_lbl = tk.Label(conf_canvas,text="           高さ:",bg="#87CEFA")
        excel_height_lbl.place(x=230,y=200)
        excel_height_text = tk.StringVar()
        if inifile.has_option('EXCEL','ROW_HEIGHT'):
            excel_height_text.set(inifile['EXCEL']['ROW_HEIGHT'])
        excel_height_entry = tk.Entry(conf_canvas,textvariable=excel_height_text,width=8,justify='center')
        excel_height_entry.place(x=290,y=200)
        # Excel_Name_Field Setting
        excel_name_lbl = tk.Label(conf_canvas,text="・氏名のフィールド名",bg="#87CEFA")
        excel_name_lbl.place(x=15,y=220)
        excel_name_text = tk.StringVar()
        if inifile.has_option('EXCEL','KEY_NAME_EXCEL'):
            excel_name_text.set(inifile['EXCEL']['KEY_NAME_EXCEL'])
        excel_name_entry = tk.Entry(conf_canvas,textvariable=excel_name_text,width=30)
        excel_name_entry.place(x=130,y=220)
        # Excel_Name_Cell Setting
        excel_name_cell_row_lbl = tk.Label(conf_canvas,text="・氏名の値を入れる氏名フィールドからの位置    行:",bg="#87CEFA")
        excel_name_cell_row_lbl.place(x=15,y=240)
        excel_name_cell_row_text = tk.StringVar()
        if inifile.has_option('EXCEL','NAME_DIFFERENCE_ROW'):
            excel_name_cell_row_text.set(inifile['EXCEL']['NAME_DIFFERENCE_ROW'])
        excel_name_cell_row_entry = tk.Entry(conf_canvas,textvariable=excel_name_cell_row_text,width=4,justify='center')
        excel_name_cell_row_entry.place(x=260,y=240)
        excel_name_cell_col_lbl = tk.Label(conf_canvas,text="  列:",bg="#87CEFA")
        excel_name_cell_col_lbl.place(x=285,y=240)
        excel_name_cell_col_text = tk.StringVar()
        if inifile.has_option('EXCEL','NAME_DIFFERENCE_COLUMN'):
            excel_name_cell_col_text.set(inifile['EXCEL']['NAME_DIFFERENCE_COLUMN'])
        excel_name_cell_col_entry = tk.Entry(conf_canvas,textvariable=excel_name_cell_col_text,width=4,justify='center')
        excel_name_cell_col_entry.place(x=310,y=240)
        # Excel_YearMonth_Field Setting
        excel_year_lbl = tk.Label(conf_canvas,text="・年月のフィールド名                年:",bg="#87CEFA")
        excel_year_lbl.place(x=15,y=260)
        excel_year_text = tk.StringVar()
        if inifile.has_option('EXCEL','KEY_YEAR_EXCEL'):
            excel_year_text.set(inifile['EXCEL']['KEY_YEAR_EXCEL'])
        excel_year_entry = tk.Entry(conf_canvas,textvariable=excel_year_text,width=10)
        excel_year_entry.place(x=175,y=260)
        excel_month_lbl = tk.Label(conf_canvas,text="       月:",bg="#87CEFA")
        excel_month_lbl.place(x=250,y=260)
        excel_month_text = tk.StringVar()
        if inifile.has_option('EXCEL','KEY_MONTH_EXCEL'):
            excel_month_text.set(inifile['EXCEL']['KEY_MONTH_EXCEL'])
        excel_month_entry = tk.Entry(conf_canvas,textvariable=excel_month_text,width=10)
        excel_month_entry.place(x=290,y=260)
        # Excel_Year_Cell Setting
        excel_year_cell_row_lbl = tk.Label(conf_canvas,text="・年の値を入れる年フィールドからの位置          行:",bg="#87CEFA")
        excel_year_cell_row_lbl.place(x=15,y=280)
        excel_year_cell_row_text = tk.StringVar()
        if inifile.has_option('EXCEL','YEAR_DIFFERENCE_ROW'):
            excel_year_cell_row_text.set(inifile['EXCEL']['YEAR_DIFFERENCE_ROW'])
        excel_year_cell_row_entry = tk.Entry(conf_canvas,textvariable=excel_year_cell_row_text,width=4,justify='center')
        excel_year_cell_row_entry.place(x=250,y=280)
        excel_year_cell_col_lbl = tk.Label(conf_canvas,text="  列:",bg="#87CEFA")
        excel_year_cell_col_lbl.place(x=275,y=280)
        excel_year_cell_col_text = tk.StringVar()
        if inifile.has_option('EXCEL','YEAR_DIFFERENCE_COLUMN'):
            excel_year_cell_col_text.set(inifile['EXCEL']['YEAR_DIFFERENCE_COLUMN'])
        excel_year_cell_col_entry = tk.Entry(conf_canvas,textvariable=excel_year_cell_col_text,width=4,justify='center')
        excel_year_cell_col_entry.place(x=300,y=280)
        # Excel_Month_Cell Setting
        excel_month_cell_row_lbl = tk.Label(conf_canvas,text="・月の値を入れる年フィールドからの位置          行:",bg="#87CEFA")
        excel_month_cell_row_lbl.place(x=15,y=300)
        excel_month_cell_row_text = tk.StringVar()
        if inifile.has_option('EXCEL','MONTH_DIFFERENCE_ROW'):
            excel_month_cell_row_text.set(inifile['EXCEL']['MONTH_DIFFERENCE_ROW'])
        excel_month_cell_row_entry = tk.Entry(conf_canvas,textvariable=excel_month_cell_row_text,width=4,justify='center')
        excel_month_cell_row_entry.place(x=250,y=300)
        excel_month_cell_col_lbl = tk.Label(conf_canvas,text="  列:",bg="#87CEFA")
        excel_month_cell_col_lbl.place(x=275,y=300)
        excel_month_cell_col_text = tk.StringVar()
        if inifile.has_option('EXCEL','MONTH_DIFFERENCE_COLUMN'):
            excel_month_cell_col_text.set(inifile['EXCEL']['MONTH_DIFFERENCE_COLUMN'])
        excel_month_cell_col_entry = tk.Entry(conf_canvas,textvariable=excel_month_cell_col_text,width=4,justify='center')
        excel_month_cell_col_entry.place(x=300,y=300)
        # Macro_List Setting
        macro_lbl = tk.Label(conf_canvas,\
                             text="・起動マクロのリスト\n※a,bのマクロを登録する場合、['a','b']のリスト形式で記入",\
                            bg="#87CEFA",anchor=tk.W,justify="left")
        macro_lbl.place(x=15,y=320)
        macro_text = tk.StringVar()
        if inifile.has_option('EXCEL','EXCEL_MACRO_NAME'):
            macro_text.set(inifile['EXCEL']['EXCEL_MACRO_NAME'])
        macro_entry = tk.Entry(conf_canvas,textvariable=macro_text,width=55)
        macro_entry.place(x=40,y=360)

        # Create Python_Range
        conf_canvas.create_rectangle(10,410,390,510,outline="#EE82EE",fill="#EE82EE")
        python_setting_lbl = tk.Label(conf_canvas,\
                                     text="Python",\
                                    font=("MSゴシック","10","bold"),\
                                    fg="#000000",\
                                    background="#D3D3D3")
        python_setting_lbl.place(x=15,y=405)
        # Add_Rows Setting
        add_rows_lbl = tk.Label(conf_canvas,text="・ADD ROWS(この値は変更できません。)",bg="#EE82EE")
        add_rows_lbl.place(x=15,y=430)
        add_rows_text = tk.StringVar()
        if inifile.has_option('PYTHON','ADD_ROW'):
            add_rows_text.set(inifile['PYTHON']['ADD_ROW'])
        add_rows_entry = tk.Entry(conf_canvas,\
                                  textvariable=add_rows_text,\
                                    width=5,\
                                    justify='center',\
                                    state='disabled',\
                                    disabledbackground="#A9A9A9")
        add_rows_entry.place(x=230,y=430)
        # Replace_Word_List Setting
        rep_word_lbl = tk.Label(conf_canvas,\
                                text="・完全一致の際、CSV側フィールド名で無視したい文字列\n※['a','b']のリスト形式で記入",\
                                bg="#EE82EE",\
                                justify='left',\
                                anchor=tk.W)
        rep_word_lbl.place(x=15,y=450)
        rep_word_text = tk.StringVar()
        if inifile.has_option('PYTHON','REPLASE_WORD'):
            rep_word_text.set(inifile['PYTHON']['REPLASE_WORD'])
        rep_word_entry = tk.Entry(conf_canvas,textvariable=rep_word_text,width=60)
        rep_word_entry.place(x=20,y=485)

        # Configration_Update_Button
        update_btn = ttk.Button(conf_canvas,\
                                text="更新",\
                                width=10,\
                                command=conf_update_click,\
                                default='active',\
                                style="office.TButton")
        update_btn.place(x=295,y=515)

        # What To Do When Closing
        conf_window.protocol("WM_DELETE_WINDOW", conf_window_close)




