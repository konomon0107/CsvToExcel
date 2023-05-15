import csv
import joblib

def scan_data():
    # Read CSVFile Path
    file_path = joblib.load('csv_file_path.txt')
    # Read CSVFile
    open_scan = open(file_path[0],"r", encoding="shift_jis")
    read_scan = csv.reader(open_scan,\
                           delimiter=",",\
                            doublequote=True,\
                            lineterminator="\r\n",\
                            quotechar='"',\
                            skipinitialspace= True)
    # List CSV Datas
    personal_infomation_data = [item for item in read_scan]
    # Read textfile
    radio_value = joblib.load('radio_value.txt')
    radiovalue = radio_value[0]

    # Header Row Extraction
    HEADER_DATA = personal_infomation_data[0]
    # Data Row Extraction
    PERSONAL_INFOMATION_DATA = personal_infomation_data[1:]
    # Sort Arry
    if radiovalue == 1:
        # 部門コード、社員コードの優先順位でソート
        PERSONAL_INFOMATION_DATA.sort(reverse=False,key=lambda x:(x[3],x[1]))
    
    return HEADER_DATA,PERSONAL_INFOMATION_DATA