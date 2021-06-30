import openpyxl
from openpyxl.chart import BarChart, Reference
import os
import math
import json

#import week_calc

file_path = "./ExcelFile/test.xlsx"

World = "29"
guild_name = "Lune"
guild_data = {}
guild_data_path = "./jsondata/" + World + "/" + guild_name + "/charData.json"
guild_imgData_path = "./img/" + World + "/" + guild_name + "/"

def get_chardata() :
    if os.path.exists(guild_data_path):
        global guild_data
        with open(guild_data_path, 'r', encoding="UTF8") as f:
            json_data = json.load(f)
        guild_data = json_data
        return True
    else :
        return False


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def createExcel_if_not_exist() :   
    if os.path.exists(file_path) :
        wb = openpyxl.load_workbook(file_path)
    else :
        wb = openpyxl.Workbook()
        wb.save(file_path)
    return wb
    
def createSheet_if_not_exist(wb, sheet_name, guild_data) :
    wb.active
    if sheet_name in wb.sheetnames :
        return wb[sheet_name]
    else :
        sheet = wb.create_sheet(sheet_name)
        sheet.cell(row=1, column=1).value = "케릭터명"
        sheet.cell(row=1, column=2).value = "주간미션"
        sheet.cell(row=1, column=3).value = "수로"
        sheet.cell(row=1, column=4).value = "플래그"

        sheet.merge_cells(start_row=1, end_row=1, start_column=7, end_column=10)
        sheet.cell(row=1, column=7).value = "이번주"

        sheet.cell(row=2, column=7).value = "구분"
        sheet.cell(row=3, column=7).value = "참가자 수"
        sheet.cell(row=4, column=7).value = "점수 평균"
        sheet.cell(row=5, column=7).value = "참가자 평균"
        sheet.cell(row=6, column=7).value = "점수 합산"

        sheet.cell(row=2, column=8).value = "주간미션"
        sheet.cell(row=2, column=9).value = "수로 "
        sheet.cell(row=2, column=10).value = "플래그"

        sheet.cell(row=3, column=8).value = '=COUNTIF(B2:B201,">0")'
        sheet.cell(row=3, column=9).value = '=COUNTIF(C2:C201,">0")'
        sheet.cell(row=3, column=10).value = '=COUNTIF(D2:D201,">0")'

        sheet.cell(row=4, column=8).value = '=AVERAGE(B2:B201)'
        sheet.cell(row=4, column=9).value = '=AVERAGE(C2:C201)'
        sheet.cell(row=4, column=10).value = '=AVERAGE(D2:D201)'

        sheet.cell(row=5, column=8).value = '=AVERAGEIF(B2:B201,">0")'
        sheet.cell(row=5, column=9).value = '=AVERAGEIF(C2:C201,">0")'
        sheet.cell(row=5, column=10).value = '=AVERAGEIF(D2:D201,">0")'

        sheet.cell(row=6, column=8).value = '=SUM(B2:B201)'
        sheet.cell(row=6, column=9).value = '=SUM(C2:C201)'
        sheet.cell(row=6, column=10).value = '=SUM(D2:D201)'
        
        if len(wb.sheetnames) != 2 :
            last_sheet_name = wb.sheetnames[len(wb.sheetnames) - 2]
            sheet.merge_cells(start_row=1, end_row=1, start_column=12, end_column=15)
            sheet.cell(row=1, column=12).value = "저번주"

            sheet.cell(row=2, column=12).value = "구분"
            sheet.cell(row=3, column=12).value = "참가자 수"
            sheet.cell(row=4, column=12).value = "점수 평균"
            sheet.cell(row=5, column=12).value = "참가자 평균"
            sheet.cell(row=6, column=12).value = "점수 합산"

            sheet.cell(row=2, column=13).value = "주간미션"
            sheet.cell(row=2, column=14).value = "수로 "
            sheet.cell(row=2, column=15).value = "플래그"

            sheet.cell(row=3, column=13).value = "='" + last_sheet_name + "'!H3"
            sheet.cell(row=3, column=14).value = "='" + last_sheet_name + "'!I3"
            sheet.cell(row=3, column=15).value = "='" + last_sheet_name + "'!J3"

            sheet.cell(row=4, column=13).value = "='" + last_sheet_name + "'!H4"
            sheet.cell(row=4, column=14).value = "='" + last_sheet_name + "'!I4"
            sheet.cell(row=4, column=15).value = "='" + last_sheet_name + "'!J4"

            sheet.cell(row=5, column=13).value = "='" + last_sheet_name + "'!H5"
            sheet.cell(row=5, column=14).value = "='" + last_sheet_name + "'!I5"
            sheet.cell(row=5, column=15).value = "='" + last_sheet_name + "'!J5"

            sheet.cell(row=6, column=13).value = "='" + last_sheet_name + "'!H6"
            sheet.cell(row=6, column=14).value = "='" + last_sheet_name + "'!I6"
            sheet.cell(row=6, column=15).value = "='" + last_sheet_name + "'!J6"



        i = 0
        for rowData in guild_data['charData'] :
            sheet.cell(row=(i + 2), column=1).value = rowData['name']
            sheet.cell(row=(i + 2), column=2).value = 0
            sheet.cell(row=(i + 2), column=3).value = 0
            sheet.cell(row=(i + 2), column=4).value = 0
            i = i + 1

        wb.save(file_path)
        return sheet

if __name__ == "__main__" :
    createFolder("./testExcelFile")
    get_chardata()

    wb = createExcel_if_not_exist()
    weekinfo = week_calc.get_weekinfo(0)
    sheet_name = (  weekinfo["first"]["year"] + weekinfo["first"]["month"] + weekinfo["first"]["day"] + "-" + 
                    weekinfo["last"]["year"] + weekinfo["last"]["month"] + weekinfo["last"]["day"])
    print(sheet_name)
    createSheet_if_not_exist(wb, sheet_name, guild_data)
    print(wb.sheetnames)
    sheet_names_reverse = wb.sheetnames[::-1]
    sheet_names_reverse.pop()
    print(wb.sheetnames)
    print(sheet_names_reverse)
