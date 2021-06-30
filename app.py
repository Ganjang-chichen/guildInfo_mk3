import re
from tkinter import *
from PIL import Image, ImageTk
import json
import os
import copy

from function import excel_ctl
from function import week_calc

World = "29"
guild_name = "Lune"
guild_data = {}
guild_data_path = "./jsondata/charData.json"
guild_imgData_path = "./img/charImg/"
excel_path = "./ExcelFile/text.xlsx"

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def get_chardata() :
    if os.path.exists(guild_data_path):
        global guild_data
        global World
        global guild_name
        global excel_path

        with open(guild_data_path, 'r', encoding="UTF8") as f:
            json_data = json.load(f)
        guild_data = json_data
        World = guild_data["World"]
        guild_name = guild_data["GuildName"]
        excel_path = "./ExcelFile/" + guild_name + ".xlsx"

        return True
    else :
        return False

# Boot seq -------------------------------------------------------------------------------------------
isCharDataExist = get_chardata()

createFolder("./ExcelFile")
excel_ctl.file_path = excel_path
wb = excel_ctl.createExcel_if_not_exist()
thisWeekString = week_calc.get_week_range(week_calc.get_weekinfo(0))
if thisWeekString in wb.sheetnames :
    1
else :
    excel_ctl.createSheet_if_not_exist(wb, thisWeekString, guild_data)

root = Tk()
root.title("guild manager")
root.iconbitmap("./img/slime.ico")
root.geometry("900x750")

main_fraim = LabelFrame(root, height=300, width=500, bg="#ffffff")
main_fraim.pack()


# img setting
setting_button_img = ImageTk.PhotoImage(Image.open("./img/gear1.png"))
create_log_button_img = ImageTk.PhotoImage(Image.open("./img/slime-export.png"))

char_img_dict = {}

def char_img_load() :
    global char_img_dict
    if isCharDataExist :
        for data in guild_data["charData"] :
            name = data["name"]
            temp_img = ImageTk.PhotoImage(Image.open(guild_imgData_path + name + ".png"))
            char_img_dict[name] = temp_img
        return True
    else : 
        return False

isImgDataExist = char_img_load()
# UI --------------------------------------------------------------------------------------
# clear window
def reset() : 
    global main_fraim
    main_fraim.forget()
    main_fraim = LabelFrame(root, height=300, width=500, bg="#ffffff")
    main_fraim.pack()

# 주간 기록 입력 창----------------------------------------------------------------------------
# 주간 기록 전용 변수
Selected_week = StringVar() # 주차 선택 값
selected_sheet = False      # 주차 엑셀 차트
char_page_num = 0           # 페이지 넘버 show_charinfo() 에서 사용

def onclick_input_log(self, row, col, value) :
    selected_sheet.cell(row=row, column=col).value = value
    wb.save(excel_ctl.file_path)
    

def onEnter_input_log(e, row, col, value) :
    selected_sheet.cell(row=row, column=col).value = int(value)
    wb.save(excel_ctl.file_path)
    

def add_line(idx, data, row_num):
    global main_fraim
    char_name = data
    char_img = char_img_dict[char_name]
    char_btn = Button(main_fraim, image=char_img, bg="#ffffff")
    char_btn.grid(row=(2 * idx) , column=0)

    char_label = Label(main_fraim, text=char_name, bg="#ffffff")
    char_label.grid(row=(2 * idx + 1), column=0)

    selected_w = IntVar()
    option_w = [0, 1, 2, 3, 4, 5]
    drop_w = OptionMenu(main_fraim, selected_w, *option_w, command=lambda self : onclick_input_log(self, row=row_num, col=2, value=selected_w.get()))
    drop_w.grid(row=(2 * idx), column=1)
    selected_w.set(int(selected_sheet.cell(row=row_num, column=2).value))

    entry_s = Entry(main_fraim, width=5)
    entry_s.bind('<Return>', lambda e : onEnter_input_log(e, row=row_num, col=3, value=entry_s.get()))
    entry_s.grid(row=(2 * idx), column=2)
    entry_s.insert(0, selected_sheet.cell(row=row_num, column=3).value)

    selected_f = IntVar()
    option_f = [0, 100, 200, 250, 350, 400, 450, 550, 650, 800, 1000]
    drop_f = OptionMenu(main_fraim, selected_f, *option_f, command=lambda self : onclick_input_log(self, row=row_num, col=4, value=selected_f.get()))
    drop_f.grid(row=(2 * idx), column=3)
    selected_f.set(int(selected_sheet.cell(row=row_num, column=4).value))

def show_charinfo(page_num) :
    global selected_sheet
    char_count = len(char_img_dict)
    if (char_count // 3) == page_num :
        if char_count % 3 == 1 :
            add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value, 3 * page_num + 2)
        elif char_count % 3 == 2 :
            add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value, 3 * page_num + 2)
            add_line(2, selected_sheet.cell(row=(3 * page_num + 3),column=1).value, 3 * page_num + 3)
    else :
        add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value, 3 * page_num + 2)
        add_line(2, selected_sheet.cell(row=(3 * page_num + 3),column=1).value, 3 * page_num + 3)
        add_line(3, selected_sheet.cell(row=(3 * page_num + 4),column=1).value, 3 * page_num + 4)

def back_btn_charinfo() :
    global char_page_num
    if char_page_num > 0 :
        char_page_num = char_page_num - 1
        show_log_window()

def forward_btn_charinfo() :
    global char_page_num
    char_count = len(char_img_dict)
    page_count = 0
    if char_count % 3 == 0 :
        page_count = char_count // 3 - 1
    else :
        page_count = char_count // 3
    
    if char_page_num < page_count :
        char_page_num = char_page_num + 1
        show_log_window()

def alter_date(self) :
    show_log_window()

def show_log_window():
    reset()

    Label(main_fraim, text="기록창", bg="#ffffff").grid(row=0, column=0)
    sheet_names_reverse = wb.sheetnames[::-1]
    sheet_names_reverse.pop()

    if len(sheet_names_reverse) == 0 :
        w = week_calc.get_weekinfo(0)
        thisweek = week_calc.get_week_range(w)
        sheet_names_reverse = [thisweek]
    
    if Selected_week.get() == '' :
        Selected_week.set(sheet_names_reverse[0])
    week_select = OptionMenu(main_fraim, Selected_week, *sheet_names_reverse, command=alter_date)
    week_select.grid(row=0, column=1, columnspan=3)
    

    Label(main_fraim, text="케릭터", bg="#ffffff").grid(row=1, column=0)
    Label(main_fraim, text="주간미션", bg="#ffffff").grid(row=1, column=1)
    Label(main_fraim, text="지하수로", bg="#ffffff").grid(row=1, column=2)
    Label(main_fraim, text="플래그", bg="#ffffff").grid(row=1, column=3)

    if isCharDataExist :
        global selected_sheet
        selected_sheet = excel_ctl.createSheet_if_not_exist(wb, Selected_week.get(), guild_data)

        show_charinfo(char_page_num)

        buttonLeft = Button(main_fraim, text="<<<", command=back_btn_charinfo)
        buttonLeft.grid(row=8, column=0,columnspan=2, sticky=W)
        buttonRight = Button(main_fraim, text=">>>", command=forward_btn_charinfo)
        buttonRight.grid(row=8, column=2,columnspan=2, sticky=E)
        
        char_count = len(selected_sheet["A"]) - 1
        page_count = 0
        if char_count % 3 == 0 :
            page_count = char_count // 3
        else :
            page_count = char_count // 3 + 1
        page_label = Label(main_fraim, text=("< " + str(char_page_num + 1) + " / " + str(page_count) + " >"))
        page_label.grid(row=9, column=0, columnspan=4)
    else :
        warningLabel = Label(main_fraim, text="케릭터 데이터가 존재하지 않습니다.")
        warningLabel.grid(row=2, column=0, columnspan=4)



def show_config() :
    reset()
    

def show_mainframe() :
    reset()
    global main_fraim
    global setting_button_img
    global create_log_button_img
    
    setting_button = Button(main_fraim, image=setting_button_img, bg="#ffffff")
    setting_button.grid(row=0, column=0)

    setting_label = Label(main_fraim, text="설정", bg="#ffffff")
    setting_label.grid(row=1, column=0)

    
    create_log_button = Button(main_fraim, image=create_log_button_img, bg="#ffffff", command=show_log_window)
    create_log_button.grid(row=0, column=1)

    create_log_label = Label(main_fraim, text="이번주 리스트", bg="#ffffff")
    create_log_label.grid(row=1, column=1)

show_mainframe()

# menu setting
main_menu = Menu(root)
root.config(menu=main_menu)

windowCtl_menu = Menu(main_menu)
main_menu.add_cascade(label="화면이동", menu=windowCtl_menu)
windowCtl_menu.add_command(label="첫 화면", command=show_mainframe)
windowCtl_menu.add_command(label="설정 화면",command=show_config)
windowCtl_menu.add_command(label="입력 화면",command=show_log_window)

root.mainloop()