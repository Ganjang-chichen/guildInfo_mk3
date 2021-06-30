from tkinter import *
from PIL import Image, ImageTk
import json
import os
import copy

from function import database
from function import excel_ctl
from function import week_calc

World = "29"
guild_name = "Lune"
guild_data = {}
guild_data_path = "./jsondata/" + World + "/" + guild_name + "/charData.json"
guild_imgData_path = "./img/" + World + "/" + guild_name + "/"

def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

def get_chardata() :
    if os.path.exists(guild_data_path):
        global guild_data
        with open(guild_data_path, 'r', encoding="UTF8") as f:
            json_data = json.load(f)
        guild_data = json_data
        return True
    else :
        return False

# Boot seq -------------------------------------------------------------------------------------------
isCharDataExist = get_chardata()

createFolder("./testExcelFile")
wb = excel_ctl.createExcel_if_not_exist()

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
Selected_week = StringVar()
selected_sheet = False

def add_line(idx, data):
    global main_fraim
    char_name = data
    char_img = char_img_dict[char_name]
    char_btn = Button(main_fraim, image=char_img, bg="#ffffff")
    char_btn.grid(row=(2 * idx) , column=0)

    char_label = Label(main_fraim, text=char_name, bg="#ffffff")
    char_label.grid(row=(2 * idx + 1), column=0)

    selected_w = IntVar()
    option_w = [0, 1, 2, 3, 4, 5]
    drop_w = OptionMenu(main_fraim, selected_w, *option_w)
    drop_w.grid(row=(2 * idx), column=1)

    entry_s = Entry(main_fraim, width=5)
    entry_s.grid(row=(2 * idx), column=2)
    entry_s.insert(0, "0")

    selected_f = IntVar()
    option_f = [0, 100, 200, 250, 350, 400, 450, 550, 650, 800, 1000]
    drop_f = OptionMenu(main_fraim, selected_f, *option_f)
    drop_f.grid(row=(2 * idx), column=3)

char_page_num = 0
def show_charinfo(page_num) :
    global selected_sheet
    char_count = len(char_img_dict)
    if (char_count // 3) == page_num :
        if char_count % 3 == 1 :
            add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value)
        elif char_count % 3 == 2 :
            add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value)
            add_line(2, selected_sheet.cell(row=(3 * page_num + 3),column=1).value)
    else :
        add_line(1, selected_sheet.cell(row=(3 * page_num + 2),column=1).value)
        add_line(2, selected_sheet.cell(row=(3 * page_num + 3),column=1).value)
        add_line(3, selected_sheet.cell(row=(3 * page_num + 4),column=1).value)

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
    week_select = OptionMenu(main_fraim, Selected_week, *sheet_names_reverse)
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

root.mainloop()