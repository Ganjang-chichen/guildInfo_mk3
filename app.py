import re
from tkinter import *
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
import numpy
import json
import os
import copy
import sys
import time

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

if isCharDataExist == False :
    print("케릭터 데이터가 존재하지 않습니다.")
    print("먼저 데이터 다운로드를 실행해주시고 다시 실행해 주세요.")
    time.sleep(3)
    sys.exit(0)

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
show_summary_chart_img = ImageTk.PhotoImage(Image.open("./img/graph.png"))

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
selected_sorting_mode = StringVar() # sorting
selected_sorting_mode.set("기본 : 직위 -> Lv 순")

sorted_list = guild_data["charData"]
sorted_by_lv = sorted(guild_data["charData"], key=lambda data : data["Lv"], reverse=True)
sorted_by_name = sorted(guild_data["charData"], key=lambda data : data["name"])


def isExist(x) :
    if x > 0 :
        return 1
    else :
        return 0

def onclick_input_log(self, row, col, value) :
    selected_sheet.cell(row=row, column=col).value = value
    wb.save(excel_ctl.file_path)
    
def onEnter_input_log(e, row, col, value) :
    selected_sheet.cell(row=row, column=col).value = int(value)
    wb.save(excel_ctl.file_path)
    
def show_charlog_graph(char_name) :
    date_list = []
    char_weekly_mission = []
    char_suro = []
    char_flag = []
    for sheet_name in wb.sheetnames :
        if sheet_name == "Sheet" :
            continue
        else :
            temp_sheet = wb[sheet_name]
            temp_idx = 0
            isLogExist = False
            for temp_A in temp_sheet['A'] :
                if temp_A.value == char_name :
                    isLogExist = True
                    break
                temp_idx = temp_idx + 1
            if isLogExist :
                date_list.append(sheet_name.split("-")[1])
                char_weekly_mission.append(temp_sheet['B' + str(temp_idx + 1)].value)
                char_suro.append(temp_sheet['C' + str(temp_idx + 1)].value)
                char_flag.append(temp_sheet['D' + str(temp_idx + 1)].value)

    isJoined_wList = list(map(isExist, char_weekly_mission))
    isJoined_sList = list(map(isExist, char_suro))
    isJoined_fList = list(map(isExist, char_flag))

    main_graph = plt.figure(constrained_layout=True, figsize=(12, 8))
    plt.rc('font', family="Malgun Gothic")
    spec4 = main_graph.add_gridspec(ncols=4, nrows=4)

    graph_char_img = mpimg.imread(guild_imgData_path + char_name + ".png")
    main_graph.add_subplot(spec4[0, 0])
    plt.imshow(graph_char_img)

    bar_distribute = numpy.arange(len(date_list))

    main_graph.add_subplot(spec4[0, 1:])
    plt.bar(bar_distribute-0.0, isJoined_wList, label="주간 미션", width=0.03)
    plt.bar(bar_distribute+0.03, isJoined_sList, label="수로", width=0.03)
    plt.bar(bar_distribute+0.06, isJoined_fList, label="플래그", width=0.03)
    plt.xticks(bar_distribute, date_list)
    plt.title(label="참가 여부")
    plt.legend()

    main_graph.add_subplot(spec4[1, :])
    plt.plot(date_list, char_weekly_mission, label="주간 미션")
    plt.title(label="주간 미션 점수")
    plt.legend()

    main_graph.add_subplot(spec4[2, :])
    plt.plot(date_list, char_suro, label="수로")
    plt.title(label="수로 점수")
    plt.legend()

    main_graph.add_subplot(spec4[3, :])
    plt.plot(date_list, char_flag, label="플래그")
    plt.title(label="플래그 점수")
    plt.legend()

    plt.suptitle(char_name)
    plt.show()

def add_line(idx, data):
    global main_fraim
    global selected_sheet

    isExist_inExcel = False
    row_num = 1
    for cell in selected_sheet["A"] :
        if cell.value == data :
            isExist_inExcel = True
            break
        else :
            row_num = row_num + 1

    if isExist_inExcel == False:
        selected_sheet["A" + str(row_num)].value = data
        selected_sheet["B" + str(row_num)].value = 0
        selected_sheet["C" + str(row_num)].value = 0
        selected_sheet["D" + str(row_num)].value = 0
        wb.save(excel_ctl.file_path)

    char_name = data

    if char_img_dict[char_name] == None :
        char_img = setting_button_img
    else : 
        char_img = char_img_dict[char_name]
    char_btn = Button(main_fraim, image=char_img, bg="#ffffff", command=lambda : show_charlog_graph(char_name))
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

    char_count = len(sorted_list)
    if (char_count // 3) == page_num :
        if char_count % 3 == 1 :
            add_line(1, sorted_list[3 * page_num]["name"])
        elif char_count % 3 == 2 :
            add_line(1, sorted_list[3 * page_num]["name"])
            add_line(2, sorted_list[3 * page_num + 1]["name"])
    else :
        add_line(1, sorted_list[3 * page_num]["name"])
        add_line(2, sorted_list[3 * page_num + 1]["name"])
        add_line(3, sorted_list[3 * page_num + 2]["name"])

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

def alter_mode(self) : 
    global sorted_list
    if selected_sorting_mode.get() == "기본 : 직위 -> Lv 순" :
        sorted_list = guild_data["charData"]
    elif selected_sorting_mode.get() == "이름순" :
        sorted_list = sorted_by_name
    elif selected_sorting_mode.get() == "Lv 순" :
        sorted_list = sorted_by_lv
    show_log_window()
    

def alter_date(self) :
    show_log_window()

def show_log_window():
    reset()

    sorting_options = ["기본 : 직위 -> Lv 순", "이름순", "Lv 순"]
    drop_sorting = OptionMenu(main_fraim, selected_sorting_mode, *sorting_options, command=alter_mode)
    drop_sorting.grid(row=0, column=0)

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



def show_summary_chart() :
    date_list = []
    joined_weekly = []
    joined_suro = []
    joined_flag = []
    avg_weekly = []
    avg_suro = []
    avg_flag = []
    avg_joined_weekly = []
    avg_joined_suro = []
    avg_joined_flag = []
    sum_weekly = []
    sum_suro =[]
    sum_flag = []
    
    for sheet_name in wb.sheetnames :
        if sheet_name == "Sheet" :
            continue
        else :
            date_list.append(sheet_name.split('-')[1])
            temp_sheet = wb[sheet_name]
            
            temp_weekly_joined = 0
            temp_weekly_sum = 0
            for w in temp_sheet["B"] :
                if w.value == "주간 미션" :
                    continue
                else :
                    temp_weekly_sum  = temp_weekly_sum + int(w.value)
                    if w.value > 0 :
                        temp_weekly_joined  = temp_weekly_joined + 1
            
            joined_weekly.append(temp_weekly_joined)
            avg_weekly.append(temp_weekly_sum / (len(temp_sheet["B"]) - 1))
            if temp_weekly_joined > 0 :
                avg_joined_weekly.append(temp_weekly_sum / temp_weekly_joined)
            else :
                avg_joined_weekly.append(0)
            sum_weekly.append(temp_weekly_sum)

            temp_suro_joined = 0
            temp_suro_sum = 0
            for s in temp_sheet["C"] :
                if s.value == "수로" :
                    continue
                else :
                    temp_suro_sum  = temp_suro_sum + int(s.value)
                    if s.value > 0 :
                        temp_suro_joined  = temp_suro_joined + 1
            
            joined_suro.append(temp_suro_joined)
            avg_suro.append(temp_suro_sum / (len(temp_sheet["C"]) - 1))
            if temp_suro_joined > 0 :
                avg_joined_suro.append(temp_suro_sum / temp_suro_joined)
            else :
                avg_joined_suro.append(0)
            sum_suro.append(temp_suro_sum)

            temp_flag_joined = 0
            temp_flag_sum = 0
            for f in temp_sheet["D"] :
                if f.value == "플래그" :
                    continue
                else :
                    temp_flag_sum  = temp_flag_sum + int(f.value)
                    if f.value > 0 :
                        temp_flag_joined  = temp_flag_joined + 1
            
            joined_flag.append(temp_flag_joined)
            avg_flag.append(temp_flag_sum / (len(temp_sheet["D"]) - 1))
            if temp_flag_joined > 0 :
                avg_joined_flag.append(temp_flag_sum / temp_flag_joined)
            else :
                avg_joined_flag.append(0)
            sum_flag.append(temp_flag_sum)
            
    if len(date_list) > 1 :
        data_count = len(date_list)

        summary_axis_label = ["점수 평균", "참여자 점수 평균"]
        bar_distribute_summary = numpy.arange(len(summary_axis_label))

        summary_sum_axis_label = ["점수 총합"]
        bar_distribute_summary_sum = numpy.arange(len(summary_sum_axis_label))

        summary_chart = plt.figure(constrained_layout=True, figsize=(12, 8))
        plt.rc('font', family="Malgun Gothic")
        spec2x4 = summary_chart.add_gridspec(ncols=6, nrows=4)

        summary_chart.add_subplot(spec2x4[0, :3])
        this_week_joined = [joined_weekly[data_count - 1], joined_suro[data_count - 1], joined_flag[data_count - 1]]
        last_week_joined = [joined_weekly[data_count - 2], joined_suro[data_count - 2], joined_flag[data_count - 2]]
        joined_axis_label = ["주간미션", "수로", "플래그"]
        bar_distribute = numpy.arange(len(joined_axis_label))
        plt.bar(bar_distribute-0.0, this_week_joined, label="이번주", width=0.05)
        plt.bar(bar_distribute+0.05, last_week_joined, label="저번주", width=0.05)
        plt.xticks(bar_distribute, joined_axis_label)
        plt.title(label="이번주/저번주 참여 인원 비교")
        plt.legend()

        summary_chart.add_subplot(spec2x4[0, 3:5])
        this_week_weekly = [avg_weekly[data_count - 1], avg_joined_weekly[data_count - 1]]
        last_week_weekly = [avg_weekly[data_count - 2], avg_joined_weekly[data_count - 2]]
        plt.bar(bar_distribute_summary-0.0, this_week_weekly, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary+0.05, last_week_weekly, label="저번주", width=0.05)
        plt.xticks(bar_distribute_summary, summary_axis_label)
        plt.title(label="주간미션 평균")
        plt.legend()

        summary_chart.add_subplot(spec2x4[0, 5])
        this_week_weekly_sum = [sum_weekly[data_count - 1]]
        last_week_weekly_sum = [sum_weekly[data_count - 2]]
        plt.bar(bar_distribute_summary_sum-0.0, this_week_weekly_sum, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary_sum+0.05, last_week_weekly_sum, label="저번주", width=0.05)
        plt.xticks(bar_distribute_summary_sum, summary_sum_axis_label)
        plt.title(label="주간미션 누계")
        plt.legend()

        summary_chart.add_subplot(spec2x4[1, :2])
        this_week_suro = [avg_suro[data_count - 1], avg_joined_suro[data_count - 1]]
        last_week_suro = [avg_suro[data_count - 2], avg_joined_suro[data_count - 2]]
        plt.bar(bar_distribute_summary-0.0, this_week_suro, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary+0.05, last_week_suro, label="저번주", width=0.05)
        plt.xticks(bar_distribute_summary, summary_axis_label)
        plt.title(label="수로 평균")
        plt.legend()
        
        summary_chart.add_subplot(spec2x4[1, 2])
        this_week_suro_sum = [sum_suro[data_count - 1]]
        last_week_suro_sum = [sum_suro[data_count - 2]]
        plt.bar(bar_distribute_summary_sum-0.0, this_week_suro_sum, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary_sum+0.05, last_week_suro_sum, label="저번주", width=0.05)
        plt.title(label="수로 누계")
        plt.legend()

        summary_chart.add_subplot(spec2x4[1, 3:5])
        this_week_flag = [avg_flag[data_count - 1], avg_joined_flag[data_count - 1]]
        last_week_flag = [avg_flag[data_count - 2], avg_joined_flag[data_count - 2]]
        plt.bar(bar_distribute_summary-0.0, this_week_flag, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary+0.05, last_week_flag, label="저번주", width=0.05)
        plt.title(label="플래그 평균")
        plt.legend()

        summary_chart.add_subplot(spec2x4[1, 5])
        this_week_flag_sum = [sum_flag[data_count - 1]]
        last_week_flag_sum = [sum_flag[data_count - 2]]
        plt.bar(bar_distribute_summary_sum-0.0, this_week_flag_sum, label="이번주", width=0.05)
        plt.bar(bar_distribute_summary_sum+0.05, last_week_flag_sum, label="저번주", width=0.05)
        plt.title(label="플래그 누계")
        plt.legend()

        summary_chart.add_subplot(spec2x4[2, :3])
        plt.plot(date_list, joined_weekly, label="주간미션 참가자수")
        plt.plot(date_list, joined_suro, label="수로 참가자수")
        plt.plot(date_list, joined_flag, label="플래그 참가자수")
        plt.title(label="참가자 수")
        plt.legend()

        summary_chart.add_subplot(spec2x4[2, 3:])
        plt.plot(date_list, sum_weekly, label="주간미션 점수")
        plt.title(label="주간미션 합계")
        plt.legend()

        summary_chart.add_subplot(spec2x4[3, :3])
        plt.plot(date_list, sum_suro, label="수로 점수")
        plt.title(label="수로 합계")
        plt.legend()
        
        summary_chart.add_subplot(spec2x4[3, 3:])
        plt.plot(date_list, sum_flag, label="플래그 점수")
        plt.title(label="플래그 합계")
        plt.legend()

        plt.suptitle("개요")
        plt.show()
    else :
        summary_chart = plt.figure(constrained_layout=True, figsize=(12, 8))
        plt.rc('font', family="Malgun Gothic")
        spec1x3 = summary_chart.add_gridspec(ncols=1, nrows=3)
        
        summary_chart.add_subplot(spec1x3[0, 0])
        plt.plot(date_list, sum_weekly, label="주간미션 점수")
        plt.title(label="주간미션 합계")
        plt.legend()

        summary_chart.add_subplot(spec1x3[1, 0])
        plt.plot(date_list, sum_suro, label="수로 점수")
        plt.title(label="수로 합계")
        plt.legend()
        
        summary_chart.add_subplot(spec1x3[2, 0])
        plt.plot(date_list, sum_flag, label="플래그 점수")
        plt.title(label="플래그 합계")
        plt.legend()

        plt.suptitle("개요")
        plt.show()

def show_mainframe() :
    reset()
    global main_fraim
    global show_summary_chart_img
    global create_log_button_img
    
    setting_button = Button(main_fraim, image=show_summary_chart_img, bg="#ffffff", command=show_summary_chart)
    setting_button.grid(row=0, column=0)

    setting_label = Label(main_fraim, text="전체 요약", bg="#ffffff")
    setting_label.grid(row=1, column=0)

    create_log_button = Button(main_fraim, image=create_log_button_img, bg="#ffffff", command=show_log_window)
    create_log_button.grid(row=0, column=1)

    create_log_label = Label(main_fraim, text="이번주 리스트", bg="#ffffff")
    create_log_label.grid(row=1, column=1)

    Label(main_fraim, text="주의사항", bg="#a1131f", foreground="#ffff00").grid(row=2, column=0, columnspan=2)
    Label(main_fraim, text="이 프로그램 실행 전 데이터 다운로드 후 실행할 것.", bg="#ffffff").grid(row=3, column=0, columnspan=2)
    Label(main_fraim, text="이 프로그램과 엑셀파일을 동시에 열지 마세요.", bg="#ffffff").grid(row=4, column=0, columnspan=2)
    Label(main_fraim, text="데이터 입력시 오류가 발생합니다.", bg="#ffffff").grid(row=5, column=0, columnspan=2)
    

show_mainframe()

# menu setting
main_menu = Menu(root)
root.config(menu=main_menu)

windowCtl_menu = Menu(main_menu)
main_menu.add_cascade(label="화면이동", menu=windowCtl_menu)
windowCtl_menu.add_command(label="첫 화면", command=show_mainframe)
windowCtl_menu.add_command(label="입력 화면",command=show_log_window)

root.mainloop()