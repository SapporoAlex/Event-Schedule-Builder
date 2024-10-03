from datetime import datetime, timedelta
import openpyxl as xl
from openpyxl.styles import Font, Alignment
import re

run = True
date_string = ""

day_in_ja_dic = {
    "Monday": "(月)",
    "Tuesday": "(火)",
    "Wednesday": "(水)",
    "Thursday": "(木)",
    "Friday": "(金)",
    "Saturday": "(土)",
    "Sunday": "(日)",
}

# Get today's date
today = datetime.now().date()
# Calculate tomorrow's date by adding 1 day
tomorrow = today + timedelta(days=1)
tomorrow_day_of_week = tomorrow.strftime("%A")
tomorrow_date_string = str(tomorrow) + " " + str(day_in_ja_dic.get(tomorrow_day_of_week))


def sort_list(list):
    sorted_list = list
    try:
        sorted_list = sorted(list, key=lambda x: datetime.strptime(x.split(' ')[0], "%H:%M"))
    except ValueError:
        print("10:30は、このような時間形式にしてください。")
    return sorted_list


def add_line_break_to_list_items(list):
    newline_list = "\n".join(list)
    return newline_list


def display_table(tokiwagi, nogiku, entei, hall):
    print(f"""


{date_string}
===ときわぎ===
{tokiwagi}

===のぎく====
{nogiku}

===園庭=====
{entei}

===ホール===
{hall}

    """)


def set_date():
    print("明日ですか？")
    print("明日の日付を設定するにはエンターキーを、日付を変更するには[n]を押します。")
    choice = input(">>")
    if choice.lower() == "n":
        print("日付を入力")
        date = input(">>")
    else:
        date = tomorrow_date_string
    return date


def enter_loc_details(loc_list):
    print(f"PONYの時間とアクティビティを入力するか、エンターキーを押します。")
    pony = input(">>")
    if pony != "":
        pony = pony + " (ポニー)"
        loc_list.append(pony)

    print(f"BEETLEの時間とアクティビティを入力するか、エンターキーを押します。")
    beetle = input(">>")
    if beetle != "":
        beetle = beetle + " (ビートル)"
        loc_list.append(beetle)

    print(f"GRASSHOPPERの時間とアクティビティを入力するか、エンターキーを押します。")
    grasshopper = input(">>")
    if grasshopper != "":
        grasshopper = grasshopper + " (グラスホッパー)"
        loc_list.append(grasshopper)

    print(f"NENSHOの時間とアクティビティを入力するか、エンターキーを押します。")
    nensho = input(">>")
    if nensho != "":
        nensho = nensho + " (年少)"
        loc_list.append(nensho)

    print(f"DOLPHINの時間とアクティビティを入力するか、エンターキーを押します。")
    dolphin = input(">>")
    if dolphin != "":
        dolphin = dolphin + " (ドルフィン)"
        loc_list.append(dolphin)

    print(f"PENGUINの時間とアクティビティを入力するか、エンターキーを押します。")
    penguin = input(">>")
    if penguin != "":
        penguin = penguin + " (ペングイン)"
        loc_list.append(penguin)

    print(f"NENCHUの時間とアクティビティを入力するか、エンターキーを押します。")
    nenchu = input(">>")
    if nenchu != "":
        nenchu = nenchu + " (年中)"
        loc_list.append(nenchu)

    print(f"GIRAFFEの時間とアクティビティを入力するか、エンターキーを押します。")
    giraffe = input(">>")
    if giraffe != "":
        giraffe = giraffe + " (ギラッフ)"
        loc_list.append(giraffe)

    print(f"PANDAの時間とアクティビティを入力するか、エンターキーを押します。")
    panda = input(">>")
    if panda != "":
        panda = panda + " (パンダ)"
        loc_list.append(panda)

    print(f"NENCHOの時間とアクティビティを入力するか、エンターキーを押します。")
    nencho = input(">>")
    if nencho != "":
        nencho = nencho + " (年長)"
        loc_list.append(nencho)

    return loc_list


def add_or_remove_request(loc_list):
    print(f"""

1. アイテムを追加する
2. アイテムを削除する

""")
    add_remove = input(">> ")
    if add_remove == "1":
        loc_list = sort_list(enter_loc_details(loc_list))
        return add_line_break_to_list_items(loc_list) if len(loc_list) > 1 else loc_list
    else:
        entry_edit_menu = True
        while entry_edit_menu:
            print(f"""
どのエントリーを削除しますか？
1. PONY
2. BEETLE
3. GRASSHOPPER
4. NENSHO
5. DOLPHIN
6. PENGUIN
7. NENCHU
8. GIRAFFE
9. PANDA
10. NENCHO
11. 戻る

""")
            remove_request = input(">>")
            labels = ["(ポニー)", "(ビートル)", "(グラスホッパー)", "(年少)", "(ドルフィン)",
                      "(ペングイン)", "(年中)", "(ギラッフ)", "(パンダ)", "(年長)"]
            try:
                idx = int(remove_request) - 1
                loc_list = [item for item in loc_list if labels[idx] not in item]
            except (IndexError, ValueError):
                entry_edit_menu = False if remove_request == "11" else entry_edit_menu

        loc_list = sort_list(enter_loc_details(loc_list))
        return add_line_break_to_list_items(loc_list)


def save_to_excel(date_string, tokiwagi_list_sorted_newline, nogiku_list_sorted_newline, entei_list_sorted_newline,
                  hall_list_sorted_newline):
    workbook = xl.load_workbook("template/template 4 rows.xlsx")
    sheet = workbook.active
    sheet['a1'] = f"日付: {date_string}"
    sheet['a1'].font = Font(size=11)
    sheet['a2'] = "ときわぎ"
    sheet['a2'].font = Font(bold=True, size=12)
    sheet['b2'] = "のぎく"
    sheet['b2'].font = Font(bold=True, size=12)
    sheet['c2'] = "園庭"
    sheet['c2'].font = Font(bold=True, size=12)
    sheet['d2'] = "ホール"
    sheet['d2'].font = Font(bold=True, size=12)

    sheet['a3'] = tokiwagi_list_sorted_newline if tokiwagi_list_sorted_newline else "No activities"
    sheet['b3'] = nogiku_list_sorted_newline if nogiku_list_sorted_newline else "No activities"
    sheet['c3'] = entei_list_sorted_newline if entei_list_sorted_newline else "No activities"
    sheet['d3'] = hall_list_sorted_newline if hall_list_sorted_newline else "No activities"

    sheet['a3'].font = Font(size=11)
    sheet['b3'].font = Font(size=11)
    sheet['c3'].font = Font(size=11)
    sheet['d3'].font = Font(size=11)

    file_name = input("ファイル名を入力: ")
    sanitized_file_name = re.sub(r'[\/:*?"<>|]', "", file_name)
    workbook.save(f'{sanitized_file_name}.xlsx')


while run:
    date_string = set_date()

    tokiwagi_list = []
    nogiku_list = []
    entei_list = []
    hall_list = []

    tokiwagi_list_sorted = sort_list(tokiwagi_list)
    nogiku_list_sorted = sort_list(nogiku_list)
    entei_list_sorted = sort_list(entei_list)
    hall_list_sorted = sort_list(hall_list)

    tokiwagi_list_sorted_newline = add_line_break_to_list_items(tokiwagi_list_sorted)
    nogiku_list_sorted_newline = add_line_break_to_list_items(nogiku_list_sorted)
    entei_list_sorted_newline = add_line_break_to_list_items(entei_list_sorted)
    hall_list_sorted_newline = add_line_break_to_list_items(hall_list_sorted)

    select_loc_screen = True
    while select_loc_screen:
        display_table(
            tokiwagi_list_sorted_newline,
            nogiku_list_sorted_newline,
            entei_list_sorted_newline,
            hall_list_sorted_newline
        )
        print(f"""

場所を選ぶ
1. ときわぎ
2. のぎく
3. 園庭
4. ホール
5. ファイルを保存して終了する
6. プログラムを終了する

""")
        choice = input(">> ")

        if choice == "1":
            tokiwagi_list_sorted_newline = add_or_remove_request(tokiwagi_list)
        elif choice == "2":
            nogiku_list_sorted_newline = add_or_remove_request(nogiku_list)
        elif choice == "3":
            entei_list_sorted_newline = add_or_remove_request(entei_list)
        elif choice == "4":
            hall_list_sorted_newline = add_or_remove_request(hall_list)
        elif choice == "5":
            save_to_excel(date_string, tokiwagi_list_sorted_newline, nogiku_list_sorted_newline,
                          entei_list_sorted_newline, hall_list_sorted_newline)
            run = False
            break
        elif choice == "6":
            run = False
            break
