import openpyxl as xl
from openpyxl.styles import Font, Alignment


def enter_title():
    title = input("イベントのタイトル: ")
    return title


def enter_leader():
    leader = input("リーダー/司会: ")
    return leader


def enter_date():
    date = input("日付: ")
    return date


def enter_location():
    location = input("場所: ")
    return location


def enter_equipment():
    equipment = input("すべての機器をリスト: ")
    return equipment


def enter_activity(activities, activity_times):
    time = input("時間: ")
    activity = input("アクティビティを入力: ")
    activities.append(activity)
    activity_times.append(str(time))
    return activities, activity_times


def print_activities(activity_text):
    return "\n".join(activity_text)


def make_activities_text(activity_times, activities):
    combined_list = ["      " + "時間: " + a + "\n" + "        " + "アクティビティ: " + b for a, b in zip(activity_times, activities)]
    return combined_list


def main():
    run = True
    main_menu = True
    choice = 0
    while run:
        while main_menu:
            print(f"""
    ****************************        
    イベントビルダー
    ****************************  
    1. 新しいスケジュールを作成します。
    2. プログラムを終了します。
    ****************************
""")
            try:
                choice = int(input(">>"))
            except:
                print("「1」または「2」を入力してください。")
            if choice == 1:
                activities = []
                activity_times = []
                title = enter_title()
                leader = enter_leader()
                date = enter_date()
                location = enter_location()
                equipment = enter_equipment()
                confirmation_menu = True
                while confirmation_menu:
                    print(f"""
        確認メニュー
        *****************
        1. アクティビティを追加します。
        2. 詳細を確認します。
        3. 編集します。
        4. ファイルを保存し、メインメニューに戻ります。
        
    """)
                    try:
                        choice = int(input(">>"))
                    except:
                        print("「1」、「2」、「3」または「4」を入力してください。")
                    if choice == 1:
                        enter_activity(activities, activity_times)
                    if choice == 2:
                        if len(activities) >= 1:
                            activity_text = make_activities_text(activity_times, activities)
                        else:
                            activity_text = ["<アクティビティなし>"]
                        activities_str = print_activities(activity_text)
                        print(f"""
        ***********************************************
        タイトル: {title}
        リーダー: {leader}
        日付: {date}       
        場所: {location}
        機器: {equipment}
        アクティビティ: 
        {activities_str}
        ***********************************************
                        """)
                    if choice == 3:
                        edit_menu = True
                        while edit_menu:
                            print("""
        編集メニュー
        *********
        1. タイトル
        2. リーダー
        3. 日付
        4. 場所
        5. 機器
        6. <- 戻る
        
        """)
                            try:
                                choice = int(input(">>"))
                            except:
                                print("「1」、「2」、「3」、「4」、「5」または「6」を入力してください。")
                            if choice == 1:
                                title = enter_title()
                                print(f"タイトルが {title} に更新されました")
                            if choice == 2:
                                leader = enter_leader()
                                print(f"リーダーが {leader} に更新されました")
                            if choice == 3:
                                date = enter_date()
                                print(f"日付が {date} に更新されました")
                            if choice == 4:
                                location = enter_location()
                                print(f"場所が {location} に更新されました")
                            if choice == 5:
                                equipment = enter_equipment()
                                print(f"機器が {equipment} に更新されました")
                            if choice == 6:
                                edit_menu = False
                                break

                    if choice == 4:
                        workbook = xl.load_workbook("template/template.xlsx")
                        sheet = workbook.active
                        sheet['a1'] = f"日付: {date}"
                        sheet['a1'].font = Font(size=11)
                        sheet['b1'] = title
                        sheet['b1'].font = Font(bold=True, size=14)
                        sheet['c1'] = f"リーダー: {leader}"
                        sheet['c1'].font = Font(size=11)
                        sheet['c2'] = f"場所: {location}"
                        sheet['c2'].font = Font(size=11)
                        sheet['c3'] = f"機器: {equipment}"
                        sheet['c3'].font = Font(size=11)

                        for row in range(1, 12):
                            sheet[f'B{row}'].alignment = Alignment(wrap_text=True)
                            sheet[f'B{row}'].font = Font(size=11)

                        for row in [2, 3]:
                            sheet[f'C{row}'].alignment = Alignment(wrap_text=True)
                            sheet[f'C{row}'].font = Font(size=11)

                        top_left_alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)

                        for row in range(1, 12):  # Apply to rows 1 to 11 in column B
                            sheet[f'A{row}'].alignment = top_left_alignment

                        for row in range(2, 12):  # Apply to rows 1 to 11 in column B
                            sheet[f'B{row}'].alignment = top_left_alignment

                        for row in range(1, 12):  # Apply to cells C2 and C3
                            sheet[f'C{row}'].alignment = top_left_alignment

                        cur_entry = 0
                        for time, activity in zip(activity_times, activities):
                            sheet[f'a{2 + cur_entry}'] = time
                            sheet[f'b{2 + cur_entry}'] = activity
                            sheet[f'a{2 + cur_entry}'].alignment = xl.styles.Alignment(wrap_text=True)
                            sheet[f'b{2 + cur_entry}'].alignment = xl.styles.Alignment(wrap_text=True)
                            cur_entry += 1
                        file_name = input("ファイル名を入力: ")
                        workbook.save(f'{file_name}.xlsx')
                        confirmation_menu = False
                        main_menu = True
                        break
            if choice == 2:
                quit()
            else:
                print("無効な入力です。")


if __name__ == "__main__":
    main()
