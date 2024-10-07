import openpyxl as xl
from openpyxl.styles import Font, Alignment


def enter_title():
    title = input("Event Title: ")
    return title


def enter_leader():
    leader = input("Leader/MC: ")
    return leader


def enter_date():
    date = input("Date: ")
    return date


def enter_location():
    location = input("Location: ")
    return location


def enter_equipment():
    equipment = input("List all equipment: ")
    return equipment


def enter_activity(activities, activity_times):
    time = input("Time: ")
    activity = input("Enter activity: ")
    activities.append(activity)
    activity_times.append(str(time))
    return activities, activity_times


def make_activities_text(activity_times, activities):
    combined_list = ["time: " + a + "\n" + "Activity: " + b for a, b in zip(activity_times, activities)]
    return combined_list


def print_activities(activity_text):
    return "\n".join(activity_text)


def main():
    run = True
    main_menu = True
    choice = 0
    while run:
        while main_menu:
            print(f"""
    **********************        
    Event Schedule Builder
    **********************        
    1. Build new schedule.
    2. Exit Program.
    **********************
""")
            try:
                choice = int(input(">>"))
            except:
                print("Please enter '1' or '2'.")
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
        Confirmation Menu
        *****************
        1. Add activity.
        2. Check details.
        3. Edit.
        4. Save file, and quit to main menu.
        
    """)
                    try:
                        choice = int(input(">>"))
                    except:
                        print("Please enter '1', '2', '3' or '4'.")
                    if choice == 1:
                        enter_activity(activities, activity_times)
                    if choice == 2:
                        if len(activities) >= 1:
                            activity_text = make_activities_text(activity_times, activities)
                        else:
                            activity_text = ["<No Activities>"]
                        activities_str = print_activities(activity_text)
                        print(f"""
        ***********************************************
        Title: {title}
        Leader: {leader}
        Date: {date}       
        Location: {location}
        Equipment: {equipment}
        Activities:
        {activities_str}
        ***********************************************
                        
                        """)
                    if choice == 3:
                        edit_menu = True
                        while edit_menu:
                            print("""
        Edit Menu
        *********
        1. Title
        2. Leader
        3. Date
        4. Location
        5. Equipment
        6. <- Back
        
        """)
                            try:
                                choice = int(input(">>"))
                            except:
                                print("Please enter '1', '2', '3', '4', '5' or '6'.")
                            if choice == 1:
                                title = enter_title()
                                print(f"Title updated to {title}")
                            if choice == 2:
                                leader = enter_leader()
                                print(f"Leader updated to {leader}")
                            if choice == 3:
                                date = enter_date()
                                print(f"Date updated to {date}")
                            if choice == 4:
                                location = enter_location()
                                print(f"Location updated to {location}")
                            if choice == 5:
                                equipment = enter_equipment()
                                print(f"Equipment updated to {equipment}")
                            if choice == 6:
                                edit_menu = False
                                break

                    if choice == 4:
                        workbook = xl.load_workbook("template/template.xlsx")
                        sheet = workbook.active
                        sheet['a1'] = f"Date: {date}"
                        sheet['a1'].font = Font(size=11)
                        sheet['b1'] = title
                        sheet['b1'].font = Font(bold=True, size=14)
                        sheet['c1'] = f"Leader: {leader}"
                        sheet['c1'].font = Font(size=11)
                        sheet['c2'] = f"Location: {location}"
                        sheet['c2'].font = Font(size=11)
                        sheet['c3'] = f"Equipment: {equipment}"
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
                        file_name = input("Enter file name: ")
                        workbook.save(f'{file_name}.xlsx')
                        confirmation_menu = False
                        main_menu = True
                        break
            if choice == 2:
                quit()
            else:
                "Invalid input."


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"There was an error: {e}")
    input("press enter to exit.")
