import datetime as d
from time import ctime
import os

while True:
    print("\t\tFile Forensics\n\t\t____ _________\n\n")
    root_path = input("Enter the path        : ")
    while True:
        while True:
            try:
                days = int(input("Enter number of days  : "))
                today = d.datetime.now()
                diff = (today - d.timedelta(days = days)).timestamp()
            except ValueError:
                print("Invalid number of days (NaN). Please enter a valid number")
                continue
            except OverflowError:
                print("Your input goes way too back in the past. Please enter a lesser value")
                continue
            else:
                break
        time_list = list()
        file_list = list()
        access_list = list()
        mod_list = list()
        size_list = list()
        for p,s,f in os.walk(root_path):
            print("\nCurrently Searching in ", p)
            for i in f:
                file_name = os.path.join(p, i)
                file_time = os.path.getctime(file_name)
                if file_time > diff:
                    print(i, "\t\twas created on", ctime(file_time))
                    time_list.append(file_time)
                    file_list.append(file_name)
                    access_list.append(ctime(os.path.getatime(file_name)))
                    mod_list.append(ctime(os.path.getmtime(file_name)))
                    size_list.append(os.path.getsize(file_name))
        if len(time_list):
            save = input("\nDo you want to save this result in an Excel Workbook? (y/n)")
            if save == 'y' or save == 'Y':

                import xlwt as x

                wb = x.Workbook()
                ws = wb.add_sheet("List of Files")
                style = x.easyxf('font:bold on')
                ws.write(0, 0, "Date of the Exercise : {0}".format(ctime(today.timestamp())), style = style)
                ws.write(1, 0, "Date {0} days ago : {1}".format(days, ctime(diff)), style = style) 
                ws.write(2, 0, "Fully qualified File Name", style = style)
                ws.write(2, 1, "File Size", style = style)
                ws.write(2, 2, "File Created Time", style = style)
                ws.write(2, 3, "Last Modified Time", style = style)
                ws.write(2, 4, "Last Accessed Time", style = style)
                ws.write(2, 5, "Days passed in between", style = style)
                for i in range(len(time_list)):
                    ws.write(i+3, 0, file_list[i])
                    ws.write(i+3, 1, "{0} bytes".format(size_list[i]))
                    ws.write(i+3, 2, ctime(time_list[i]))
                    ws.write(i+3, 3, mod_list[i])
                    ws.write(i+3, 4, access_list[i])
                    ago = today - d.datetime.fromtimestamp(time_list[i])
                    if ago.days:
                        ws.write(i+3, 5, "{0} days".format(ago.days))
                    else:
                        ws.write(i+3, 5, "{0} seconds".format(ago.seconds))
                save_path = input("Enter the save directory                     : ")
                while True:
                    try:
                        save_name = input("Enter the workbook name (Without extension)  : ")
                        wb.save(os.path.join(save_path, save_name) + ".xls")
                    except OSError:
                        print("Can't use special characters while naming files [or] File is currently in use. Please rename your file")
                        continue
                    else:
                        break
        else:
            print("\nNo files created in {0} during the past {1} days".format(root_path, days))
        cont = input("\nDo you want to check for files created within a different period in the same path? (y/n)")
        if cont == 'y' or cont == 'Y':
            continue
        else:
            break
    cont = input("\nDo you want to check in a different path? (y/n)")
    if cont == 'y' or cont == 'Y':
        continue
    else:
        break
