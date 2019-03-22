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
        write_list = list()
        for p,s,f in os.walk(root_path):
            print("\nCurrently Searching in ", p)
            for i in f:
                file_name = os.path.join(p, i)
                file_time = os.path.getctime(file_name)
                if file_time > diff:
                    ago = today - d.datetime.fromtimestamp(file_time)
                    if ago.days:
                        ago = "{0} days".format(ago.days)
                    else:
                        ago = "{0} seconds".format(ago.seconds)
                    file_time = ctime(file_time)
                    print(i, "\t\twas created on {0} - {1} ago".format(file_time, ago))
                    temp_list = [file_name, os.path.getsize(file_name), file_time, ctime(os.path.getmtime(file_name)), ctime(os.path.getatime(file_name)), ago]
                    write_list.append(temp_list)
        if len(write_list):
            save = input("\nDo you want to save this result in an Excel Workbook? (y/n)")
            if save == 'y' or save == 'Y':

                import xlwt as x

                wb = x.Workbook()
                ws = wb.add_sheet("List of Files")
                title_list = [["Date of the Exercise : {0}".format(ctime(today.timestamp()))], ["Date {0} days ago : {1}".format(days, ctime(diff))], ["Fully qualified File Name", "File Size", "File Created Time", "Last Modified Time", "Last Accessed Time", "Days passed in between"]]
                for i, row in enumerate(title_list):
                    for j, col in enumerate(row):
                        ws.write(i, j, col, style = x.easyxf('font:bold on'))
                for i, row in enumerate(write_list):
                    for j, col in enumerate(row):
                        ws.write(i+3, j, col)
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
