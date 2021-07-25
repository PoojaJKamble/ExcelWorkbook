import xlsxwriter
import datetime

drive_letter = r'C:\\'
folder_name = r'Users\pkamble7\PycharmProjects'
folder_time = datetime.datetime.now().strftime("%m-%d-%Y")
folder_to_save_files = drive_letter + folder_name + folder_time

workbook = xlsxwriter.Workbook(folder_time + 'Billing.xlsx')

worksheet1 = workbook.add_worksheet("Summary")

worksheet1.set_column(0, 0, 20)
worksheet1.set_column(1, 1, 20)
worksheet1.set_column(2, 2, 15)
worksheet1.set_column(3, 4, 5)
worksheet1.set_column(5, 5, 15)
worksheet1.set_column(6, 6, 6)
worksheet1.set_column(7, 7, 15)
worksheet1.set_column(8, 8, 15)
worksheet1.set_column(10, 10, 19)

bold = workbook.add_format({'bold': True})
full_border = workbook.add_format(
    {
        "border" : 1,
        "border_color" : "#000000"
    }
)
worksheet1.write(
    'A1',
    'B1',
    full_border
)


size = workbook.add_format({'size':20})


worksheet1.write('A1','B1','C1' "Billing Volume" )
worksheet1.write('A1', "First Name", size)
worksheet1.write('A1', "First Name", bold)
worksheet1.write('B1', "Last Name", size)
worksheet1.write('B1', "Last Name", bold)
worksheet1.write('C1', 'E-mail ID', size)
worksheet1.write('C1', 'E-mail ID', bold)

print("Workbook created Successfully!")

workbook.close()
