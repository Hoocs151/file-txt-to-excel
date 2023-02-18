import glob
import os
import win32com.client

os.chdir(os.path.dirname(os.path.abspath(__file__)))

xl = win32com.client.Dispatch("Excel.Application")

workbook = xl.Workbooks.Add()
worksheet = workbook.Sheets("Sheet1")

txt_files = glob.glob("*.txt")
num_files = len(txt_files)
for i, file in enumerate(txt_files):
    with open(file) as f:
        lines = f.readlines()
        data = [line.strip().split(":") for line in lines]

    for j, row in enumerate(data):
        worksheet.Cells(j+1, 1).Value = row[0]
        worksheet.Cells(j+1, 2).NumberFormat = "@"
        worksheet.Cells(j+1, 2).Value = "" + row[1]
        worksheet.Cells(j+1, 3).Value = row[2] + ' ' + row[3]

    output_filename = os.path.join(os.getcwd(), os.path.splitext(os.path.basename(file))[0] + ".xlsx")
    workbook.SaveAs(output_filename)
    print(f""" 
 ▄  █ ████▄ ████▄ ▄█▄      ▄▄▄▄▄   
█   █ █   █ █   █ █▀ ▀▄   █     ▀▄ 
██▀▀█ █   █ █   █ █   ▀ ▄  ▀▀▀▀▄   
█   █ ▀████ ▀████ █▄  ▄▀ ▀▄▄▄▄▀    
   █              ▀███▀            
  ▀                                
                                    """)
    print(f"{i+1} trên {num_files}: {file} -> {output_filename}")
workbook.Close()
xl.Quit()
