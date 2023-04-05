import sys
import subprocess
import os
import csv
import glob
import time

try:
   from PIL import Image, ImageDraw, ImageFont
except:
    subprocess.run(['pip','install','Pillow'])
    from PIL import Image, ImageDraw, ImageFont

try:
   import openpyxl
except:
   subprocess.run(['pip','install','openpyxl'])
   import openpyxl

# function to read names from CSV file
def read_csv_file(file_path):
    names = []
    with open(file_path, 'r') as file:
        reader = csv.reader(file)
        for row in reader:
            names.append(row[0])
    return names

# function to read names from XLSX file
def read_xlsx_file(file_path):
    names = []
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
    for row in ws.iter_rows(min_row=1, max_row=1):
        for cell in row:
            names.append(cell.value)
    return names

def coupons(names: list, certificate: str, font_path: str,op_directory):
   
    for name in names:
          
        # adjust the position according to 
        # your sample
        text_y_position = 480
   
        # opens the image
        img = Image.open(certificate, mode ='r')
          
        # gets the image width
        image_width = img.width
          
        # gets the image height
        image_height = img.height 
   
        # creates a drawing canvas overlay 
        # on top of the image
        draw = ImageDraw.Draw(img)
   
        # gets the font object from the 
        # font file (TTF)
        font = ImageFont.truetype(
            font_path,
            180 # change this according to your needs
        )
   
        # fetches the text width for 
        # calculations later on
        text_width = font.getlength(name)
   
        draw.text(
            (
                # this calculation is done 
                # to centre the image
                (image_width - text_width) / 2,
                text_y_position
            ),
            name,
            font = font        )
   
        # saves the image in png format
        img.save(f"./output/{op_directory}/{name}.png")
    print("Task Completed successfully")
    time.sleep(10)
  
# Driver Code
if __name__ == "__main__":
   # Create Output file
   opd = input("Enter Output Directory Name : ")
   try:
      os.mkdir("./output/" + opd)
      print("Directory created successfully!")
   except OSError:
      print("Directory already exists or couldn't be created!")

      
   # Get Name SpreadSheet
   xlsx_files = glob.glob(os.path.join("./data source/", '*.xlsx'))
   csv_files = glob.glob(os.path.join("./data source/", '*.csv'))
   data_list = xlsx_files + csv_files
   data_list_disp = [s.replace("./data source\\",'') for s in data_list ]
   print("List of available Data Sources : ")
   for i,s in enumerate(data_list_disp):
      print(i,") ",s)
   fileno = int(input("Enter the file number to be selected : "))
   filepath = f'./data source/{data_list_disp[fileno]}'
   print("Selected: ",filepath)
    
   if filepath.endswith('.csv'):
      NAMES = read_csv_file(filepath)
   elif filepath.endswith('.xlsx'):
      NAMES = read_xlsx_file(filepath)
   else:
      print('Unsupported file format')
      
   # path to font
   FONT = "C:/Windows/Fonts/ITCEDSCR.ttf"
      
   # path to sample certificate
   pngfiles = glob.glob(os.path.join("./designs/", '*.png'))
   pngfiles_disp = [s.replace("./designs\\",'') for s in pngfiles ]
   print("List of available Designs Sources : ")
   for i,s in enumerate(pngfiles_disp):
      print(i,") ",s)
   imgno = int(input("Enter the design number to be selected : "))
   CERTIFICATE = f'./designs/{pngfiles_disp[imgno]}'
   print("Selected: ",CERTIFICATE)
   
   coupons(NAMES, CERTIFICATE, FONT,opd)
