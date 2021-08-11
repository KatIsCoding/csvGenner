from tkinter import filedialog, Entry, simpledialog
import tkinter
import xlrd
import csv
from fuzzysearch import find_near_matches_in_file, find_near_matches
# All the GUI prompts
emails_file = filedialog.askopenfile(title= "Select the file that contains the emails")
payouts_file = filedialog.askopenfilename(title= "Select the file that cointains the payouts information")
output_filename = simpledialog.askstring(title="Enter the output filename", prompt="Output Filename")
output_directory = filedialog.askdirectory(title="Open the directory where you want to store the output file")
info = emails_file.read().splitlines()

for index, value in enumerate(info):
    info[index] = value.strip("\t")
    
book = xlrd.open_workbook(payouts_file)
sh = book.sheet_by_index(0)
imprints = []
totals = []
currency = []

# Populating information | Ignoring csv processing method
for x in range(sh.nrows):
    imprints.append(sh.row(x)[0].value)
    totals.append(sh.row(x)[6].value)
    currency.append(sh.row(x)[7].value)
output_rows = []

# Formatting data
for index, imprint in enumerate(imprints):
    for email in info:
        if len(find_near_matches(imprint, email, max_l_dist=1)) > 0:
            start_point = email.find("-")
            output_rows.append([email[start_point + 2:], str(int(totals[index]) * 0.8), currency[index]])

# Creating and storing output information
with open(f"{output_directory}/{output_filename}.csv", "w") as payout_file:
    payout_writer = csv.writer(payout_file, delimiter=",")
    for x in output_rows:
        payout_writer.writerow(x)
            
        

