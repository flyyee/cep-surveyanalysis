import xlrd
import csv

def csv_from_excel():
    wb = xlrd.open_workbook('responses.xlsx')
    sh = wb.sheet_by_name('Form responses 1')
    your_csv_file = open('csv.csv', 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()

# runs the csv_from_excel function:
csv_from_excel()
with open('csv.csv', 'r') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=' ', quotechar='|')
    count = 0
    for row in spamreader:
        count += 1
        print(count)
        print(", " + str(row))