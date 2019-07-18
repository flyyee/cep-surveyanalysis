import openpyxl
import statistics
from collections import Counter
import warnings
warnings.filterwarnings(action='ignore', category=UserWarning, module='gensim')
# filters warning that importing gensim for the first time causes
from gensim.summarization import keywords
import datetime
import matplotlib.pyplot as plt
import os
import sys


def is_number(s):  # tests if input is a number
    try:
        float(s)
        return True
    except ValueError:
        return False


class result:  # result class, contains data about each column in an excel sheet, aka one result object per question
    def __init__(self):  # initialises values to be used later
        self.mean = 0
        self.median = 0
        self.mode = None
        self.optcount = None
        self.isnumericaldataset = True
        self.isdatetimedataset = False
        self.datasetname = ""
        self.datasetlength = 0
        self.keywords = []

    def getdata(self, data):  # gets the name of the column, aka the question from the entire column of data
        self.datasetname = data[0]
        for x in range(1, len(data)):
            if not is_number(str(data[x])):  # also checks if the column, aka dataset only contains numbers
                self.isnumericaldataset = False
                break  # checks every element in a dataset, or until it is
                # established that the dataset does not only contain numbers
        for x in range(1, len(data)):
            if isinstance(data[x], datetime.datetime):  # also checks if the dataset contains datetime objects
                self.isdatetimedataset = True
                break  # checks every element in a dataset, or until it is
                # established that the dataset does not contain datetime objects
        # checking the types on inputs in a dataset are necessary for
        # giving users the correct options to view the data from
        self.datasetlength = len(data) - 1  # first element of a dataset is
        # the question, so the actual length of the data in a set is one less

    def revealmode(self, verbose=True):  # prints the mode
        print("Mode: " + str(self.mode))
        print("\n")

    def revealmeanandmedianandmode(self, verbose=True):  # prints the mean and median and mode
        print("Mean: " + str(self.mean))
        print("Median: " + str(self.median))
        print("Mode: " + str(self.mode))
        print("\n")
        if verbose:  # only displays graphs if verbose switch is not false
            labels = ["Mean", "Median", "Mode"]
            values = [float(self.mean), float(self.median), float(self.mode)]
            plt.bar(labels, values)
            plt.title("Mean, median and mode of options for: " + self.datasetname + "\n \n")
            plt.show()

        # Note: separate functions are needed for mean, median and mode and mode alone as non-numerical datasets may
        # have a mode, but will not have a mean or median, while numerical datasets contain all three

    def revealoptcount(self, verbose=True):  # prints the popularity of options
        print("Popularity of options: ")
        for key in self.optcount.keys():  # this is a counter object that has a dictionary of a
            # value in the dataset and the corresponding number of times it is present in the dataset
            # the keys method contains all the keys, while values contains all the values in the dictionary
            percent = int(self.optcount[key]) / self.datasetlength * 100
            # calculates the percentage that option weights
            # the loop iterates over every key and gets the corresponding number of times it is present
            # in the list by retrieving it from the counter dictionary
            percent = '%.2f' % percent  # rounds the percentage to two decimal places
            print("Option " + str(key) + ": " + str(self.optcount[key]) +
                  " times: " + str(percent) + "% of total responses")
        print("\n")
        if verbose:  # only displays graphs if verbose switch is not false
            exp_vals = self.optcount.values()
            exp_labels = self.optcount.keys()
            plt.axis("equal")
            plt.pie(exp_vals, labels=exp_labels, autopct='%1.1f%%', radius=1.2)
            # shows the percentage in the pie chart as well
            plt.title("Popularity of options for: " + self.datasetname + "\n \n")
            plt.show()

    def revealkeywords(self, verbose=True):
        print("Keywords: ")
        if self.keywords == "Too few keywords. Index error." or self.keywords == "No keywords for a datetime dataset."\
        or self.keywords == "No keywords. List empty.":
            print(self.keywords)
        else:
            for kw in self.keywords:  # this contains a keyword object that contains a two-dimensional
                # list of each keyword and each corresponding score
                print("Score of keyword " + kw[0] + ": " + str(kw[1]))
        print("\n")

    def set(self, data):
        del data[0]
        if not self.isnumericaldataset:  # if the dataset is non-numerical, a mean cannot be derived from it
            self.mean = "Unable to get mean from dataset. Type incompatible."
        else:
            self.mean = statistics.mean(data)
            self.mean = '%.2f' % float(self.mean)  # rounds mean to two decimal places

        if not self.isnumericaldataset:  # if the dataset is non-numerical, a median cannot be derived from it
            self.median = "Unable to get median from dataset. Type incompatible."
        else:
            self.median = statistics.median(data)

        try:
            self.mode = statistics.mode(data)
        except statistics.StatisticsError:  # catches errors if the dataset is unsuitable for a mode
            self.mode = "No unique mode."
        except TypeError:  # catches errors if the dataset is unsuitable for a mode
            self.mode = "Unable to get mode from dataset. Type incompatible."

        if not self.isnumericaldataset and not self.isdatetimedataset:
            # only accepts datasets that are strings and not datetimes or integers
            stringifieddata = " "
            for word in data:
                stringifieddata += word + " "  # converts the dataset into a plaintext form from a list
            try:
                self.keywords = keywords(text=stringifieddata, words=10, scores=True, lemmatize=True)
                if not self.keywords:  # returns false if list is empty
                    self.keywords = "No keywords. List empty."
            except IndexError:  # catches error if there are too few keywords
                self.keywords = "Too few keywords. Index error."
        elif self.isdatetimedataset:  # no keywords are present in a datetime dataset
            self.keywords = "No keywords for a datetime dataset."

        self.optcount = Counter(data)  # creates a counter object that counts
        # how many times each unique instance of a value in a list is present

    def exportall(self, path, datasetnum):
        try:
            f = open(path, "a+")  # opens file provided in path, appends, and creates if non-existent
            stdout = sys.stdout  # keeps original stdout
            sys.stdout = f  # pipes system output from console to file in f
            print("Dataset " + str(datasetnum) + ": " + self.datasetname)
            # reveals in non-verbose mode so matplotlib graphs do not show up
            if self.isnumericaldataset:
                self.revealmeanandmedianandmode(False)
                self.revealoptcount(False)
            else:
                self.revealkeywords(False)
                self.revealmode(False)
                self.revealoptcount(False)
            print("\n\n\n")
        finally:
            f.close()  # closes file
            sys.stdout = stdout  # reset stdout to original

def export_main(results):
    currentdir = os.getcwd()  # gets the directory the program was started in
    currentdatetime = datetime.datetime.now()  # gets the current date and time
    currentdatetime = str(currentdatetime).replace(':', "-")  # replaces : in the time with -
    # so that is in an acceptable file name
    newfolderpath = os.path.join(currentdir, str(currentdatetime))  # creates the new directory path
    if not os.path.exists(newfolderpath):  # checks if the path exists
        os.makedirs(newfolderpath)  # creates the folder

    newfolderpath += "\\analysis.txt"  # edits the path to include the file name
    for x, result in enumerate(results):
        result.exportall(newfolderpath, x)



while True:  # keeps prompting user for the file name until one that exists is entered
    wbname = input("What is the name of the excel workbook? (.xls or .xlsx): ")
    try:
        with open(wbname) as f:
            break  # exits the loop
    except FileNotFoundError:  # checks if excel workbook exists
        print("File does not exist. Check if you entered the right extension.")

wb = openpyxl.load_workbook(wbname)

while True:  # keeps prompting user for the sheet name until an existing one is enter
    sheetname = input("What is the name of the sheet?: ")
    if sheetname in wb.get_sheet_names():
        sheet = wb.get_sheet_by_name(sheetname)  # loads the excel sheet
        break  # exits the loop
    else:
        print("No such sheet. Check the capitalisation.")

print("Parsing " + sheetname + " from " + wbname)

numdatasets = len(sheet['1'])  # gets the number of datasets, or columns by calculating the length of column 1
numrows = len(sheet['A'])  # gets the number of rows by calculating the length of row a

data = []
results = [result() for x in range(numdatasets)]  # creates as many result objects as there are datasets

for x in range(1, numdatasets + 1):  # loops over the datasets, or columns
    data.clear()  # clears the dataset from the previous run

    cures = results[x - 1]  # currentresult

    for n in range(1, numrows + 1):  # loops over each row in a dataset
        data.append(sheet.cell(row=n, column=x).value)  # appends value of the cell into the dataset

    # calls the methods of the current result object to use the dataset
    cures.getdata(data)
    cures.set(data)

# the loop over the sheet is complete
print("Parsing of " + str(numdatasets) + " datasets complete!")

# export to txt document
export_main(results)

# keeps prompting the user until user exits
while True:
    print("Which dataset would you like to view more information about?")
    for x in range(numdatasets):
        print("Dataset " + str(x + 1) + ": " + results[x].datasetname)  # prints all datasets
    print("\nTyping exit exits the program.")

    datasetchoice = input("Option: ")
    # sanitises user input
    if is_number(datasetchoice):
        datasetchoice = int(datasetchoice)
        if datasetchoice in range(1, numdatasets + 1):  # checks if choice corresponds to one of the datasets available
            while True:
                print("Choice: Dataset " + str(datasetchoice) + ": " + results[datasetchoice - 1].datasetname)
                # user input will be one more than the element number of the result object in the results list
                if results[datasetchoice - 1].isnumericaldataset:
                    print("This is a numerical dataset.")
                    print("Viewing mean, median and mode (1) are recommended.")
                    print("Viewing the number of responses (2) per option is also available.")
                    print("Typing change brings you back to the dataset selection menu.")
                    choice = input()
                    if is_number(choice):  # sanitises input
                        choice = int(choice)
                        if choice in range(1, 3):
                            if choice == 1:
                                results[datasetchoice - 1].revealmeanandmedianandmode()
                            elif choice == 2:
                                results[datasetchoice - 1].revealoptcount()
                    elif choice == "change":
                        break  # breaks inner loop, redirects to outer loop asking for the dataset choice
                    else:
                        print("Unknown option.")

                else:
                    print("This dataset contains strings.")
                    print("Viewing summary of keywords (1) is recommended.")
                    print("Viewing number of responses per option (2) or the mode response (3) are also available.")
                    choice = input("Option: ")
                    if is_number(choice):  # sanitises input
                        choice = int(choice)
                        if choice in range(1, 4):
                            if choice == 1:
                                results[datasetchoice - 1].revealkeywords()
                            elif choice == 2:
                                results[datasetchoice - 1].revealoptcount()
                            elif choice == 3:
                                results[datasetchoice - 1].revealmode()
                    elif choice == "change":
                        break  # breaks inner loop, redirects to outer loop asking for the dataset choice
                    else:
                        print("Unknown option.")
    elif datasetchoice == "exit":
        break  # breaks the outer loops, exits program
    else:
        print("Unknown option.")
print("Program exiting.")
