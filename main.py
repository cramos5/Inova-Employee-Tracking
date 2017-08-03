import queue


from Excel import *
from Menu import *
import datetime
from multiprocessing import Queue

Menu_Instructions()

# Loading Excel Files
oldfile, newfile = getXLfiles()
oldfile = oldfile
newfile = newfile

# Open Excel Files
OldExcel = Excel(oldfile)
NewExcel = Excel(newfile)

# Set Up Changes Tracked File
currentdata = x = datetime.datetime.now().strftime("%m.%d.%Y")
ChangeExcel = Excel(currentdata + " Changes Tracked")


ChangeExcel.setupfile()


print("Loading", oldfile, "....")
OldExcel.loadfile()
OldExcel.getIDs()
print(oldfile, "has been loaded!\n")
print("Loading", newfile, "....")
NewExcel.loadfile()
NewExcel.getIDs()
print(newfile, "has been loaded!\n")


# Get keys from New Excel (Employee ID)

for key in NewExcel.ids.keys():

    newrow = NewExcel.ids[key]

    try:
        oldrow = OldExcel.ids[key]
    except KeyError:
        # Key not not present in old excel, meaning that its a new entry
        ChangeExcel.add_new(newrow)
        continue

    if oldrow != newrow:
        Change = []

        oldrow.insert(0, "Old")
        newrow.insert(0, "New")

        for i in range(0,len(oldrow)):
            if oldrow[i] == newrow[i]:
                continue
            else:
                Change.append(i)
        ChangeExcel.add_updated(oldrow, newrow, Change)

# Checking for rows that are present in old, but not in new
for key in OldExcel.ids.keys():
    oldrow = OldExcel.ids[key]
    try:
        newrow = NewExcel.ids[key]
    except KeyError:
        ChangeExcel.add_removed(oldrow)
        continue

# Write Excel File
try:
    ChangeExcel.savefile()
    print("Comparison is complete\nExcel File has been written")
    print("Program will close in 10 seconds")
    time.sleep(10)

except:
    print("Can't save file. You must close the 'Changes Tracked' excel file!")
    print("Program will close in 10 seconds")
    time.sleep(10)





