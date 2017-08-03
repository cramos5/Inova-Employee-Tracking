import os

def Menu_Instructions():

    print("******************************************************************************")
    print("Instructions:\n")
    print("1) Place Excel Files you want to compare in same directory as this program")
    print("***Be sure Excel files have different names\n***Be sure first column in both Excels is the employee id\n")
    print("2) Select Old and New Excel File you want to compare \n")
    print("3) Wait for program to finish\n")
    print("******************************************************************************")

    name = input("Please enter any key to continue")

    os.system('cls')


