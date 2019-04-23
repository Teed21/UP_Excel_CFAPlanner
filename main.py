# Written and Developed by: Tyler Wright
# Email: tylerwright17@yahoo.com / Tyler.Wright@hitachirail.com
# Date started: 12/13/2018
# Date when workable: 02/05/2019
# Last Updated: 04/23/2019

import ExcelReader
import ExcelWriter
import ExcelPlanner
# This is for open dialog window.
import tkinter as tk
from tkinter import filedialog
import os

# These variables set up an invisible window that hosts the file dialog.
root = tk.Tk()
root.withdraw()

# Example of file dialog. Use when necessary for getting files.
#file_path = filedialog.askopenfilename()

spacing = "\n"


# This function helps check if a CP exists before adding it to a list.
def check_if_cp_exists(cfa, cp_name):
    planner = ExcelPlanner.ExcelPlanner(cfa)

    cp_exists = planner.check_if_cp_exists(cp_name)

    return cp_exists


# This function gets all cp names in the selected CFA/Excel doc.
def get_cp_names(cfa):

    planner = ExcelPlanner.ExcelPlanner(cfa)

    cp_names = planner.get_cp_names()

    return cp_names


# This function will handle multiple CFA/ExcelPlanner objects.
def multiple_cfa(cfa, cp_names):
    # Create planner object using cfa name.
    planner = ExcelPlanner.ExcelPlanner(cfa)

    # Generate info based on cp names.
    list_of_gen_cp_info = planner.get_cp_cords(cp_names)
    list_of_cp_cords = list_of_gen_cp_info[0]
    list_of_cp_names = list_of_gen_cp_info[1]

    # Create list of cp objects to manipulate and combine texts from later on.
    cp_objects = planner.get_all_cp_information(list_of_cp_cords, list_of_cp_names)

    # This list is the combination of all text needed for a new CFA.
    lists = planner.return_combined_lists(cp_objects)

    return lists

# This code block gets the amount of CFAs from the user. Will continue to ask until correct input is met.
is_number = False
while is_number is not True:
    try:
        cfa_input = int(input("How many CFAs are there? 1-3\n:"))
        if cfa_input > 3 or cfa_input <= 0:
            print("Too many CFAs. Can only handle 1 to 3 CFAs.", spacing)
        else:
            print(cfa_input, "CFAs will be processed.", spacing)
            is_number = True
    except ValueError:
        print("Data Type incorrect for cfa_input. Must be a number.", spacing)

# This code block gets the names of the CFAs and the CPs related to them.
list_of_planners = []
processed = False
while processed is not True:
    # For every cfa, ask for a file name.
    for cfa_num in range(cfa_input):
        list_of_cps = []
        # User is prompted by pop-up to select a file.
        print("Please select your CFA file.", spacing)
        file_path = filedialog.askopenfilename()
        file_path_list = file_path.split("/")
        cfa_file = file_path_list[-1]
        print("File selected:", cfa_file, " in path:", file_path, spacing)
        # Checking if file exists.
        if os.path.isfile(file_path):
            # Output all possible CP names
            print(get_cp_names(file_path))
            # Asking for all CPs for specific CFA
            print("Type in all the CPs for this CFA file one-by-one. When done, type the letter 'd' to continue."
                  , spacing)
            user_input = ""
            while "d" not in user_input or "D" not in user_input and len(user_input) > 1:
                user_input = input("CP: ")
                if user_input == "":
                    print("Blank provided. Not sending blank to CP stack.")
                elif user_input != "d" and user_input != "D":
                    # Check if CP exists in file.
                    if check_if_cp_exists(file_path, user_input) is False:
                        print("CP name not found in file provided: ", file_path)
                        print("Please enter a different CP name...", spacing)
                        continue
                    list_of_cps.append(user_input)
                else:
                    print("Continuing...", spacing)

            list_of_planners.append(multiple_cfa(file_path, list_of_cps))
        else:
            # This should pull the user back into the while loop, and go through listing CFAs again.
            print("It looks like the file you tried to input does not exist. Try again.", spacing)
            processed = False
            break

        # This will break the user out of the while loop so they can continue.
        processed = True

# Process all planner information into one list for the writer to handle
master_list = []
first_time = False
for planner_list in list_of_planners:
    if first_time is False:
        first_time = True
        master_list += planner_list
    else:
        master_list[0] += planner_list[0]
        master_list[1] += planner_list[1]

# Create new file
print("Please select where you want to save your new CFA file.", spacing)
new_file_path = filedialog.askdirectory()
print("Create new file:\nWrite the request name (EX: S031) and the program will do the rest.")
request_name = input("Request Name: ")
new_file = new_file_path + "/" + request_name + "_REO_Testplan-First.xlsx"

temp_planner = ExcelPlanner.ExcelPlanner("new")

temp_planner.process_and_write_all_cp_information(master_list, new_file)

print("Done.")

"""
planner = ExcelPlanner.ExcelPlanner("OR05CFA.xls")

#planner = ExcelPlanner.ExcelPlanner("NV12CFA.xls", "multiple_cp_test3.xlsx")

list_of_cps = ["N199", "N201", "N211", "N214"]
#list_of_cps = ["F487", "F488", "F497"]

list_of_gen_cp_info = planner.get_cp_cords(list_of_cps)
list_of_cp_cords = list_of_gen_cp_info[0]
list_of_cp_names = list_of_gen_cp_info[1]

cp_objects = planner.get_all_cp_information(list_of_cp_cords, list_of_cp_names)

lists = planner.return_combined_lists(cp_objects)

planner.process_and_write_all_cp_information(lists, "multiple_cp_test6.xlsx")

"""