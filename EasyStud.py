import pyfiglet
import sys
import os
import pandas as pd
from tabulate import tabulate
from colorama import init, Fore, Back, Style

# Initialize colorama
init(autoreset=True)

# Print the logo and menu
logo = pyfiglet.figlet_format("EASYSTUD")
print(logo)
print(Fore.YELLOW + "📃 A script built to help you with student records 📃")

menu_message = """What would you like to do today??
--> Add more students: new
--> Add new data to a student: add
--> Get student detail: get
--> Remove student: rmv
--> Do basic maths: math
--> Get this guide: help
--> Exit the program: exit"""
print(Fore.CYAN + menu_message)

def load_students_from_excel(filename="data.xlsx"):
    if os.path.exists(filename):
        return pd.read_excel(filename).to_dict(orient='records')
    return []

def save_to_excel(students, filename="data.xlsx"):
    try:
        df_combined = pd.DataFrame(students)
        df_combined.to_excel(filename, index=False)
    except PermissionError:
        print(Fore.RED + f"Permission denied: '{filename}'. Please close the file if it is open and try again.")
    except Exception as e:
        print(Fore.RED + f"An error occurred while saving to Excel: {e}")

def calculate_student_marks(student):
    marks = {
        'English': student['English'], 'Hindi': student['Hindi'], 
        'SST': student['SST'], 'Science': student['Science'], 
        'Maths': student['Maths'], 'Computer': student['Computer']
    }
    return marks

students = load_students_from_excel()
print()
while True:
    print(Style.BRIGHT + Back.RED + Fore.BLACK + "-> Enter your choice: ", end="")
    choice = input(Style.BRIGHT + Fore.BLUE).strip().lower()
    print()
    if choice == "new":
        print(Fore.YELLOW + "📃📃 YOU ARE GOING TO CREATE A LIST")
        name = input(Style.BRIGHT + Fore.RED + "Enter Student Name: " + Style.BRIGHT + Fore.BLUE).strip()
        fname = input(Style.BRIGHT + Fore.RED + "Enter Student's Father Name: " + Style.BRIGHT + Fore.BLUE).strip()
        age = int(input(Style.BRIGHT + Fore.RED + "Enter Student Age: " + Style.BRIGHT + Fore.BLUE).strip())
        marks_input = input(Style.BRIGHT + Fore.RED + "Enter marks in this format - English, Hindi, SST, Science, Maths, Computer: " + Style.BRIGHT + Fore.BLUE).strip()
        marks = list(map(float, marks_input.split(',')))
        if len(marks) != 6:
            print(Fore.RED + "Error: Please enter marks for all six subjects.")
            continue
        student = {
            'Name': name, 'Father Name': fname, 'Age': age, 
            'English': marks[0], 'Hindi': marks[1], 'SST': marks[2], 
            'Science': marks[3], 'Maths': marks[4], 'Computer': marks[5]
        }
        students.append(student)
        print()
        print(Style.BRIGHT + Fore.BLACK + Back.GREEN + "==> STUDENT ADDED SUCCESSFULLY:")
        print(tabulate([student], headers="keys", tablefmt="grid"))
        save_to_excel(students)
        print()  # Adding a line gap

    elif choice == "exit":
        print(Fore.YELLOW + "==> Exiting the program.")
        sys.exit()

    elif choice == "help":
        print(Fore.CYAN + menu_message)
        print()  # Adding a line gap

    elif choice == "add":
        which_student = input(Style.BRIGHT + Fore.RED + "Enter name of the student: " + Style.BRIGHT + Fore.BLUE).strip()
        found = False
        for student in students:
            if student["Name"].lower() == which_student.lower():
                found = True
                print(Fore.GREEN + "==> Student found!!")
                what_add = input(Style.BRIGHT + Fore.RED + "Okay! What to add (e.g., 'Age:20', 'Maths:95'): " + Style.BRIGHT + Fore.BLUE).strip()
                key, value = what_add.split(':')
                if key in student:
                    student[key] = type(student[key])(value)
                    print(Fore.GREEN + f"==> Updated {which_student}'s {key} to {value}.")
                else:
                    print(Fore.RED + f"==> Invalid attribute '{key}'.")
                break
        if not found:
            print(Fore.RED + "==> Student not found.")
        save_to_excel(students)
        print()  # Adding a line gap

    elif choice == "get":
        get_data = input(Style.BRIGHT + Fore.RED + "Enter the student name: " + Style.BRIGHT + Fore.BLUE).strip()
        for student in students:
            if student["Name"].lower() == get_data.lower():
                print(Fore.GREEN + f"Here's {get_data}'s detail:")
                print(tabulate([student], headers="keys", tablefmt="grid"))
                break
        else:
            print(Fore.RED + "==> No Student record was found! Try again")
        print()  # Adding a line gap

    elif choice == "math":
        student_name = input(Style.BRIGHT + Fore.RED + "Enter the student name: " + Style.BRIGHT + Fore.BLUE).strip()
        found_student = None
        for student in students:
            if student["Name"].lower() == student_name.lower():
                found_student = student
                break
        
        if found_student:
            print(Fore.GREEN + "==> Student found!!")
            marks = calculate_student_marks(found_student)
            print(Fore.CYAN + f"Marks of {student_name}:")
            print(tabulate([marks], headers="keys", tablefmt="grid"))
            total_marks = int(input(Style.BRIGHT + Fore.RED + "Enter total marks of the test: " + Style.BRIGHT + Fore.BLUE).strip())
            include_computer = input(Style.BRIGHT + Fore.RED + "Include Computer subject in percentage calculation? (yes/no): " + Style.BRIGHT + Fore.BLUE).strip().lower()
            total_marks_exclude_computer = total_marks - marks['Computer'] if include_computer == 'no' else total_marks
            total_marks_include_computer = total_marks if include_computer == 'yes' else total_marks_exclude_computer
            total_obtained_marks = sum(marks.values())
            percentage_exclude_computer = (total_obtained_marks / total_marks_exclude_computer) * 100
            percentage_include_computer = (total_obtained_marks / total_marks_include_computer) * 100

            print(Fore.CYAN + "Percentage Calculation:")
            print(Fore.CYAN + f"Total Obtained Marks: {total_obtained_marks}")
            print(Fore.CYAN + f"Total Marks (excluding Computer): {total_marks_exclude_computer}")
            print(Fore.CYAN + f"Percentage (excluding Computer): {percentage_exclude_computer}%")
            if include_computer == 'yes':
                print(Fore.CYAN + f"Total Marks (including Computer): {total_marks_include_computer}")
                print(Fore.CYAN + f"Percentage (including Computer): {percentage_include_computer}%")

            found_student['Total Marks'] = total_obtained_marks
            found_student['Percentage (excluding Computer)'] = percentage_exclude_computer
            if include_computer == 'yes':
                found_student['Percentage (including Computer)'] = percentage_include_computer
            
            save_to_excel(students)
        else:
            print(Fore.RED + "==> Student not found.")
        print()  # Adding a line gap
