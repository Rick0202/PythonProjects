import pandas as pd
import sys


def add_number(number):


    # List to store dataframe
    results = []

    workNum = number

    i=0

    header1 = "even"
    header2 = "odd"

    print(f"{'':<20}{'even':<20}{'odd':<15}\n")


    #count = number
    
    while i <= 10:

        
           
        if workNum % 2 == 0:
            #print(f"{workNum} is an even number.")
            print(f"{'':<20}{workNum}")
            results.append({'Number': workNum,'Even': workNum, 'Odd': ''})

        else:
            #print(f"{workNum} is an odd number.")
            print(f"{'':<40}{workNum}")
            results.append({'Number': workNum,'Even': '', 'Odd': workNum})
            

        workNum += 1

        i+=1
        #print(f"\n{i} This is the i count.\n")

    file_name = input("\nEnter the name you want the Excel spreadsheet to be called, don't add the extension: ") + ".xlsx"

    df = pd.DataFrame(results)

    df.to_excel(file_name, index=False)

    print(f"\nThe file has been saved under the name: {file_name}")



# Enter a number to start the program.

number = int(input("Enter a number and the program will start with the entered number and\n and increase that number by 10 and then print\n that those numbers to the screen and to an excel spreadsheet.\n"))

# Call the method

add_number(number)

sys.exit()

 

















