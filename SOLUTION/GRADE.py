import pandas as pd
import os
import xlsxwriter

# Reading the csv file
df_new = pd.read_csv('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Maths.csv')

# saving xlsx file
GFG = pd.ExcelWriter('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Maths.xlsx')
df_new.to_excel(GFG, index=False)

GFG.save()

# Reading the csv file
df_new = pd.read_csv('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\English.csv')

# saving xlsx file
GFG = pd.ExcelWriter('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\English.xlsx')
df_new.to_excel(GFG, index=False)

GFG.save()

# Reading the csv file
df_new = pd.read_csv('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Economics.csv')

# saving xlsx file
GFG = pd.ExcelWriter('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Economics.xlsx')
df_new.to_excel(GFG, index=False)

GFG.save()

# Reading the csv file
df_new = pd.read_csv('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Accountancy Result.csv')

# saving xlsx file
GFG = pd.ExcelWriter('C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\Accountancy Result.xlsx')
df_new.to_excel(GFG, index=False)

GFG.save()

# Reading the csv file
df_new = pd.read_csv('C:\\Users\\USER\\PycharmProjects\\pythonProject2\\daas\\Business Studies.csv')

# saving xlsx file
GFG = pd.ExcelWriter('C:\\Users\\USER\\PycharmProjects\\pythonProject2\\daas\\Business Studies.xlsx')
df_new.to_excel(GFG, index=False)

GFG.save()

# import pandas package
import pandas as pd
import numpy as np
import smtplib
import smtplib as smtp

# Location or Folder wher your files are stored e.g "C:\\", "D:\\" or "C:\\user\\user_name\\Documents"

fileFolder = "C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas\\"
excelWorkbookName = "C:\\Users\\USER\\ycharmProjects\\pythonProject2\\daas\\Grade.xlsx"


# read "Accountancy Result.csv" into datframe
dfAccountancyResultCSV = pd.read_csv(fileFolder + "Accountancy Result.csv")

print("Accountancy Result")
print("------------------")
print(dfAccountancyResultCSV.to_string())
print("")

# read "Business Studies.csv" into dataframe
dfBusinessStudiesCSV = pd.read_csv(fileFolder + "Business Studies.csv")
print("Business Studies")
print("----------------")
print(dfBusinessStudiesCSV.to_string())
print("")

# read "Economics.csv" into dataframe
dfEconomicsCSV = pd.read_csv(fileFolder + "Economics.csv")
print("Economics")
print("---------")
print(dfEconomicsCSV.to_string())
print("")

# read "English.csv" into dataframe
dfEnglishCSV = pd.read_csv(fileFolder + "English.csv")
print("English")
print("-------")
print(dfEnglishCSV.to_string())
print("")

# read "Maths.csv" into dataframe
dfMathsCSV = pd.read_csv(fileFolder + "Maths.csv")
print("Maths")
print("-----")
print(dfMathsCSV.to_string())
print("")

# Create an array of data_frames created from individually read CSV files
data_frames_CSV = [dfAccountancyResultCSV, dfEnglishCSV, dfEconomicsCSV, dfMathsCSV, dfBusinessStudiesCSV]

# Concatenate all the marks for each student into one file
Grade = pd.concat(data_frames_CSV)
print("Grade")
print("---------")
print(Grade.to_string())
print("")

# Rename Accountancy "Mark" Column label to "Accountancy"
dfAccountancyResult = dfAccountancyResultCSV.copy()
dfAccountancyResult.rename(columns={'Mark': 'Accountancy'}, inplace=True)

# Rename Business Studies "Mark" Column label to "Business Studies"
dfBusinessStudies = dfBusinessStudiesCSV.copy()
dfBusinessStudies.rename(columns={'Mark': 'Business Studies'}, inplace=True)

# Rename Economics "Mark" Column label to "Economics"
dfEconomics = dfEconomicsCSV.copy()
dfEconomics.rename(columns={'Mark': 'Economics'}, inplace=True)

# Rename English "Mark" Column label to "English"
dfEnglish = dfEnglishCSV.copy()
dfEnglish.rename(columns={'Mark': 'English'}, inplace=True)

# Rename Maths "Mark" Column label to "Maths"
dfMaths = dfMathsCSV.copy()
dfMaths.rename(columns={'Mark': 'Maths'}, inplace=True)

data_frames = [dfAccountancyResult, dfEnglish, dfEconomics, dfMaths, dfBusinessStudies]

# Merge data frames in data frames by Name with each students mark in a column matching the course name
GradeBySubject = pd.merge(dfAccountancyResult, dfEnglish, on="Name")
# print(Grade.to_string())
GradeBySubject = pd.merge(GradeBySubject, dfMaths, on="Name")
# print(Grade.to_string())
GradeBySubject = pd.merge(GradeBySubject, dfEconomics, on="Name")
# print(Grade.to_string())
GradeBySubject = pd.merge(GradeBySubject, dfBusinessStudies, on="Name")
print(GradeBySubject.to_string())

# Create the Sum of all course marks per student i.e. sum(Accountancy + .... + Maths) and store in Column named "Total"
Total = pd.pivot_table(Grade, index=['Name'], values=['Mark'], aggfunc={'Mark': np.sum})
Total.rename(columns={'Mark': 'Total'}, inplace=True)
print("Total")
print("-----")
print(Total.to_string())
print("")

# Create the Average of all course marks per student i.e. Average(Accountancy + .... + Maths) and store in Column named "Total"
Average = pd.pivot_table(Grade, index=['Name'], values=['Mark'], aggfunc={'Mark': np.mean})
Average.rename(columns={'Mark': 'Average'}, inplace=True)
print("Average")
print("------")
print(Average.to_string())
print("")

# Merge All
Summary = pd.merge(GradeBySubject, Total, on="Name")
Summary = pd.merge(Summary, Average, on="Name")

print("Summary")
print("-------")
print(Summary.to_string())

# Write all subject CSVs dataframe into an Excel workbook

with pd.ExcelWriter("Grade.xlsx") as writer:
    # Write Accountancy CSV dataframe into an Excel workbook
    dfAccountancyResultCSV.to_excel(writer, sheet_name='Accountancy')

    # Write all English Studies CSV dataframe into an Excel workbook
    dfEnglishCSV.to_excel(writer, sheet_name='English')

    # Write all Maths Studies Studies CSV dataframe into an Excel workbook
    dfMathsCSV.to_excel(writer, sheet_name='Maths')

    # Write all Economics CSV dataframe into an Excel workbook
    dfEconomicsCSV.to_excel(writer, sheet_name='Economics')

    # Write all Business Studies CSV dataframe into an Excel workbook
    dfBusinessStudiesCSV.to_excel(writer, sheet_name='Business Studies')

    # Write Summary dataframe into an Excel workbook
    Summary.to_excel(writer, sheet_name='Summary')

with pd.ExcelWriter("Summary.xlsx") as writer:
    # Write Accountancy CSV dataframe into an Excel workbook
    dfAccountancyResultCSV.to_excel(writer, sheet_name='Accountancy')

    # Write all English Studies CSV dataframe into an Excel workbook
    dfEnglishCSV.to_excel(writer, sheet_name='English')

    # Write all Maths Studies Studies CSV dataframe into an Excel workbook
    dfMathsCSV.to_excel(writer, sheet_name='Maths')

    # Write all Economics CSV dataframe into an Excel workbook
    dfEconomicsCSV.to_excel(writer, sheet_name='Economics')

    # Write all Business Studies CSV dataframe into an Excel workbook
    dfBusinessStudiesCSV.to_excel(writer, sheet_name='Business Studies')

    # Write Summary dataframe into an Excel workbook
    Summary.to_excel(writer, sheet_name='Summary')

    # OS: A library to access the system folder:
    import os

    excel_file_list = os.listdir("C:\\Users\\USER\PycharmProjects\\pythonProject2\\daas")

    # Print this variable to see the names of the files stored within the folder. All files stored within the folder are displayed once you use the print function.

    print(excel_file_list)


