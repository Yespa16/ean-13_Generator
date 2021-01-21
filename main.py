import pandas as pd
import math


# Function for defining the control number
def generateLastNumber(code):
    zuyg = 0
    kent = 0
    for i in range(len(code)):
        if i%2==0:      kent += int(code[i])
        elif i%2==1:    zuyg += int(code[i])

    zuyg *= 3
    summ = zuyg+kent
    result = math.ceil(summ/10.0)*10 - summ

    return code + str(result)


# Reading excel file
df = pd.read_excel("data.xlsx")

rows = df.shape[0]

# Getting old data here
data = []
for i in range(rows):
    if len(str(df.iloc[i,0]))!=12:
        print("INVALID DATA SIZE IN LINE ", str(i), ': ', df.iloc[i,0])
        input(" ")
        break
    else:
        data.append(str(df.iloc[i, 0]))




# Appending ean-13 codes to new list
new = []
for i in data:
    new.append(generateLastNumber(i))


# Writing ean-13 codes in excel file
for i in range(len(new)):
        df.iloc[i, 0] = new[i]


# Saving changes to new file
df.to_excel("data.xlsx")

