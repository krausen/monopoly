import xlrd
import numpy as np
#np.set_printoptions(threshold=np.nan)


# spreadsheet -> matrix
workbook = xlrd.open_workbook('monopoly.xls')
sheet = workbook.sheet_by_name('probability_matrix')
ncols = sheet.ncols
nrows = sheet.nrows
probMatrix = np.zeros((40,40))

for i in range(1,nrows):
    for j in range(1,ncols-1):
        probMatrix[i-1][j-1] = float(sheet.cell(i,j).value)

# extract labels
labels = []
for i in range(1,ncols-1):
    labels.append(sheet.cell(0,i).value)

# matrix to power n
n = 100
probMatrix = np.linalg.matrix_power(probMatrix, n)

# print result, sorted on ascending probability
tuples = zip(probMatrix[0],labels)
tuples.sort()
tuples.reverse()
print tuples