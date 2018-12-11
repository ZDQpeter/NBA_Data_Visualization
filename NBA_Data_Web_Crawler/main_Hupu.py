from lxml import html
import requests
import numpy as np
import xlsxwriter
import datetime
from function_main_Plotting import function_main_Plotting

# Crawl from the website
page = requests.get('https://nba.hupu.com/standings')
tree = html.fromstring(page.content)

#This will create a list of teams:
teams = tree.xpath('//td/text()')
names = tree.xpath('//a[@target="_blank"]/text()')

# #This will create a list of prices
# prices = tree.xpath('//span[@class="item-price"]/text()')

# ==============================================
# Define parameters
teams_header_east = teams[3:15]
teams_header_west = teams[205:217]
NBA_East_Name = names[55:70]
NBA_West_Name = names[70:85]
rank = np.arange(0, 16, 1)
rank = np.transpose(np.matrix(rank))
m = len(teams)

NBA_East = []
NBA_West = []
a = 0

# -------------------------------------------
# Store results into variables
NBA_East = np.append(NBA_East, teams_header_east)
NBA_West = np.append(NBA_West, teams_header_west)

# ==============================================
print('Teams: ', teams_header_east)

for idx in range(15, 111):
    if (((idx%12)-3) == 0):
        print('Teams: ', teams[idx:(idx+12)])
        NBA_East = np.append(NBA_East, teams[idx:(idx + 12)])

for idx in range(112, 203):
    if (((idx%13)-8) == 0):
        print('Teams: ', teams[idx:(idx+12)])
        NBA_East = np.append(NBA_East, teams[idx:(idx + 12)])

print('Teams: ', teams_header_west)

for idx in range(217, 313):
    if (((idx%12)-1) == 0):
        print('Teams: ', teams[idx:(idx+12)])
        NBA_West = np.append(NBA_West, teams[idx:(idx + 12)])

for idx in range(314, m):
    if (((idx%13)-2) == 0):
        print('Teams: ', teams[idx:(idx+12)])
        NBA_West = np.append(NBA_West, teams[idx:(idx + 12)])


NBA_East = np.matrix(NBA_East)
NBA_East = np.reshape(NBA_East, (16, 12))
NBA_West = np.matrix(NBA_West)
NBA_West = np.reshape(NBA_West, (16, 12))

# Get NBA team names into the result matrix
NBA_East_Name = np.matrix(NBA_East_Name)
NBA_East_Name = np.insert(NBA_East_Name, 0, a)
NBA_East_Name = np.transpose(NBA_East_Name)
NBA_West_Name = np.matrix(NBA_West_Name)
NBA_West_Name = np.insert(NBA_West_Name, 0, a)
NBA_West_Name = np.transpose(NBA_West_Name)
East = np.hstack((NBA_East_Name, NBA_East))

East = np.hstack((rank, East))
West = np.hstack((NBA_West_Name, NBA_West))
West = np.hstack((rank, West))

print('---------------------------')
print(East)
print('---------------------------')
print(West)


# ============================================
# Save the matrix variable results in Excel file
now = datetime.datetime.now().date()
filename = './NBA_Data/' + str(now) + '.xlsx'

workbook = xlsxwriter.Workbook(filename)
worksheet1 = workbook.add_worksheet()
worksheet2 = workbook.add_worksheet()

(p, q) = np.shape(East)

for i in range(0, p):
    for j in range(0, q):
        worksheet1.write(i, j, East[i, j])
        worksheet2.write(i, j, West[i, j])

workbook.close()

# ============================================
# Call plotting function
function_main_Plotting()


