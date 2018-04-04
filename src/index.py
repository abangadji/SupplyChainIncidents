# created by Vedant Jha on 31/03/2018
# under the mentorship of Dr. Akshit Singh
# Task to perform sentiment analysis on the supply chain incidents present in the sample.xlsx file

# first importing all the required packages to be used

import math  # For doing mathematical operations
import re  # For the calculation of regular expressions
import matplotlib.pyplot as plt  # For plotting the graphs and charts
import pandas as pd  # To handle data
import xlwt  # For creating new spread-sheets and saving the data into the sheets
from textblob import TextBlob  # For getting the quantitative value for the polarity and subjectivity
import numpy as np  # For number computing
import collections  # For performing operations on collection framework


# a function to clean the tweets using regualr expression
def clean_tweet(tweet):
    '''
    Utility function to clean the text in a tweet by removing
    links and special characters using regex.
    '''
    return ' '.join(re.sub("(@[A-Za-z0-9]+)|([^0-9A-Za-z \t])|(\w+:\/\/\S+)", " ", tweet).split())


# opening the sample data file and reading the tweets data

list_sheetnames = pd.ExcelFile("sample.xlsx").sheet_names
total_num_sheets = len(list_sheetnames)

# Preprocessing tweets in each of the sheets and saving it in the new xls file "preprocessed.xls"

style_blue = xlwt.easyxf('font: name Times New Roman, color-index blue, bold on', num_format_str='#, ##0.00') # Style for headers in the excel sheet
wb = xlwt.Workbook()

avg_pol_sheet = []  # list containing average polarity of the sheets
avg_sub_sheet = []  # list containing average subjectivity of the sheets

for i in range(total_num_sheets):
    sheet_read = pd.read_excel("sample.xlsx", list_sheetnames[i])
    df = pd.DataFrame(sheet_read)
    tweets = sheet_read["Tweets"].values.tolist()
    user_date = sheet_read["Twitter user & Date"].values.tolist()
    ws = wb.add_sheet(list_sheetnames[i])
    ws.write(0, 0, "S.No.", style_blue)
    ws.write(0, 1, "Tweets", style_blue)
    ws.write(0, 2, "User & Date", style_blue)
    ws.write(0, 3, "Polarity", style_blue)
    ws.write(0, 4, "Subjectivity", style_blue)
    j = 0
    pol_sum = 0
    sub_sum = 0
    for tweet in tweets:
        analysis = TextBlob(clean_tweet(tweet))
        pol = analysis.sentiment.polarity
        sub = analysis.subjectivity
        pol_round = '%.3f' % pol
        sub_round = '%.3f' % sub
        pol_sum = pol_sum + pol
        sub_sum = sub_sum + sub
        ws.write(j + 1, 0, j + 1)
        ws.write(j + 1, 1, clean_tweet(tweet))
        ws.write(j + 1, 2, user_date[j])
        ws.write(j + 1, 3, pol_round)
        ws.write(j + 1, 4, sub_round)
        j = j + 1

    avg_pol_sheet.append('%.3f' % (pol_sum / len(tweets)))
    avg_sub_sheet.append(('%.3f' % (sub_sum / len(tweets))))

wb.save("preprocessed.xls")  # saving all the stuffs in the workbook preprocessed.xls

# calculating standard deviation for the polarity and subjectivity for each of the sheets

sd_polarity = []  # list containing standard deviation values for each of the sheets for polarity
sd_subjectivity = []  # list containing standard deviation values for each for the sheets for subjectivity

for i in range(len(list_sheetnames)):
    sheet_read = pd.read_excel("preprocessed.xls", list_sheetnames[i])
    polariti = sheet_read["Polarity"].values.tolist()
    subjecti = sheet_read["Subjectivity"].values.tolist()
    sum_diff_polariti = 0
    sum_diff_subjecti = 0
    for j in range(len(polariti)):
        sum_diff_polariti = sum_diff_polariti + ((float(polariti[j]) - float(avg_pol_sheet[i])) ** 2)
        sum_diff_subjecti = sum_diff_subjecti + ((float(subjecti[j]) - float(avg_sub_sheet[i])) ** 2)

    sd_polarity.append('%.3f' % (math.sqrt((sum_diff_polariti) / len(polariti))))
    sd_subjectivity.append('%.3f' % (math.sqrt((sum_diff_subjecti) / len(polariti))))

# now saving the average polarity, average subjectivity, standard-deviation for polarity and subjectivity in output.xls

wb = xlwt.Workbook()
ws = wb.add_sheet("output")
ws.write(0, 0, "S.No.", style_blue)
ws.write(0, 1, "Supply-Chain-Incident", style_blue)
ws.write(0, 2, "Avg. Polarity", style_blue)
ws.write(0, 3, "S.D. Polarity", style_blue)
ws.write(0, 4, "Avg. Subjectivity", style_blue)
ws.write(0, 5, "S.D. Subjectivity", style_blue)

for i in range(len(list_sheetnames)):
    ws.write(i + 1, 0, i + 1)
    ws.write(i + 1, 1, list_sheetnames[i])
    ws.write(i + 1, 2, avg_pol_sheet[i])
    ws.write(i + 1, 3, sd_polarity[i])
    ws.write(i + 1, 4, avg_sub_sheet[i])
    ws.write(i + 1, 5, sd_subjectivity[i])

wb.save("output.xls")

# plotting the line-chart for the average polarity and a
average_polarity_sheets = []
average_subjectivity_sheets = []
for k in range(len(avg_pol_sheet)):
    average_polarity_sheets.append(float(avg_pol_sheet[k]))
    average_subjectivity_sheets.append(float(avg_sub_sheet[k]))

# plotting the line-chart for the average polarity with the supply-chain-incidents
plt.title("Average Polarity of the Supply-Chain-Incidents")
plt.xlabel("Supply Chain Incidents ------------------------------------>")
plt.ylabel("Avg. Polarity  ------------------------------------------>")
plt.ylim(-0.3, 0.3)
plt.plot(list_sheetnames, average_polarity_sheets)
plt.show()

# plotting the line-chart for the average subjectivity with the supply-chain-incidents
plt.title("Average Subjectivity of the Supply-Chain-Incidents")
plt.xlabel("Supply Chain Incidents---------------------------->")
plt.ylabel("Avg. Subjectivity---------------------------------->")
plt.ylim(0.0, 1.0)
plt.plot(list_sheetnames, average_subjectivity_sheets)
plt.show()

# Plotting the bar-graph for the average polarity with the supply-chain-incidents
objects = list_sheetnames
y_pos = np.arange(len(list_sheetnames))
plt.barh(y_pos, average_polarity_sheets, align='center', alpha=0.5)
plt.yticks(y_pos, list_sheetnames)
plt.xlabel('Average Sentiment------------------>')
plt.title('Polarity analysis')

plt.show()

# Plotting the bar-graph for the average subjectivity with the supply-chain-incidents
plt.barh(y_pos, average_subjectivity_sheets, align='center', alpha=0.5)
plt.yticks(y_pos, list_sheetnames)
plt.xlabel("Average Subjectivity------------------------->")
plt.ylabel("Chain-Supply-Incidents----------------------->")
plt.title("Subjectivity Analysis")
plt.show()

# plotting the pie-chart for the average polarity with the supply-chain-incidents


colors = []
labels = list_sheetnames
sizes = []
total_sum = 0.0
col = ['a', 'b', 'c', 'd', 'e', 'f', 'g']

for i in range(len(list_sheetnames)):
    sizes.append(abs(average_polarity_sheets[i]))
    total_sum = total_sum + abs(average_polarity_sheets[i])
    if average_polarity_sheets[i] >= 0.10:
        colors.append("darkolivegreen")  # most positive
    elif 0.10 > average_polarity_sheets[i] > 0.05:
        colors.append("forestgreen")  # more positive
    elif 0.05 >= average_polarity_sheets[i] > 0.00:
        colors.append("lightgreen")  # less positive
    elif average_polarity_sheets[i] == 0.00:
        colors.append("blue")  # neutral
    elif 0.00 > average_polarity_sheets[i] > -0.05:
        colors.append("tomato")  # less negative
    elif -0.05 >= average_polarity_sheets[i] > -0.10:
        colors.append("red")  # more negative
    elif average_polarity_sheets[i] <= -0.10:
        colors.append("firebrick")  # most negative sentiment
    else:
        colors.append("red")
percentage = []
for i in range(len(sizes)):
    percentage.append((sizes[i] / total_sum) * 100)
explode = (0.2, 0.15, 0.1, 0.05, 0.01, 0.005, 0.001, 0.001, 0.001, 0.001, 0.001)  # To make slices in the pie chart
plt.axis('equal')
plt.title("Pie Chart for the sentiment positive/negative % ")
plt.pie(percentage, explode=explode, labels=labels, colors=colors,
        autopct='%1.1f%%', shadow=True, startangle=140)
lab = ["a", "b", "b", "c", "d", "e", "f"]
plt.show()

# now printing the supply chain incidents in increasing positivity
print("The Supply Chain Incidents in increasing positivity: ")

d = dict(zip(average_polarity_sheets, list_sheetnames))  # combining the sheet names with its polarity
od = collections.OrderedDict(sorted(d.items()))  # Sorting by hashing
for i in od:
    print(od.get(i))
