# Libraries
import pandas as pd
import numpy as np
import xlrd
import xlwt
import os
import datetime as dt
from collections import defaultdict

# for loop which looks for the column with Grand Total, and then finds the row where the Grand Total we are looking for
def find_grand_total(first_sheet):
    new_list = []
    for i in range(len(first_sheet.col_values(1))):
        if first_sheet.col_values(1)[i] == "Grand Total":
            new_list.append(i)
    return(new_list[1])

# Get the cell range
def get_cell_range(start_col, start_row, end_col, end_row):
    return [sheet.row_slice(row, start_colx=start_col, end_colx=end_col+1) for row in range(start_row, end_row+1)]

# Read dates
def read_date(date):
    return xlrd.xldate.xldate_as_datetime(date, 0)

# File Path for data
source = "/Users/kevinchoi/Desktop/Projects/Planet Fitness/Data Wrangling/2017/Fresno Shaw"
source2 = "/Users/kevinchoi/Desktop/Projects/Planet Fitness/Data Wrangling/2017/marketing"
WorkingFile = "2017 Fresno CoOp ROI Analysis 1.17.18.xlsx"

# Creates a list of dataframes with each worksheet of join data
def join_workbook(source):
    dir_list = os.listdir(source)
    os.chdir(source)
    #first create empty appended_data table to store the info.
    appended_data = []

    for WorkingFile in os.listdir(source):
        if not WorkingFile.startswith('.') and os.path.isfile(WorkingFile):
            print(WorkingFile)
            # Import the excel file and call it xlsx_file
            #xlsx_file = pd.ExcelFile(WorkingFile)
            temp_file = source + "/" + WorkingFile
            workbook = xlrd.open_workbook(temp_file)

            #.... do so reformating, call finished sheet reformatedDataSheet
            # j = row
            # i = columns
            first_sheet = workbook.sheet_by_index(0)

            grand_total_loc = find_grand_total(first_sheet)

            start_col = 1
            start_row = 20
            end_col = 27
            end_row = grand_total_loc

            my_dict = defaultdict(list)
            #my_dict = pd.DataFrame()

            for i in range(start_col, end_col):
                counter = 0
                for j in range(start_row, end_row):
                    counter += 1
                    if(counter == 1):
                        my_dict[first_sheet.cell_value(rowx=j, colx=i)]
                        key = first_sheet.cell_value(rowx=j, colx=i)
                    else:
                        my_dict[key].append(first_sheet.cell_value(rowx=j, colx=i))
            my_dict = pd.DataFrame(my_dict)
            appended_data.append(my_dict)
    appended_data = pd.concat(appended_data)
    appended_data['DOB'] = pd.to_datetime(appended_data['Date'].apply(read_date), errors='coerce')
    return(appended_data)

# Breakdown of the product tiers
def product_tiers(appended_data):
    # Labeling the web join (product tier)
    in_club_prior = list(appended_data.loc[:, 'BCM':'Total'].columns)
    in_club = [item + '_club' for item in in_club_prior]
    in_club_dict = dict(zip(in_club_prior, in_club))
    appended_data.rename(columns=in_club_dict, inplace=True)

    # Labeling the in web joins (product tier)
    web_prior = list(appended_data.loc[:, 'BCM ':'Total '].columns)
    web = [item + '_web' for item in web_prior]
    web_dict = dict(zip(web_prior, web))
    appended_data.rename(columns=web_dict, inplace=True)
    # Rename the Total column and rename $ column to revenue
    appended_data.rename(columns={'Total  ': 'Total', '$': 'revenue'}, inplace=True)
    # Drop date column
    appended_data.drop(columns=['Date'], inplace=True)
    return(appended_data)

# Creates a list of dataframes for the marketing dataframes
def marketing_workbook(source, WorkingFile):
    # Data source for marketing data
    dir_list = os.listdir(source)
    os.chdir(source)

    #first create empty appended_data table to store the info.
    appended_data2 = []

    temp_file = source + "/" + WorkingFile
    workbook = xlrd.open_workbook(temp_file)

    second_sheet = workbook.sheet_by_index(5)


    start_col = 1
    start_row = 1
    end_col = 13
    end_row = 16

    my_dict = defaultdict(list)
    for i in range(start_row, end_row):
        counter = 0
        for j in range(start_col, end_col):
            counter += 1
            if(counter == 1):
                my_dict[second_sheet.cell_value(rowx=i, colx=j)]
                key = second_sheet.cell_value(rowx=i, colx=j)
            else:
                my_dict[key].append(second_sheet.cell_value(rowx=i, colx=j))
    my_dict = pd.DataFrame(my_dict)
    return(my_dict)

# Split media dates for the marketing reformatedDataSheet
def split_media_dates(my_dict):
    # Split media campaign dates such that each campaign date is a row
    s = my_dict['Media Campaign Dates'].str.split(',').apply(pd.Series, 1).stack()
    s.index = s.index.droplevel(-1) # to line up with df's index
    s.name = 'Media Campaign Dates' # needs a name to join
    del my_dict['Media Campaign Dates']
    my_dict = my_dict.join(s)
    return(my_dict)

# Split the marketing metrics by start and end dates to later use so we can align join and marketing_workbook
def star_end_date(my_dict):
    # Divide date ranges with start and end dates
    my_dict["year"] = "2017"
    my_dict[["start_date", "end_date"]]= my_dict['Media Campaign Dates'].str.split("-", n = 2, expand = True)
    # Trims the start and end date for all white spaces
    my_dict[["start_date", "end_date"]] = my_dict[["start_date", "end_date"]].apply(lambda x: x.str.strip())
    my_dict["start_date"] = my_dict[["start_date", 'year']].apply(lambda x: '/'.join(x), axis=1)
    my_dict["end_date"] = my_dict[["end_date", 'year']].apply(lambda x: '/'.join(x), axis=1)
    my_dict["start_date"] = pd.to_datetime(my_dict["start_date"], format="%m/%d/%Y")
    my_dict["end_date"] = pd.to_datetime(my_dict["end_date"], format="%m/%d/%Y")
    # converts days into an integer
    my_dict["sales_length"] = (my_dict["end_date"] - my_dict["start_date"]).dt.days + 1
    return(my_dict)


# find all rows that have start-end date and for the next x number of rows duplicate
#repeat rows
appended_data = appended_data.assign(key=1)
my_dict = my_dict.assign(key=1)
df_merge = pd.merge(appended_data, my_dict, on='key').drop('key',axis=1)
# These are all the dates that will have the days between start-end date. Now I need to join this data to "appended_data"
df_merge2 = df_merge.query('DOB >= start_date and DOB <= end_date')
df_merge3 = df_merge2.loc[:,"DOB":]
df_out = pd.merge(appended_data, df_merge3, how="left", on="DOB").drop(["key"], axis=1)

# Drop unnecessary columns
df_out = df_out.drop(["Upgrades", "Downgrades", "No Impact", "Net Impact",
                      "ACH %", "CC %", "Agency Fee - 6.5% of Spend", "Extreme Reach Trafficking Fee",
                     "Fresno Bee Post-Its",], 1)
# Rename df
df_out = df_out.rename(columns={"Total  ": "Join_Daily",
                                "$": "Total_revenue", "month_year_x": "month_year", "DOB": "Date", " Fresno Co-Op Media": "Fresno Co-Op Media",
                               " Fresno Co-Op Promos": "Fresno Co-Op Promos", "Display / Mobile / Social" : "Display_Social"})
# turn all NaN into 0.
df = df_out.fillna(0.0)

# Change marketing columns into int
cols = ["TV / Cable", "Radio", "Pandora", "Display_Social", "DMV Ads", "Mobile Billboard", "Media Investment"]
df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')

# Create join/day columns and marketing/day columns
df["Pandora_Day"] = np.where(df["Pandora"] > 0, df["Pandora"]/df["sales_length"], df["Pandora"])
df["TV_Day"] = np.where(df["TV / Cable"] > 0, df["TV / Cable"]/df["sales_length"], df["TV / Cable"])
df["Radio_Day"] = np.where(df["Radio"] > 0, df["Radio"]/df["sales_length"], df["Radio"])
df["Display_Day"] = np.where(df["Display_Social"] > 0, df["Display_Social"]/df["sales_length"], df["Display_Social"])
df["Media_Day"] = np.where(df["Media Investment"] > 0, df["Media Investment"]/df["sales_length"], df["Media Investment"])
