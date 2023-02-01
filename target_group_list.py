import xlwings as xw
import pandas as pd

# Turn Excel File into a DataFrame
df = pd.read_excel("TVCPPs_FY2324.xlsx", sheet_name="TV CPPs DE")

# Create a dictionary of all the target groups
target_group_list = df["Adults"].tolist()

# Calculate the average CPP for each target group
target_group_CPP_dict = {}

# Set the first column as the index
df.set_index("Adults", inplace=True)

# Calculate the average of each row
avg_by_row = df.mean(axis=1)

# Create a dictionary of the target groups and their average CPP
for i in range(len(target_group_list)):
    target_group_CPP_dict[target_group_list[i]] = avg_by_row[i]

