
import pandas as pd
import numpy as np
import xlwings as xw
from target_group_list import target_group_list, target_group_CPP_dict

core_target_group = input("Enter the target group: ")

# open the excel file
wb = xw.Book("020123_GrossContactsCalculatorTVFY2223.xlsm")
ws = wb.sheets["Manual TV"]  # ws is the worksheet object and the second Worksheet (Manual TV) in the Workbook

# select the core target group
ws.range("B6").value = core_target_group
ws.range("B3").value = "Germany"
ws.range("D3").value = "Germany"

# create a dataframe to store the reach curves
reach_curve_df = pd.DataFrame(columns=["core_target_group",
                                       "buying_target_group",
                                       "GRP",
                                       "Reach",
                                       "Budget",
                                       "CostPerReachPoint"])

reach_curve_list = []

for tg in target_group_list:
    ws.range("C10").value = tg
    cpp = target_group_CPP_dict[tg]
    for grp in range(10, 800, 10):  # only every 10th GRP until 800 Max
        ws.range("H10").value = grp
        current_reach = ws.range("H15").value

        try:
            # create a list with dictionaries to store the values
            reach_curve_dict = {"core_target_group": core_target_group,
                                    "buying_target_group": tg,
                                    "GRP": grp,
                                    "Reach": current_reach,
                                    "Budget": grp * cpp,
                                    "CPP": cpp,
                                    "CostPerReachPoint": grp * cpp / current_reach}
        except ZeroDivisionError:
            reach_curve_dict = {"core_target_group": core_target_group,
                                "buying_target_group": tg,
                                "GRP": grp,
                                "Reach": current_reach,
                                "Budget": grp * cpp,
                                "CPP": cpp,
                                "CostPerReachPoint": 0}

        reach_curve_list.append(reach_curve_dict)

# append the dictionary to the dataframe
reach_curve_df = reach_curve_df.append(reach_curve_list, ignore_index=True)

# save the dataframe as a csv file
reach_curve_df.to_csv("reach_curve.csv", index=False)

pd.set_option('display.max_columns', None)
print(reach_curve_df)