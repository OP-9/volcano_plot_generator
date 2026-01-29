import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt
from colour_dictionaries import lectin_color_dict, glycan_dict


# FUNCTION TO RETRIEVE A CELL'S ADDRESS WITHOUT THE DEFAULT $ NOTATION
def get_cell_wo_symbol(cell):
    column_letter = ""
    for letter in cell.address.split('$'):
        if letter.isalnum():
            column_letter += letter
    return column_letter


# FUNCTION TO FILTER THE DATA ACCORDING TO USER INPUT
def data_filter(df):
    filter_name = input('Enter the name of the filter: ')
    values = len(df.loc[filter_name].unique())
    filter_value = ""
    if values > 1:
        print('\nWhich value you would like to filter?')
        print("Your options are:")
        print(df.loc[filter_name].unique())
        filter_value = input('Enter the value you would like to filter: ')
    else:
        pass

    num = 0
    filter_list = []
    ids = df.columns
    for i in df.loc[filter_name]:
        if i == filter_value:
            filter_list.append(ids[num])
        num += 1

    df = df[filter_list]

    response = input('\nDo you want to filter again? (y/n): ').lower()
    if response == 'y':
        df = data_filter(df)
    elif response == 'n':
        print("\nExiting...")
    else:
        print("\nInput not recognised.")
        print("\nExiting...")
    return df


# FUNCTION TO COMPARE DATA ACCORDING TO USER INPUT
def data_comparison(df, wb):
    lengths = []
    print('\nStarting Data Comparison...')
    comparison = input('\nEnter the name of the comparison: ')
    print("\nYour options for comparison are: ")
    print(df.loc[comparison].unique())
    compare_values = input('\nEnter the values you would like to compare (enter multiple values with spaces only): ')
    compare_values_list = compare_values.split()
    ids = df.columns
    final_comparison_list = []

    for n in range(len(compare_values_list)):
        num = 0
        count = 0
        comparison_list = []
        for i in df.loc[comparison]:
            if i == compare_values_list[n]:
                #print(ids[num])
                comparison_list.append(ids[num])
                count += 1
            num += 1
        lengths.append(count)
        final_comparison_list += comparison_list

    df = df[final_comparison_list]

    sheet_name = input('\nEnter the name of the new sheet: ')
    wb.sheets.add(sheet_name)

    new_sheet = wb.sheets[sheet_name]
    add_to_excel_sheet(lengths, new_sheet, df, compare_values_list)
    return df


# MAIN FUNCTION, THE ONLY FUNCTION THAT NEEDS TO BE CALLED BY THE USER
def gather_and_add_data(sheet, current_book):       #SHEET IS THE EXCEL SHEET THAT HAS THE DATA TO BE READ
    last_row = sheet.range('A101203').end('up')
    last_cell = sheet.range(last_row.address).end('right')
    wb = current_book

    df = sheet.range('A1:'+str(last_cell.address)).options(pd.DataFrame, header=1, index=True).value
    response = input('\nDo you want to filter the data? (y/n): ').lower()
    if response == 'y':
        df = data_filter(df)
    else:
        print("\nNot filtering...")

    data_comparison(df, wb)

    response = input('\nDo you want perform another filter/comparison? (y/n): ').lower()
    if response == 'y':
        gather_and_add_data(sheet, current_book)
    else:
        print("\nExiting...")


# FUNCTION TO ADD THE FILTERED AND COMPARED DATA TO AN EXCEL SHEET
def add_to_excel_sheet(lengths, sheet, df, compare_values_list):
    sheet.range('A1').value = df
    first_row = int(input("\nEnter the first row with a protein: "))
    comp_1_first_cell = sheet.range((first_row, 2))
    comp_1_last_cell = sheet.range((first_row, 2 + lengths[0] - 1))
    comp_2_first_cell = sheet.range((first_row, 2 + lengths[0]))
    last_column = sheet.range('A12').end('right')
    new_column = last_column.column + 2
    last_row = sheet.range('A1').end('down').row

    variable_1 = compare_values_list[0]
    variable_2 = compare_values_list[1]

    #RETRIEVING THE TWO VARIABLE'S FIRST AND LAST CELLS' ADDRESSES
    variable_1_first_cell = get_cell_wo_symbol(comp_1_first_cell)
    variable_1_last_cell = get_cell_wo_symbol(comp_1_last_cell)
    variable_2_first_cell = get_cell_wo_symbol(comp_2_first_cell)
    variable_2_last_cell = get_cell_wo_symbol(last_column)

    range_1 = str(variable_1_first_cell) + ":" + str(variable_1_last_cell)
    range_2 = str(variable_2_first_cell) + ":" + str(variable_2_last_cell)

    variable_1_last_cell_column = "".join(letter for letter in variable_1_last_cell if letter.isalpha())
    variable_2_last_cell_column = "".join(letter for letter in variable_2_last_cell if letter.isalpha())

    # ADD COLOURS TO THE VARIABLES TO DIFFERENTIATE THEM
    sheet.range(str(variable_1_first_cell)+":"+ str(variable_1_last_cell_column)+str(last_row)).color = '#d2e6c8'
    sheet.range(str(variable_2_first_cell)+":"+ str(variable_2_last_cell_column)+str(last_row)).color = '#d1e6eb'

    sheet.range((first_row - 1, new_column)).value = "Avg. of " + str(variable_1)

    sheet.range((first_row, new_column), (last_row, new_column)).formula = f"=AVERAGE({range_1})"

    sheet.range((first_row - 1, new_column + 1)).value = "Avg. of " + str(variable_2)
    sheet.range((first_row, new_column + 1), (last_row, new_column + 1)).formula = f"=AVERAGE({range_2})"

    avg_1 = sheet.range((first_row, new_column))
    avg_2 = sheet.range((first_row, new_column + 1))

    avg_1_address = get_cell_wo_symbol(avg_1)
    avg_2_address = get_cell_wo_symbol(avg_2)

    #DIFFERENCE COLUMN
    sheet.range((first_row-1, new_column + 2)).value = "Difference"
    sheet.range((first_row, new_column + 2), (last_row, new_column + 2)).formula = f"=2^({avg_1_address} - {avg_2_address})"

    # RETRIEVING THE CELL ADDRESS OF THE FIRST CELL IN THE DIFFERENCE COLUMN
    diff_cell = sheet.range((first_row, new_column + 2))
    diff_cell_address = get_cell_wo_symbol(diff_cell)

    # T.TEST COLUMN
    sheet.range((first_row-1, new_column + 3)).value = "TTest"
    sheet.range((first_row, new_column + 3), (last_row, new_column + 3)).formula = f"=T.TEST({range_1},{range_2}, 2, 3)"

    #RETRIEVING THE CELL ADDRESS OF THE FIRST CELL IN THE T.TEST COLUMN
    ttest_cell = sheet.range((first_row, new_column + 3))
    ttest_cell_address = get_cell_wo_symbol(ttest_cell)

    #LOG 2 DIFFERENCE COLUMN
    sheet.range((first_row-1, new_column + 4)).value = "Log2diff"
    sheet.range((first_row, new_column + 4), (last_row, new_column + 4)).formula = f"=LOG({diff_cell_address},2)"

    #-LOG 10 P COLUMN
    sheet.range((first_row-1, new_column + 5)).value = "-Log10p"
    sheet.range((first_row, new_column + 5), (last_row, new_column + 5)).formula = f"=-LOG({ttest_cell_address},10)"

    #INDEX COLUMN
    sheet.range((first_row-1, new_column + 6)).value = "Index"
    sheet.range((first_row, new_column + 6)).options(transpose=True).value = list(df.index[10:])

    sheet.range((first_row-1, new_column),(first_row-1, new_column + 6)).font.bold = True

    log_df = sheet.range((first_row-1, new_column + 4), (last_row, new_column + 6)).options(pd.DataFrame, header=1,
                                                                                    index=False).value
    log_df.set_index("Index", inplace=True)
    log_df = log_df.reindex(columns=["Log2diff", "-Log10p"])
    log_df = log_df.sort_values("-Log10p", ascending=False)

    sheet.range((1, new_column + 9)).value = log_df
    sheet.range((1, new_column+9), (1, new_column + 11)).font.bold = True
    sheet.autofit()

    #PLOT GENERATOR
    plt.figure(figsize=(8, 6))
    cutoff = float(input("Enter the cutoff value: "))

    plt.axhline(y= cutoff, color='red', linestyle='--', linewidth=0.8)
    plt.axvline(x=0, color='gray', linestyle='-', linewidth=0.8)

    plt.xlabel('Log Fold Change')
    plt.ylabel('-log10(p-value)')
    fig_title = input("Enter the name of the plot: ")
    plt.title(fig_title)
    min_cutoff = []
    color = 'gray' # Default colour for points

    for i in range(len(log_df['-Log10p'])):
        lectin_id = log_df.index[i]

        if log_df['-Log10p'].iloc[i] >= cutoff:
            min_cutoff.append(log_df['-Log10p'].iloc[i])
            color = lectin_color_dict.get(lectin_id, 'gray')  # Use assigned colour if above cutoff, gray if not found

        plt.scatter(x=log_df['Log2diff'].iloc[i], y=log_df['-Log10p'].iloc[i], color=color, s=10, alpha=0.7)

    for i in range(len(min_cutoff)):
        plt.annotate(log_df.index[i], (log_df['Log2diff'].iloc[i], log_df['-Log10p'].iloc[i] + 0.05))

    legend_elements = []

    for desc, color in glycan_dict.items():
        legend_elements.append(plt.Line2D([0], [0], marker='o', color='w',
                                          markerfacecolor=color, markersize=10, label=desc))

    # LEGEND
    plt.legend(handles=legend_elements, loc='best' , bbox_to_anchor=(1, 0.15))

    plt.tight_layout()

    fig_name = input("Enter the name of the file to save the plot on your computer: ")
    plt.grid(True, linestyle=':', alpha=0.6)
    plt.savefig(fig_name+".png", dpi=300)
    plt.show()


