try:
    import traceback
    import time
    import os
    import pandas as pd
    import datetime
    import xlsxwriter
except ImportError:
    print("Error importing python libriries \nCheck traceback error message below \n\n")
    traceback.print_exc()
    exit(1)


def read_input_file(file_path):
    if not os.path.exists(file_path):
        print(f'Input file path [{file_path}]  is not  valid ')
        return False, None
    try:
        df = pd.read_excel(file_path, index_col="Item Name")
        return True, df
    except:
        print(f"Error reading input file \nCheck traceback error message below \n\n")
        traceback.print_exc()
        return False, None


def get_item_index(list, level):
    '''
    Get the index of all items which has same level of nesting
    also return data on further nesting is there or not
    '''
    matched_index = []
    has_nested_data = False
    for idx, each_item in enumerate(list):
        if level == int(str(each_item[0]).replace(".", "")):
            matched_index.append(idx)
        elif level < int(str(each_item[0]).replace(".", "")):
            # Used to check if data has a nesting greater than the current level
            has_nested_data = True
    return matched_index, has_nested_data


def get_sheet_data(data, level, index_list, workbook, finished_good):
    if level == 1:
        print_data = []
        for index in index_list:
            print_data.append(data[index])
        #Save first level
        dump_sheet_data(workbook, print_data, finished_good, 1, "Pc")

    else:
        prev_index = -99
        print_data = []
        for index in index_list:
            if prev_index + 1 == index:

                print_data.append(data[index])
            else:
                ## The nested data for first time
                if len(print_data) != 0:
                    # Save  previous data
                    dump_sheet_data(workbook, print_data, finished_good, quantity, unit)
                    print_data = []
                finished_good = data[index - 1][1]
                quantity = data[index - 1][2]
                unit = data[index - 1][3]

                print_data.append(data[index])
                prev_index = index
        # saving last data
        dump_sheet_data(workbook, print_data, finished_good, quantity, unit)


def dump_sheet_data(workbook, data, finishd_product, quantity, unit):
    data_format1 = workbook.add_format({'bg_color': '#FFFF00',"border":1})
    df_border_blue  = workbook.add_format({"border":5,"bg_color":"0000FF"})
    df_border_bottom  = workbook.add_format({"bottom":5})
    worksheet = workbook.add_worksheet(finishd_product)
    worksheet.write_row(0, 0, ["Finished Good List"])
    worksheet.write_row(1, 0, ["#", "Item Description", "Quantity", "Unit"],cell_format=df_border_blue)
    worksheet.write_row(2, 0, ["1"])
    worksheet.write_row(2, 1, [finishd_product, quantity, unit],cell_format=data_format1)

    worksheet.write_row(3, 0, ["End of FG","","",""],cell_format=df_border_bottom)
    worksheet.write_row(4, 0, ["Raw Material list"])
    worksheet.write_row(5, 0, ["#", "Item Description", "Quantity", "Unit"],cell_format=df_border_blue)
    row = 6
    number = 1
    for each_data in data:
        worksheet.write_row(row, 0, [number])
        worksheet.write_row(row, 1, [ each_data[1], each_data[2], each_data[3]],cell_format=data_format1)
        row += 1
        number += 1
    worksheet.write_row(row, 0, ["End of RM"])
    worksheet.set_column(0, row, 30)

def main():
    # hard code input file for now
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    input_file = r"BOM file for Data processing.xlsx"

    # Read file and checks
    _has_data, file_content = read_input_file(file_path=os.path.join(BASE_DIR, input_file))
    if not _has_data:
        return
    print('Data from file ')
    print(file_content)

    output_file = datetime.datetime.now().strftime(input_file[:-5] + '_%H_%M_%d_%m_%Y.xlsx')
    print(f"Creating new excel file {output_file}")

    # dump source data
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    file_content.to_excel(writer, sheet_name='Source')
    workbook = writer.book

    # Divide data frame based on item
    items = set(file_content.index.values)
    for idx, item in enumerate(items):
        finished_good = item
        list = file_content.loc[item].values.tolist()
        level = 1
        has_more_level = True
        print(f"Executing for {item}")
        # Recurse till all the nested level are read
        while has_more_level:
            print(f"Dumping data for level {level}")
            index_list, has_more_level = get_item_index(list, level)
            get_sheet_data(list, level, index_list, workbook, finished_good=finished_good)
            level += 1
    writer.save()


if __name__ == '__main__':
    t_start = time.perf_counter()
    main()
    t_end = time.perf_counter()
    print("Time taken by the main module in seconds:", t_end - t_start)
    print("Completed all task ")
