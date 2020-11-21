try:
    import traceback
    import time
    import os
    import pandas as pd
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


def get_item_index(list,level):
    matched_idex = []
    has_nested_data  = False
    for idx,each_item in enumerate(list):
        # print(f"Each item is  {each_item}")
        # print(each_item[0])
        if level == int(str(each_item[0]).replace(".", "")):
            # print(f"Same level  at index {idx}")
            matched_idex.append(idx)
        elif level < int(str(each_item[0]).replace(".", "")):
            # print(f"found a nested level {level}  < {each_item[0]} at  index  {idx}")
            has_nested_data = True
    return matched_idex,has_nested_data


def dump_sheet_data(data,level,index_list,finished_good):
    print(data,level,index_list,finished_good)
    # wait  = input("Check data bofeo print  ")
    print(f"Printing data for {level} with index  {index_list}")
    if level == 1:
        print("Printing data for first level ")
        print(f"Creating new sheet for {finished_good}")
        for index  in index_list:
            print(data[index])

    else:
        print(f"print data for level {level}")
        prev_index = -99
        for index in index_list:
            if prev_index+1 == index:
                # print("data is continues now printing data")
                print(data[index])
            else:
                print(f"creating new sheet for {data[index-1][1]}")
                print(data[index])
                prev_index = index

def main():
    # hard code input file for now
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    input_file = r"BOM file for Data processing.xlsx"

    # Read file and checks
    _has_data, file_content = read_input_file(file_path=os.path.join(BASE_DIR, input_file))
    if not _has_data:
        return

    print(file_content)
    # Divide data frame based on item
    items = set(file_content.index.values)
    for idx, item in enumerate(items):
        finished_good  = item
        list = file_content.loc[item].values.tolist()
        # print(list,idx)
        level = 1
        has_more_level = True
        while has_more_level:
            print(f"Data for level {level}")
            index_list,has_more_level = get_item_index(list,level)
            print(f"indx list is  {index_list}")
            dump_sheet_data(list,level,index_list,finished_good= finished_good)
            level +=1


if __name__ == '__main__':
    t_start = time.perf_counter()
    main()
    t_end = time.perf_counter()
    print("Time taken by the main module in seconds:", t_end - t_start)
    print("Completed all task ")
