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


def add_new_sheet(df, finished_product, level, quantity=1, unit="Pc"):
    print(f"Creating sheet for product : {finished_product}")

    # create a new df with first line as data provided and second line as all data which has level = level
    print(df)
    df_print = pd.DataFrame(
        {"Level": [level - 1], "Raw material": [finished_product], "Quantity": [quantity], "Unit": [unit]})
    print(df_print)
    for i in range(len(df)):
        if level == int(str(df.iloc[i, 0]).replace(".", "")):
            # print(df.iloc[i])
            df_print[len(df)] = [level-1,[df.iloc[i,1]], [df.iloc[i,2]],[df.iloc[i,3]]]



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
    for item in items:
        list = file_content.loc[item].values.tolist()
        new_data = file_content.loc[item]
        # print(new_data)
        add_new_sheet(df=new_data, finished_product=item, level=1)


if __name__ == '__main__':
    t_start = time.perf_counter()
    main()
    t_end = time.perf_counter()
    print("Time taken by the main module in seconds:", t_end - t_start)
    print("Completed all task ")
