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
        return False,None
    try:
        df = pd.read_excel(file_path)
        return True,df
    except:
        print(f"Error reading input file \nCheck traceback error message below \n\n")
        traceback.print_exc()
        return False,None

def main():
    # hard code input file for now
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    input_file = r"BOM file for Data processing.xlsx"

    # Read file and checks
    _has_data,file_content  = read_input_file(file_path = os.path.join(BASE_DIR,input_file))
    if not _has_data:
        return
    print(file_content)


if __name__ == '__main__':
    t_start  = time.perf_counter()
    main()
    t_end = time.perf_counter()
    print("Time taken by the main module in seconds:",t_end - t_start)
    print("Completed all task ")