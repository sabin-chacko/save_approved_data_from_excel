import os
import pandas as pd

#location of excel files
data_folder = "data_folder"

def get_xlsx_filename() -> list:
    """to get xlsx file from the folder

    Returns:
        list: list of xlsx file name
    """    
    file_list = []

    if os.path.exists(data_folder):
        for f in os.listdir(data_folder):
            #check f is of type file and f is of xlsx format
            if os.path.isfile(os.path.join(data_folder, f)) and f.split(".")[1] == "xlsx":
                file_list.append(f)
    return file_list

def get_data_from_each_file(df: object, filename: str) -> object:
    """getdata from each files

    Args:
        writer (object): dataframe
        filename (str): file name

    Returns:
        object: dataframe
    """
    data = pd.read_excel(os.path.join(data_folder, filename), sheet_name="Sheet1", engine='openpyxl', dtype=str)
    #iterate each row
    for index, row in data.iterrows():
        #check if status is Approved
        status = row.get("Status")
        if row.get("Status") == "Approved":
            df.loc[len(df.index)] = [row.get("Person"), row.get("Status")] 
    
    return df

def process_xlsx_file(file_list: list) -> None:
    """proccess xlsx files

    Args:
        file_list (list): list of xlsx filename
    """
    #initialise empty xlsx file
    df = pd.DataFrame({'Person': [],
                   'Status': []})
    for f in file_list:
        df = get_data_from_each_file(df, f)
    # write the final dataframe to final_output.xlsx
    with pd.ExcelWriter("final_output.xlsx") as writer:  
        df.to_excel(writer,
                              engine='openpyxl',
                              sheet_name="Data",
                              index=False)



def main() -> None:
    """main method
    """
    file_list = get_xlsx_filename()
    process_xlsx_file(file_list)

if __name__ == "__main__":
    main()