import pandas as pd
from typing import Dict


def read_input_file(file_name: str, sheet_name: str, to_df: bool = False) -> pd.DataFrame:
    """
    Takes a filename (path to a file) and returns a pandas dataframe.

    Arguments:
        file_name {String} -- Path to an Excel file

    Returns:
        [pd.DataFrame] -- A pandas Dataframe representation of the Excel file
    """

    try:
        excel = pd.ExcelFile(file_name)
        dataframe = excel.parse(sheet_name)
        print(dataframe.to_string())
        if to_df:
            return dataframe
        else:
            return dataframe.to_dict(orient="records")
    except FileNotFoundError as error:
        return error


def write_to_excel(file_path: str, dataframe: pd.DataFrame) -> None:
    """Writes a dataframe to an excel file

    Arguments:
        file_path {str} -- The name of the file you want to write to.
        dataframe {pd.DataFrame} -- A pandas Dataframe created by using
        pandas to convert an Excel document into a DataFrame.

    Returns:
        None -- Writes a new Excel file for each row in the original Excel file.
        This is to help with testing.
    """

    for i, row in dataframe.iterrows():
        dataframe.iloc[i:i + 1].to_excel(file_path + '_line_' + str(i) + '_new.xlsx', index=False)

