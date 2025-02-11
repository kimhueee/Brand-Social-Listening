import pandas as pd

def to_excel_file(file_name, dict_sheet_df):

    file_directory = "C" + ":" + "\ ".strip() + "Users" + "\ ".strip() + "Kompa-KTelt056" + "\ ".strip() + "Desktop" \
                     + "\ ".strip() + str(file_name) + ".xlsx"
    with pd.ExcelWriter(file_directory) as writer:
        for slide, dataframe in dict_sheet_df.items():

            if type(dataframe) == list:
                start_column = 0
                for each_df in dataframe:
                    each_df.to_excel(writer, sheet_name=slide, startcol=start_column)
                    start_column += len(each_df.axes[1]) + 3

            if (type(dataframe) == pd.DataFrame) or (type(dataframe) == pd.Series):
                dataframe.to_excel(writer, sheet_name=slide)