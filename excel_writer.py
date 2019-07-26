import pandas as pd
import os
import glob


def write_excel_file(path='./', separator=',', outputfilename='output.xlsx'):
    shortfilename, fileExtension = os.path.splitext(outputfilename)
    if fileExtension == '' or fileExtension != '.xlsx':
        print('File extension must be xlsx, changing')
        fileExtension = '.xlsx'
    outputfilename = shortfilename + fileExtension
    if os.path.isdir(path) is False:
        print('Path is not valid, exiting')
        return
    if path[-1:] != '/':
        path = path + '/'
    all_files = glob.glob(os.path.join(path, "*.csv"))
    filename = path + outputfilename
    writer = pd.ExcelWriter(filename)
    # remove 0 size files
    for f in all_files:
        try:
            if os.path.getsize(f) == 0:
                print(f, 'is empty, removing')
                os.remove(f)
        except:
            continue
    # reinitialize all_files now empty files have been removed
    all_files = glob.glob(os.path.join(path, "*.csv"))
    # loop through files and add tabs to excel file
    for f in all_files:
        df = pd.read_csv(f, sep=separator, engine='python')
        # tab name is the name of the CSV file, minus the CSV, in title case with spaces instead of underscores
        tabname = os.path.basename(f)
        tabname = tabname.replace('.csv', '')
        tabname = tabname.replace('_', ' ')
        tabname = tabname.title()
        # Strip trailing whitespace from columns
        df.columns = df.columns.str.strip()
        df.to_excel(writer, sheet_name=tabname, index=False)
        # Indicate worksheet for formatting
        worksheet = writer.sheets[tabname]
        # Iterate through each column and set the width == the max length in that column.
        # A padding length of 2 is also added.
        for i, col in enumerate(df.columns):
            # find length of column i
            column_len = df[col].astype(str).str.len().max()
            # If the header is longer than the data in the column use the length of the header
            column_len = max(column_len, len(col)) + 2
            # But don't go above 30
            column_len = min(column_len, 30)
            # set the column length
            worksheet.set_column(i, i, column_len)
    writer.save()
