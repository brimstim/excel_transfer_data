from openpyxl import load_workbook
import pandas as pd

wb = load_workbook(filename='scwc_grab.xlsx',
                   read_only=True)
ws_pits = wb['Sep Meter Pits']
ws_samps = wb['Sep Cl2 Sample']

def calculate_gallons(row):
    return row / 1000

# Read the cell values into a list of lists
data_rows1 = []
for row in ws_pits['K3':'K33']:
    data_cols = []
    for cell in row:
        data_cols.append(cell.value)
    data_rows1.append(data_cols)
# Read the cell values into a list of lists
data_rows2 = []
for row in ws_samps['B2':'B32']:
    data_cols = []
    for cell in row:
        data_cols.append(cell.value)
    data_rows2.append(data_cols)



# Transform into dataframe

df1 = pd.DataFrame(data_rows1).div(1000)
df2 = pd.DataFrame(data_rows2)
testdf = ([1000])


with pd.ExcelWriter('scwc_go.xlsx', engine='openpyxl',mode='a') as writer:
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df2.to_excel(writer,sheet_name='sep20',startrow=11, startcol=1, header=False, index=False)# startrow=13, startcol=1
    writer.save()
with pd.ExcelWriter('scwc_go.xlsx', engine='openpyxl',mode='a') as writer:
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df1.to_excel(writer,sheet_name='sep20',startrow=11, startcol=4, header=False, index=False)# startrow=13, startcol=1
    writer.save()
