
# Note - I installed an older version of xlrd, 1.2.0, in order to have compatibility with xlsx file
import xlrd
import xlsxwriter

# Define some lists we will use to parse old and build new output
AGE_POSTFIXES = ['_cb', '_7', '_14', '_22', '_28']
CONTAMINATE_PREFIX = ['ppdde', 'pcb118', 'pcb153', 'pcb138', 'pcb180', 'pfoa', 'pfhxs', 'pfna', 'pfda','pfostotal', 'totalpfos']
OUTPUT_HEADERS = ['slideid', 'sample_age', 'adultsmoker', 'adultbmi', 'ppdde', 'pcb118', 'pcb153', 'pcb138', 'pcb180', 'pfoa', 'pfhxs', 'pfna', 'pfda','pfostotal']

# Open original spreadsheet
workbook = xlrd.open_workbook('/Users/afreisthler/PycharmProjects/reshape/CompleteK1S.xlsx')
worksheet = workbook.sheet_by_index(0)

# Get header values into a list
headers = []
for col in range(worksheet.ncols):
    headers.append(worksheet.cell_value(0, col))

# Read existing data into a list of dictionaries
data = []
for row in range(1, worksheet.nrows):
    dict_representing_row = {}
    for col in range(worksheet.ncols):
        dict_representing_row[headers[col]] = worksheet.cell_value(row, col)
    data.append(dict_representing_row)

# Loop through each of our existing rows and expand as desired
new_data = []
for oldrow in data:
    for age in AGE_POSTFIXES:
        new_data_row = {}
        new_data_row['slideid'] = oldrow['slideid']
        new_data_row['sample_age'] = int(age.replace('_', '').replace('cb', '0'))
        new_data_row['adultsmoker'] = oldrow['adultsmoker']
        new_data_row['adultbmi'] = oldrow['adultbmi']

        for prefix in CONTAMINATE_PREFIX:
            old_value = oldrow.get(prefix + age, '')
            if '22' in age and not old_value:
                old_value = oldrow.get(prefix, '')
            new_data_row[prefix] = old_value

        # Fix the totalpfos issue
        totalpfos = new_data_row['totalpfos']
        del (new_data_row['totalpfos'])
        if totalpfos:
            new_data_row['pfostotal'] = totalpfos

        new_data.append(new_data_row)

# Output new spreadsheet
workbook = xlsxwriter.Workbook('/tmp/CompleteK1S_reshaped.xlsx')
worksheet = workbook.add_worksheet()

# Widen columns
worksheet.set_column('A:Z', 20)

# Output headers
for i, header in enumerate(OUTPUT_HEADERS):
    worksheet.write(0, i, header)

# Output data
for row_index, row in enumerate(new_data):
    for header_index, header in enumerate(OUTPUT_HEADERS):
        worksheet.write(row_index + 1, header_index, row.get(header))

workbook.close()
