import pandas as pd
import openpyxl as xl
import openpyxl.styles as style

##########################################################################
# free people request
def find_free_people(day, start_time, end_time):
    start_time = pd.to_datetime(start_time, format='%H:%M').time()
    end_time = pd.to_datetime(end_time, format='%H:%M').time()

    busy_people = all_schedules[(all_schedules['Zi'] == day) &
                                ((all_schedules['Ora incepere'] <= end_time) &
                                (all_schedules['Ora sfarsit'] >= start_time))]

    busy_names = busy_people['Nume'].unique()

    all_people = all_schedules['Nume'].unique()

    free_people = [person for person in all_people if person not in busy_names]

    df_free_people = pd.DataFrame(free_people, columns=['Persoane Libere'])
    df_free_people.to_excel(output_file, index=False)

#################################################################################
# column style
def adjust_column(column, output_file):
    wb = xl.load_workbook(output_file)
    ws = wb.active
    max_length = 0
    for cell in ws[column]:
        if cell.value:
            cell.alignment = style.Alignment(horizontal='center')
            max_length = max(max_length, len(str(cell.value)))

    ws.column_dimensions[column].width = max_length + 2
    wb.save(output_file)

##################################################################################
# read and format data
file_path = 'Orare_facultate.xlsx'
excel_file = pd.ExcelFile(file_path, engine='openpyxl')
all_schedules = pd.DataFrame()
output_file = 'excel_modificat.xlsx'

for sheet_name in excel_file.sheet_names:
    df_sheet = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    df_sheet['Nume'] = sheet_name
    all_schedules = pd.concat([all_schedules, df_sheet])

all_schedules['Ora incepere'] = pd.to_datetime(all_schedules['Ora incepere'], format='%H:%M:%S').dt.time
all_schedules['Ora sfarsit'] = pd.to_datetime(all_schedules['Ora sfarsit'], format='%H:%M:%S').dt.time

all_schedules = all_schedules[['Nume', 'Zi', 'Ora incepere', 'Ora sfarsit', 'Curs']]
all_schedules = all_schedules.dropna(subset=['Zi'])

print(all_schedules)

##################################################
# Example
ziua = 'Luni'
ora_inceput = '11:30'
ora_sfarsit = '13:30'
find_free_people(ziua, ora_inceput, ora_sfarsit)

adjust_column('A', output_file)
