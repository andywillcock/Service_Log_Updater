import pandas as pd
from datetime import datetime, date
from easygui import fileopenbox, msgbox

today = datetime.today().strftime('%Y-%b-%d')

writer = pd.ExcelWriter('('+today+') '+'SO Centralized Support Log.xlsx', engine='xlsxwriter',
                        datetime_format='dd mmm yyyy',
                        date_format='dd mmm yyyy'
                        )


def create_cell_fill_format(bg_color):
    cell_format = workbook.add_format()
    cell_format.set_pattern(1)
    cell_format.set_bg_color(bg_color)
    return cell_format


current_log_file = fileopenbox(msg = "Select Current SO Centralized Support Log Excel Workbook")
current_log = pd.read_excel(current_log_file, sheet_name =2, header=0, skiprows=4)
current_log = current_log[pd.isna(current_log['Study ID']) == False]

so_file = fileopenbox(msg="Select SO Export")
so = pd.read_excel(so_file, sheet_name=0, header=0, skiprows=8)
so = so.drop('Column Headers:', axis=1)
so = so[pd.isna(so['Study ID']) == False]

biw_file = fileopenbox(msg="Select BIW Export")
biw = pd.read_excel(biw_file, sheet_name=0, header=0, skiprows=2)
biw = biw.rename(index=str, columns={'Study':'Study ID'})

agg_file = fileopenbox(msg="Select SO Aggregate Performance Report")
agg = pd.read_excel(agg_file, sheet_name=0, header=0, skiprows=5)
agg = agg.rename(index=str, columns={'SID' : 'Study ID'})
agg = agg[['Study ID','PFSAD','PLSE','PrLSE']]
agg = agg.astype(str)
agg = agg.replace('--', '')
for col in ['PFSAD','PLSE','PrLSE']:
    for i in range(len(agg)):
        if agg[col][i] != '' and agg[col][i] != 'nan':
            agg[col][i] = datetime.strptime(agg[col][i], '%d-%b-%Y')
            agg[col][i] = agg[col][i].date()

in_study_list = []
for row in range(len(so)):
    if so['Study ID'][row] in current_log['Study ID'].values:
        in_study_list.append('Yes')
    else:
        in_study_list.append('No')
so['In Study List'] = pd.Series(in_study_list)

# Filter Study Optimizer Export
so = so[so['Molecule'] != 'NO SOURCE']

so = so[so['Plan Status'] != 'Approved Plan']
so = so[so['Plan Status'] != 'No Plan Needed']

so = so[so['Molecule'] != 'N/A']

so = so[so['Phase'] != 'Phase I']
so = so[so['Phase'] != 'Phase IV']
so = so[so['Phase'] != 'Source Unavailable']

# Select studies that are not already in the SO Centralized Support Log
new_studies = so[so['In Study List'] == 'No'].reset_index(drop=True)
new_studies = new_studies.rename(index=str, columns={'StudyOptimizer Status': 'SO Status'})

# Create a copy of the current SO Centralized Support Log and append new studies to the dataframe
new_log = current_log
new_log = new_log.append(new_studies)

# Check to see if Current SO has the correct columns, if not add columns so that the rest of the update runs correctly
if 'PFSAD (SO)' not in new_log.columns.get_values():
    new_log['PFSAD (SO)'] = ''
if 'PFSAD (BIW)' not in new_log.columns.get_values():
    new_log['PFSAD (BIW)'] = ''
if 'PLSE (SO)' not in new_log.columns.get_values():
    new_log['PLSE (SO)'] = ''
if 'PrLSE (SO)' not in new_log.columns.get_values():
    new_log['PrLSE (SO)'] = ''
if 'Most Recent Adjustment or Scenario Date (Formula)' not in new_log.columns.get_values():
    new_log['Most Recent Adjustment or Scenario Date (Formula)'] = ''

# Select and reorder columns to match the current SO Centralized Support Log format
new_log = new_log[['Study ID','OIL','OIL Comment','Action','Code','Study Name',
                   'Therapeutic Area','Indication','Description','Theme','Molecule','Plan Status','SO Status','Sponsor',
                   'Phase','Study Manager','PFSAD (SO)','PFSAD (BIW)','PLSE (SO)','PrLSE (SO)',
                   'Weeks Ahead (+) or Behind (-)','Actual FPI','Actual LPI','Enrollment Status','Approved Plan?',
                   'Approved Plan Mod. Date','Most Recent Adjustment or Scenario Date',
                   'Most Recent Adjustment or Scenario Date (Formula)','Most Recent Adjustment/Scenario Validity',
                   'Comments','Number of New Scenarios Requested','Date Requested','Requestor Name','Assigned to',
                   'Time Required to Meet with SMT & Draft Scenarios\n(approx hours)']]

# Add PFSAD, PLSE, PrLSE from Study Optimizer and PFSAD from BIW (performs like VLOOKUP)
new_log = new_log.merge(agg[['Study ID','PFSAD','PLSE','PrLSE']], how='inner', on='Study ID', suffixes=('_SO',''))
new_log = new_log.merge(biw[['Study ID','First Site Activated (Planned)']],
                        how='inner', on='Study ID', suffixes=('_BIW',''))
new_log = new_log[['Study ID','OIL','OIL Comment','Action','Code','Study Name',
                   'Therapeutic Area','Indication','Description','Theme','Molecule','Plan Status','SO Status','Sponsor',
                   'Phase','Study Manager','PFSAD','First Site Activated (Planned)','PLSE','PrLSE',
                   'Weeks Ahead (+) or Behind (-)','Actual FPI','Actual LPI','Enrollment Status','Approved Plan?',
                   'Approved Plan Mod. Date','Most Recent Adjustment or Scenario Date',
                   'Most Recent Adjustment or Scenario Date (Formula)','Most Recent Adjustment/Scenario Validity',
                   'Comments','Number of New Scenarios Requested','Date Requested','Requestor Name','Assigned to',
                   'Time Required to Meet with SMT & Draft Scenarios\n(approx hours)']]

# Rename some columns
new_log = new_log.rename(index=str, columns = {'PFSAD': 'PFSAD (SO)', 'First Site Activated (Planned)': 'PFSAD (BIW)',
                                               'PLSE':'PLSE (SO)','PrLSE':'PrLSE (SO)'})


new_log = new_log.replace(r'[^\x00-\x7F]', '', regex=True)
new_log = new_log.replace('--', '')
new_log['PFSAD (SO)'] = pd.to_datetime(new_log['PFSAD (SO)'], format='%d/%b/%Y', errors='ignore')
new_log['PFSAD (BIW)'] = pd.to_datetime(new_log['PFSAD (BIW)'], format='%d/%b/%Y', errors='ignore')
new_log['PLSE (SO)'] = pd.to_datetime(new_log['PLSE (SO)'], format='%d/%b/%Y', errors='ignore')
new_log['PrLSE (SO)'] = pd.to_datetime(new_log['PrLSE (SO)'], format='%d/%b/%Y', errors='ignore')

sheet_name = 'SO Centralized Support Log'
             # +str(datetime.today()).split()[0]
new_log.to_excel(writer,sheet_name=sheet_name, index = False)
workbook = writer.book
worksheet = writer.sheets[sheet_name]

#Format Header row to blue fill w/ white text and wrap text
cell_format = workbook.add_format()
cell_format.set_pattern(1)
cell_format.set_bg_color('#336699')
cell_format.set_font_color('white')
cell_format.set_text_wrap()
worksheet.write_row('A1:AL1', new_log.columns.values,cell_format)

format1 = create_cell_fill_format('yellow')
format_gt12 = create_cell_fill_format('red')
format_8to12 = create_cell_fill_format('orange')
format_4to8 = create_cell_fill_format('yellow')
format_0to4 = create_cell_fill_format('green')
format_blank = create_cell_fill_format('no color')

# Apply formulas and conditional formatting
for row_num in range(2, len(new_log)):

    # Formula to fill Code column based on contents of other columns of interest
    worksheet.write_formula(row_num - 1, 6, '=IF(F%d="Hide","Hide",IF(OR(AND(S%d<>"--",S%d>TODAY()),AND(T%d<>"--",T%d>TODAY())),'
                                           '"First Site Not Activated",IF(AA%d="No","No Approved Plan",IF(AA%d="Yes",'
                                           '"Approved Plan","N/A"))))'
                                            % (row_num, row_num, row_num, row_num, row_num, row_num, row_num))

    # Formula to calculate Weeks Ahead or Behind Schedule
    worksheet.write_formula(row_num - 1, 20, '=IF(OR(T%d="",S%d=""),"",ROUND((T%d-S%d)/7,0))' %
                            (row_num, row_num, row_num, row_num))


    # Conditional formatting for Study ID column to show if its PFSAD is within 8 weeks of today
    worksheet.conditional_format('A%d' % (row_num), {'type':     'formula',
                                                     'criteria': '=OR(AND(Q%d<>0,AND($Q%d > TODAY(),$Q%d<(TODAY()+56))),'
                                                                 'AND(R%d<>0,AND($R%d > TODAY(),$R%d<(TODAY()+56))))'
                                                     % (row_num, row_num, row_num, row_num, row_num, row_num),
                                                     'format':   format1,
                                                     'stop_if_true': True})

    # Conditional formatting for Weeks Ahead or Behind column
    worksheet.conditional_format('U%d' % (row_num), {'type': 'formula',
                                                     'criteria': 'U%d=""' % (row_num),
                                                     'format': format_blank})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria': '>=',
                                                     'value': 12,
                                                     'format': format_gt12})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria' : 'between',
                                                     'minimum': 8,
                                                     'maximum': 12,
                                                     'format': format_8to12})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria' : 'between',
                                                     'minimum': 4,
                                                     'maximum': 8,
                                                     'format': format_4to8})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria' : 'between',
                                                     'minimum': 0,
                                                     'maximum': 4,
                                                     'format': format_0to4})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria': '<=',
                                                     'value': -12,
                                                     'format': format_gt12})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria': 'between',
                                                     'minimum': -8,
                                                     'maximum': -12,
                                                     'format': format_8to12})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria': 'between',
                                                     'minimum': -4,
                                                     'maximum': -8,
                                                     'format': format_4to8})
    worksheet.conditional_format('U%d' % (row_num), {'type': 'cell',
                                                     'criteria': 'between',
                                                     'minimum': 0,
                                                     'maximum': -4,
                                                     'format': format_0to4})

worksheet.set_column(0,37,19)
worksheet.set_column('L:P',37, None, {'hidden' : 1})
worksheet.set_column('Z:AA',37, None, {'hidden' : 1})
worksheet.set_column('AF:AF',37, None, {'hidden' : 1})


msgbox('SO Service Log Update Complete....file saved to the folder containing this program')
writer.save()
