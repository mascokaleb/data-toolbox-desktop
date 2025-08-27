"""
name: Time & Labor Audit
description: Accepts two Excel files P1 and P2.
required_files:
  P1: "P1"
  P2: "P2"
file_filters:
  P1: "Excel Files (*.xlsx *.xls);;All Files (*)"
  P2: "Excel Files (*.xlsx *.xls);;All Files (*)"
"""
from pathlib import Path
import pandas as pd
from datetime import datetime

def main(P1: Path, P2: Path):
    # Load Excel files into DataFrames
    df_p1 = pd.read_excel(P1)
    df_p2 = pd.read_excel(P2)

    errors = []

    # Check in P2: 'Time and Labor Badge Number' == 'Employee Id'
    for idx, row in df_p2.iterrows():
        badge_num = row.get('Time and Labor Badge Number')
        emp_id = row.get('Employee Id')
        if badge_num != emp_id:
            errors.append({
                'File': 'P2',
                'Row': idx + 2,  # +2 for 1-based Excel row and header row
                'ErrorType': 'Time & Labor Badge Number',
                'ErrorDetail': f"Badge Number '{badge_num}' does not match Employee Id '{emp_id}'",
                'Cols': ['Time and Labor Badge Number', 'Employee Id']
            })

    # Additional P2 policy rules (Work State CA)
    if all(col in df_p2.columns for col in ['Work State', 'Default Hours', 'Payroll Policy Name', 'Employment Type Description']):
        for idx, row in df_p2.iterrows():
            ws = str(row.get('Work State')).strip() if pd.notna(row.get('Work State')) else ''
            dh = row.get('Default Hours')
            pol = str(row.get('Payroll Policy Name')).strip() if pd.notna(row.get('Payroll Policy Name')) else ''
            et = str(row.get('Employment Type Description')).strip() if pd.notna(row.get('Employment Type Description')) else ''

            if ws == 'CA':
                # Rule 1: CA + 80 hours -> California Full Time 8
                if pd.notna(dh) and float(dh) == 80.0 and pol != 'California Full Time 8':
                    errors.append({
                        'File': 'P2',
                        'Row': idx + 2,
                        'ErrorType': 'Payroll Policy',
                        'ErrorDetail': "CA + Default Hours 80 requires Payroll Policy 'California Full Time 8'",
                        'Cols': ['Payroll Policy Name', 'Default Hours']
                    })
                # Rule 2: CA + 75 hours -> California Full Time 7.5
                if pd.notna(dh) and float(dh) == 75.0 and pol != 'California Full Time 7.5':
                    errors.append({
                        'File': 'P2',
                        'Row': idx + 2,
                        'ErrorType': 'Payroll Policy',
                        'ErrorDetail': "CA + Default Hours 75 requires Payroll Policy 'California Full Time 7.5'",
                        'Cols': ['Payroll Policy Name', 'Default Hours']
                    })
                # Rule 3: CA + Temporary FT/PT -> California Part Time
                if et.lower() in ['temporary full time', 'temporary part time'] and pol != 'California Part Time':
                    errors.append({
                        'File': 'P2',
                        'Row': idx + 2,
                        'ErrorType': 'Payroll Policy',
                        'ErrorDetail': "CA + Temporary Employment Type requires Payroll Policy 'California Part Time'",
                        'Cols': ['Payroll Policy Name', 'Employment Type Description']
                    })
            # New rules to add after existing CA rules
            # Check for Job Title and Employment Type Description rules related to Holiday List Code
            if all(col in df_p2.columns for col in ['Job Title', 'Holiday List Code', 'Employment Type Description']):
                job_title = str(row.get('Job Title')).strip() if pd.notna(row.get('Job Title')) else ''
                holiday_list_code = row.get('Holiday List Code')
                emp_type_desc = str(row.get('Employment Type Description')).strip() if pd.notna(row.get('Employment Type Description')) else ''
                if job_title in ['Summer Clerk', 'Partner']:
                    if pd.notna(holiday_list_code) and str(holiday_list_code).strip() != '':
                        errors.append({
                            'File': 'P2',
                            'Row': idx + 2,
                            'ErrorType': 'Holiday List',
                            'ErrorDetail': "Job Title 'Summer Clerk' or 'Partner' requires blank Holiday List Code",
                            'Cols': ['Holiday List Code', 'Job Title']
                        })
                if emp_type_desc in ['Temporary Full Time', 'Temporary Part Time']:
                    if pd.notna(holiday_list_code) and str(holiday_list_code).strip() != '':
                        errors.append({
                            'File': 'P2',
                            'Row': idx + 2,
                            'ErrorType': 'Holiday List',
                            'ErrorDetail': "Temporary Employment Type requires blank Holiday List Code",
                            'Cols': ['Holiday List Code', 'Employment Type Description']
                        })
            # Rule: If Pay Type Code is Hourly then Allow Clock In or Out on Web? must be Yes
            if all(col in df_p2.columns for col in ['Pay Type Code', 'Allow Clock In or Out on Web?']):
                pay_type = str(row.get('Pay Type Code')).strip().lower() if pd.notna(row.get('Pay Type Code')) else ''
                allow_web = str(row.get('Allow Clock In or Out on Web?')).strip().lower() if pd.notna(row.get('Allow Clock In or Out on Web?')) else ''
                hourly_aliases = {'hourly', 'h', 'hrly', 'hour'}
                yes_aliases = {'yes', 'y', 'true', '1'}
                if pay_type in hourly_aliases and allow_web not in yes_aliases:
                    errors.append({
                        'File': 'P2',
                        'Row': idx + 2,
                        'ErrorType': 'Web Clock',
                        'ErrorDetail': "Hourly employees must have 'Allow Clock In or Out on Web?' = Yes",
                        'Cols': ['Allow Clock In or Out on Web?', 'Pay Type Code']
                    })

    # Prepare Error columns for P1 and P2
    df_p1['Error'] = ''
    df_p2['Error'] = ''
    df_p2['ErrorType'] = ''
    df_p2['ErrorDetail'] = ''

    # Map errors back to df_p2 Error columns
    for error in errors:
        if error['File'] == 'P2':
            row_idx = error['Row'] - 2  # convert back to zero-based index
            current_error_type = df_p2.at[row_idx, 'ErrorType']
            current_error_detail = df_p2.at[row_idx, 'ErrorDetail']
            new_error_type = error['ErrorType']
            new_error_detail = error['ErrorDetail']
            if current_error_type:
                df_p2.at[row_idx, 'ErrorType'] = f"{current_error_type}; {new_error_type}"
            else:
                df_p2.at[row_idx, 'ErrorType'] = new_error_type
            if current_error_detail:
                df_p2.at[row_idx, 'ErrorDetail'] = f"{current_error_detail}; {new_error_detail}"
            else:
                df_p2.at[row_idx, 'ErrorDetail'] = new_error_detail

    # Create error DataFrames for P1 and P2 with all original columns plus Error
    df_p1_errors = df_p1[df_p1['Error'] != ''].copy()
    df_p2_errors = df_p2[df_p2['ErrorType'] != ''].copy()

    # Prepare output directory and filename
    out_dir = Path(P1).parent
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_path = out_dir / f"time_labor_audit_{timestamp}.xlsx"

    # Write to Excel with formatting
    with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
        # Write error sheets with full data plus Error columns
        df_p1_errors.to_excel(writer, sheet_name='P1_Errors', index=False)
        df_p2_errors.to_excel(writer, sheet_name='P2_Errors', index=False)

        # Write full sheets
        df_p1.to_excel(writer, sheet_name='P1', index=False)
        df_p2.to_excel(writer, sheet_name='P2', index=False)

        workbook  = writer.book
        red_format = workbook.add_format({'bg_color': '#FFC7CE'})

        # Add tables with autofilter to each sheet
        def add_table_with_autofilter(worksheet, df):
            nrows = len(df)
            ncols = len(df.columns)
            if ncols == 0:
                return
            worksheet.add_table(0, 0, nrows, ncols - 1, {
                'autofilter': True,
                'columns': [{'header': col} for col in df.columns]
            })

        worksheet_p1_err = writer.sheets['P1_Errors']
        add_table_with_autofilter(worksheet_p1_err, df_p1_errors)

        worksheet_p2_err = writer.sheets['P2_Errors']
        add_table_with_autofilter(worksheet_p2_err, df_p2_errors)

        worksheet_p1 = writer.sheets['P1']
        add_table_with_autofilter(worksheet_p1, df_p1)

        worksheet_p2 = writer.sheets['P2']
        add_table_with_autofilter(worksheet_p2, df_p2)

        # Highlight errored cells in P1 full sheet
        error_col_p1 = df_p1.columns.get_loc('Error')
        for row_idx, error_msg in enumerate(df_p1['Error']):
            if error_msg:
                # Highlight the Error cell red
                worksheet_p1.write(row_idx + 1, error_col_p1, error_msg, red_format)

        # Highlight errored cells in P2 full sheet
        error_type_col_p2 = df_p2.columns.get_loc('ErrorType')
        error_detail_col_p2 = df_p2.columns.get_loc('ErrorDetail')

        # First highlight the error columns themselves when any error present on that row
        for row_idx, (etype, edetail) in enumerate(zip(df_p2['ErrorType'], df_p2['ErrorDetail'])):
            if etype or edetail:
                worksheet_p2.write(row_idx + 1, error_type_col_p2, etype, red_format)
                worksheet_p2.write(row_idx + 1, error_detail_col_p2, edetail, red_format)

        # Then for each recorded error, highlight the specific columns involved
        for err in errors:
            if err.get('File') != 'P2':
                continue
            excel_row = err['Row'] - 1
            for col_name in err.get('Cols', []):
                if col_name in df_p2.columns:
                    cidx = df_p2.columns.get_loc(col_name)
                    val = df_p2.iat[excel_row - 1, cidx]
                    if pd.isna(val):
                        worksheet_p2.write_blank(excel_row, cidx, None, red_format)
                    else:
                        worksheet_p2.write(excel_row, cidx, val, red_format)

        # Highlight errored cells in P1_Errors sheet
        if not df_p1_errors.empty:
            error_col_p1_err = df_p1_errors.columns.get_loc('Error')
            for row_idx, error_msg in enumerate(df_p1_errors['Error']):
                if error_msg:
                    worksheet_p1_err.write(row_idx + 1, error_col_p1_err, error_msg, red_format)

        # Highlight errored cells in P2_Errors sheet
        if not df_p2_errors.empty:
            error_type_col_p2_err = df_p2_errors.columns.get_loc('ErrorType')
            error_detail_col_p2_err = df_p2_errors.columns.get_loc('ErrorDetail')

            # Map original Excel row number to row index within the filtered error sheet
            error_rows_set = set(df_p2_errors.index.tolist())
            for err in errors:
                if err.get('File') != 'P2':
                    continue
                original_row = err['Row'] - 2  # zero-based index in df_p2
                if original_row not in error_rows_set:
                    continue
                # Find the relative row in the error sheet
                rel_row = df_p2_errors.index.get_loc(original_row)
                excel_row = rel_row + 1
                for col_name in err.get('Cols', []):
                    if col_name in df_p2_errors.columns:
                        cidx = df_p2_errors.columns.get_loc(col_name)
                        val = df_p2_errors.iat[rel_row, cidx]
                        if pd.isna(val):
                            worksheet_p2_err.write_blank(excel_row, cidx, None, red_format)
                        else:
                            worksheet_p2_err.write(excel_row, cidx, val, red_format)
            # Also highlight the error columns themselves
            for row_idx, (etype, edetail) in enumerate(zip(df_p2_errors['ErrorType'], df_p2_errors['ErrorDetail'])):
                if etype or edetail:
                    worksheet_p2_err.write(row_idx + 1, error_type_col_p2_err, etype, red_format)
                    worksheet_p2_err.write(row_idx + 1, error_detail_col_p2_err, edetail, red_format)

    return out_path
