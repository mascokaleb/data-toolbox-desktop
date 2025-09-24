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

    # P1 rule: If Employment Type is RFT or RPT & Pay Type is hourly & Default Hours = 80 THEN Time Off Type must be PTO
    if all(col in df_p1.columns for col in ['Employment Type', 'Pay Type', 'Default Hours', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            pay_type = str(row.get('Pay Type')).strip().lower() if pd.notna(row.get('Pay Type')) else ''
            def_hours = row.get('Default Hours')
            time_off = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if emp_type in {'RFT', 'RPT'} and pay_type in {'hourly', 'h', 'hrly', 'hour'} and pd.notna(def_hours) and float(def_hours) == 80.0:
                if time_off != 'PTO':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'PTO - Paid Time Off 40 Hours',
                        'ErrorDetail': "RFT/RPT hourly with Default Hours = 80 requires Time Off Type 'PTO'",
                        'Cols': ['Time Off Type', 'Employment Type', 'Pay Type', 'Default Hours']
                    })

    # P1 rule: If Employment Type is RFT or RPT & Pay Type is hourly & Default Hours = 75 THEN Time Off Type must be PTO75
    if all(col in df_p1.columns for col in ['Employment Type', 'Pay Type', 'Default Hours', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            pay_type = str(row.get('Pay Type')).strip().lower() if pd.notna(row.get('Pay Type')) else ''
            def_hours = row.get('Default Hours')
            time_off = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if emp_type in {'RFT', 'RPT'} and pay_type in {'hourly', 'h', 'hrly', 'hour'} and pd.notna(def_hours) and float(def_hours) == 75.0:
                if time_off != 'PTO75':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'PTO75 - PAID TIME OFF 37.5 HOURS',
                        'ErrorDetail': "RFT/RPT hourly with Default Hours = 75 requires Time Off Type 'PTO75'",
                        'Cols': ['Time Off Type', 'Employment Type', 'Pay Type', 'Default Hours']
                    })

    # P1 rule: If Pay Type is NOT 'Attorney' and Position is 'Staff' THEN Time Off Type must be 'VAC'
    if all(col in df_p1.columns for col in ['Pay Type', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            pay_type_val = str(row.get('Pay Type')).strip().lower() if pd.notna(row.get('Pay Type')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if position_val == 'staff' and pay_type_val != 'attorney':
                if time_off_val != 'VAC':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'VAC - VACATION',
                        'ErrorDetail': "Non-attorney Staff requires Time Off Type 'VAC'",
                        'Cols': ['Time Off Type', 'Pay Type', 'Position']
                    })

    # P1 rule: If Pay Type is salary and Position is NOT Attorney THEN Time Off Type must be SICK
    if all(col in df_p1.columns for col in ['Pay Type', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            pay_type_val = str(row.get('Pay Type')).strip().lower() if pd.notna(row.get('Pay Type')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if pay_type_val == 'salary' and position_val != 'attorney':
                if time_off_val != 'SICK':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'SICK - SICK',
                        'ErrorDetail': "Salary + non-Attorney requires Time Off Type 'SICK'",
                        'Cols': ['Time Off Type', 'Pay Type', 'Position']
                    })

    # P1 rule: If Work Location is SEATTLE or REMOTE-WASHINGTON (Seattle) and Position is Attorney THEN Time Off Type must be SICK
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'SEATTLE', 'REMOTE-WASHINGTON (SEATTLE)'} and position_val == 'attorney':
                if time_off_val != 'SICK':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'SICK - SICK',
                        'ErrorDetail': "Seattle-based Attorneys require Time Off Type 'SICK'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: If Work Location is SEATTLE or REMOTE-WASHINGTON (Seattle) and Employment Type is NOT RPT or RFT THEN Time Off Type must be SICK
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'SEATTLE', 'REMOTE-WASHINGTON (SEATTLE)'} and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'SICK':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'SICK - SICK',
                        'ErrorDetail': "Seattle-based non-RPT/RFT employees require Time Off Type 'SICK'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })

    # P1 rule: If Work Location is NEW JERSEY or REMOTE-NEW JERSEY (Berkley Heights) and Position is Attorney THEN Time Off Type must be SICK
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'NEW JERSEY', 'REMOTE-NEW JERSEY (BERKLEY HEIGHTS)'} and position_val == 'attorney':
                if time_off_val != 'SICK':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'SICK - SICK',
                        'ErrorDetail': "New Jersey-based Attorneys require Time Off Type 'SICK'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: Massachusetts S_MA - If Work Location is REMOTE-MASSACHUSETTS (Boston) or Boston AND Position is Attorney THEN Time Off Type must be S_MA
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-MASSACHUSETTS (BOSTON)', 'BOSTON'} and position_val == 'attorney':
                if time_off_val != 'S_MA':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_MA - Massachusetts Sick',
                        'ErrorDetail': "Massachusetts-based Attorneys require Time Off Type 'S_MA'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: Massachusetts S_MA - If Work Location is REMOTE-MASSACHUSETTS (Boston) or Boston AND Employment Type is NOT RPT or RFT THEN Time Off Type must be S_MA
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-MASSACHUSETTS (BOSTON)', 'BOSTON'} and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_MA':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_MA - Massachusetts Sick',
                        'ErrorDetail': "Massachusetts-based non-RPT/RFT employees require Time Off Type 'S_MA'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })

    # P1 rule: Arizona S_AZ1 - If Work Location is REMOTE-ARIZONA (Phoenix) or Phoenix AND Position is Attorney THEN Time Off Type must be S_AZ1
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-ARIZONA (PHOENIX)', 'PHOENIX'} and position_val == 'attorney':
                if time_off_val != 'S_AZ1':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_AZ1 - Arizona Sick Pay',
                        'ErrorDetail': "Arizona-based Attorneys require Time Off Type 'S_AZ1'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: Arizona S_AZ1 - If Work Location is REMOTE-ARIZONA (Phoenix) or Phoenix AND Employment Type is NOT RPT or RFT THEN Time Off Type must be S_AZ1
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-ARIZONA (PHOENIX)', 'PHOENIX'} and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_AZ1':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_AZ1 - Arizona Sick Pay',
                        'ErrorDetail': "Arizona-based non-RPT/RFT employees require Time Off Type 'S_AZ1'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })

    # P1 rule: Chicago S_CPL - If Employment Type is NOT RPT or RFT AND Work Location is REMOTE-ILLINOIS (Chicago) or Chicago THEN Time Off Type must be S_CPL
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-ILLINOIS (CHICAGO)', 'CHICAGO'} and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_CPL':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_CPL - Chicago Paid Leave',
                        'ErrorDetail': "Chicago-based non-RPT/RFT employees require Time Off Type 'S_CPL'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })


    # P1 rule: Maryland S_MD - If Work Location is REMOTE-MARYLAND (Baltimore) or BALTIMORE AND Position is Attorney THEN Time Off Type must be S_MD
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-MARYLAND (BALTIMORE)', 'BALTIMORE'} and position_val == 'attorney':
                if time_off_val != 'S_MD':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_MD - Maryland Sick',
                        'ErrorDetail': "Maryland-based Attorneys require Time Off Type 'S_MD'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: Maryland S_MD - If Employment Type is NOT RPT or RFT AND Work Location is REMOTE-MARYLAND (Baltimore) or BALTIMORE THEN Time Off Type must be S_MD
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-MARYLAND (BALTIMORE)', 'BALTIMORE'} and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_MD':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_MD - Maryland Sick',
                        'ErrorDetail': "Maryland-based non-RPT/RFT employees require Time Off Type 'S_MD'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })


    # P1 rule: Chicago S_CHI - If Work Location is REMOTE-ILLINOIS (Chicago) or Chicago AND Position is Attorney THEN Time Off Type must be S_CHI
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in {'REMOTE-ILLINOIS (CHICAGO)', 'CHICAGO'} and position_val == 'attorney':
                if time_off_val != 'S_CHI':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_CHI - Chicago Sick',
                        'ErrorDetail': "Chicago-based Attorneys require Time Off Type 'S_CHI'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: California S_CA - If Work Location is in specified CA locations AND Position is Attorney THEN Time Off Type must be S_CA
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        CA_LOCS = {
            'SAN DIEGO', 'SACRAMENTO', 'LOS ANGELES', 'IRVINE', 'WOODLAND HILLS',
            'REMOTE-CALIFORN (LOS ANGELES)', 'REMOTE CALIFORN (IRVINE)', 'REMOTE-CALIFORN (SACRAMENTO)'
        }
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in CA_LOCS and position_val == 'attorney':
                if time_off_val != 'S_CA':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_CA - California Sick',
                        'ErrorDetail': "California-based Attorneys in specified locations require Time Off Type 'S_CA'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: California S_CA - If Work Location is in specified CA locations AND Employment Type is NOT RPT or RFT THEN Time Off Type must be S_CA
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        CA_LOCS = {
            'SAN DIEGO', 'SACRAMENTO', 'LOS ANGELES', 'IRVINE', 'WOODLAND HILLS',
            'REMOTE-CALIFORN (LOS ANGELES)', 'REMOTE CALIFORN (IRVINE)', 'REMOTE-CALIFORN (SACRAMENTO)'
        }
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in CA_LOCS and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_CA':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_CA - California Sick',
                        'ErrorDetail': "California-based non-RPT/RFT employees in specified locations require Time Off Type 'S_CA'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })

    # P1 rule: San Francisco CSFPH - If Work Location is San Francisco or REMOTE-CALIF (San Francisco) THEN Time Off Type must be CSFPH
    if all(col in df_p1.columns for col in ['Work Location', 'Time Off Type']):
        SF_LOCS = {'SAN FRANCISCO', 'REMOTE-CALIF (SAN FRANCISCO)'}
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in SF_LOCS:
                if time_off_val != 'CSFPH':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'CSFPH - San Francisco',
                        'ErrorDetail': "San Francisco locations require Time Off Type 'CSFPH'",
                        'Cols': ['Time Off Type', 'Work Location']
                    })


    # P1 rule: San Francisco S_SF - If Work Location is San Francisco or REMOTE-CALIF (San Francisco) AND Position is Attorney THEN Time Off Type must be S_SF
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        SF_LOCS = {'SAN FRANCISCO', 'REMOTE-CALIF (SAN FRANCISCO)'}
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in SF_LOCS and position_val == 'attorney':
                if time_off_val != 'S_SF':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_SF - San Francisco Sick',
                        'ErrorDetail': "San Francisco-based Attorneys require Time Off Type 'S_SF'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: Oregon S_OR1 - If Work Location is PORTLAND OR THEN Time Off Type must be S_OR1
    if all(col in df_p1.columns for col in ['Work Location', 'Time Off Type']):
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc == 'PORTLAND OR':
                if time_off_val != 'S_OR1':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_OR1 - Oregon Sick',
                        'ErrorDetail': "Portland OR locations require Time Off Type 'S_OR1'",
                        'Cols': ['Time Off Type', 'Work Location']
                    })


    # P1 rule: Washington DC S_DC1 - If Work Location is REMOTE-WASHINGTON DC or WASHINGTON DC (DC) AND Position is Attorney AND Job Title is NOT Partner THEN Time Off Type must be S_DC1
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Job Title', 'Time Off Type']):
        DC_LOCS = {'REMOTE-WASHINGTON DC', 'WASHINGTON DC (DC)'}
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            job_title_val = str(row.get('Job Title')).strip().lower() if pd.notna(row.get('Job Title')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in DC_LOCS and position_val == 'attorney' and job_title_val != 'partner':
                if time_off_val != 'S_DC1':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_DC1 - Washington DC Sick',
                        'ErrorDetail': "Washington DC-based non-Partner Attorneys require Time Off Type 'S_DC1'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position', 'Job Title']
                    })

    # P1 rule: New York City S_NYC - If Work Location is REMOTE-NEW YORK (New York City) or NEW YORK AND Position is Attorney THEN Time Off Type must be S_NYC
    if all(col in df_p1.columns for col in ['Work Location', 'Position', 'Time Off Type']):
        NYC_LOCS = {'REMOTE-NEW YORK (NEW YORK CITY)', 'NEW YORK'}
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            position_val = str(row.get('Position')).strip().lower() if pd.notna(row.get('Position')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in NYC_LOCS and position_val == 'attorney':
                if time_off_val != 'S_NYC':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_NYC - New York City Sick',
                        'ErrorDetail': "New York City-based Attorneys require Time Off Type 'S_NYC'",
                        'Cols': ['Time Off Type', 'Work Location', 'Position']
                    })

    # P1 rule: New York City S_NYC - If Work Location is REMOTE-NEW YORK (New York City) or NEW YORK AND Employment Type is NOT RPT or RFT THEN Time Off Type must be S_NYC
    if all(col in df_p1.columns for col in ['Work Location', 'Employment Type', 'Time Off Type']):
        NYC_LOCS = {'REMOTE-NEW YORK (NEW YORK CITY)', 'NEW YORK'}
        for idx, row in df_p1.iterrows():
            work_loc = str(row.get('Work Location')).strip().upper() if pd.notna(row.get('Work Location')) else ''
            emp_type = str(row.get('Employment Type')).strip().upper() if pd.notna(row.get('Employment Type')) else ''
            time_off_val = str(row.get('Time Off Type')).strip().upper() if pd.notna(row.get('Time Off Type')) else ''
            if work_loc in NYC_LOCS and emp_type not in {'RPT', 'RFT'}:
                if time_off_val != 'S_NYC':
                    errors.append({
                        'File': 'P1',
                        'Row': idx + 2,
                        'ErrorType': 'S_NYC - New York City Sick',
                        'ErrorDetail': "New York City-based non-RPT/RFT employees require Time Off Type 'S_NYC'",
                        'Cols': ['Time Off Type', 'Work Location', 'Employment Type']
                    })

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
    df_p1['ErrorType'] = ''
    df_p1['ErrorDetail'] = ''
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

    # Map errors back to df_p1 Error columns
    for error in errors:
        if error['File'] == 'P1':
            row_idx = error['Row'] - 2
            current_error_type = df_p1.at[row_idx, 'ErrorType']
            current_error_detail = df_p1.at[row_idx, 'ErrorDetail']
            new_error_type = error['ErrorType']
            new_error_detail = error['ErrorDetail']
            df_p1.at[row_idx, 'ErrorType'] = f"{current_error_type}; {new_error_type}".strip('; ').strip() if current_error_type else new_error_type
            df_p1.at[row_idx, 'ErrorDetail'] = f"{current_error_detail}; {new_error_detail}".strip('; ').strip() if current_error_detail else new_error_detail

    # Create error DataFrames for P1 and P2 with all original columns plus ErrorType/ErrorDetail
    df_p1_errors = df_p1[(df_p1['ErrorType'] != '') | (df_p1['ErrorDetail'] != '')].copy()
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
        error_type_col_p1 = df_p1.columns.get_loc('ErrorType')
        error_detail_col_p1 = df_p1.columns.get_loc('ErrorDetail')
        for row_idx, (etype, edetail) in enumerate(zip(df_p1['ErrorType'], df_p1['ErrorDetail'])):
            if etype or edetail:
                worksheet_p1.write(row_idx + 1, error_type_col_p1, etype, red_format)
                worksheet_p1.write(row_idx + 1, error_detail_col_p1, edetail, red_format)
        # Highlight specific columns per P1 error
        for err in errors:
            if err.get('File') != 'P1':
                continue
            excel_row = err['Row'] - 1
            for col_name in err.get('Cols', []):
                if col_name in df_p1.columns:
                    cidx = df_p1.columns.get_loc(col_name)
                    val = df_p1.iat[excel_row - 1, cidx]
                    if pd.isna(val):
                        worksheet_p1.write_blank(excel_row, cidx, None, red_format)
                    else:
                        worksheet_p1.write(excel_row, cidx, val, red_format)

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
            error_type_col_p1_err = df_p1_errors.columns.get_loc('ErrorType')
            error_detail_col_p1_err = df_p1_errors.columns.get_loc('ErrorDetail')
            # Highlight error columns
            for row_idx, (etype, edetail) in enumerate(zip(df_p1_errors['ErrorType'], df_p1_errors['ErrorDetail'])):
                if etype or edetail:
                    worksheet_p1_err.write(row_idx + 1, error_type_col_p1_err, etype, red_format)
                    worksheet_p1_err.write(row_idx + 1, error_detail_col_p1_err, edetail, red_format)
            # Highlight specific columns per P1 error
            error_rows_set_p1 = set(df_p1_errors.index.tolist())
            for err in errors:
                if err.get('File') != 'P1':
                    continue
                original_row = err['Row'] - 2
                if original_row not in error_rows_set_p1:
                    continue
                rel_row = df_p1_errors.index.get_loc(original_row)
                excel_row = rel_row + 1
                for col_name in err.get('Cols', []):
                    if col_name in df_p1_errors.columns:
                        cidx = df_p1_errors.columns.get_loc(col_name)
                        val = df_p1_errors.iat[rel_row, cidx]
                        if pd.isna(val):
                            worksheet_p1_err.write_blank(excel_row, cidx, None, red_format)
                        else:
                            worksheet_p1_err.write(excel_row, cidx, val, red_format)

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
