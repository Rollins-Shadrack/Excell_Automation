import datetime
import os

import numpy as np
import pandas as pd
import streamlit as st

processed_files = []
global excel_name


def team_summary(productivity, team_leader):
    # Remove leading and trailing spaces in column names
    team_leader.columns = team_leader.columns.str.strip()
    productivity.columns = productivity.columns.str.strip()

    # Group by 'TEAM LEADER' in team_leader dataset and calculate the total ACMs and Files Target
    team_leader_summary = team_leader.groupby('TEAM LEADER').agg(
        {'Account  Managers': lambda x: x.nunique(), 'TARGET': 'sum'}).reset_index()

    # Group by 'TEAM LEADER' in productivity dataset and calculate the total sum of ACMs and Phone calls made
    productivity_summary = productivity.groupby('TEAM LEADER').agg(
        {'ACM': lambda x: x.nunique(), 'Phone calls Made': 'sum'}).reset_index()

    # Merge the two summaries based on 'TEAM LEADER'
    team_summary = team_leader_summary.merge(productivity_summary, on='TEAM LEADER')

    # Add the 'Inactive ACMs' column
    team_summary['Inactive ACMs'] = team_summary['Account  Managers'] - team_summary['ACM']

    # Add the 'Files Deficit' column
    team_summary['Files Deficit'] = team_summary['TARGET'] - team_summary['Phone calls Made']

    team_summary = team_summary.rename(columns={
        'Account  Managers': 'TOTAL ACMS',
        'ACM': 'ACMS ACTIVE',
        'Inactive ACMs': 'INACTIVE ACMS',
        'TARGET': 'FILES TARGET',
        'Phone calls Made': 'FILES WORKED',
        'Files Deficit': 'FILES DEFICIT'
    })
    sum_of_acms = team_summary['TOTAL ACMS'].sum()
    sum_of_active_acms = team_summary['ACMS ACTIVE'].sum()
    sum_of_inactive_acms = team_summary['INACTIVE ACMS'].sum()
    sum_of_files_target = team_summary['FILES TARGET'].sum()
    sum_of_files_worked = team_summary['FILES WORKED'].sum()
    sum_of_deficit_files = team_summary['FILES DEFICIT'].sum()
    total_team_summary = pd.DataFrame(
        {'TEAM LEADER': ['Total'], 'TOTAL ACMS': [sum_of_acms], 'ACMS ACTIVE': [sum_of_active_acms],
         'INACTIVE ACMS': [sum_of_inactive_acms], 'FILES TARGET': [sum_of_files_target],
         'FILES WORKED': [sum_of_files_worked], 'FILES DEFICIT': [sum_of_deficit_files]})
    team_summary = team_summary.reindex(
        columns=['TEAM LEADER', 'TOTAL ACMS', 'ACMS ACTIVE', 'INACTIVE ACMS', 'FILES TARGET', 'FILES WORKED',
                 'FILES DEFICIT'])

    team_summary = pd.concat([team_summary, total_team_summary], ignore_index=True)

    return team_summary


def branch_summary(dataset):
    # Group by 'Branch' and calculate the number of unique ACMs and the sum of 'Phone calls made'
    grouped_data = dataset.groupby('Branch').agg({'ACM': 'nunique', 'Phone calls Made': 'sum'}).reset_index()

    # Rename the columns for clarity
    grouped_data = grouped_data.rename(columns={'ACM': 'Active ACMs', 'Phone calls Made': 'Files Worked'})

    # Calculate totals
    active_acms_totals = grouped_data['Active ACMs'].sum()
    files_worked_totals = grouped_data['Files Worked'].sum()

    # Create totals rows
    total_branch_summary = pd.DataFrame(
        {'Branch': ['Total'], 'Active ACMs': [active_acms_totals], 'Files Worked': [files_worked_totals]})

    # Concatenate totals rows with the grouped data
    result = pd.concat([grouped_data, total_branch_summary], ignore_index=True)

    return result


def filter_dataset_by_current_date(ptp_with_due_dates):
    # Get the current date
    current_date = datetime.date.today()

    # Remove the time portion from the "created" column
    ptp_with_due_dates['created'] = ptp_with_due_dates['created'].str.split('T').str[0]

    # Convert the "created" column to datetime
    ptp_with_due_dates['created'] = pd.to_datetime(ptp_with_due_dates['created'])

    # Extract the date portion from the "created" column and compare with the current date
    filtered_dataset = ptp_with_due_dates[ptp_with_due_dates['created'].dt.date == current_date]

    return filtered_dataset


def remove_duplicates_and_sum_by_acm(productivity_dt):
    # Convert 'Ptp Amount' column to numeric
    productivity_dt['Ptp Amount'] = pd.to_numeric(productivity_dt['Ptp Amount'], errors='coerce')

    # Group the DataFrame by ACM and sum the 'Ptp Amount'
    summed_data = productivity_dt.groupby('ACM').agg({'Ptp Amount': 'sum'}).reset_index()

    # Merge summed data with original data
    cleaned_data = pd.merge(productivity_dt, summed_data, on='ACM', suffixes=('', '_sum'))

    # Remove duplicates based on ACM column
    cleaned_data = cleaned_data.drop_duplicates(subset='ACM')

    cleaned_data.drop('Ptp Amount', axis=1, inplace=True)

    cleaned_data.rename(columns={'Ptp Amount_sum': 'Ptp Amount'}, inplace=True)

    return cleaned_data


def filter_remove_zero_calls(df):
    # Filter rows where "Phone calls made" is not equal to 0
    filtered_df = df.loc[df["Phone calls Made"] != 0]

    return filtered_df


def calculate_column_sum(dataset, column_name):
    # Convert dataset to DataFrame
    productivity_dt = pd.DataFrame(dataset)

    # Remove duplicates
    productivity_dt = productivity_dt.drop_duplicates()

    # Remove spaces in the column titles
    productivity_dt.columns = productivity_dt.columns.str.strip()

    # Convert from string to numeric
    productivity_dt[column_name] = pd.to_numeric(productivity_dt[column_name], errors='coerce')

    # Calculate the sum of the column
    total_sum = productivity_dt[column_name].sum()

    return total_sum


def upload_excel_files():
    st.header("Productivity excel files")

    # setting allowed extensions and allowed files
    ALLOWED_EXTENSIONS = ['.xlsx', '.xls']
    DESIRED_FILENAMES = ['DAILY ACM PERFORMANCE.xlsx', 'TEAM LEADER ACMS.xlsx', 'ACM PRODUCTIVITY TEMPLATE.xlsx',
                         'PTPs Created by ACMs.xlsx', 'Debtors spoken to.xlsx',
                         'PTPS CREATED WITH THEIR DUE DATES.xlsx']

    # allow only excel files
    def allowed_file(filename):
        return any(filename.endswith(ext) for ext in ALLOWED_EXTENSIONS)

    # allow only files with specific names
    def desired_file(filename):
        return filename in DESIRED_FILENAMES

    # now upload your files
    uploaded_files = st.file_uploader("Upload Excel files required to calculate Productivity",
                                      accept_multiple_files=True)

    if uploaded_files:
        for file in uploaded_files:
            if allowed_file(file.name) and desired_file(file.name):
                df = pd.read_excel(file)
                # st.write(f"Successfully processed file: {file.name}")
                processed_files.append((file.name, df))
            else:
                st.write(
                    f"Invalid file: {file.name}. Please upload an Excel file with one of the desired filenames: {DESIRED_FILENAMES}")

    if processed_files:
        st.write("Uploaded files:")
        for index, (file_name, df) in enumerate(processed_files):
            st.write(file_name)
            # st.dataframe(df)
        if st.button(f"Next "):
            next_function()
    else:
        st.write("No files were uploaded.")


cleaned_productivity_data = None
BranchSummury = None
TeamSummary = None
dataset = None


def do_a_vlookup_and_insertion():
    if processed_files:
        calls_file = None
        team_leader_file = None

        for file_name, df in processed_files:
            if file_name == 'DAILY ACM PERFORMANCE.xlsx':
                calls_file = df
            elif file_name == 'TEAM LEADER ACMS.xlsx':
                team_leader_file = df
            elif file_name == 'Debtors spoken to.xlsx':
                debtors_spoken_to_file = df
            elif file_name == 'PTPs Created by ACMs.xlsx':
                promise_to_pay = df
            elif file_name == 'PTPS CREATED WITH THEIR DUE DATES.xlsx':
                ptp_with_due_dates = df
        # st.dataframe(team_leader_file)
        # st.dataframe(debtors_spoken_to_file)
        if calls_file is not None and team_leader_file is not None and debtors_spoken_to_file is not None and promise_to_pay is not None and ptp_with_due_dates is not None:
            # vlookup process
            merged_df = pd.merge(calls_file, team_leader_file[['Account  Managers', 'TEAM LEADER']], left_on='acmname',
                                 right_on='Account  Managers', how='left')

            merged_df.rename(columns={'TEAM LEADER': 'TEAM LEADER'}, inplace=True)

            # insert the team leader column next to acmname
            columns = merged_df.columns.tolist()
            acmname_index = columns.index('acmname')
            team_leader_index = columns.index('TEAM LEADER')
            columns.insert(acmname_index + 1, columns.pop(team_leader_index))
            merged_df = merged_df[columns]

            # Replace values in the "shiftId" column
            merged_df['shiftid'] = merged_df['shiftid'].replace({1: 'DAY', 2: 'EVENING', 3: ''})

            # remove the additional column
            merged_df.drop('Account  Managers', axis=1, inplace=True)

            # an array to match my two files
            column_mapping = {
                'acmname': 'ACM',
                'TEAM LEADER': 'TEAM LEADER',
                'name': 'Branch',
                'remoteuser': 'Remote user',
                'shiftid': 'Shif',
                'count': 'Phone calls Made'
            }

            # rename the column
            merged_df.rename(columns=column_mapping, inplace=True)
            # st.dataframe(merged_df)

            debtors_dt = pd.merge(merged_df, debtors_spoken_to_file[['acmname', 'count']], left_on='ACM',
                                  right_on='acmname', how='left')
            debtors_dt.rename(columns={'count': 'Debtors spoken to'}, inplace=True)
            # remove the additional column
            debtors_dt.drop('acmname', axis=1, inplace=True)

            # calculate the spoke rate
            debtors_dt['Spoke rate'] = debtors_dt['Debtors spoken to'] / debtors_dt['Phone calls Made']

            # VLOOKUP promise to pay
            promise_dt = pd.merge(debtors_dt, promise_to_pay[['acmname', 'count']], left_on='ACM', right_on='acmname',
                                  how='left')
            promise_dt.rename(columns={'count': 'Promise to Pay'}, inplace=True)
            promise_dt.drop('acmname', axis=1, inplace=True)

            # calculate the ptp conversion rate
            promise_dt['Ptp conversion Rate'] = promise_dt['Promise to Pay'] / promise_dt['Debtors spoken to']

            ptp_with_due_dates['ptpamount'] = ptp_with_due_dates['ptpamount'].str.replace('"', '')

            ptp_with_due_dates = filter_dataset_by_current_date(ptp_with_due_dates)

            # VLOOKUP promise to pay
            productivity_dt = pd.merge(promise_dt, ptp_with_due_dates[['ptpamount', 'acmname']], left_on='ACM',
                                       right_on='acmname', how='left')
            productivity_dt.rename(columns={'ptpamount': 'Ptp Amount'}, inplace=True)
            productivity_dt.drop('acmname', axis=1, inplace=True)

            cleaned_productivity_data = remove_duplicates_and_sum_by_acm(productivity_dt)

            cleaned_productivity_data = filter_remove_zero_calls(cleaned_productivity_data)

            active_acms = np.count_nonzero(cleaned_productivity_data['ACM'].notnull())

            files_worked_on = calculate_column_sum(cleaned_productivity_data, 'Phone calls Made')

            total_ptps_created = calculate_column_sum(cleaned_productivity_data, 'Promise to Pay')

            total_ptps_amount = calculate_column_sum(cleaned_productivity_data, 'Ptp Amount')

            total_debtors_spoken_to = calculate_column_sum(cleaned_productivity_data, 'Debtors spoken to')

            dataset = {
                'SNAPSHOT': ['active_acms', 'files_worked_on', 'total_ptps_created', 'total_ptps_amount',
                             'total_debtors_spoken_to'],
                'Data': [active_acms, files_worked_on, total_ptps_created, total_ptps_amount, total_debtors_spoken_to]
            }

            BranchSummury = branch_summary(cleaned_productivity_data)

            TeamSummary = team_summary(cleaned_productivity_data, team_leader_file)

            st.dataframe(cleaned_productivity_data)
            st.write(BranchSummury)
            st.write(TeamSummary)
            st.dataframe(dataset)

            now = datetime.datetime.now()
            formatted_date = now.strftime("%dth %B")
            formatted_time = now.strftime("%I%p").lstrip("0")
            output_string = f"ACM PRODUCTIVITY {formatted_date} AS AT {formatted_time}"

            st.write(output_string)

            desktop_path = os.path.expanduser("~/Productivity")
            file_path = os.path.join(desktop_path, f"{output_string}.xlsx")

            startcol_cleaned = 0
            startcol_branch = cleaned_productivity_data.shape[1] + 1
            startcol_team = startcol_branch + BranchSummury.shape[1] + 1
            startcol_dataset = startcol_team + TeamSummary.shape[1] + 1

            # Save the Excel file
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                # Write the data frames to the Excel file
                cleaned_productivity_data.to_excel(writer, sheet_name='Productivity', index=False,
                                                   startcol=startcol_cleaned)
                BranchSummury.to_excel(writer, sheet_name='Productivity', index=False, startcol=startcol_branch)
                TeamSummary.to_excel(writer, sheet_name='Productivity', index=False, startcol=startcol_team)
                pd.DataFrame(dataset).to_excel(writer, sheet_name='Productivity', index=False,
                                               startcol=startcol_dataset)

            st.write(f"File saved successfully at: {file_path}")



        else:
            st.write("The 'CALLS.xlsx' or 'TEAM LEADER ACMS.xlsx' file was not uploaded.")
    else:
        st.write("No files were uploaded.")


def next_function():
    do_a_vlookup_and_insertion()


# Run the Streamlit app
if __name__ == "__main__":
    upload_excel_files()
