import pandas as pd

def convert_to_datetime(date_str):
    # Convert date string (e.g., '09/23/2023') to datetime
    return pd.to_datetime(date_str, format='%m/%d/%Y', errors='coerce')

def convert_to_timedelta(time_str):
    # Convert time string or float (e.g., '3:45' or 3.75) to timedelta
    if pd.isna(time_str):
        return pd.Timedelta(0)
    if isinstance(time_str, float):
        hours, minutes = divmod(int(time_str * 60), 60)
    else:
        hours, minutes = map(int, str(time_str).split(':'))
    return pd.Timedelta(hours=hours, minutes=minutes)

def analyze_employee_data(file_path):
    # Read Excel file into a DataFrame
    df = pd.read_excel(file_path)

    # Sort the DataFrame by employee ID and date
    df.sort_values(by=['File Number', 'Time'], inplace=True)

    # Initialize variables to store results
    consecutive_days_7 = []
    short_shifts = []
    long_shifts = []

    # Iterate through each employee
    for _, employee_data in df.groupby('File Number'):
        consecutive_count = 1
        prev_shift_end = None

        # Iterate through each row for the employee
        for _, row in employee_data.iterrows():
            # Convert Time and Time Out columns to datetime objects
            row['Time'] = convert_to_datetime(row['Time'])
            row['Time Out'] = convert_to_datetime(row['Time Out'])

            # Convert Pay Cycle Start Date and Pay Cycle End Date to datetime objects
            row['Pay Cycle Start Date'] = convert_to_datetime(row['Pay Cycle Start Date'])
            row['Pay Cycle End Date'] = convert_to_datetime(row['Pay Cycle End Date'])

            # Convert Timecard Hours (as Time) to timedelta
            timecard_hours = convert_to_timedelta(row['Timecard Hours (as Time)'])

            # Check consecutive days
            if prev_shift_end and row['Pay Cycle Start Date'] == prev_shift_end + pd.DateOffset(days=1):
                consecutive_count += 1
            else:
                consecutive_count = 1

            # Check short shifts
            if prev_shift_end and not pd.isna(row['Time']) and (row['Time'] - prev_shift_end).seconds < 10 * 3600 and (row['Time'] - prev_shift_end).seconds > 3600:
                short_shifts.append((row['File Number'], row['Employee Name'], row['Position ID']))

            # Check long shifts
            if not pd.isna(row['Time']) and (row['Time Out'] - row['Time']).seconds > 14 * 3600:
                long_shifts.append((row['File Number'], row['Employee Name'], row['Position ID']))

            # Update prev_shift_end
            prev_shift_end = row['Time Out'] + timecard_hours

            # Store consecutive 7 days
            if consecutive_count == 7:
                consecutive_days_7.append((row['File Number'], row['Employee Name'], row['Position ID']))

    # Print results to console
    print("Employees who worked for 7 consecutive days:")
    for emp in consecutive_days_7:
        print(emp)

    print("\nEmployees with less than 10 hours between shifts but greater than 1 hour:")
    for emp in short_shifts:
        print(emp)

    print("\nEmployees who worked for more than 14 hours in a single shift:")
    for emp in long_shifts:
        print(emp)

    # Write console output to a file
    with open('output.txt', 'w') as output_file:
        output_file.write("Employees who worked for 7 consecutive days:\n")
        output_file.write("\n".join(str(emp) for emp in consecutive_days_7))

        output_file.write("\n\nEmployees with less than 10 hours between shifts but greater than 1 hour:\n")
        output_file.write("\n".join(str(emp) for emp in short_shifts))

        output_file.write("\n\nEmployees who worked for more than 14 hours in a single shift:\n")
        output_file.write("\n".join(str(emp) for emp in long_shifts))

if __name__ == "__main__":
    # Replace 'your_file_path.xlsx' with the actual path to your Excel file
    analyze_employee_data('test.xlsx')
