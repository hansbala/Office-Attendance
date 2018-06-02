import openpyxl

# Global Variables
NUM_EMPLOYEES_KDS       = 0
NUM_EMPLOYEES_HDC       = 0
NUM_EMPLOYEES_OFF       = 0
MONTH_TOTAL_DAYS        = 0
NORMAL_WORKING_DAYS     = 0
DATES_SUNDAYS_LIST      = []
DATES_HOLIDAYS_LIST     = []
ATTENDANCE_BOOK         = openpyxl.load_workbook('Attendance.xlsx')
ATTENDANCE_SHEET_OFF    = ATTENDANCE_BOOK['Office']
ATTENDANCE_SHEET_KDS    = ATTENDANCE_BOOK['KDS']
ATTENDANCE_SHEET_HDC    = ATTENDANCE_BOOK['HDC']

def acceptValues():
    # Function that accepts all needed values to run the script
    global NUM_EMPLOYEES_OFF, NUM_EMPLOYEES_KDS, NUM_EMPLOYEES_HDC, MONTH_TOTAL_DAYS, NORMAL_WORKING_DAYS, DATES_SUNDAYS_LIST, DATES_HOLIDAYS_LIST
    NUM_EMPLOYEES_OFF   = int(input('Enter number of employees in OFF: '))    
    NUM_EMPLOYEES_KDS   = int(input('Enter number of employees in KDS: '))
    NUM_EMPLOYEES_HDC   = int(input('Enter number of employees in HDC: '))
    MONTH_TOTAL_DAYS    = int(input('Enter number of days of the month: '))
    DATES_SUNDAYS_LIST  = list(map(int, input('Enter Sunday dates (with spaces in between): ').split(' ')))
    resp                = input('Any Holidays this month? ')
    if resp == "yes":
        DATES_HOLIDAYS_LIST = list(map(int, input('Enter Holiday dates (with spaces in between): ').split(' ')))
    NORMAL_WORKING_DAYS = MONTH_TOTAL_DAYS - len(set(DATES_SUNDAYS_LIST) | set(DATES_HOLIDAYS_LIST))

def getExcelColumnName(col_num):
    # Works upto ZZ (not tested after that simply cause I don't require it)
    # Thankfully we don't have more than 31 days in a month (I'm happy AF. Get it? AF ;) )
    if col_num > 26:
        return str(chr(int(col_num / 26) + ord('A') - 1)) + str(chr(int(col_num % 26) + ord('A') - 1))
    else:
        return str(chr(col_num + ord('A') - 1))

def computeDaysWorked(NUM_EMPLOYEES, ATTENDANCE_SHEET):
    global MONTH_TOTAL_DAYS
    # Computation of days absent / present / extra shift for KDS
    # Salary Colums span from   [B-* to AF-*] (for a 31 day month)
    # Salary Rows span from     [*-2 to *-xx] (for (xx - 1) number of employees)
    for row_num in range(2, NUM_EMPLOYEES + 2):
        # Using floats to account for half-days, 1.5 shifts, and so on
        num_sundays_worked      = 0.0
        num_holidays_worked     = 0.0
        num_extra_shift_worked  = 0.0
        num_absent              = 0.0
        for col_num in range(2, MONTH_TOTAL_DAYS + 2):
            col_name = getExcelColumnName(col_num)
            curr_cell_value = float(ATTENDANCE_SHEET[col_name + str(row_num)].value)
            # If he is absent but it is a Sunday, or a Holiday, don't do anything
            if curr_cell_value == 0 and (((col_num -1) in DATES_SUNDAYS_LIST) or ((col_num - 1) in DATES_HOLIDAYS_LIST)):
                continue
            # It is not a Sunday or a Holiday, so mark him absent
            elif curr_cell_value == 0:
                num_absent += 1.0
                continue
            # He has come for duty (for some duration > 0)
            # He has come for duty on a Holiday / (Holiday + Sunday)
            elif (col_num - 1) in DATES_HOLIDAYS_LIST:
                num_holidays_worked += curr_cell_value
                continue
            # He has come for duty on a Sunday
            elif (col_num - 1) in DATES_SUNDAYS_LIST:
                num_sundays_worked += curr_cell_value
                continue
            # Duty on a regular day (also accounted for extra shift)
            if curr_cell_value > 1.0:
                # He has worked some number of extra shifts
                num_extra_shift_worked += (curr_cell_value - 1.0)
            else:
                # Something like half-day
                num_absent += (1 - curr_cell_value)
        # -- end inner for --#
        # After computation of duration worked write it to the necessary cells
        ATTENDANCE_SHEET['AI' + str(row_num)] = NORMAL_WORKING_DAYS         # Normal Working Days
        ATTENDANCE_SHEET['AJ' + str(row_num)] = num_sundays_worked          # Sundays worked
        ATTENDANCE_SHEET['AK' + str(row_num)] = num_holidays_worked         # Holidays worked
        ATTENDANCE_SHEET['AL' + str(row_num)] = num_extra_shift_worked      # Extra Shifts worked
        ATTENDANCE_SHEET['AM' + str(row_num)] = num_absent                  # Days absent
    # -- end outer for -- #
    # Nothing more to do for each worker

def main():
    global NUM_EMPLOYEES_OFF, NUM_EMPLOYEES_KDS, NUM_EMPLOYEES_HDC, NORMAL_WORKING_DAYS, DATES_SUNDAYS_LIST, DATES_HOLIDAYS_LIST
    global ATTENDANCE_BOOK, ATTENDANCE_SHEET_KDS, ATTENDANCE_SHEET_HDC
    acceptValues()    
    computeDaysWorked(NUM_EMPLOYEES_OFF, ATTENDANCE_SHEET_OFF)      # Compute days worked for OFF (Office Staff)
    computeDaysWorked(NUM_EMPLOYEES_KDS, ATTENDANCE_SHEET_KDS)      # Compute days worked for KDS (Kolkata Dock)
    computeDaysWorked(NUM_EMPLOYEES_HDC, ATTENDANCE_SHEET_HDC)      # Compute days worked for HDC (Haldia Dock Complex)
    ATTENDANCE_BOOK.save('Attendance_computed.xlsx')                # Save the workbook once everything is done with a different file name (for safety)
    print('\n\nAll done! You saved a ton of time with Python! Have a great day!\n\n')

# ??Why do I have so much documentation for a hacky script?? ;)
if __name__ == '__main__':
    main()
    input('Press [Enter] to exit')
