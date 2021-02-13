import csv
from datetime import datetime
import calendar
import xlsxwriter
import argparse
from decimal import Decimal
import copy

EURO_FORMAT = {'num_format': '€#,##0.00'}
DATE_FORMAT = {'num_format': 'dd-mm-yyyy'}
INPUT_DATE_FORMAT = '%Y%m%d'

def main():
    args = parse_arguments()

    retailers = args.retailers
    if(len(retailers) == 0):
        retailers = [retailer[0] for retailer in read_file(args.retailers_file)]
        print(retailers)
    file_path = args.input_file_path

    transactions = read_file(file_path)
    # reverse, so we start at january
    transactions.reverse()

    transactions_filtered = filter_transactions(retailers, transactions)
    structured_data = convert_to_structure(transactions_filtered)
    
    workbook = xlsxwriter.Workbook('budget_analysis.xlsx')

    transform_to_workbook_by_date(structured_data, workbook, "Retailer Expenditure by date")

    accumulative_by_retailer = calculate_retailer_accumulative(transactions_filtered)
    transform_to_workbook( accumulative_by_retailer, workbook, "Retailer Accumulative")

    monthly_cost_by_retailer = calculate_retailer_cost_per_month(transactions_filtered)
    month_sheet = transform_to_workbook( monthly_cost_by_retailer, workbook,  "Retailer cost by month")
    add_chart(workbook, month_sheet, 1, len(monthly_cost_by_retailer.keys()), "Retailer cost by month")

    workbook.close()


def calculate_retailer_accumulative(transactions):
    accumulative_cost = {}

    for transaction in transactions:
        retailer = transaction[1]
        value = convert_to_decimal(transaction[6])
        if(retailer in accumulative_cost):
            accumulative_cost[retailer] += value
        else:
            accumulative_cost[retailer] = value
        
    return accumulative_cost

def calculate_retailer_cost_per_month(transactions):
    accumulated_month_view = {}

    for transaction in transactions:
        time = datetime.strptime(transaction[0], INPUT_DATE_FORMAT)
        month_name = calendar.month_name[time.month]
        value = convert_to_decimal(transaction[6])
        if month_name in accumulated_month_view:
            accumulated_month_view[month_name] += value
        else:
            accumulated_month_view[month_name] = value

    return accumulated_month_view

def convert_to_structure(transactions):
    structured_data = { 0: {}, 1: [] }
    date_counts = {}

    offset=0
    for transaction in transactions:
        date = transaction[0]
        value = convert_to_decimal(transaction[6])
        retailer = transaction[1]

        if date in date_counts:
            date_counts[date] += 1
        else:
            date_counts[date] = offset
        
        structured_data[1].insert(offset, { retailer: value })

        offset+=1

    rowOffset=0
    for date_entry, date_count in date_counts.items():
        entry = { "bounds": [rowOffset, date_count] }
        structured_data[0][date_entry] = entry
        rowOffset = date_count

    return structured_data

def convert_to_decimal(str_value):
    culture_version = str_value.replace(',','.')
    decimal_version = Decimal(culture_version)

    return decimal_version

def filter_transactions(retailers, transactions):
    filtered_transactions=[]
    
    for transaction in transactions:
        name = transaction[1]
        for retailer in retailers:
            if name.lower().find(retailer.lower()) > -1:
                # easier lookup, make option?
                transaction[1] = retailer
                filtered_transactions.append(transaction)

    return filtered_transactions

def read_file(file_path):
    records = []

    with open(file_path, newline='') as csvfile:
        linereader = csv.reader(csvfile, delimiter=';')
        for row_columns in linereader:
            records.append(row_columns)

    return records

def transform_to_workbook_by_date(structured_data, workbook, sheetname):

    worksheet = workbook.add_worksheet(sheetname)
    date_format = workbook.add_format(DATE_FORMAT)
    currency_format = workbook.add_format(EURO_FORMAT)

    max_characters_0 = get_column_width_by_max_chars(structured_data[0].keys())
    #todo: fix below column widths
    max_characters_1 = get_column_width_by_max_chars(structured_data[1])
    max_characters_2 = get_column_width_by_max_chars(structured_data[1])

    worksheet.set_column(0, 0, max_characters_0)
    worksheet.set_column(1, 1, max_characters_1)
    worksheet.set_column(2, 2, max_characters_2)

    rowIndex = 0
    for date in structured_data[0].keys():
        lower_bound = structured_data[0][date]["bounds"][0]
        upper_bound = structured_data[0][date]["bounds"][1]

        date_time = datetime.strptime(date, INPUT_DATE_FORMAT)
        worksheet.write_datetime(lower_bound, 0, date_time, date_format)

        rowSpan = upper_bound - lower_bound
        if rowSpan > 0:
            for spanIndex in range(lower_bound, upper_bound):
                retailer = list(structured_data[1][spanIndex].keys())[0]
                value = list(structured_data[1][spanIndex].values())[0]
                worksheet.write(spanIndex, 1, retailer)
                worksheet.write_number(spanIndex, 2, value, currency_format)
        
        rowIndex += 1

def transform_to_workbook(view, workbook, sheetname):

    worksheet = workbook.add_worksheet(sheetname)
    currency_format= workbook.add_format(EURO_FORMAT)

    max_characters_key = get_column_width_by_max_chars(view.keys())
    max_characters_value = get_column_width_by_max_chars(view.values())

    worksheet.set_column(0, 0, max_characters_key)
    worksheet.set_column(1, 1, max_characters_value)

    rowIndex = 0
    for key, value in view.items():
        worksheet.write(rowIndex, 0, key) 

        worksheet.write_number(rowIndex, 1, value, currency_format)

        rowIndex += 1
    
    return worksheet

def get_column_width_by_max_chars(collection):
    def get_max(item):
        if type(item) is type(str):
            return len(item)
        return len(str(item))

    max_item = max(collection, key=get_max)
    if type(max_item) is type(str):
        return len(max_item) + 1
    
    return len(str(max_item)) + 1

def add_chart(workbook, worksheet, start, end, sheet_name):
    chart = workbook.add_chart({'type': 'line'})
    sheet_name_quoted = f'\'{sheet_name}\''

    chart.add_series({'values': f'={sheet_name_quoted}!$B${start}:$B${end}'})

    worksheet.insert_chart('C1', chart)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='Based on input csv containing transactions, generate structured excel to allow detailed analysis and budgeting.')

    parser.add_argument("input_file_path", metavar='str', help='Absolute path to input file, csv extension')
    parser.add_argument("--retailers-file", dest="retailers_file", metavar='str', help='Csv file with retailers to extract transactions for')
    parser.add_argument('--retailers', nargs='*', help="list of retailers to filter transactions on", default=[])

    args = parser.parse_args()

    return args

if __name__ == '__main__':
    main()