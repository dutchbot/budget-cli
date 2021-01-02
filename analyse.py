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
    file_path = args.input_file_path

    transactions = extract_records(file_path)
    # reverse, so we start at january
    transactions.reverse()

    transactions_filtered = filter_transactions(retailers, transactions)
    grouped_by_date = group_by_date(transactions_filtered)
    
    workbook = xlsxwriter.Workbook('budget_analysis.xlsx')

    transform_to_workbook_by_date(grouped_by_date, workbook, "Retailer Expenditure by date")

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

def group_by_date(transactions):
    expenditure_by_date_by_retailer = {}

    for transaction in transactions:
        date = transaction[0]
        value = convert_to_decimal(transaction[6])
        retailer = transaction[1]

        if date in expenditure_by_date_by_retailer:
            expenditure_by_date_by_retailer[date][retailer] = value
        else:
            expenditure_by_date_by_retailer[date] = { retailer: value}

    return expenditure_by_date_by_retailer

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

def extract_records(file_path):
    records = []

    with open(file_path, newline='') as csvfile:
        linereader = csv.reader(csvfile, delimiter=';')
        for row_columns in linereader:
            records.append(row_columns)

    return records


def transform_to_workbook_by_date(grouped_by_date_transactions, workbook, sheetname):

    worksheet = workbook.add_worksheet(sheetname)
    date_format = workbook.add_format(DATE_FORMAT)
    currency_format= workbook.add_format(EURO_FORMAT)

    rowIndex = 0
    for date, accumulated_transactions in grouped_by_date_transactions.items():

        date_time = datetime.strptime(date, INPUT_DATE_FORMAT)
        worksheet.write_datetime(rowIndex, 0, date_time, date_format)

        rowSpan = rowIndex
        for retailer, value in accumulated_transactions.items():

            worksheet.write(rowSpan, 1, retailer)
            worksheet.write_number(rowSpan, 2, value, currency_format)

            rowSpan += 1
        rowIndex += len(accumulated_transactions)

def transform_to_workbook(view, workbook, sheetname):

    worksheet = workbook.add_worksheet(sheetname)
    currency_format= workbook.add_format(EURO_FORMAT)

    rowIndex = 0
    for key, value in view.items():
        worksheet.write(rowIndex, 0, key) 

        worksheet.write_number(rowIndex, 1, value, currency_format)

        rowIndex += 1
    
    return worksheet

def add_chart(workbook, worksheet, start, end, sheet_name):
    chart = workbook.add_chart({'type': 'line'})
    sheet_name_quoted = f'\'{sheet_name}\''

    chart.add_series({'values': f'={sheet_name_quoted}!$B${start}:$B${end}'})

    worksheet.insert_chart('C1', chart)


def parse_arguments():
    parser = argparse.ArgumentParser(
        description='Based on input csv containing transactions, generate structured excel to allow detailed analysis and budgeting.')

    parser.add_argument("input_file_path", metavar='str', help='Absolute path to input file, csv extension')
    parser.add_argument('--retailers', nargs='*', help="list of retailers to filter transactions on")

    args = parser.parse_args()

    print(args)

    return args

if __name__ == '__main__':
    main()