import csv
from datetime import datetime
import calendar
import xlsxwriter
import argparse
from decimal import Decimal

EURO_FORMAT = {'num_format': '€#,##0.00'}
INPUT_DATE_FORMAT = '%Y%m%d'

def main():
    args = parse_arguments()

    retailers = args.retailers
    file_path = args.input_file_path

    transactions = extract_records(file_path)

    transactions_filtered = filter_transactions(retailers, transactions)
    grouped_by_date = group_by_date(transactions_filtered)
    
    workbook = xlsxwriter.Workbook('budget_analysis.xlsx')
    currency_format = workbook.add_format(EURO_FORMAT)

    worksheet1 = get_worksheet(workbook, "Retailer Expenditure by date")
    transform_to_workbook_by_date(grouped_by_date, worksheet1,currency_format)

    worksheet2 = get_worksheet(workbook, "Retailer Accumulative")
    accumulative_by_retailer = calculate_retailer_accumulative(transactions_filtered)
    transform_to_workbook(worksheet2, accumulative_by_retailer,currency_format)

    worksheet3 = get_worksheet(workbook, "Retailer cost by month")
    monthly_cost_by_retailer = calculate_retailer_cost_per_month(transactions_filtered)
    transform_to_workbook(worksheet3, monthly_cost_by_retailer,currency_format)

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
        spamreader = csv.reader(csvfile, delimiter=';')
        for row_columns in spamreader:
            records.append(row_columns)

    return records

def get_worksheet(workbook, name):
    return workbook.add_worksheet(name)

def transform_to_workbook_by_date(grouped_by_date_transactions, worksheet, currency_format):
    rowIndex = 1
    for record, accumulated_transactions in grouped_by_date_transactions.items():
        date = record
        time = datetime.strptime(date, INPUT_DATE_FORMAT)
        strformat = time.strftime("%d-%m-%Y")
        worksheet.write('A' + str(rowIndex), strformat) 
        rowSpan = rowIndex
        for retailer, value in accumulated_transactions.items():
            worksheet.write('B' + str(rowSpan), retailer)
            worksheet.write('C' + str(rowSpan), value, currency_format)
            rowSpan += 1
        rowIndex += len(accumulated_transactions)

def transform_to_workbook(worksheet, view, currency_format):
    rowIndex = 1
    for key, value in view.items():
        worksheet.write('A' + str(rowIndex), key) 
        worksheet.write('B' + str(rowIndex), value, currency_format)
        rowIndex += 1

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