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

    stores = args.stores
    file_path = args.input_file_path

    transactions = extract_records(file_path)

    transactions_filtered = filter_transactions(stores, transactions)
    grouped_by_date = group_by_date(transactions_filtered)
    
    workbook = xlsxwriter.Workbook('budget_analysis.xlsx')
    currency_format = workbook.add_format(EURO_FORMAT)

    worksheet1 = get_worksheet(workbook, "ExpenditureByDate")
    transform_to_workbook_by_date(grouped_by_date, worksheet1,currency_format)

    worksheet2 = get_worksheet(workbook, "AccumulusByStore")
    accumulus_store = calculate_accumulated_cost_by_store(transactions_filtered)
    print(accumulus_store)
    transform_to_workbook(worksheet2, accumulus_store,currency_format)

    worksheet3 = get_worksheet(workbook, "StoreCostPerMonth")
    storecost_month =calculated_store_cost_per_month(transactions_filtered)
    print(storecost_month)
    transform_to_workbook(worksheet3, storecost_month,currency_format)

    workbook.close()


def calculate_accumulated_cost_by_store(transactions):
    store_accumulated_cost = {}

    for transaction in transactions:
        store = transaction[1]
        value = convert_to_decimal(transaction[6])
        if(store in store_accumulated_cost):
            store_accumulated_cost[store] += value
        else:
            store_accumulated_cost[store] = value
        
    return store_accumulated_cost

def calculated_store_cost_per_month(transactions):
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
    store_transaction_by_date = {}

    for transaction in transactions:
        date = transaction[0]
        value = convert_to_decimal(transaction[6])
        store = transaction[1]

        if date in store_transaction_by_date:
            store_transaction_by_date[date][store] = value
        else:
            store_transaction_by_date[date] = { store: value}

    return store_transaction_by_date

def convert_to_decimal(str_value):
    culture_version = str_value.replace(',','.')
    decimal_version = Decimal(culture_version)

    return decimal_version

def filter_transactions(stores, transactions):
    filtered_transactions=[]
    
    for transaction in transactions:
        name = transaction[1]
        for store in stores:
            if name.lower().find(store.lower()) > -1:
                # easier lookup, make option?
                transaction[1] = store
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
        for store, value in accumulated_transactions.items():
            worksheet.write('B' + str(rowSpan), store)
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
    parser.add_argument('--stores', nargs='*', help="list of stores to filter transactions on")

    args = parser.parse_args()

    print(args)

    return args

if __name__ == '__main__':
    main()