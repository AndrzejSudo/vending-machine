import sys, openpyxl, os

print('openpyxl imported successfully')


def getSnacks(): #extracting vending content to arrays

    try:
        snacks_wb = openpyxl.load_workbook('snacks.xlsx')
    except FileNotFoundError:
        print('Snacks spreadsheet not found, ensure its in the same directory as program')
        sys.exit()
    
    snacks_sh = snacks_wb['Sheet1']
    snacks = {}
    quantity = {}

    for snack_name, snack_price in zip(snacks_sh['A'], snacks_sh['B']):
        snacks[snack_name.value] = snack_price.value
    
    for snack_name, snack_quantity in zip(snacks_sh['A'], snacks_sh['C']):
        quantity[snack_name.value] = snack_quantity.value
    
    return snacks, quantity

print(getSnacks()[1])



