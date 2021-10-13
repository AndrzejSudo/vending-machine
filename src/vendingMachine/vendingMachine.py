"""Python vending machine"""

import random, sys

try:
    import openpyxl
except ModuleNotFoundError:
    print('This program requires openpyxl module, install it with pip')
    sys.exit()

# this vending machine is going to use Polish national currency called "Polish zloty" (code: PLN)
# available coin nominals are 1gr, 2gr, 5gr, 10gr, 20gr, 50gr, 1zl, 2zl, 5zl
# this vending machine wont accept coins with denominations smaller than 10gr
COINS = [0.1, 0.2, 0.5, 1, 2, 5]

def main():
    
    print('Welcome to vending machine.')
    print('Pick snacks using their corresponding numbers.\nWe have:')
    snacks = getSnacks()[1].keys()
    for index, item in enumerate(snacks):
        print(index,':', item, end=' | ')
    print('')

    wallet = 20
        
    while True:
        
        if wallet == 0:
            print('You are out of money')
            sys.exit()
        
        print('What do you want to buy?')
        print(f'You have {WALLET}PLN left')
        echo = input('> ')

        if echo.isdecimal() and 0 <= int(echo) <= len(snacks)-1:
            snack = int(echo)
            break
        else:
            print('Invalid input, try again')
        
    order = buySnacks(snack, wallet)

def getSnacks():

    try:
        snacks_wb = openpyxl.load_workbook('snacks.xlsx')
    except FileNotFoundError:
        print('Snacks spreadsheet not found, ensure its in the same directory as program')
        sys.exit()
    
    snacks_sh = snacks_wb['Sheet1']
    snacks_price = {}
    snacks_quantity = {}

    for name, price, quantity in zip(snacks_sh['A'], snacks_sh['B'], snacks_sh['C']):
        snacks_price[name.value] = price.value
        snacks_quantity[name.value] = quantity.value
        
    return snacks_price, snacks_quantity

def buySnacks(purchase, wallet):




main()
    
