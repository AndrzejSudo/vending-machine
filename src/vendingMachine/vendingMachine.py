"""Python vending machine"""

import time, sys

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
    snacks_pc, snacks_qt = getSnacks()[0], getSnacks()[1]
    snacks = snacks_pc.keys()
    snacks_list = []

    for index, item in enumerate(snacks):
        snacks_list.append(item)
        print(index,':', item, end=' | ')
    print('')
    
    WALLET = 20
        
    while True:
        
        if WALLET == 0:
            print('You are out of money')
            sys.exit()
        
        print('What do you want to buy?')
        print(f'You have {WALLET} PLN left')
        echo = input('> ')

        if echo.isdecimal() and 0 <= int(echo) <= len(snacks)-1:
            SNACK_ID = int(echo)
            if snacks_qt[snacks_list[SNACK_ID]] == 0:
                print(f'Machine is out of {snacks_list[SNACK_ID]}.\nPick something else.')
                continue
            break
        else:
            print('Invalid input, try again')
    
    snack_name = snacks_list[SNACK_ID] 
    snack_pc = round(float(snacks_pc[snacks_list[SNACK_ID]]), 2)

    order = buySnacks(snack_name, snack_pc, snacks_qt, WALLET)
    print(order, type(order))
    

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

def buySnacks(snack_name, snack_pc, snacks_qt, wallet):

    print(f'You\'ve chosen {snack_name}, which costs {snack_pc} PLN')
    print('Please insert proper amount of coins or (C)ancel your purchase')
    
    pay = 0

    while pay != snack_pc:
        if pay < snack_pc:
            while True:
                print(f'{snack_pc-pay} PLN left to pay')
                amount = input('> ')
                if amount.upper().startswith('C'):
                    return None

                try:
                    amount = round(float(amount), 2)
                except ValueError:
                    print('Invalid amount, try again')
                    continue

                if amount in COINS:
                    pay += amount
                    wallet -= pay
                    break
                else:
                    print('Unrecognized coin. Insert Polish zloty nominals')
        else:
            print(f'You\'ve inserted {pay-snack_pc} PLN too much.\nReturning change.')
            pay -= pay-snack_pc
            
            wallet += pay-snack_pc
    
    snacks_qt[snack_name] -= 1
    return wallet

if __name__ == '__main__':
    main()