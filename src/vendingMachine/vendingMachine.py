"""Python vending machine"""

import time, sys

try:
    import openpyxl
except ModuleNotFoundError:
    print('This program requires openpyxl module, install it with pip')
    sys.exit()

"""Vending machine is using Polish national currency called "Polish zloty" (code: PLN).
Available coin nominals are 1gr, 2gr, 5gr, 10gr, 20gr, 50gr, 1zl, 2zl, 5zl.
Denominations smaller than 10gr are not accepted"""

COINS = [0.1, 0.2, 0.5, 1, 2, 5]
WALLET = 20
bought_snacks = []

def main():
    
    snacks_pc, snacks_qt = getSnacks()[0], getSnacks()[1]
    snacks = snacks_pc.keys()
    snacks_list = []

    print(' VENDING MACHINE '.center(41, '$'))
    print('Purchase snacks using their corresponding numbers:')

    for index, item in enumerate(snacks):
        snacks_list.append(item)
        print(index,':', item, end=' | ')
    print('')
    
    pickSnack(bought_snacks, snacks, snacks_pc, snacks_qt, snacks_list)

    print('What would you like to do now?')
    while True:
        print('''Available options:
    1. Make another purchase
    2. Refund your purchase
    3. Enter service mode
    4. Quit''')
        response = input('> ')

        if response.isdecimal() and int(response) in range(1, 5):
            response = int(response)
            if response == 1:
                main()
            elif response == 2:
                getRefund(bought_snacks, snacks, snacks_list, snacks_pc, snacks_qt)
            elif response == 3:
                print('Entering service mode')
                for dot in range(3):
                    time.sleep(1)
                    print('*', end=' ')
                print('')
                serviceMode()
            else:
                print('Goodbye, have a nice day')
                sys.exit()

        else:
            print('Invalid input, try again')
            continue
    
def getSnacks():

    try:
        snacks_wb = openpyxl.load_workbook('snacks.xlsx', data_only=True)
    except FileNotFoundError:
        print('Snacks spreadsheet not found, ensure its in the proper working directory')
        sys.exit()
    
    snacks_sh = snacks_wb['Sheet1']
    snacks_price = {}
    snacks_quantity = {}

    for name, price, quantity in zip(snacks_sh['A'], snacks_sh['B'], snacks_sh['C']):
        snacks_price[name.value] = price.value
        snacks_quantity[name.value] = quantity.value
    
    snacks_wb.close()

    return snacks_price, snacks_quantity

def pickSnack(bought_snacks, snacks, snacks_pc, snacks_qt, snacks_list):

    global WALLET

    while True:
        if WALLET == 0:
            print('You are out of money')
            sys.exit()
        
        print('What do you want to buy?')
        print(f'You have {round(WALLET, 2)} PLN left')
        echo = input('> ')

        if echo.isdecimal() and 0 <= int(echo) <= len(snacks)-1:
            snack_id = int(echo)
            if snacks_qt[snacks_list[snack_id]] == 0:
                print(f'We are out of {snacks_list[snack_id]}.\nPick something else.')
                continue
            break
        else:
            print('Invalid input, try again')
    
    snack_name = snacks_list[snack_id] 
    snack_pc = float(snacks_pc[snacks_list[snack_id]])

    buySnacks(bought_snacks, snack_id, snack_name, snack_pc, snacks_qt)

    return snack_name, snack_pc

def buySnacks(bought_snacks, snack_id, snack_name, snack_pc, snacks_qt):

    global WALLET

    print(f'You\'ve chosen {snack_name}, which costs {snack_pc} PLN')
    print('Please insert proper amount of coins or (C)ancel your purchase')
    
    pay = 0

    while pay != snack_pc:
        print(f'{round(snack_pc-pay, 2)} PLN left to pay')
        amount = input('> ')

        if amount.upper().startswith('C'):
            print('Canceling purchase')
            WALLET += pay
            main()

        try:
            amount = round(float(amount), 2)
        except ValueError:
            print('Invalid amount, try again')
            continue

        if amount in COINS:
            pay += amount
            WALLET -= amount
            if pay < snack_pc:
                continue
            else:
                change = round(pay-snack_pc, 2)
                print(f'You\'ve inserted {change} PLN too much.')
                print('Returning change', end=' ')
                for dot in range(3):
                    time.sleep(1)
                    print('.', end=' ')
                print('')
                pay -= change
                WALLET += change
                
        else:
            print('Unrecognized coin. Insert Polish zloty nominals')            

    
    snacks_qt[snack_name] -= 1
    bought_snacks += [snack_id]

    print('Purchase confirmed.\nPlease hold', end=' ')
    for dot in range(3):
        time.sleep(1)
        print('.', end=' ')
    print('\nHere is your snack!')
    print(f'>>>> {snack_name.upper()} <<<<')
    return WALLET

def getRefund(bought_snacks, snacks, snacks_list, snacks_pc, snacks_qt):
    
    global WALLET

    print('What snack do you want to refund?')
    for index, item in enumerate(snacks):
        print(index,':', item, end=' | ')
    print('')

    while True:
        refunded_snack = input('> ')

        if refunded_snack.isdecimal():
            if int(refunded_snack) in bought_snacks:
                refunded_snack = int(refunded_snack)
                break
            elif not int(refunded_snack) in bought_snacks and int(refunded_snack) in range(len(snacks)+1):
                print('You didn\'t bought this snack yet.\nYou have only bought:')
                for id in bought_snacks:
                    print(f'{snacks_list[id]}, ', end=' ')
                print('\nPick again')
                continue

            else:
                print('There is no such snack under this number.\nTry again.')
                continue 
        else:
            print('Invalid input, pick correct snack number')
            continue
    
    print('Refund accepted.\nPlease hold', end=' ')
    for dot in range(3):
        time.sleep(1)
        print('.', end=' ')

    print(f'\nYou\'ve returned {snacks_list[refunded_snack]}')
    snack_pc = float(snacks_pc[snacks_list[refunded_snack]])
    WALLET += snack_pc
    print(f'{snack_pc} PLN refunded.')
    snacks_qt[snacks_list[refunded_snack]] += 1
    
    if len(bought_snacks) > 0:
        print('Do you want to refund another snack? (Y)es or (N)o')
        while True:
            response = input('> ')
            if response.upper().startswith('Y'):
                getRefund(bought_snacks, snacks, snacks_list, snacks_pc, snacks_qt)
            elif response.upper().startswith('N'):
                main()
            else:
                print('Unrecognized input, try again.')
    else:
        main()

def serviceMode():
    
    while True:
        print('SERVICE MODE'.center(40, '*'), end='')
        print('''
    1. Increase your wallet ballance
    2. Change snacks prices
    3. Change snacks quantities
    4. Add new snacks
    5. Remove snacks
    6. Exit service mode''')
        print(''.center(40, '*'))

        mode = input('> ')

        if mode.isdecimal() and int(mode) in range(7):
            mode = int(mode)
            snacks_wb = openpyxl.load_workbook('snacks.xlsx')
            snacks_sh = snacks_wb['Sheet1']
            snacks_list = []
            snacks_pc = {}
            snacks_qt = {}

            n = 1
            for i in snacks_sh:
                    snacks_list.append(snacks_sh[f'A{n}'].value)
                    n+=1

            for name, price, quantity in zip(snacks_sh['A'], snacks_sh['B'], snacks_sh['C']):
                snacks_pc[name.value] = price.value
                snacks_qt[name.value] = quantity.value

            if mode == 1:
                snacks_wb.close()
                print('Enter how much PLN would you like to add to your wallet')

                while True:
                    try:
                        amount = float(input('> '))
                    except ValueError:
                        print('Invalid amount, try again')
                        continue
                    if float(amount) > 0:
                        global WALLET
                        WALLET += float(amount)
                        print(f'New wallet ballance: {WALLET}')
                        break
                    else:
                        print('Added amount, can\'t be negative')

            elif mode == 2:
                counter = 0
                print(f'Currently there are {len(snacks_pc)} snacks priced as follows:')
                for name, price in snacks_pc.items():
                    print(f'{counter}. {name}: {price} PLN ; ', end=' ')
                    counter += 1
                print('')
                    
                while True:
                    print('For what snack would you like to change the price?')
                    item_id = input('> ')
                    if item_id.isdecimal() and int(item_id) in range(len(snacks_list)+1):
                        item_id = int(item_id)
                        break
                    else:
                        print('Invalid input, use snack numbers')

                print(f'You\'ve picked {snacks_list[item_id]} which currently costs {snacks_pc[snacks_list[item_id]]} PLN')

                while True:
                    print(f'What will be new price for {snacks_list[item_id]}?')
                    try:
                        price = float(input('> '))
                        if price > 0:
                            break
                        else:
                            print('Price can\'t be negative')
                    except ValueError:
                        print('Invalid price, try again')
                        continue

                snacks_sh[f'B{item_id+1}'] = price             
                snacks_wb.save('snacks.xlsx')
                snacks_wb.close()
                print('Price changed')

            elif mode == 3:
                counter = 0
                print(f'Currently there are {len(snacks_qt)} snacks in an amounts as follows:')
                for name, quantity in snacks_qt.items():
                    print(f'{counter}. {name}: {quantity}; ', end=' ')
                    counter += 1
                print('')
                    
                while True:
                    print('For what snack would you like to change quantity?')
                    item_id = input('> ')
                    if item_id.isdecimal() and int(item_id) in range(len(snacks_list)+1):
                        item_id = int(item_id)
                        break
                    else:
                        print('Invalid input, use snack numbers')

                print(f'You\'ve picked {snacks_list[item_id]} which is currently in quantity of {snacks_qt[snacks_list[item_id]]}.')

                while True:
                    print(f'What will be new quantity for {snacks_list[item_id]}?')
                    try:
                        qty = int(input('> '))
                        if qty > 0:
                            break
                        else:
                            print('Quantity can\'t be negative')
                            continue
                    except ValueError:
                        print('Invalid quantity, try again')
                        continue

                snacks_sh[f'C{item_id+1}'] = qty             
                snacks_wb.save('snacks.xlsx')
                snacks_wb.close()
                print('Quantity changed')

            elif mode == 4:
                counter = 0
                print(f'Currently there are {len(snacks_list)} snacks:')
                for index, name in enumerate(snacks_list):
                    print(f'{index}. {name}; ', end=' ')
                print('')

                while True:
                    print('What\'s name of new snack would you like to add?')
                    new_snack = input('> ')
                    if not len(new_snack) > 3:
                        print('Snack name must be longer than 3 characters')
                        continue
                    else:
                        break
                
                new_snack_id = f'A{len(snacks_list)+1}'
                snacks_sh[new_snack_id] = new_snack
                
                while True:
                    print(f'What PLN price {new_snack} will have?')
                    try:
                        new_price = float(input('> '))
                        if new_price > 0:
                            break
                        else:
                            print('Price can\'t be negative')
                            continue
                    except ValueError:
                        print('To set price, use float numbers')
                        continue
                
                new_price_id = f'B{len(snacks_list)+1}'
                snacks_sh[new_price_id] = new_price

                while True:
                    print(f'What will be quantity of {new_snack}?')
                    try:
                        new_qty = int(input('> '))
                        if new_qty > 0:
                            break
                        else:
                            print('Quantity can\'t be negative')
                            continue
                    except ValueError:
                        print('To set quantity, use integer numbers')
                        continue
                
                new_qty_id = f'C{len(snacks_list)+1}'
                snacks_sh[new_qty_id] = new_qty

                snacks_wb.save('snacks.xlsx')
                snacks_wb.close()
                print('New entry for your snack:')
                print(f'{len(snacks_list)+1}. {snacks_sh[new_snack_id].value}: {snacks_sh[new_price_id].value} PLN, {snacks_sh[new_qty_id].value} qt.')

            elif mode == 5:
                counter = 0
                print(f'Currently there are {len(snacks_list)} snacks:')
                for index, name in enumerate(snacks_list):
                    print(f'{index}. {name}; ', end=' ')
                print('')

                while True:
                    print('Which snack do you want to remove?')
                    del_snack = input('> ')
                    if del_snack.isdecimal() and 0 < int(del_snack) <= len(snacks_list):
                        del_snack = int(del_snack)
                        break
                    else:
                        print('Invalid input, use integer snacks index numbers')
                                
                # del_snack_wb_id = del_snack + 1
                # for letter in ('A', 'B', 'C'):
                #     snacks_sh[f'{letter}{del_snack_wb_id}'] = None
                #     for index in range(del_snack_wb_id, len(snacks_list)+1):
                #         snacks_sh[f'{letter}{index}'] = snacks_sh[f'{letter}{index+1}'].value
                # basically all of this is done with delete_rows method, 
                # but at that time I didn't knew about it
                # so I'm just going to leave this, beacuse I like it

                snacks_sh.delete_rows(del_snack+1)
                snacks_wb.save('snacks.xlsx')
                snacks_wb.close()
                print(f'"{snacks_list[del_snack]}" removed')

            elif mode == 6:
                print('Quitting service mode')
                for dot in range(3):
                    time.sleep(1)
                    print('*', end=' ')
                print('')
                main()

        else:
            print('Invalid input, pick proper mode number.')
   
            
if __name__ == '__main__':
    main()
