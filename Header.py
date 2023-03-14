import pandas as pd
from tabulate import tabulate



class Budget:
    filename = ''
    writer = ''

    def __init__(self, filename='products.xlsx'):
        self.filename = filename

    def start(self):
        exel_file = pd.ExcelFile(self.filename, engine='openpyxl')
        table_main = pd.read_excel(exel_file, sheet_name='main')
        for index in range(len(table_main.index) - 1):#перебор и обработка каждой строки в главной таблице
            table_main.at[index, 'amount_sell_all'] = table_main.at[index, 'amount_sell_good']
            if table_main.at[index, 'myself']:
                table_main.at[index, 'all_purchase_price'] = table_main.at[index, 'amount_buy'] * \
                                                             table_main.at[index, 'purchase_price'] + \
                                                             table_main.at[index, 'advertising'] +\
                                                             table_main.at[index, 'delivery']
                table_main.at[index, 'profit'] = -table_main.at[index, 'all_purchase_price'] +\
                                              table_main.at[index, 'amount_sell_good'] * table_main.at[index, 'price']
            else:
                table_main.at[index, 'all_purchase_price'] = (table_main.at[index, 'amount_buy'] *
                                                              table_main.at[index, 'purchase_price']) / 2 + \
                                                             table_main.at[index, 'advertising'] + \
                                                             table_main.at[index, 'delivery']
                table_main.at[index, 'profit'] = -table_main.at[index, 'all_purchase_price'] + \
                                              (table_main.at[index, 'amount_sell_good'] * table_main.at[index, 'price']) / 2

        table_operations = pd.read_excel(exel_file, sheet_name='operations')
        for index in range(len(table_operations.index)):#перебор до таблицы и обновл основной
            ID = table_operations.at[index, 'ID'] # id данного объекта
            main_product = table_main.query('ID==@ID') # получение объекта в главной табл
            table_operations.at[index, 'name'] = main_product['name'].values[0]
            if table_main.at[main_product.index[0], 'myself']:
                table_operations.at[index, 'profit_operations'] = (table_operations.at[index, 'new_price']
                                                                   - main_product['purchase_price'].values[0]) * \
                                                                  table_operations.at[index, 'amount_sell']
                table_main.at[main_product.index[0], 'profit'] += table_operations.at[index, 'amount_sell'] * \
                                                                  table_operations.at[index, 'new_price']
            else:
                table_operations.at[index, 'profit_operations'] = (table_operations.at[index, 'new_price'] -
                                                                   main_product['purchase_price'].values[0]) / 2 \
                                                                  * table_operations.at[index, 'amount_sell']
                table_main.at[main_product.index[0], 'profit'] += (table_operations.at[index, 'amount_sell'] * \
                                                                  table_operations.at[index, 'new_price']) / 2
            table_main.at[main_product.index[0], 'amount_sell_all'] += table_operations.at[index, 'amount_sell']

        all_profit = 0
        for index in range(len(table_main.index) - 1):
            all_profit += table_main.at[index, 'profit']

        _sum = table_main.query('name=="ВСЕГО"')
        table_main.at[_sum.index[0], 'amount_buy'] = all_profit
        writer = pd.ExcelWriter(self.filename, engine='xlsxwriter', mode='w')
        table_main.to_excel(writer, sheet_name='main', index=False)
        table_operations.to_excel(writer, sheet_name='operations', index=False)
        writer.save()









