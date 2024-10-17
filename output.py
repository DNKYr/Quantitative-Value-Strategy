from SImple_Value import final_dataframe
import xlsxwriter
import pandas as pd

#label different dataframe
simple_value_dataframe = final_dataframe

#Turn dataframe into xlsx doc
with pd.ExcelWriter('Value_Trade.xlsx') as writer:
    simple_value_dataframe.to_excel(writer, sheet_name = 'Simple Value Trade', index = False)
    string_format = writer.book.add_format({
        'border': 1
    })

    dollar_format = writer.book.add_format({
        'num_format': '$0.00',
        'border': 1
    })

    integer_format = writer.book.add_format({
        'num_format': '0',
        'border': 1
    })

    float_format = writer.book.add_format({
        'num_format': '0.000',
        'border': 1
    })

    column_format = {
        'A': ['Ticker', string_format],
        'B': ['Price', dollar_format],
        'C': ['Price-to-Earning Ratio', float_format],
        'D': ['Number of Shares to Buy', integer_format]
    }

    worksheet = writer.sheets['Recommended Trade']

    for column in column_format.keys():
        worksheet.set_column(f'{column}:{column}', 18, column_format[column][1])
        worksheet.write(f'{column}1', column_format[column][0], string_format)