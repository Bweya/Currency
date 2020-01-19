import xlsxwriter
import pandas as pd
import os
import math

def trade():

    files = os.listdir('files')

    year = '2019'

    workbook = xlsxwriter.Workbook(year+'currency.xlsx')
    worksheet = workbook.add_worksheet()

    border = workbook.add_format({'border':1,'align':'center'})
    bolds = workbook.add_format({'bold': True, 'font_size':18, 'border': 1})

    worksheet.set_column('A:F', 45)

    worksheet.merge_range('A1:F1', 'FX Trades by Currency 01 JAN - 31 DEC '+year, bolds)

    bold = workbook.add_format({'bold':True, 'border':1})
    gtotal =  workbook.add_format({'bold': True, 'align':'center', 'bg_color':'#A9A9A9', 'border': 1})
    footer = workbook.add_format({'border':1})

    worksheet.write("A2", "Currency", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("B2", "Total Amount", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("C2", "Base Currency Equiv: Market Rate (US$)", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("D2", "Base Currency Equiv: CO Rate (US$)", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("E2", "Number of FX Trades", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))
    worksheet.write("F2", "Total Value Add (US$)", workbook.add_format({'bold':True, 'align':'center', 'border':1, 'bg_color':'#A9A9A9'}))

    check_curr = []
    grand_valueAdd = 0
    grand_amount = 0
    grand_count = 0
    grand_base_market = 0
    grand_base_CO = 0
    row_record = 3

    #Build currency dictionary first

    for f in files:

        if f[:3] == '.DS' or f[:1] == 't':

            print('DS File Store')

        else:

            TradeData = pd.read_excel("files/"+f, sheet_name = 0, header = None, skiprows=3)

            for x in TradeData[6]:
                if isinstance(x, str) == True:

                    if x not in check_curr:
                        check_curr.append(x.strip())


    #==========================================================================


    #==========================================================================

    #valueAdd_sorted.sort(reverse = True)
    #v = valueAdd_sorted
    #for val in v:


    for curr in check_curr:

        count = 0
        valueadd_bp = 0
        amount = 0
        base_market = 0
        base_CO = 0


        for file in files:

            if file[:3] == '.DS' or file[:1] == 't':

                print('DS File Store')

            else:

                TradeData = pd.read_excel("files/"+file, sheet_name = 0, header = None, skiprows=3)

                #for x in TradeData[6]:
                    #if isinstance(x, str) == True:


                index = 0
                for y in TradeData[6]:
                    if isinstance(y, str) == True:

                        if curr == y.strip():
                            count += 1
                            valueadd_bp += float( (TradeData[16][index]).replace(",","") )
                            amount += float( TradeData[5][index].replace(",","") )
                            base_market += float( TradeData[14][index].replace(",","") )
                            base_CO += float( TradeData[13][index].replace(",","") )
                    index += 1

        #if val ==  valueadd_bp:

        grand_valueAdd += valueadd_bp
        grand_amount += amount
        grand_count += count
        grand_base_CO += base_CO
        grand_base_market += base_market

                #print(file[:-25],' Grand value add equals: ', "{:,.2f}".format(grand_valueAdd) )
        worksheet.write("A"+str(row_record), curr, border)
        worksheet.write("B"+str(row_record), "{:,.2f}".format(amount), border)
        worksheet.write("C"+str(row_record), "{:,.2f}".format(base_market), border)
        worksheet.write("D"+str(row_record), "{:,.2f}".format(base_CO), border)
        worksheet.write("E"+str(row_record), "{:,.2f}".format(count), border)

        worksheet.write("F"+str(row_record), "{:,.2f}".format(valueadd_bp), border)

        row_record += 1


        print('\n')

    worksheet.write("A"+str(row_record + 1), "TOTAL", gtotal)
    worksheet.write("B"+str(row_record + 1), "", gtotal)
    worksheet.write("C"+str(row_record + 1), "{:,.2f}".format(grand_base_market), gtotal)
    worksheet.write("D"+str(row_record + 1), "{:,.2f}".format(grand_base_CO), gtotal)
    worksheet.write("E"+str(row_record + 1), "{:,.2f}".format(grand_count), gtotal)
    worksheet.write("F"+str(row_record + 1), "{:,.2f}".format(grand_valueAdd), gtotal)
    worksheet.merge_range('A'+str(row_record+2)+':F'+str(row_record+2), "Compiled by: Louisa Tinga - Treasury Unit", footer)

    workbook.close()
    print('Process Completed successfully')
