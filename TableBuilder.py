import string

import xlsxwriter
from template import templateByBU





class TableItem:
    def __init__(self, items):
        self.items = items
        self.rowItems = []
        self.headings = [{'header': 'Line Number'}, {'header': 'Part Number'}, {'header': 'Description'},
                         {'header': 'Unit List Price'}, {'header': 'Qty'}, {'header': 'Disc(%)'},
                         {'header': 'Unit Net Price'}, {'header': 'Extended Net Price'}, {'header': 'WPA Disc(%)'},
                         {'header': 'WPA Net Price'}
                         ]

    def buildAssumptions(self, wb, fileName):
        formats = wb.formats
        right_justify = wb.add_format({'align': 'right'})
        percent_format = wb.add_format({'num_format' : '#0%'})
        bold_format = wb.add_format({'bold' : True,
                                     'underline': True,
                                     'align' : 'center'})
        merge_format = wb.add_format({
            'bold' : True,
            'font_size' : 16,
            'align' : 'center',
            'valign' : 'vcenter',
            'fg_color' : '#00B0F0'})

        col = dict(zip(string.ascii_uppercase, range(0, 26)))

        ws = wb.add_worksheet('Assumptions-Variables')

        rownum = 0
        ws.write(rownum, col['C'], 'WPA HW Discount Delayer Factor', right_justify)
        ws.write(rownum, col['D'], .25, percent_format)
        formula = f"='{ws.name}'!$D${rownum+1}"
        wb.define_name(f'WPA_Delayer_Factor', formula)

        rownum += 1

        ws.write(rownum, col['C'], 'FW Discount Factor', right_justify)
        ws.write(rownum, col['D'], .04, percent_format)
        formula = f"='{ws.name}'!$D${rownum+1}"
        wb.define_name(f'FW_Delayer_Factor', formula)

        rownum += 1

        agency = ''
        if '-' in fileName:
            start = fileName.index('- ') + 2
            end = fileName.index('.')
            agency = fileName[start:end]
        ws.merge_range(f'A{rownum + 1}:D{rownum + 1}', agency, merge_format)

        rownum += 1

        row = ['', '', 'SW', 'HW']
        ws.write_row(rownum, 0, row, bold_format)

        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'Historical Discount', 'Hist_Disc')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'Historical Discount SP Routing', 'Hist_SP_Disc')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'Historical Discount Meraki', 'Hist_Meraki_Disc')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'Historical Discount IoT', 'Hist_IoT_Disc')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA Discount', 'WPA_Disc', 'Hist_Disc_HW + (1 - Hist_Disc_HW) * '
                                                                             'WPA_Delayer_Factor')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA SP Discount', 'WPA_SP_Disc', 'Hist_SP_Disc_HW + (1 - '
                                                                                   'Hist_SP_Disc_HW) * '
                                                                             'WPA_Delayer_Factor')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA Meraki Discount', 'WPA_Meraki_Disc', 'Hist_Meraki_Disc_HW + (1 - '
                                                                                   'Hist_Meraki_Disc_HW) * '
                                                                                   'WPA_Delayer_Factor')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA IoT Discount', 'WPA_IoT_Disc', 'Hist_IoT_Disc_HW + (1 - '
                                                                                   'Hist_IoT_Disc_HW) * '
                                                                                   'WPA_Delayer_Factor')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA Discount FW HW', 'WPA_FW_Disc', 'Hist_Disc_HW + '
                                                                                      'FW_Delayer_Factor')
        rownum += 1

        self.writeAssumptionLine(wb, ws, rownum, 'WPA Discount Collab HW', 'WPA_Collab_Disc', 'Hist_Disc_HW')
        rownum += 1

        ws.autofit()
        ws.set_column(2, 2, 8)
        ws.set_column(3, 3, 6)

    def writeAssumptionLine(self, wb, ws, rownum, label, defName, formula=''):
        right_justify = wb.add_format({'align' : 'right'})
        percent_format = wb.add_format({'num_format' : '#0.0%'})
        col = dict(zip(string.ascii_uppercase, range(0, 26)))

        disc = 0
        if 'WPA' in label :
            disc = 1

        ws.write(rownum, col['B'], label, right_justify)
        ws.write(rownum, col['C'], disc, percent_format)

        if formula == '':
            ws.write(rownum, col['D'], 0, percent_format)
        else:
            ws.write_formula(f'D{rownum + 1}', formula, percent_format)

        f = f"='{ws.name}'!$C${rownum + 1}"
        wb.define_name(f'{defName}_SW', f)

        f = f"='{ws.name}'!$D${rownum + 1}"
        wb.define_name(f'{defName}_HW', f)

    def writeBlankRow(self, worksheet, rownum):
        row = ['']
        worksheet.write_row(rownum, 0, row)
        rownum += 1

    def getColLetter(self, col):
        letters = dict(zip(range(1, 27), string.ascii_uppercase))

        letterItem = col

        colLetter = ''

        try:
            if letterItem > 26:
                letterItem -= 26
                colLetter = f'A{letters[letterItem]}'
            else:
                colLetter = letters[letterItem]
        except:
            print(letterItem)

        return colLetter


    def printRows(self, fileName='LDOS Info.xlsx'):
        path = '../Output/' + str(fileName)

        BES =[
            'Cloud and Compute',
            'Cloud Networking',
            'Collaboration',
            'Enterprise Routing',
            'Enterprise Switching',
            'IOT',
            'Meraki',
            'Other',
            'Security',
            'Service Provider Routing',
            'Wireless'
        ]

        wb = xlsxwriter.Workbook(path)

        currency_format = wb.add_format({'num_format' : '$###,##0.00'})
        percent_format = wb.add_format({'num_format' : '#0%'})

        color = '#92D050'
        currency_format_green = wb.add_format({'num_format' : '$###,##0.00', 'bg_color' : color})
        percent_format_green = wb.add_format({'num_format' : '#0%', 'bg_color' : color})
        bold_format_green = wb.add_format({'bold' : True, 'bg_color' : color})

        currItem = 1

        self.buildAssumptions(wb, fileName)

        for BE in BES:

            ws = wb.add_worksheet(BE)

            #Filter for only the current BE
            BETemplate = list(filter(lambda x : x.BU == BE, self.items))


            if len(BETemplate) < 1:
                continue
            else:
                BETemplate = BETemplate[0]
            if len(BETemplate.templates) < 1:
                continue

            discounts = [0, 0]
            if BE == 'Service Provider Routing':
                discounts = ['Hist_SP_Disc_HW', 'WPA_SP_Disc_HW', 'Hist_SP_Disc_SW', 'WPA_SP_Disc_SW']
            elif BE == 'Collaboration':
                discounts = ['Hist_Disc_HW', 'WPA_Collab_Disc_HW', 'Hist_Disc_SW', 'WPA_Collab_Disc_SW']
            elif BE == 'Meraki':
                discounts = ['Hist_Meraki_Disc_HW', 'WPA_Meraki_Disc_HW', 'Hist_Meraki_Disc_SW', 'WPA_Meraki_Disc_SW']
            elif BE == 'IOT':
                discounts = ['Hist_IoT_Disc_HW', 'WPA_IoT_Disc_HW', 'Hist_IoT_Disc_SW', 'WPA_IoT_Disc_SW']
            elif BE == 'Cloud and Compute':
                discounts = ['Hist_Disc_HW', 'Hist_Disc_HW', 'Hist_Disc_SW', 'Hist_Disc_SW']
            else:
                discounts = ['Hist_Disc_HW', 'WPA_Disc_HW', 'Hist_Disc_SW', 'WPA_Disc_SW']

            rownum = 0
            startrow = 0


            for template in BETemplate.templates:
                row = [template.name]
                ws.write_row(rownum, 0, row)
                rownum += 2
                startrow = rownum

                templateHeadings = []

                if 'FW' in template.name and BE == 'Security':
                    discounts = ['Hist_Disc_HW', 'WPA_FW_Disc_HW', 'Hist_Disc_SW', 'WPA_FW_Disc_SW']

                firstLine = True
                totalCol = 0

                for item in template.items:
                    row = [item.LineNumber, item.PartNumber, item.Description, item.UnitListPrice, item.Qty,
                           item.Disc, item.UnitNetPrice, item.ExtendedNetPrice, 0, 0
                           ]
                    col = 0
                    for c in row:
                        if col == 3:
                            ws.write(rownum, col, c, currency_format)
                        elif col == 5:
                            #ws.write(rownum, col, c, percent_format)
                            ws.write_formula(f'F{rownum+1}', discounts[0], percent_format)
                        elif col == 6:
                            ws.write_formula(f'G{rownum+1}', f'=D{rownum+1}*(1-F{rownum+1})', currency_format)
                        elif col == 7:
                            ws.write_formula(f'H{rownum+1}', f'=G{rownum+1}*E{rownum+1}', currency_format)
                        elif col == 8:
                            ws.write_formula(f'I{rownum + 1}', discounts[1], percent_format)
                        elif col == 9:
                            #J=D*(1-I)*E
                            #D = Unit List, I = WPA Disc, E = Qty, J = WPA Price
                            ws.write_formula(f'J{rownum+1}', f'=D{rownum+1}*(1-I{rownum+1})*E{rownum+1}',
                                             currency_format)
                        else:
                            ws.write(rownum, col, c)
                        col += 1


                    if firstLine:
                        firstLine = False
                        if len(template.dates) < 1:
                            totalCol = col
                        for dateItem in template.dates.items():
                            date, extra = str(dateItem[0]).split(".")
                            date = str(date)

                            if {'header': date} not in templateHeadings:
                                templateHeadings.append({'header': date})
                            ws.write(rownum, col, dateItem[1])

                            col += 1
                            if col > totalCol:
                                totalCol = col

                    rownum += 1

                templateHeadings = self.headings + templateHeadings

                finalCol = self.getColLetter(totalCol)
                tableName = str(template.name).replace(' ', '') + 'Tbl'
                ws.add_table(f'A{startrow}:{finalCol}{rownum}', {'columns': templateHeadings,
                                                                 'name': tableName})
                rownum += 1
                ws.write_formula(f'H{rownum}', f'=sum(H{startrow+1}:H{rownum - 1})', currency_format)
                ws.write_formula(f'J{rownum}', f'=sum(J{startrow + 1}:J{rownum - 1})', currency_format)

                adjName = str.replace(str(template.name), ' ', '_')
                adjName = str.replace(adjName, '/', '')
                subName = f'{adjName}_SubTotal'
                formula = f"='{ws.name}'!$H${rownum}"

                wb.define_name(f'{subName}', formula)

                wpaSubName = f'{adjName}_WPA_SubTotal'
                wpaFormula = f"='{ws.name}'!$J${rownum}"

                wb.define_name(f'{wpaSubName}', wpaFormula)

                rownum += 1

                # Build Savings Estimation
                ws.write(rownum - 1, 8, 'Savings', bold_format_green)
                ws.write_formula(f'J{rownum}', f'H{rownum - 1} - J{rownum - 1}', currency_format_green)
                ws.write_formula(f'K{rownum}', f'IFERROR(J{rownum} / H{rownum - 1}, 0)', percent_format_green)

                rownum += 1

            ws.autofit()
            ws.set_column(10, 50, 7)

        wb.set_size(2000, 1500)
        wb.close()
        print(f'Completed export of {fileName}')








