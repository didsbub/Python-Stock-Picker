import xlwt
import os
from datetime import date, datetime


class ExcelExporter:
    def __init__(self, filename = "excelexport.xls"):
        self.filename = filename
        self.wb = None

    def GetFilename(self):
        return self.filename

    def SetWorkBook(self, wb):
        self.wb = wb

    def GetWorkBook(self):
        if self.wb != None:
            return self.wb
        else:
            self.wb = xlwt.Workbook(encoding="utf-8")
            return self.wb

    def SaveWorkBook(self):
        self.GetWorkBook().save(self.GetFilename())

    def exporttoexcel(self, data, bestkeys, sheetname):

        def writecolumnheadings(sheet, columnlist):
            colno = -1
            for column in columnlist:
                colno += 1
                title, numberformat, width, style = column
                myformat = xlwt.easyxf(strg_to_parse=bgtitles+";"+style)

                sheet.write(0,colno,title, myformat)
                sheet.col(colno).width = 256*width


        #.........................Number Style Definitions.........................
        dollarformatstr = '_($* #,##0_);_($* (#,##0);_($* "-"??_);_(@_)'
        datetimeformatstr = 'm/d/yyyy h:mm'
        dateformatstr = 'm/d/yyyy'
        textformatstr = ""

        #........Typical Column Widths ( units are in number of characters )........
        datecolumnswidth = 15
        datetimecolumnswidth = 20
        quarterrevenuecolumnswidth = 20
        quarterepscolumnswidth = 10
        annualrevenuecolumnswidth = 25
        annualepscolumnswidth = 10
        symbolcolumnswidth = 17
        companycolumnswidth = 50
        industrycolumnswidth = 30
        pricecolumnswidth = 10

        #..............................Style Definitions...........................
        bgtitles = "font: bold on; align: wrap on, vert centre, horiz center"
        bgn = "pattern:pattern solid, fore-color %s"
        bg1a = bgn%"white"
        bg1b = bgn%"light_turquoise"
        bg2a = bgn%"tan"
        bg2b = bgn%"light_yellow"
        bg3a = bgn%"pale_blue"
        bg3b = bgn%"turquoise"
        bg4 = bgn%"lime"

        #............List of column names and corresponding formatting..............
        #........(title, numberformatting, columnwidth, background style)...........
        columnlist = (
        ("Symbols", textformatstr, symbolcolumnswidth, bg1a),
        ("Company", textformatstr, companycolumnswidth, bg1a),
        ("Industry", textformatstr, industrycolumnswidth, bg1a),
        ("Price", dollarformatstr, pricecolumnswidth, bg1b),
        ("Price Date", datetimeformatstr, datetimecolumnswidth, bg1b),
        #1st quarter:
        ("Quarter Date", dateformatstr, datecolumnswidth, bg2a),
        ("Quarterly Revenue", dollarformatstr, quarterrevenuecolumnswidth, bg2a),
        ("Quarterly EPS", dollarformatstr, quarterepscolumnswidth, bg2a),
        #2nd quarter:
        ("Quarter Date", dateformatstr, datecolumnswidth, bg2b),
        ("Quarterly Revenue", dollarformatstr, quarterrevenuecolumnswidth, bg2b),
        ("Quarterly EPS", dollarformatstr, quarterepscolumnswidth, bg2b),
        #3rd quarter:
        ("Quarter Date", dateformatstr, datecolumnswidth, bg2a),
        ("Quarterly Revenue", dollarformatstr, quarterrevenuecolumnswidth, bg2a),
        ("Quarterly EPS", dollarformatstr, quarterepscolumnswidth, bg2a),
        #4th quarter:
        ("Quarter Date", dateformatstr, datecolumnswidth, bg2b),
        ("Quarterly Revenue", dollarformatstr, quarterrevenuecolumnswidth, bg2b),
        ("Quarterly EPS", dollarformatstr, quarterepscolumnswidth, bg2b),
        #5th quarter:
        ("Quarter Date", dateformatstr, datecolumnswidth, bg2a),
        ("Quarterly Revenue", dollarformatstr, quarterrevenuecolumnswidth, bg2a),
        ("Quarterly EPS", dollarformatstr, quarterepscolumnswidth, bg2a),
        #1st year:
        ("Annum Date", dateformatstr, datecolumnswidth, bg3a),
        ("Annual Revenue", dollarformatstr, annualrevenuecolumnswidth, bg3a),
        ("Annual EPS", dollarformatstr, annualepscolumnswidth, bg3a),
        #2nd year:
        ("Annum Date", dateformatstr, datecolumnswidth, bg3b),
        ("Annual Revenue", dollarformatstr, annualrevenuecolumnswidth, bg3b),
        ("Annual EPS", dollarformatstr, annualepscolumnswidth, bg3b),
        #3rd year:
        ("Annum Date", dateformatstr, datecolumnswidth, bg3a),
        ("Annual Revenue", dollarformatstr, annualrevenuecolumnswidth, bg3a),
        ("Annual EPS", dollarformatstr, annualepscolumnswidth, bg3a),
        #4th year:
        ("Annum Date", dateformatstr, datecolumnswidth, bg3b),
        ("Annual Revenue", dollarformatstr, annualrevenuecolumnswidth, bg3b),
        ("Annual EPS", dollarformatstr, annualepscolumnswidth, bg3b),
        #Extra Info (projected values etc.):
        ("Projected EPS", dollarformatstr, annualepscolumnswidth, bg4),
        ("Projected Revenue", dollarformatstr, annualrevenuecolumnswidth, bg4),
        ("Average EPS Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Average Revenue Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Years of EPS Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Years of Revenue Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("PE", dollarformatstr, annualepscolumnswidth, bg4),
        )
        

        wb = self.GetWorkBook()
        sheet = wb.add_sheet(sheetname)

        #.....................Writing column headings..............................
        writecolumnheadings(sheet, columnlist)
        
        #.......................Generating Style Cache.............................
        styledict = {}
        for col in columnlist:
            numberformat = col[1]
            color = col[3]
            style = (numberformat, color)
            if not styledict.has_key(style):
                styledict[style] = xlwt.easyxf(num_format_str=numberformat, strg_to_parse=color)


        #.......................Writing column data.................................
        for rowno in range(1, len(data)+1):
            rowdata = data[rowno-1]

            (symbol, company, industry, price, pricedate, (quarterdates,
                quarterrevenues, quarterepses, annualdates, annualrevenues,
                annualepses),extrainfovalues) = rowdata

            print "Writing %s"%symbol

            #print quarterdates
            #print quarterrevenues
            #print quarterepses

            for colno in range(0, len(columnlist)):
                try:
                    if colno < 5:
                        value = (symbol, company, industry, price, pricedate)[colno]
                    elif colno < 20:
                        value = (quarterdates, quarterrevenues,
                                quarterepses)[(colno-5)%3][(colno-5)/3]
                    elif colno < 32:
                        value = (annualdates, annualrevenues,
                                annualepses)[(colno-20)%3][(colno-20)/3]
                    elif colno < 39:
                        extrainfokey = columnlist[colno][0].lower().replace(" ","")
                        value = extrainfovalues[extrainfokey]
                    else:
                        value = "N/A"
                except:
                    value = "N/A"

                try:
                    numberformat = columnlist[colno][1]
                    color = columnlist[colno][3]
                    sheet.write(rowno, colno, value,
                            styledict[numberformat, color])
                except:
                    print colno
                    print value

        #..............GENERATING SUMMARY SHEET FOR BEST STOCKS.....................
        print "Genarating summary sheet for %s"%sheetname
        sheet = wb.add_sheet(sheetname+"_Summary")
        #............List of column names and corresponding formatting..............
        #........(title, numberformatting, columnwidth, background style)...........
        columnlist = (
        ("Symbols", textformatstr, symbolcolumnswidth, bg1a),
        ("Price", dollarformatstr, pricecolumnswidth, bg1b),
        ("Average Revenue Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Years of Revenue Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Average EPS Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("Years of EPS Growth", dollarformatstr, annualepscolumnswidth, bg4),
        ("PE", dollarformatstr, annualepscolumnswidth, bg4),
        ("Company", textformatstr, companycolumnswidth, bg1a),
        ("Industry", textformatstr, industrycolumnswidth, bg1a),
        )

        #.....................Writing column headings..............................
        writecolumnheadings(sheet, columnlist)


        #...............Find Index Numbers of Best Stocks in [data].................
        bestkeyindex={}
        for i in range(0, len(data)):
            if data[i][0] in bestkeys:
                bestkeyindex[data[i][0]] = i

        #.......................Writing column data.................................
        rowno = 0
        for symbol in bestkeys:
            rowno += 1
            dataindexno = bestkeyindex[symbol]
            rowdata = data[dataindexno]

            (symbol, company, industry, price, pricedate, (quarterdates,
                quarterrevenues, quarterepses, annualdates, annualrevenues,
                annualepses),extrainfovalues) = rowdata

            if symbol not in bestkeys: 
                continue

            for colno in range(0, len(columnlist)):
                try:
                    if colno < 2:
                        value = (symbol, price)[colno]
                    elif colno < 7:
                        #Obtain the key from titles of columns, ie. title = "Average Revenue Growth"
                        #=> key = averagerevenuegrowth
                        extrainfokey = columnlist[colno][0].lower().replace(" ","")
                        value = extrainfovalues[extrainfokey]
                    elif colno < 9:
                        value = (company, industry)[colno-7]
                    else:
                        value = "N/A"
                except:
                    value = "N/A"

                try:
                    numberformat = columnlist[colno][1]
                    color = columnlist[colno][3]
                    sheet.write(rowno, colno, value,
                            styledict[numberformat, color])
                except:
                    print colno
                    print value
        
        self.SaveWorkBook()
        print "Genaration of excel file for %s is complete."%sheetname


if __name__ == "__main__":
    #colorpalette()
    exporttoexcel("excelexporter.xls")




