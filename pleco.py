#!/usr/bin/python
import os, shutil
import sys
import re
from BeautifulSoup import BeautifulSoup
import sqlite3
import urllib
import hashlib
import time
import datetime
import excelexporter

SCHEMA = """
CREATE TABLE {0}_COMPANIES (
    symbol TEXT PRIMARY KEY,
    company TEXT,
    industry TEXT
);

CREATE TABLE {0}_PRICES (
    symbol TEXT,
    date INTEGER,
    price INTEGER
);

CREATE TABLE {0}_FINANCIALS (
    symbol TEXT,
    type TEXT,
    date TEXT,
    value INTEGER
);

"""


DATABASE_NAME = "pleco.db"
CACHE_FOLDER = "cache"


class Database:
    def __init__(self, stockexchangename):
        create = not os.path.exists( DATABASE_NAME )

        self.STOCKEXCHANGENAME=stockexchangename

        self.conn = sqlite3.connect( DATABASE_NAME, timeout=99.0 )
        if create:
            c = self.conn.cursor()
            c.executescript( SCHEMA.format(stockexchangename) )
        else:
            c = self.conn.cursor()
            #...Obtaining the existing stock exchange markets in the database..
            c.execute("SELECT name FROM sqlite_master WHERE type='table'")
            existingtablelist = c.fetchall()
            existingstockexchangelist = []
            for item in existingtablelist:
                tablename = item[0]
                if tablename.find("COMPANIES") != -1:
                    existingstockexchangelist.append(tablename.split("_")[0])
            #..Introducing new stock exchange markets to database if necessary..
            if stockexchangename not in existingstockexchangelist:
                c.executescript( SCHEMA.format(stockexchangename) )



    def addCompany( self, symbol, company, industry ):
        c = self.conn.cursor()
        c.execute( "DELETE FROM %s_COMPANIES WHERE symbol=?"%self.STOCKEXCHANGENAME, (symbol,));
        c.execute( "INSERT INTO %s_COMPANIES values ( ?, ?, ? )"%self.STOCKEXCHANGENAME, 
                (symbol, company, industry ) )
        self.conn.commit()

    def clearCompanies(self):
        c = self.conn.cursor()
        c.execute( "DELETE FROM %s_COMPANIES"%self.STOCKEXCHANGENAME)
        self.conn.commit()

    def getCompanies(self):
        c = self.conn.cursor()
        c.execute( "SELECT * FROM %s_COMPANIES"%self.STOCKEXCHANGENAME )
        return c.fetchall();

    def setPrice(self, symbol, date, price):
        c = self.conn.cursor()
        c.execute( "INSERT INTO %s_PRICES VALUES (?, ?, ?)"%self.STOCKEXCHANGENAME,
                ( symbol, date, price ) )
        self.conn.commit()

    def getPrice(self, symbol ):
        c = self.conn.cursor()
        c.execute( "SELECT price FROM %s_PRICES WHERE symbol=? ORDER BY DATE DESC"%self.STOCKEXCHANGENAME,
                ( symbol, ) )
        return c.fetchone()[0]

    def clearPrices(self):
        c = self.conn.cursor()
        c.execute( "DELETE FROM %s_PRICES"%self.STOCKEXCHANGENAME)
        self.conn.commit()

    def clearFinancials(self):
        c = self.conn.cursor()
        c.execute( "DELETE FROM %s_FINANCIALS"%self.STOCKEXCHANGENAME)
        self.conn.commit()

    def setFinancials( self, symbol, type, date, value ):
        c = self.conn.cursor()
        c.execute("DELETE FROM %s_FINANCIALS WHERE symbol=? AND type=? and date=?"%self.STOCKEXCHANGENAME,
                (symbol, type, date))
        c.execute("INSERT INTO %s_FINANCIALS VALUES (?, ?, ?, ?)"%self.STOCKEXCHANGENAME,
                ( symbol, type, date, value ) )
        self.conn.commit()

    def getFinancials( self, symbol, type ):
        c = self.conn.cursor()
        c.execute( "SELECT * FROM %s_FINANCIALS WHERE symbol=? AND type=? ORDER BY DATE DESC"%self.STOCKEXCHANGENAME,
                (symbol, type))
        return c.fetchall()

    def getEverything( self ):
        c = self.conn.cursor()
        c.execute( """
                SELECT {0}_COMPANIES.symbol, company, industry, type, value, price from
                {0}_COMPANIES, {0}_PRICES, {0}_FINANCIALS WHERE
                {0}_COMPANIES.symbol = {0}_PRICES.symbol AND {0}_PRICES.symbol =
                {0}_FINANCIALS.symbol""".format(self.STOCKEXCHANGENAME))

        return c.fetchall()

    def getEverythingIncludingDates( self ):
        c = self.conn.cursor()
        c.execute( """
                SELECT {0}_COMPANIES.symbol, {0}_COMPANIES.company,
                {0}_COMPANIES.industry, {0}_PRICES.price, {0}_PRICES.date,
                {0}_FINANCIALS.type, {0}_FINANCIALS.date, {0}_FINANCIALS.value
                from {0}_COMPANIES, {0}_PRICES, {0}_FINANCIALS WHERE
                {0}_COMPANIES.symbol = {0}_PRICES.symbol AND {0}_PRICES.symbol
                = {0}_FINANCIALS.symbol ORDER BY {0}_COMPANIES.symbol""".format(self.STOCKEXCHANGENAME))

        return c.fetchall()

    def close( self ):
        self.conn.close()

class PageCache:
    def __init__(self):
        if not os.path.exists( CACHE_FOLDER ):
            os.mkdir( CACHE_FOLDER )

    def get( self, url, fname = None ):
        if fname == None:
            fname = hashlib.sha1(url).hexdigest()
        filename = os.path.join( CACHE_FOLDER, fname )

        if os.path.exists( filename ):
            return open( filename, "rt" ).read()
        else:
            print "Retrieve %s" % url
            try:
                f = urllib.urlopen(url)

                content = f.read()
                f.close()
            except IOError:
                print >>sys.stderr, "Unable to connect to %s, retrying in 10 seconds. Please check your internet connection."%url
                time.sleep(10)
                return self.get(url, fname)

            f = open( filename, "w" );
            f.write( content );
            f.close()

            return content

    def EmptyCache(self):
        shutil.rmtree(CACHE_FOLDER)


class Pleco_TSX:
    def __init__(self):
        self.db = Database(self.GetStockExchangeName())
        self.webCache = PageCache()

    def GetStockExchangeName(self):
        return "TSX"

    def convertToGlobeAndMailFormat(self, symbol):
        # convert symbol from employed format in database to globeandmail.com
        # format for scraping industries
        symbol = symbol.upper()
        if symbol.startswith("TSE:"):
            symbol = symbol[4:]+"-T"
        return symbol

    def scrapeIndustryForSymbol( self, symbol ):
        # lookup file, otherwise retrieve the url
        url = "http://www.theglobeandmail.com/globe-investor/markets/stocks/summary/?q=%s"
        page = self.webCache.get( url%self.convertToGlobeAndMailFormat(symbol))
        item = BeautifulSoup(page,
                convertEntities=BeautifulSoup.HTML_ENTITIES).find( 'li', {"class":"industry last"})

        if item == None or item.string == None or len(item.string.strip(" "))==0:
            print "Warning: Cannot find industry in %s" % url%self.convertToGlobeAndMailFormat(symbol)
            return "N/A"
        else:
            return item.string

    def scrapeCompanyNameForSymbol( self, symbol ):
        url = "http://www.google.com/finance?q=%s&fstype=ii" % symbol.upper()
        page = self.webCache.get( url )

        expr = re.compile(r"""Financial Statements for (.*?) - Google Finance""")
        m = expr.search(page)
        if m:
            return BeautifulSoup(m.group(1),
                    convertEntities=BeautifulSoup.HTML_ENTITIES).contents[0].string
        else:
            return None

    def scrapeCompanies( self ):
        self.db.clearCompanies()
        letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        PageExpr = re.compile("""Page \d+ of (\d+)""")
        SymExpr = re.compile("""symbol=([^"&]+)""")
        found = {}

        def process(page):
            m = SymExpr.findall( page )
            for a in m:
                symbol = "TSE:" + str(a)
                if symbol in found: continue
                found[symbol] = 1
                name = self.scrapeCompanyNameForSymbol( symbol )
                industry = self.scrapeIndustryForSymbol( symbol )
                if name == None or industry == None: continue
                print "Found %s (%s) - %s" % (name, symbol, industry)
                self.db.addCompany( symbol, name, industry )

        for s in letters:
            url = "http://www.tmx.com/HttpController?GetPage=ListedCompaniesViewPage&SearchCriteria=Name&SearchKeyword=%s&SearchType=StartWith&Page=%d&SearchIsMarket=Yes&Market=T&Language=en" % (s, 1)
            page = self.webCache.get( url )
            m = PageExpr.search( page )
            if m:
                numPages = int(m.group(1))
            else:
                numPages = 1

            process(page)

            for p in range( 1, numPages ):
                url = "http://www.tmx.com/HttpController?GetPage=ListedCompaniesViewPage&SearchCriteria=Name&SearchKeyword=%s&SearchType=StartWith&Page=%d&SearchIsMarket=Yes&Market=T&Language=en" % (s, p)
                page = self.webCache.get( url )
                process(page)

    def scrapePrices(self):
        self.db.clearPrices()
        symbols = []
        for company in self.db.getCompanies():
            symbols.append(company[0])

        self.scrapePricesforSymbols(symbols)

    def convertToYahooFormat( self, list ):
        # convert from google to yahoo format.
        ret = []
        for symbol in list:
            symbol = symbol[4:] # remove tse:
            symbol = symbol.lower().replace('.', '-') + ".to"
            ret.append( symbol )

        return ret

    def scrapePricesforSymbols( self, companysymbols ):
        date = int(time.time())
        def getPrices(stocks, list):
            prices = requestYahooPrices( self.convertToYahooFormat( list ) )

            for i in range(len(prices)):
                self.db.setPrice( list[i], date, prices[i] )
                print "%s = $%.2f" % (list[i], float(prices[i]) / 1000)

        def requestYahooPrices(symbols):
            # form HTTP request
            url = "http://finance.yahoo.com/d/quotes.csv?s=%s&f=l1&e=.csv" % \
                ( ",".join(symbols) )
            prices = []

            content = self.webCache.get(url)

            # for each line,
            for line in content.split("\n"):
                line = line.strip()
                if line == "": continue
                prices.append(int(float(line) * 1000))

            # return the list.
            return prices

        stocks = {}
        date = int(time.time())
        for symbol in companysymbols:
            stocks[symbol] = 0

        # for each chunk of 64 stocks,
        array = []
        for key in stocks.keys():
            array.append( key )
            if len( array ) == 64:
                getPrices( stocks, array )
                array = []

        if len( array ) > 0:
            getPrices( stocks, array )
            array = []

    def scrapeFinancials( self ):
        self.db.clearFinancials()
        for company in self.db.getCompanies():
            self.scrapeFinancialsForSymbol( company[0] )

    def convertToGoogleFormat(self, symbol):
        # convert symbol from employed format in database to yahoo format.
        # for scraping prices
        return symbol # google format is employed in database for tsx

    def scrapeFinancialsForSymbol( self, symbol ):
        date = int(time.time())

        def checkPresence( page, pattern ):
            for line in page:
                if line.find(pattern) != -1:
                    return True

            return False

        def extractRow( soup, text ):
            def byname(tag):
                return str(tag.string).rstrip() == text and tag.name == 'td'

            tag = soup.find(byname)
            contents = []
            while tag:
                tag = tag.findNextSibling('td')
                if tag == None: break
                contents.append(str(tag.find(text=True)))
            return moneyToNumber(contents)

        def moneyToNumber( arr ):
            ret = []
            for a in arr:
                if a == '-':
                    ret.append(0)
                else:
                    ret.append(int(float(a.replace(",", "")) * 1000 ))

            return ret

        def extractDates( lines ):
            values = []
            expr = re.compile(r"""(\d\d\d\d-\d\d-\d\d)""")
            for line in lines:
                m = expr.search(line)
                if m:
                    values.append( m.group(0) )
                else:
                    values.append("")

            return values

        def findLinesLike( page, pattern ):
            lines = []
            skipped = -1
            pattern = re.compile(pattern)
            for line in page:
                if pattern.search(line):
                    lines.append( line )
                    skipped = 0
                elif skipped >= 0:
                    skipped += 1
                    if skipped >= 5:
                        break
            return lines

        print "Scraping financials for %s" % symbol

        # retrieve the web page
        url = "http://www.google.com/finance?q=%s&fstype=ii" % self.convertToGoogleFormat(symbol)
        page = self.webCache.get( url )
        soup = BeautifulSoup(page)
        page = page.split('\n')
        quarterlyPage = soup.find( "div", { "id" : "incinterimdiv" } )
        annualPage = soup.find( "div", { "id" : "incannualdiv" } )

        qstr = str(quarterlyPage).split('\n')
        astr = str(annualPage).split('\n')

        # Look for "In Millions of". If not there, error!
        if not checkPresence( page, "In Millions of" ):
            print >>sys.stderr, "While processing %s could not find 'In Millions of' at %s" % (symbol, url)
            return False

        # Set multiplier to 1000000
        multiplier = 1000000

        # build array of all lines like "3 months Ending"
        quarterlyDates = extractDates(findLinesLike( qstr, """\d+ (months|weeks) ending""" ))

        # Build array of all lines like "12 months Ending"
        annualDates = extractDates(findLinesLike( astr, """\d+ (months|weeks) ending""" ))

        # Look for td containing "Total Revenue"
        # Extract all td elements in siblings that contain only a number

        # Build table for revenue
        quarterlyRevenue = extractRow( quarterlyPage, "Revenue" )
        annualRevenue = extractRow( annualPage, "Revenue" )

        # Build table for ";Diluted EPS Normalized EPS&"
        quarterlyEPS = extractRow( quarterlyPage, "Diluted Normalized EPS" )
        annualEPS = extractRow( annualPage, "Diluted Normalized EPS" )

        for i in range( len(quarterlyRevenue) ):
            self.db.setFinancials( symbol, "QuarterlyRevenue", quarterlyDates[i],
                    quarterlyRevenue[i] * multiplier )
            self.db.setFinancials( symbol, "QuarterlyEPS", quarterlyDates[i],
                    quarterlyEPS[i] )

        for i in range( len(annualRevenue) ):
            self.db.setFinancials( symbol, "AnnualRevenue", annualDates[i],
                    annualRevenue[i] * multiplier )
            self.db.setFinancials( symbol, "AnnualEPS", annualDates[i],
                    annualEPS[i] )

    def addProjected( self, symbol, type ):
        financials = self.db.getFinancials( symbol, "Quarterly%s" % type )
        if len(financials) < 4:
            return

        projected = financials[0][3] + financials[1][3] + financials[2][3] + \
                    financials[3][3]

        self.db.setFinancials( symbol, "Projected%s" % type, 0, projected )

    def addAverageGrowth( self, symbol, type ):
        financials = self.db.getFinancials( symbol, "Annual%s" % type )
        avgGrowth = 0.0
        if len(financials) > 1:
            projected = self.db.getFinancials( symbol, "Projected%s" % type )
            financials.extend( projected )
            financials.reverse()
            first = financials[0][3]
            count = 0
            for val in financials:
                if first > 0:
                    growth = float((val[3] - first)) / first
                    avgGrowth += growth
                    count += 1
                else:
                    avgGrowth = 0.0
                    count = 0
                first = val[3]

            if count < 2:
                avgGrowth = 0.0
            else:
                avgGrowth /= count
        
        self.db.setFinancials( symbol, "Average%sGrowth" % type, 0, 
                round( avgGrowth * 100 ) )

    def addYearsOfGrowth( self, symbol, type ):
        financials = self.db.getFinancials( symbol, "Annual%s" % type )
        count = 0
        if len(financials) > 0:
            last = financials[0]
            for line in financials[1:]:
                if line[3] < last:
                    count += 1
                else:
                    break

        self.db.setFinancials( symbol, "YearsOf%sGrowth" % type, 0, count )

    def addPE( self, symbol ):
        price = self.db.getPrice( symbol )
        financials = self.db.getFinancials( symbol, "ProjectedEPS" )
        if len(financials) == 0:
            return

        earnings = financials[0][3]
        if earnings > 0:
            pe = round(float(price)/float(earnings) * 10)
        else:
            pe = 0

        self.db.setFinancials( symbol, "PE", 0, pe );

    def addExtraInfo( self ):
        for company in self.db.getCompanies():
            symbol = company[0]
            print "Processing %s...    \r" % symbol,
            sys.stdout.flush()
            self.addProjected(symbol, "EPS")
            self.addProjected(symbol, "Revenue")
            self.addAverageGrowth( symbol, "EPS" )
            self.addAverageGrowth( symbol, "Revenue" )
            self.addYearsOfGrowth( symbol, "EPS" )
            self.addYearsOfGrowth( symbol, "Revenue" )
            self.addPE( symbol )

        print

    def process(self, returnsymbols=False):
        stocks = {}
        for record in self.db.getEverything():
            symbol = record[0]
            company = record[1]
            industry = record[2]
            type = record[3]
            value = record[4]
            price = record[5]
            if symbol not in stocks:
                stock = { "symbol": symbol, 
                    "price": price, 
                    "company": company,
                    "industry": industry}
                stocks[symbol] = stock
            else:
                stock = stocks[symbol]

            stock[type] = value

        stocks = filter( self.filt, stocks.values() )

        stocks.sort( key = lambda stock: stock["AverageRevenueGrowth"] )
        if returnsymbols:
            symbols = []
            for stock in stocks:
                symbols.append(stock["symbol"])
            return symbols
        else:
            self.printTable(stocks)

    def filt(self, stock):
        return \
            stock["YearsOfRevenueGrowth"] >= 1 and \
            stock["YearsOfEPSGrowth"] >= 1 and \
            stock["AverageRevenueGrowth"] >= 3 and \
            stock["AverageEPSGrowth"] >= 3 and \
            "PE" in stock and \
            stock["PE"] >= 0 and \
            stock["PE"] <= 50 \
            and stock["ProjectedEPS"] > 0 \
          

    def printTable(self, stocks):
        print "symbol, AverageRevenueGrowth, YearsOfRevenueGrowth, AverageEPSGrowth, YearsOfEPSGrowth, PE, Company"
        for stock in stocks:
            print stock["symbol"].ljust(13),
            print str(stock["AverageRevenueGrowth"]).ljust(5),
            print str(stock["YearsOfRevenueGrowth"]).ljust(3),
            print str(stock["AverageEPSGrowth"]).ljust(5),
            print str(stock["YearsOfEPSGrowth"]).ljust(3),
            print str(stock["PE"]).ljust(5),
            print stock["company"],
            print

    def exportToExcel(self, exporterobject=excelexporter.ExcelExporter()):
        def processme():
            stocks = {}
            print "Querying database for %s in order to export to Excel."%self.GetStockExchangeName()
            records = self.db.getEverythingIncludingDates()
            print "Database query finished. Processing..."
            for record in records:
                symbol = record[0]
                company = record[1]
                industry = record[2]
                price = record[3]
                pricedate = record[4]
                type = record[5]
                date = record[6]
                value = record[7]
                if symbol not in stocks:
                    stock = { "symbol": symbol, 
                        "company": company,
                        "industry": industry,
                        "price": price,
                        "pricedate": pricedate,
                        "financials":[]
                        }
                    stocks[symbol] = stock
                else:
                    stock = stocks[symbol]

                stock["financials"].append([type, date, value])

            return stocks

        data = [] #this will contain all the information required to fill the Excel sheet
        stocks = processme() # query the database for entire database content
        stocksymbols = stocks.keys()
        stocksymbols.sort() # required in order to populate the [data] with ordered symbols
        for stocksymbol in stocksymbols:
            print stocksymbol
            stock = stocks[stocksymbol]
            symbol, company, industry = stock["symbol"],stock["company"],stock["industry"]
            price, pricedate = stock["price"], stock["pricedate"]
            financials = stock["financials"]

            #convert time of the price from epoch to human readable format
            ts = time.localtime(pricedate)
            pricedate = datetime.datetime(ts[0],ts[1],ts[2],ts[3],ts[4],ts[5])

            #convert dates of financials from plain text to datetime.date
            for i in range(0,len(financials)):
                datestring = financials[i][1]
                try:
                    year, month, day = map(int, datestring.split("-"))
                    financials[i][1] = datetime.date(year,month,day)
                except:
                    year, month, day = (1900,01,01) #Error in date
                    #print >>sys.stderr, "Error in date of %s"%stocksymbol
                    financials[i][1] = datetime.date(year,month,day)

            #Sort each financial value by date (increasing order) and by type:
            #by date:
            financials.sort(lambda x,y: 1 if x[1]>y[1] else -1 if x[1]<y[1] else 0)
            #by type:
            financials.sort(lambda x,y: 1 if x[0]>y[0] else -1 if x[0]<y[0] else 0)

            #Group financials according to financial data type
            quarterlyrevenue = []
            quarterlyeps = []
            annualrevenue = []
            annualeps = []
            extrainfovalues = {}
            erronousfinancials = []
            for item in financials:
                if item[0].lower().find("quarterlyrevenue") != -1:
                    quarterlyrevenue.append(item)
                elif item[0].lower().find("quarterlyeps") != -1:
                    quarterlyeps.append(item)
                elif item[0].lower().find("annualrevenue") != -1:
                    annualrevenue.append(item)
                elif item[0].lower().find("annualeps") != -1:
                    annualeps.append(item)
                elif item[0] in ("ProjectedEPS", "ProjectedRevenue",
                        "AverageEPSGrowth", "AverageRevenueGrowth",
                        "YearsOfEPSGrowth", "YearsOfRevenueGrowth", "PE"):
                    extrainfovalues[item[0].lower()] = item[2] #these items do not contain dates
                else:
                    #An unknown (new?) financial data type is detected.
                    erronousfinancials.append(item)
            if len(erronousfinancials) != 0:
                print >>sys.stderr, \
                "Unknown (new?) financial data type(s) encountered (%s) while processing symbol %s"%(erronousfinancials, symbol)

            #print quarterlyrevenue
            #print quarterlyeps
            #print annualrevenue 
            #print annualeps

            quarterdates = []
            quarterrevenues = []
            quarterepses = []
            annualdates = []
            annualrevenues = []
            annualepses = []
                    
            #populate quarter dates, revenues, and epses
            for i in range(0, max(len(quarterlyrevenue),len(quarterlyeps))):
                try:
                    quarterdates.append(quarterlyrevenue[i][1])
                    quarterrevenues.append(quarterlyrevenue[i][2])
                    quarterepses.append(quarterlyeps[i][2])
                    assert quarterlyrevenue[i][1] == quarterlyeps[i][1]
                except AssertionError:
                    print >>sys.stderr, \
                        "Dates of quarterly revenue and quarterly eps are inconsistent for symbol %s"%(symbol)
                except IndexError:
                    print >>sys.stderr, \
                        "Missing or redundant elements are detected in quarterly revenue or quarterly eps of symbol %s"%(symbol)

                #print quarterdates
                #print quarterrevenues
                #print quarterepses

            #populate annual dates, revenues, and epses
            for i in range(0, max(len(annualrevenue),len(annualeps))):
                try:
                    annualdates.append(annualrevenue[i][1])
                    annualrevenues.append(annualrevenue[i][2])
                    annualepses.append(annualeps[i][2])
                    assert annualrevenue[i][1] == annualeps[i][1]
                except AssertionError:
                    print >>sys.stderr, \
                        "Dates of annual revenue and annual eps are inconsistent for symbol %s"%(symbol)
                except IndexError:
                    print >>sys.stderr, \
                        "Missing or redundant elements are detected in annual revenue or annual eps of symbol %s"%(symbol)

                #print annualdates
                #print annualrevenues
                #print annualepses

            data.append((symbol, company, industry, price, pricedate,
                (quarterdates, quarterrevenues, quarterepses, annualdates,
                    annualrevenues, annualepses),extrainfovalues))

        print "Computing summary data..."
        symbolsofbeststocks = self.process(True)
        print "Finished processing %s data. Now generating the Excel file."%self.GetStockExchangeName()
        exporterobject.exporttoexcel(data, symbolsofbeststocks, self.GetStockExchangeName())

    def run(self, arguments=sys.argv):
        for i in range(1, len(arguments)):
            if arguments[i] == "--companies":
                self.scrapeCompanies()
            elif arguments[i] == "--prices":
                self.scrapePrices()
            elif arguments[i] == "--financials":
                self.scrapeFinancials()
            elif arguments[i] == '--extra':
                self.addExtraInfo()
            elif arguments[i] == "--all":
                self.scrapeCompanies()
                self.scrapeFinancials()
                self.scrapePrices()
                self.addExtraInfo()
            elif arguments[i] == "--test": 
                self.addPE("tse:g")
            elif arguments[i] == "--process":
                self.process()
            elif arguments[i] == '--excelexport':
                self.exportToExcel()

    def exit(self):
        self.db.close()

class Pleco_NASDAQ(Pleco_TSX):
    def __init__(self):
        Pleco_TSX.__init__(self)

    def GetStockExchangeName(self):
        return "NASDAQ"

    def scrapeCompanies( self ):
        self.db.clearCompanies()
        url = "http://www.nasdaq.com/screening/companies-by-industry.aspx?exchange=NASDAQ&render=download"
        content = self.webCache.get(url)
        lines = content.split("\n")

        for i in range(1, len(lines)):
            line = lines[i].strip()
            if line == "": continue
            datalist = line.split(",")
            symbol = "NASDAQ:" + datalist[0].strip('"').strip(" ")
            name = datalist[1].strip('"')
            industry = datalist[7].strip('"')
            print "Found %s (%s) - %s" % (name, symbol, industry)
            self.db.addCompany( symbol, name, industry )

    def convertToYahooFormat( self, list ):
        # convert symbol from employed format in database to yahoo format.
        # for scraping prices
        ret = []
        for symbol in list:
            symbol = symbol[7:] # remove NASDAQ:
            #symbol = symbol.lower().replace('.', '-')
            ret.append( symbol )

        return ret

    def convertToGoogleFormat( self, symbol ):
        # convert symbol from employed format in database to google format.
        # for scraping prices
        symbol = symbol[7:] # remove NASDAQ:
        return symbol

class Pleco_NYSE(Pleco_TSX):
    def __init__(self):
        Pleco_TSX.__init__(self)

    def GetStockExchangeName(self):
        return "NYSE"

    def convertToGlobeAndMailFormat(self, symbol):
        symbol = symbol.upper()
        symbol = symbol[5:]+"-N"
        return symbol

    def convertToYahooFormat( self, list ):
        # convert symbol from employed format in database to yahoo format.
        # for scraping prices
        ret = []
        for symbol in list:
            symbol = symbol[5:] # remove NYSE:
            #symbol = symbol.lower().replace('.', '-')
            ret.append( symbol )

        return ret

    def convertToGoogleFormat( self, symbol ):
        # convert symbol from employed format in database to google format.
        # for scraping prices
        return symbol #google format is used in database for nyse

    def scrapeCompanies(self):
        self.db.clearCompanies()
        letters = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")+["Other"]
        symbolexpr = re.compile("""\["(.*?)",""")
        companynameexpr = re.compile("""\[.*?,"(.*?)",""")

        url = "http://www.nyse.com/about/listed/lc_ny_name_%s.js"
        for letter in letters:
            content = self.webCache.get(url%letter)
            symbollist = symbolexpr.findall(content)
            companynamelist = companynameexpr.findall(content)
            #add NYSE: to the beginning of the symbols
            symbollist = map(lambda s: "NYSE:%s"%s, symbollist)
            #convert company name to unicode
            companynamelist = map(lambda s: unicode(s,"UTF-8","replace"), companynamelist)

            

            #Populating Industrylist from globeandmail.com
            industrylist = []
            for symbol in symbollist:
                try:
                    industrylist.append(self.scrapeIndustryForSymbol( symbol ))
                except:
                    industrylist.append("N/A")

            #Writing all entries starting with the same letter to the database:
            for i in range(0, len(symbollist)):
                symbol, name, industry = symbollist[i], companynamelist[i], industrylist[i]
                self.db.addCompany( symbol, name, industry )


class Pleco_HKG(Pleco_TSX):
    def __init__(self):
        Pleco_TSX.__init__(self)

    def GetStockExchangeName(self):
        return "HKG"

    def convertToYahooFormat( self, list ):
        # convert symbol from employed format in database to yahoo format.
        # for scraping prices
        ret = []
        for symbol in list:
            symbol = symbol[4:]+".HK" # remove HKG: and add ".HK"
            #symbol = symbol.lower().replace('.', '-')
            ret.append( symbol )

        return ret

    def convertToGoogleFormat( self, symbol ):
        # convert symbol from employed format in database to google format.
        # for scraping prices
        return symbol #google format is used in database for HKG 

    def scrapeCompanies( self ):
        self.db.clearCompanies()
        url = "http://www.hkex.com.hk/eng/market/sec_tradinfo/stockcode/eisdeqty_pf.htm"
        content = self.webCache.get(url)

        symbols = re.compile("""WidCoID=0(.*?)\&amp""").findall(content)
        names = re.compile("""target.*?>(.*?)<""").findall(content)

        #add HKG: to the beginning of the symbols
        symbols = map(lambda s: "HKG:%s"%s, symbols)

        industries = []
        for symbol in symbols:
            try:
                industries.append(self.scrapeIndustryForSymbol(symbol))
            except:
                industries.append("N/A")

            
        for i in range(0, len(symbols)):
            symbol, name, industry = \
                    symbols[i], names[i], industries[i]
            self.db.addCompany( symbol, name, industry )

    def convertToBloombergFormat(self, symbol):
        return symbol[4:].lstrip("0")+":HK"

    def scrapeIndustryForSymbol(self, symbol):
        url = "http://www.bloomberg.com/quote/%s"
        content = self.webCache.get(url%self.convertToBloombergFormat(symbol))

        return re.findall(""">Industry:</span>\n*.*>(.*)""", content)[0]


    def scrapeFinancialsForSymbol( self, symbol ):
        date = int(time.time())

        def checkPresence( page, pattern ):
            for line in page:
                if line.find(pattern) != -1:
                    return True

            return False

        def extractRow( soup, text ):
            def byname(tag):
                return str(tag.string).rstrip() == text and tag.name == 'td'

            tag = soup.find(byname)
            contents = []
            while tag:
                tag = tag.findNextSibling('td')
                if tag == None: break
                contents.append(str(tag.find(text=True)))
            return moneyToNumber(contents)

        def moneyToNumber( arr ):
            ret = []
            for a in arr:
                if a == '-':
                    ret.append(0)
                else:
                    ret.append(int(float(a.replace(",", "")) * 1000 ))

            return ret

        def extractDates( lines ):
            values = []
            expr = re.compile(r"""(\d\d\d\d-\d\d-\d\d)""")
            for line in lines:
                m = expr.search(line)
                if m:
                    values.append( m.group(0) )
                else:
                    values.append("")

            return values

        def findLinesLike( page, pattern ):
            lines = []
            skipped = -1
            pattern = re.compile(pattern)
            for line in page:
                if pattern.search(line):
                    lines.append( line )
                    skipped = 0
                elif skipped >= 0:
                    skipped += 1
                    if skipped >= 5:
                        break
            return lines

        print "Scraping financials for %s" % symbol

        # retrieve the web page
        url = "http://www.google.com/finance?q=%s&fstype=ii" % self.convertToGoogleFormat(symbol)
        page = self.webCache.get( url )
        soup = BeautifulSoup(page)
        page = page.split('\n')
        quarterlyPage = soup.find( "div", { "id" : "incinterimdiv" } )
        annualPage = soup.find( "div", { "id" : "incannualdiv" } )

        qstr = str(quarterlyPage).split('\n')
        astr = str(annualPage).split('\n')

        # Look for "In Millions of". If not there, error!
        if not checkPresence( page, "In Thousands of" ):
            print >>sys.stderr, "While processing %s could not find 'In Thousands of' at %s" % (symbol, url)
            return False

        # Set multiplier to 1000 (one thousand)
        multiplier = 1000

        # build array of all lines like "3 months Ending"
        quarterlyDates = extractDates(findLinesLike( qstr, """\d+ (months|weeks) ending""" ))

        # Build array of all lines like "12 months Ending"
        annualDates = extractDates(findLinesLike( astr, """\d+ (months|weeks) ending""" ))

        # Look for td containing "Total Revenue"
        # Extract all td elements in siblings that contain only a number

        # Build table for revenue
        quarterlyRevenue = extractRow( quarterlyPage, "Turnover" )
        annualRevenue = extractRow( annualPage, "Turnover" )

        # Build table for ";Diluted EPS Normalized EPS&"
        quarterlyEPS = extractRow( quarterlyPage, "Diluted EPS (HKD)" )
        annualEPS = extractRow( annualPage, "Diluted EPS (HKD)" )

        for i in range( len(quarterlyRevenue) ):
            self.db.setFinancials( symbol, "QuarterlyRevenue", quarterlyDates[i],
                    quarterlyRevenue[i] * multiplier )
            self.db.setFinancials( symbol, "QuarterlyEPS", quarterlyDates[i],
                    quarterlyEPS[i] )

        for i in range( len(annualRevenue) ):
            self.db.setFinancials( symbol, "AnnualRevenue", annualDates[i],
                    annualRevenue[i] * multiplier )
            self.db.setFinancials( symbol, "AnnualEPS", annualDates[i],
                    annualEPS[i] )

#keys should be composed of upper case letters only:
STOCKEXCHANGE_CLASSES = {"TSX":Pleco_TSX,
                        "NASDAQ":Pleco_NASDAQ,
                        "NYSE":Pleco_NYSE,
                        "HKG":Pleco_HKG
                        } 


#Change path to the exported excel document
xlsfilename = "pleco.xls"
def SetXlsFilename(filename):
    global xlsfilename 
    xlsfilename = filename
def GetXlsFilename():
    return xlsfilename


def run(arguments=sys.argv):
    def parsestockexchange(argumentstring):
        returnlist = []
        try:
            stockexchange=argumentstring.split("@")[1]
        except IndexError:
            print >> sys.stderr, "Invalid Syntax. See: pleco.py --companies@ALL --financials@NASDAQ,NYSE"
            return returnlist
            
        if stockexchange.upper() == "ALL":
            return STOCKEXCHANGE_CLASSES.values()
        thislist = stockexchange.split(",")
        for name in thislist:
            if name.upper() in STOCKEXCHANGE_CLASSES.keys():
                returnlist.append(STOCKEXCHANGE_CLASSES[name.upper()])
            else:
                print >> sys.stderr, "Ignoring unknown stock exchange name: %s"%name

        return returnlist

    for i in range(1, len(arguments)):
        if arguments[i].find("--companies") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.scrapeCompanies()
                    stockexchangeobject.exit()
        elif arguments[i].find("--prices") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.scrapePrices()
                    stockexchangeobject.exit()
        elif arguments[i].find("--financials") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.scrapeFinancials()
                    stockexchangeobject.exit()
        elif arguments[i].find('--extra') != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.addExtraInfo()
                    stockexchangeobject.exit()
        elif arguments[i].find("--all") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.scrapeCompanies()
                    stockexchangeobject.scrapeFinancials()
                    stockexchangeobject.scrapePrices()
                    stockexchangeobject.addExtraInfo()
                    stockexchangeobject.exit()
        elif arguments[i].find("--test") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeclass.addPE("tse:g")
                    stockexchangeclass.exit()
        elif arguments[i].find("--process") != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.process()
                    stockexchangeobject.exit()
        elif arguments[i].find('--excelexport') != -1:
            stockexchanges = parsestockexchange(arguments[i])
            if stockexchanges != False:
                filename = GetXlsFilename()
                exporter = excelexporter.ExcelExporter(filename)
                for stockexchangeclass in stockexchanges:
                    stockexchangeobject = stockexchangeclass()
                    stockexchangeobject.exportToExcel(exporter)
                    stockexchangeobject.exit()
        else:
            print >> sys.stderr, "Ignoring unknown argument: %s"%arguments[i]

if __name__ == "__main__":
    run()

