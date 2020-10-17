#
# Copyright (c) 2020 Pearce Software Solutions. All rights reserved.
#
import sys
import csv
import stock_cache
import argparse
import datetime

#
#  This is a stand alone process to load the database 'funds' with the mutual 'fund_record'.
#
if __name__ == "__main__":

    # Set up the arguments
    parser = argparse.ArgumentParser("Load Mutual Funds Data")
    parser.add_argument("--input","-i",help="Input filename",required=True)
    parser.add_argument("--symbol","-s",help="Fund Symbol or Ticker",required=True)
    parser.add_argument("--splitdate","-p",help="Any Split Dates in for mm/dd/YYYY i.e 08/10/2018")
    parser.add_argument("--splitratio","-r",help="Required if split-date is present. Form new:old")

    args = parser.parse_args()

    # Initialize the cache
    cache = stock_cache.StockCache()

    fundsReader = csv.reader(
        open(args.input, newline=""), delimiter=",", quotechar='"'
    )

    headers = []

    if args.splitdate != None:
        if args.splitratio == None:
            print("Warning, missing argument split ratio")
            parser.print_usage()
            sys.exit(-1)

        splitDate = datetime.datetime.strptime(args.splitdate,"%m/%d/%Y")

        parts = args.splitratio.split(':')
        if len(parts) != 2:
            print("Warning, invalid ratio")
            parser.print_usage()
            sys.exit(-1)

        splitRatio = float(parts[0])/float(parts[1])
        print("Split Data {} : {}".format(splitDate,splitRatio))


    i = 0

    for row in fundsReader:

        print("{}:{}:{}".format(i,len(row),row))

        i = i + 1

        if len(row) < 5:
            continue

        if 'Date' in row:
            print("Found Header!")
            for r in row:
                headers.append(r.lower())

            print(headers)
            continue

        if len(headers) == 0:
            # headers have not been reached, any data before them is ignored.
            continue

        fundData = {}
        fundData["symbol"] = args.symbol
        for i in range(0,len(headers)):
            fundData[headers[i]] = row[i]

        print(fundData)

        # Get the Julian Date and Year
        theDate = fundData['date']
        fundDT = datetime.datetime.strptime(theDate, "%m/%d/%Y")

        if args.splitdate != None and fundDT < splitDate:
            print("Before the split! {}".format(fundData))

            for lbl in ['open', 'high', 'low', 'close'] :
                value = fundData.get(lbl)
                if value != None:
                    nLbl = 'u'+lbl
                    fundData[nLbl] = value
                    nValue = round(float(value) * splitRatio,3)
                    fundData[lbl] = str(nValue)

            print(fundData)

        fundTT = fundDT.timetuple()
        jul = str(fundTT.tm_yday)
        jDate = str(fundTT.tm_year) + jul.zfill(3)

        print("key {}:{}".format(args.symbol,jDate))
        #if len(row) < 3:
        #    continue
        record = cache.fundsDataRead(args.symbol,jDate)
        if record != None:
            print("Record exsists: {}".format(record))
            continue

        fundRecord = {}
        fundRecord["fund_record"] = fundData
        print(cache.fundsDataSave(args.symbol,jDate,fundRecord))



    print("All Done")
