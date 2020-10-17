#!/usr/bin/env python3
#
# Copyright (c) 2020 Pearce Software Solutions. All rights reserved.
#
import argparse
import stock_cache
import portfoliovalue as pv



if __name__ == "__main__":

    # Set up the arguments
    parser = argparse.ArgumentParser("Load Portfolio Value CSV Files")
    parser.add_argument( "--lookup", "-l", help="File containing lookups for translations", default="lookup.csv" )
    parser.add_argument( "--input", "-i", help="Portfolio Values Lookups CSV", default="portfolio_value.csv" )

    args = parser.parse_args()

    cache = stock_cache.StockCache()

    # load up the lookup table
    # lookUps = dict()
    # LoadLookup(lookupFilename, lookUps)
    #
    # load the portfolio value and lookups
    #
    # Load the Portfolio Value CSV file.  This provides the last price when it's not available
    # through iexdata.
    pValues = pv.PortfolioValue(args.input, args.lookup)
    print("Creation Date: {}",pValues.created)
    cache.LoadCacheFromJson("portfolio_value", pValues.data, date=pValues.created)