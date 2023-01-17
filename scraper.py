import argparse
from main import CollectData
import time


if __name__ == "__main__":
    # time.sleep(2)
    parser = argparse.ArgumentParser(description='Money Control Scraper')

    parser.add_argument('-s', '--stocks', nargs='+',
                        dest="stocks",
                        help="List the stocks you want to scrape for details")

    parser.add_argument("-f", '--funds', nargs='+',
                        dest="funds",
                        help="List the mutual funds you want to scrape for details")


    args = parser.parse_args()

    if not args.funds and not args.stocks:
        print("Something went wrong!")
        print(parser.print_help())
        exit()


    if args.funds:
        C = CollectData(ids=args.funds)
        C.collect("funds")
    elif args.stocks:
        C = CollectData(ids=args.stocks)
    
        C.collect("stocks")