import openpyxl
from datetime import datetime
from smd_stock import *
from smd_sales import *
from bsc_stock import *


start_time = datetime.now()


def main():
    # Get user input for which hotsheet to update
    while True:
        hotsheet = input("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")

        if hotsheet.lower() == 'smd':
            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? (everyday, holiday, both, exit): ")

                if section.lower() == 'everyday':
                    smd_stock()
                    smd_sales()
                    break
                elif section.lower() == 'holiday':
                    smd_stock_holiday()
                    smd_sales_holiday()
                    break
                elif section.lower() == 'both':
                    smd_stock()
                    smd_stock_holiday()
                    smd_sales()
                    smd_sales_holiday()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'holiday', 'both', or 'exit'.")

        # TODO: Add support for bsc
        elif hotsheet.lower() == 'bsc':
            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? (everyday, winter, notecards, spring, all, exit): ")

                if section.lower() == 'everyday':
                    bsc_stock()
                    #bsc_sales()
                    break
                elif section.lower() == 'winter':
                    bsc_stock_winter()
                    #bsc_sales_winter()
                    #break
                elif section.lower() == 'notecards':
                    pass
                    #bsc_stock_notecards()
                    #bsc_sales_notecards()
                    #break
                elif section.lower() == 'spring':
                    pass
                    #bsc_stock_spring()
                    #bsc_sales_spring()
                    #break
                elif section.lower() == 'all':
                    bsc_stock()
                    bsc_stock_winter()
                    #bsc_stock_notecards()
                    #bsc_stock_spring()
                    #bsc_sales()
                    #bsc_sales_winter()
                    #bsc_sales_notecards()
                    #bsc_sales_spring()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'winter', 'notecards', 'spring', 'all', or 'exit'.")

        # TODO: Add support for 21c
        elif hotsheet.lower() == '21c':
            pass

        elif hotsheet.lower() == 'exit':
            break
        else:
            print("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit'.")


        time_elapsed = datetime.now() - start_time
        print("Done!\nElapsed time: %s" % time_elapsed)
    



if __name__ == "__main__":
    main()
