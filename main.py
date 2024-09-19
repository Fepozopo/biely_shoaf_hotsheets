import openpyxl
from datetime import datetime
from smd_stock import *
from smd_sales import *


start_time = datetime.now()


def main():
    # Get user input for which hotsheet to update
    while True:
        hotsheet = input("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")

        if hotsheet.lower() == 'smd':
            # Get user input for which season to update
            while True:
                season = input("Which season do you want to update? (everyday, holiday, both, exit): ")

                if season.lower() == 'everyday':
                    smd_stock()
                    smd_sales()
                    break
                elif season.lower() == 'holiday':
                    smd_stock_holiday()
                    smd_sales_holiday()
                    break
                elif season.lower() == 'both':
                    smd_stock()
                    smd_stock_holiday()
                    smd_sales()
                    smd_sales_holiday()
                    break
                elif season.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'holiday', 'both', or 'exit'.")

        # TODO: Add support for bsc
        elif hotsheet.lower == 'bsc':
            pass

        # TODO: Add support for 21c
        elif hotsheet.lower == '21c':
            pass

        elif hotsheet.lower() == 'exit':
            break
        else:
            print("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit'.")


        time_elapsed = datetime.now() - start_time
        print("Done!\nElapsed time: %s" % time_elapsed)
    



if __name__ == "__main__":
    main()
