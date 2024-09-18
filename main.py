import openpyxl
from datetime import datetime
from smd_stock import *
from smd_sales import *


start_time = datetime.now()


def main():

    # Get user input for whether to update the Shannon Martin Hotsheet
    while True:
        update_smd = input("Do you want to update the Shannon Martin Hotsheet? (y/n): ")

        if update_smd.lower() == 'y':

            # Get user input for whether to update everyday product
            while True:
                everyday_stock = input("Do you want to update everyday product? (y/n): ")

                if everyday_stock.lower() == 'y':
                    smd_stock()
                    smd_sales()
                    break
                elif everyday_stock.lower() == 'n':
                    break
                else:
                    print("Invalid input. Please enter 'y' or 'n'.")

            # Get user input for whether to update holiday product
            while True:

                holiday_stock = input("Do you want to update holiday product? (y/n): ")

                if holiday_stock.lower() == 'y':
                    smd_stock_holiday()
                    smd_sales_holiday()
                    break
                elif holiday_stock.lower() == 'n':
                    break
                else:
                    print("Invalid input. Please enter 'y' or 'n'.")
            break
        elif update_smd.lower() == 'n':
            break
        else:
            print("Invalid input. Please enter 'y' or 'n'.")

    # Get user input for whether to update the Biely & Shoaf Hotsheet


    time_elapsed = datetime.now() - start_time
    print("Done!\nElapsed time: %s" % time_elapsed)
    



if __name__ == "__main__":
    main()
