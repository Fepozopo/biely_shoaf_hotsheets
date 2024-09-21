import openpyxl
from datetime import datetime
from stock import Update_Stock
from sales import Update_Sales


start_time = datetime.now()


def main():
    # HOTSHEET | SECTION | REPORT | START | SKU | ON HAND | ON PO | ON SO/BO
    smd_stock = Update_Stock("./xlsx/SMD_HOTSHEET.xlsx", "EVERYDAY", "./xlsx/SMD-IM_StockStatus.xlsx", 3, "E", "F", "I", "K")
    smd_stock_holiday = Update_Stock("./xlsx/SMD_HOTSHEET.xlsx", "HOLIDAY", "./xlsx/SMD-IM_StockStatus.xlsx", 2, "C", "D", "D", "H")
    bsc_stock = Update_Stock("./xlsx/BSC_HOTSHEET.xlsx", "Everyday", "./xlsx/BSC-IM_StockStatus.xlsx", 2, "D", "E", "F", "H")
    bsc_stock_winter = Update_Stock("./xlsx/BSC_HOTSHEET.xlsx", "Winter Holiday", "./xlsx/BSC-IM_StockStatus.xlsx", 2, "E", "F", "I", "G")
    bsc_stock_notecards = Update_Stock("./xlsx/BSC_HOTSHEET.xlsx", "A2 Notecards", "./xlsx/BSC-IM_StockStatus.xlsx", 2, "D", "F", "G", "I")
    bsc_stock_spring = Update_Stock("./xlsx/BSC_HOTSHEET.xlsx", "Spring holiday", "./xlsx/BSC-IM_StockStatus.xlsx", 2, "D", "E", "H", "J")
    c21_stock = Update_Stock("./xlsx/21C_HOTSHEET.xlsx", "EVERYDAY", "./xlsx/21C-IM_StockStatus.xlsx", 2, "C", "D", "E", "G")

    # HOTSHEET | SECTION | REPORT | START | SKU | YTD
    smd_sales = Update_Sales("./xlsx/SMD_HOTSHEET.xlsx", "EVERYDAY", "./xlsx/SMD-IM_SalesAnalysisCondensed.xlsx", 3, "E", "P")
    smd_sales_holiday = Update_Sales("./xlsx/SMD_HOTSHEET.xlsx", "HOLIDAY", "./xlsx/SMD-IM_SalesAnalysisCondensed.xlsx", 2, "C", "N")
    bsc_sales = Update_Sales("./xlsx/BSC_HOTSHEET.xlsx", "Everyday", "./xlsx/BSC-IM_SalesAnalysisCondensed.xlsx", 2, "D", "K")
    #bsc_sales_winter = Update_Sales("./xlsx/BSC_HOTSHEET.xlsx", "Winter Holiday", "./xlsx/BSC-IM_SalesAnalysisCondensed.xlsx", 2, "E", "L")
    #bsc_sales_notecards = Update_Sales("./xlsx/BSC_HOTSHEET.xlsx", "A2 Notecards", "./xlsx/BSC-IM_SalesAnalysisCondensed.xlsx", 2, "D", "L")
    bsc_sales_spring = Update_Sales("./xlsx/BSC_HOTSHEET.xlsx", "Spring holiday", "./xlsx/BSC-IM_SalesAnalysisCondensed.xlsx", 2, "D", "L")
    #c21_sales = Update_Sales("./xlsx/21C_HOTSHEET.xlsx", "EVERYDAY", "./xlsx/21C-IM_SalesAnalysisCondensed.xlsx", 2, "C", "M")
    #c21_sales_boxedcards = Update_Sales("./xlsx/21C_HOTSHEET.xlsx", "boxed card unit sales", "./xlsx/21C-IM_SalesAnalysisCondensed.xlsx", 2, "C", "H")

    # Get user input for which hotsheet to update
    while True:
        hotsheet = input("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")

        if hotsheet.lower() == 'smd':
            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? (everyday, holiday, all, exit): ")
                if section.lower() == 'everyday':
                    print('Updating everyday stock...')
                    smd_stock.update()
                    print('Updating everyday sales...')
                    smd_sales.update()
                    break
                elif section.lower() == 'holiday':
                    print('Updating holiday stock...')
                    smd_stock_holiday.update()
                    print('Updating holiday sales...')
                    smd_sales_holiday.update()
                    break
                elif section.lower() == 'all':
                    print('Updating everyday stock...')
                    smd_stock.update()
                    print('Updating holiday stock...')
                    smd_stock_holiday.update()
                    print('Updating everyday sales...')
                    smd_sales.update()
                    print('Updating holiday sales...')
                    smd_sales_holiday.update()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'holiday', 'all', or 'exit'.")

        # TODO: Finish support for bsc
        elif hotsheet.lower() == 'bsc':
            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? (everyday, winter, notecards, spring, all, exit): ")
                if section.lower() == 'everyday':
                    print('Updating everyday stock...')
                    bsc_stock.update()
                    print('Updating everyday sales...')
                    bsc_sales.update()
                    break
                elif section.lower() == 'winter':
                    print('Updating winter holiday stock...')
                    bsc_stock_winter.update()
                    #bsc_sales_winter.update()
                    break
                elif section.lower() == 'notecards':
                    pass
                    print('Updating notecards stock...')
                    bsc_stock_notecards.update()
                    #bsc_sales_notecards.update()
                    break
                elif section.lower() == 'spring':
                    pass
                    print('Updating spring holiday stock...')
                    bsc_stock_spring.update()
                    print('Updating spring holiday sales...')
                    bsc_sales_spring.update()
                    break
                elif section.lower() == 'all':
                    print('Updating everyday stock...')
                    bsc_stock.update()
                    print('Updating winter holiday stock...')
                    bsc_stock_winter.update()
                    print('Updating notecards stock...')
                    bsc_stock_notecards.update()
                    print('Updating spring holiday stock...')
                    bsc_stock_spring.update()
                    print('Updating everyday sales...')
                    bsc_sales.update()
                    #bsc_sales_winter.update()
                    #bsc_sales_notecards.update()
                    print('Updating spring holiday sales...')
                    bsc_sales_spring.update()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'winter', 'notecards', 'spring', 'all', or 'exit'.")

        # TODO: Finish support for 21c
        elif hotsheet.lower() == '21c':
            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? ('everyday', 'boxedcards', all, 'exit): ")
                if section.lower() == 'everyday':
                    print('Updating everyday stock...')
                    c21_stock.update()
                    #c21_sales.update()
                    break
                elif section.lower() == 'boxedcards':
                    print('No stock kept...')
                    #c21_sales_boxedcards.update()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'boxedcards', 'all', or 'exit'.")

        elif hotsheet.lower() == 'exit':
            break
        else:
            print("Invalid input. Please enter 'smd', 'bsc', '21c', or 'exit'.")


        time_elapsed = datetime.now() - start_time
        print("Done!\nElapsed time: %s" % time_elapsed)
    



if __name__ == "__main__":
    main()
