import tkinter.filedialog
from datetime import datetime
from stock import Update_Stock
from sales import Update_Sales


start_time = datetime.now()


def main():
    # Get user input for which hotsheet to update
    while True:
        hotsheet = input("Which hotsheet do you want to update? (smd, bsc, 21c, exit): ")

        if hotsheet.lower() == 'smd':
            # Select files
            file_smd_hotsheet = tkinter.filedialog.askopenfilename(title='Select the SMD HOTSHEET...', filetypes=[("Excel", "*.xlsx")])
            file_smd_stock_report = tkinter.filedialog.askopenfilename(title='Select the SMD Stock Report...', filetypes=[("Excel", "*.xlsx")])
            file_smd_sales_report = tkinter.filedialog.askopenfilename(title='Select the SMD Sales Report...', filetypes=[("Excel", "*.xlsx")])

            # HOTSHEET | SECTION | REPORT | START | SKU | ON HAND | ON PO | ON SO/BO
            smd_stock = Update_Stock(file_smd_hotsheet, "EVERYDAY", file_smd_stock_report, 3, "E", "F", "I", "K")
            smd_stock_holiday = Update_Stock(file_smd_hotsheet, "HOLIDAY", file_smd_stock_report, 2, "C", "D", "D", "H")
            # HOTSHEET | SECTION | REPORT | START | SKU | YTD
            smd_sales = Update_Sales(file_smd_hotsheet, "EVERYDAY", file_smd_sales_report, 3, "E", "P")
            smd_sales_holiday = Update_Sales(file_smd_hotsheet, "HOLIDAY", file_smd_sales_report, 2, "C", "N")
            
            # Check if files were selected
            if not file_smd_hotsheet or not file_smd_stock_report or not file_smd_sales_report:
                print("One or more files were not selected. Please try again.")
                continue

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

        elif hotsheet.lower() == 'bsc':
            # Select files
            file_bsc_hotsheet = tkinter.filedialog.askopenfilename(title='Select the BSC HOTSHEET...', filetypes=[("Excel", "*.xlsx")])
            file_bsc_stock_report = tkinter.filedialog.askopenfilename(title='Select the BSC Stock Report...', filetypes=[("Excel", "*.xlsx")])
            file_bsc_sales_report = tkinter.filedialog.askopenfilename(title='Select the BSC Sales Report...', filetypes=[("Excel", "*.xlsx")])

            # HOTSHEET | SECTION | REPORT | START | SKU | ON HAND | ON PO | ON SO/BO
            bsc_stock = Update_Stock(file_bsc_hotsheet, "Everyday", file_bsc_stock_report, 2, "D", "E", "F", "H")
            bsc_stock_winter = Update_Stock(file_bsc_hotsheet, "Winter Holiday", file_bsc_stock_report, 2, "E", "F", "I", "G")
            bsc_stock_notecards = Update_Stock(file_bsc_hotsheet, "A2 Notecards", file_bsc_stock_report, 2, "D", "F", "G", "I")
            bsc_stock_spring = Update_Stock(file_bsc_hotsheet, "Spring holiday", file_bsc_stock_report, 2, "D", "E", "H", "J")
            # HOTSHEET | SECTION | REPORT | START | SKU | YTD
            bsc_sales = Update_Sales(file_bsc_hotsheet, "Everyday", file_bsc_sales_report, 2, "D", "K")
            bsc_sales_winter = Update_Sales(file_bsc_hotsheet, "Winter Holiday", file_bsc_sales_report, 2, "E", "L")
            bsc_sales_winter_kits = Update_Sales(file_bsc_hotsheet, "Winter Holiday Kits", file_bsc_sales_report, 2, "E", "L")
            bsc_sales_notecards = Update_Sales(file_bsc_hotsheet, "A2 Notecards", file_bsc_sales_report, 2, "D", "L")
            bsc_sales_spring = Update_Sales(file_bsc_hotsheet, "Spring holiday", file_bsc_sales_report, 2, "D", "L")

            # Check if files were selected
            if not file_bsc_hotsheet or not file_bsc_stock_report or not file_bsc_sales_report:
                print("One or more files were not selected. Please try again.")
                continue

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
                    print('Updating winter holiday sales...')
                    bsc_sales_winter.update()
                    bsc_sales_winter_kits.update()
                    break
                elif section.lower() == 'notecards':
                    pass
                    print('Updating notecards stock...')
                    bsc_stock_notecards.update()
                    print('Updating notecards sales...')
                    bsc_sales_notecards.update()
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
                    print('Updating winter holiday sales...')
                    bsc_sales_winter.update()
                    bsc_sales_winter_kits.update()
                    print('Updating notecards sales...')
                    bsc_sales_notecards.update()
                    print('Updating spring holiday sales...')
                    bsc_sales_spring.update()
                    break
                elif section.lower() == 'exit':
                    break
                else:
                    print("Invalid input. Please enter 'everyday', 'winter', 'notecards', 'spring', 'all', or 'exit'.")

        elif hotsheet.lower() == '21c':
            # Select files
            file_c21_hotsheet = tkinter.filedialog.askopenfilename(title='Select the C21 HOTSHEET...', filetypes=[("Excel", "*.xlsx")])
            file_c21_stock_report = tkinter.filedialog.askopenfilename(title='Select the C21 Stock Report...', filetypes=[("Excel", "*.xlsx")])
            file_c21_sales_report = tkinter.filedialog.askopenfilename(title='Select the C21 Sales Report...', filetypes=[("Excel", "*.xlsx")])

            # HOTSHEET | SECTION | REPORT | START | SKU | ON HAND | ON PO | ON SO/BO
            c21_stock = Update_Stock(file_c21_hotsheet, "EVERYDAY", file_c21_stock_report, 2, "C", "D", "E", "G")
            # HOTSHEET | SECTION | REPORT | START | SKU | YTD
            c21_sales = Update_Sales(file_c21_hotsheet, "EVERYDAY", file_c21_sales_report, 2, "C", "M")
            c21_sales_boxedcards = Update_Sales(file_c21_hotsheet, "boxed card unit sales", file_c21_sales_report, 2, "C", "H")

            # Check if files were selected
            if not file_c21_hotsheet or not file_c21_stock_report or not file_c21_sales_report:
                print("One or more files were not selected. Please try again.")
                continue

            # Get user input for which sections to update
            while True:
                section = input("Which section do you want to update? ('everyday', 'boxedcards', all, 'exit): ")
                if section.lower() == 'everyday':
                    print('Updating everyday stock...')
                    c21_stock.update()
                    print('Updating everyday sales...')
                    c21_sales.update()
                    break
                elif section.lower() == 'boxedcards':
                    print('No stock kept for boxedcards...')
                    print('Updating boxedcards sales...')
                    c21_sales_boxedcards.update()
                    break
                elif section.lower() == 'all':
                    print('Updating everyday stock...')
                    c21_stock.update()
                    print('No stock kept for boxedcards...')
                    print('Updating everyday sales...')
                    c21_sales.update()
                    print('Updating boxedcards sales...')
                    c21_sales_boxedcards.update()
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
