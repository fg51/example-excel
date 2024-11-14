from pathlib import Path

import win32com.client


def main() -> None:
    try:
        if win32com.client.GetObject(Class="Excel.Application"):
            print("Close all Excel applications!")
            input("please press the 'ENTER' key to exit.")
            raise RuntimeError("Close all Excel applications!")
    except win32com.client.pywintypes.com_error:
        pass

    p = Path("log.csv")

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(p.resolve())

        ws = wb.Worksheets[0]

        ws.Rows(1).EntireColumn.AutoFit()

        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = 1
        ws.PageSetup.FitToPagesWide = 1

        wb.PrintOut()

    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    main()
