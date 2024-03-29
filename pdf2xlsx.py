import tabula, pandas as pd
import re

def main():

    pdf_path = input("""\nNew Excel file will be created in the same directory as the pdf file. 
\nPlease enter pdf file path: """)
    
    # pdf_path = "C:\Python\Python310\Programs\Calculus_Demystified.pdf"

    match = re.search(r"(.*)\\(\w+)\.pdf$",pdf_path)
    excel_path = match.group(1) + '\\' + match.group(2) + '.xlsx'
    print(f'\nExcel path: {excel_path}')

    opt = input("\nHit enter to convert all tables in the pdf file into excel sheets, enter number to convert specific page, and for multiple pages, enter the numbers separated by commas, or hyphen for a range of pages. Example: '1-5' or '1,3,5': ")
    if not opt:
        opt = 'all'
    elif opt.isdigit():
        opt = int(opt)

    dfs =  tabula.read_pdf(pdf_path, stream = True, pages = opt)
    with pd.ExcelWriter(excel_path) as wf:
        index = 1
        for df in dfs:
            df.to_excel(wf, sheet_name = f'Sheet{index}',index = False)
            index+=1
    
    print("\nDone!")

if __name__ == "__main__":
    main()
