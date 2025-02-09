import converter
def main():
    file_name=input("File name: ")
    if file_name.endswith(".ics"):
        converter.ics_to_csv(file_name,input("CSV file name: "))
    if file_name.endswith(".xlsx"):
        converter.xls_to_ics(file_name,input("ICS file name: "))
if __name__=="__main__":
    main()