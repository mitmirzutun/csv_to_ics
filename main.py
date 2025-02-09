import converter
def main():
    import tkinter.filedialog
    import os
    file_name=tkinter.filedialog.askopenfilename(title="Open file",filetypes=[("CSV Files",".csv"),("Carlender File",".ics"),("",".xlsx")])
    if file_name.endswith(".ics"):
        save_file_name = os.path.basename(file_name)
        save_file_name = ".".join(save_file_name.split(".")[:-1])+".csv"
        save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="csv",filetypes=[("CSV Files",".csv")],title="Save file as",initialfile=save_file_name)
        converter.ics_to_csv(file_name,save_file_name)
    if file_name.endswith(".xlsx"):
        save_file_name = os.path.basename(file_name)
        save_file_name = ".".join(save_file_name.split(".")[:-1])+".ics"
        save_file_name = tkinter.filedialog.asksaveasfilename(defaultextension="ics",filetypes=[("Carlender File",".ics")],title="Save file as",initialfile=save_file_name)
        converter.xls_to_ics(file_name,save_file_name)
if __name__=="__main__":
    main()