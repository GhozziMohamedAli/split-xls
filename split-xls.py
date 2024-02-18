import os
import pandas as pd
import shutil
from tkinter import filedialog
import customtkinter
from tkinter import messagebox

# the appearance mode of the system
customtkinter.set_appearance_mode("System")   
 
# Sets the color of the widgets in the window 
customtkinter.set_default_color_theme("dark-blue")    
 
# Dimensions of the window
appWidth, appHeight = 600, 700

# App Class
class App(customtkinter.CTk):
    # The layout of the window will be written
    # in the init function itself
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
         # Sets the title of the window to "App"
        self.title("Excel Split ")   
        # Sets the dimensions of the window to 600x700
        self.geometry(f"{appWidth}x{appHeight}")    

        # Create a button to open the file dialog
        self.select_button = customtkinter.CTkButton(self, text="Select File",font=('Inter',14) ,command=self.open_file_dialog)
        self.select_button.grid(row=0, column=0,padx=20, pady=20,sticky="ew")

        # number Label
        self.number_label=customtkinter.CTkLabel(self, text="Split Number")
        self.number_label.grid(row=1,column=0, pady=20,sticky="ew")

        self.number =customtkinter.CTkEntry(self,placeholder_text="ex:100")
        self.number.grid(row=1, column=1, pady=20,sticky="ew")

        # Generate Button
        self.generateResultsButton = customtkinter.CTkButton(self,text="Generate Results",command=self.split_excel)
        self.generateResultsButton.grid(row=5, column=1,columnspan=2, padx=20, pady=20, sticky="ew")
        

    def open_file_dialog(self):
        # Open a file dialog box
        self.file_path = filedialog.askopenfilename()
        ext = self.file_path.split('.')
        if ext[1].lower()!='xls' and ext[1].lower()!='xlsx' :
            messagebox.showerror('Error', 'Error: File must be of type Excel')
            exit()
        self.file_name = customtkinter.CTkLabel(self, text=self.file_path)
        self.file_name.grid(row=0, column=1, padx=20, pady=20,sticky="ew")


    def split_excel(self):
        df = pd.read_excel(self.file_path)
        folder_path = './output'
        shutil.rmtree(folder_path)
        
        os.mkdir('./output')
        num = int(self.number.get())
        num_split = df.shape[0]//num

        for i in range(num_split):
            start_index = i * num_split
            end_index = (i + 1) * num_split
            split_df = df.iloc[start_index:end_index]

            # Write chunk to separate file
            output_file = f"./output/output_{i}.xlsx"
            split_df.to_excel(output_file, index=False)

            # Write the remaining rows to the last file if any
        if df.shape[0] % num != 0:
            remaining_df = df.iloc[num_split * num:]
            output_file = f"./output/output_{num_split}.xlsx"
            remaining_df.to_excel(output_file, index=False)

if __name__ == "__main__":
    app = App()
    # Used to run the application
    app.mainloop()      