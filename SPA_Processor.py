import tkinter as tk
from tkinter import filedialog
import os, csv, xlsxwriter, subprocess



class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        
        self.master = master
        self.grid()
        self.create_widgets()


    def create_widgets(self):

        self.TopLabel = tk.Label(text = "Select directory of files to process-------------------------------------------------->", font = '18')
        self.TopLabel.grid(row=0, column = 0)
        self.fetch_directory = tk.Button(text = "Browse", font = '18', fg = 'blue', command = self.browse_button)
        self.fetch_directory.grid(row = 0, column = 1, sticky = tk.E)
        self.DirectoryLabel = tk.Label(text = ' ', font = '11')
        self.DirectoryLabel.grid(row=1, column = 0, columnspan='2')
        self.process_data = tk.Button(text = "Process Data", font = '18', fg = 'blue', state  ='disabled', command = self.process_data_button)
        self.process_data.grid(row=2, column=0)
        
        self.useful_info = tk.Label(text= ' ', font = "14")
        self.useful_info.grid(row=4, column = 0)
        self.open_file = tk.Button(text = "Open Created Excel doc", font = '18', fg = 'blue', state = 'disabled', command = self.open_file_button)
        self.open_file.grid(row=5, column = 0)

    def browse_button(self):
        global directory_name
        directory_name = filedialog.askdirectory()
        self.DirectoryLabel["text"] = directory_name
        self.process_data["state"] = 'normal'

    def process_data_button(self):
            #new strategy - just force open every file as a csv
            global my_processed_data
            self.LoadingLabel = tk.Label(text = "Thinking......", fg = 'red', font = '18')
            self.LoadingLabel.grid(row = 3, column = 0)
            try:
                my_processed_data = directory_name + "\\my_processed_data.xlsx"
                workbook = xlsxwriter.Workbook(directory_name + "\\my_processed_data.xlsx") 
                worksheet = workbook.add_worksheet()
                #setup basic column headers and whatnot
                self.initialize_excel_doc(workbook, worksheet)
                row = 1
                count = 0

                
            
                #Fill out info
                for filename in os.listdir(directory_name):
                    if filename.endswith(".spa"):
                        count = count + 1
                        relevant_info = self.extract_info(directory_name + "\\" + filename)
                        #strip out the file extension, keep the name
                        worksheet.write(row, 0, filename[:-4])
                        #write everything else
                        worksheet.write(row, 1, relevant_info["[\'START_FREQ"]) 
                        worksheet.write(row, 2, relevant_info["[\'STOP_FREQ"])
                        worksheet.write(row, 3, relevant_info["[\'CENTER_FREQ"])
                        worksheet.write(row, 4, relevant_info["[\'CH_PWR_WIDTH"])
                        worksheet.write(row, 5, relevant_info["[\'CH_PWR_VALUE"])
                        worksheet.write(row, 6, relevant_info["[\'MKR_SPA_FREQN0"])
                        worksheet.write(row, 7, relevant_info["[\'MKR_SPA_MAGNT0"])
                        worksheet.write(row, 8, relevant_info["[\'MKR_SPA_FREQN2"])
                        worksheet.write(row, 9, relevant_info["[\'MKR_SPA_MAGNT2"])
                        worksheet.write(row, 10, relevant_info["[\'MKR_SPA_FREQN4"])
                        worksheet.write(row, 11, relevant_info["[\'MKR_SPA_MAGNT4"])
                        worksheet.write(row, 12, relevant_info["[\'MKR_SPA_FREQN6"])
                        worksheet.write(row, 13, relevant_info["[\'MKR_SPA_MAGNT6"])
                        worksheet.write(row, 14, relevant_info["[\'MKR_SPA_FREQN8"])
                        worksheet.write(row, 15, relevant_info["[\'MKR_SPA_MAGNT8"])
                        worksheet.write(row, 16, relevant_info["[\'MKR_SPA_FREQN10"])
                        worksheet.write(row, 17, relevant_info["[\'MKR_SPA_MAGNT10"])

                        row = row + 1
                        
            
                key_words= {
                    "[\'CH_PWR_VALUE" : "3,5",
                    "[\'MKR_SPA_MAGNT0" : "6,7",
                    "[\'MKR_SPA_MAGNT2" : "8,9",
                    "[\'MKR_SPA_MAGNT4" : "10,11",
                    "[\'MKR_SPA_MAGNT6" : "12,13",
                    "[\'MKR_SPA_MAGNT8" : "14,15",
                    "[\'MKR_SPA_MAGNT10" : "16,17"
                }
                #format the excel document to only display columns with useful information
                #a columns width is set to 18 if it's got helpful info, 0 if it doesn't.
                worksheet.set_column(0, 0, 18) #FILENAME WILL ALWAYS BE DISPLAYED
                worksheet.set_column(1, 2, 0) #START FREQ, STOP FREQ SHOULD NEVER BE DISPLAYED BY DEFAULT
                
                for word in key_words:
                    the_key = str(word)
                    
                    range = key_words[the_key].split(',')
                    
                    start = range[0]
                    end = range[1]
                    
                    if self.compare_to_default(relevant_info[the_key]):
                        worksheet.set_column(int(start), int(end), 0)
                    else:
                        worksheet.set_column(int(start), int(end), 18)

                
                #need to close workbook so it saves
                workbook.close()

                
                useful_info_text = ""
                if not self.compare_to_default(relevant_info["[\'CH_PWR_VALUE"]):
                    useful_info_text = useful_info_text + "Channel power was centered on " + relevant_info["[\'CENTER_FREQ"] + ", with a width of " + relevant_info["[\'CH_PWR_WIDTH"] +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT0"]):
                    useful_info_text = useful_info_text + "Marker 1 was set to " + relevant_info["[\'MKR_SPA_FREQN0"]  +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT2"]):
                    useful_info_text = useful_info_text + "Marker 2 was set to " + relevant_info["[\'MKR_SPA_FREQN2"]  +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT4"]):
                    useful_info_text = useful_info_text + "Marker 3 was set to " + relevant_info["[\'MKR_SPA_FREQN4"]  +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT6"]):
                    useful_info_text = useful_info_text + "Marker 4 was set to " + relevant_info["[\'MKR_SPA_FREQN6"]  +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT8"]):
                    useful_info_text = useful_info_text + "Marker 5 was set to " + relevant_info["[\'MKR_SPA_FREQN8"]  +"\n"
                if not self.compare_to_default(relevant_info["[\'MKR_SPA_MAGNT10"]):
                    useful_info_text = useful_info_text + "Marker 6 was set to " + relevant_info["[\'MKR_SPA_FREQN10"]  +"\n"
                self.LoadingLabel['text'] = "Finished processing " + str(count) + " files"
                self.useful_info['text'] = useful_info_text
                self.open_file['state'] = "normal"
            
                #self.LoadingLabel['text'] = "I couldn't find any .spa files in the specified directory"
            
            except xlsxwriter.exceptions.FileCreateError:
                self.LoadingLabel['text'] = "Oops - it looks like you have an existing file called \n my_processed_data.xlsx open.  Please close it and try again"
            except UnboundLocalError:
                self.LoadingLabel['text'] = "I couldn't find any .spa files in selected directory"

            

            
    #pass a csv, extracts relevant info, returns as dictionary
    def extract_info(self, filename):
        csv_dict = {
                "[\'START_FREQ" : " ",
                "[\'STOP_FREQ" : " ",
                "[\'CENTER_FREQ" : " ",
                "[\'CH_PWR_WIDTH" : " ",
                "[\'CH_PWR_VALUE" : " ",
                "[\'MKR_SPA_FREQN0" : " ",
                "[\'MKR_SPA_FREQN2" : " ",
                "[\'MKR_SPA_FREQN4" : " ",
                "[\'MKR_SPA_FREQN6" : " ",
                "[\'MKR_SPA_FREQN8" : " ",
                "[\'MKR_SPA_FREQN10" : " ",
                "[\'MKR_SPA_MAGNT0" : " ",
                "[\'MKR_SPA_MAGNT2" : " ",
                "[\'MKR_SPA_MAGNT4" : " ",
                "[\'MKR_SPA_MAGNT6" : " ",
                "[\'MKR_SPA_MAGNT8" : " ",
                "[\'MKR_SPA_MAGNT10" : " "
            }
        with open(filename) as csv_file:
            csv_reader = csv.reader(csv_file,delimiter = '\n')
            
            for row in csv_reader:
                
                row_string = str(row)
                row_split = row_string.split('=')
                
                if row_split[0] in csv_dict and csv_dict[row_split[0]] == " ":
                    #splicing out the last two characters to deal with formatting nonsense
                    csv_dict[row_split[0]] = row_split[1][:-2]
                    
            return csv_dict

    #open the excel doc and close the tkinter window
    def open_file_button(self):
        subprocess.Popen([my_processed_data], shell = True)
        self.master.destroy()

    #help display stuff
    def compare_to_default(self, string1):
        return string1 == "0.000000"

    #initalize the excel document, fill out column headers. This will only be called if file writes successfully
    def initialize_excel_doc(self, workbook, worksheet):
        bold = workbook.add_format({'bold' : True})
        worksheet.write("A1", "FILE NAME", bold)
        worksheet.write("B1", "START FREQUENCY")
        worksheet.write("C1", "STOP FREQUENCY")
        worksheet.write("D1", "CENTER FREQUENCY")
        worksheet.write("E1", "CHANNEL WIDTH")
        worksheet.write("F1", "CHANNEL POWER")
        worksheet.write("G1", "M1 FREQ")
        worksheet.write("H1", "M1 MAG")
        worksheet.write("I1", "M2 FREQ")
        worksheet.write("J1", "M2 MAG")
        worksheet.write("K1", "M3 FREQ")
        worksheet.write("L1", "M3 MAG")
        worksheet.write("M1", "M4 FREQ")
        worksheet.write("N1", "M4 MAG")
        worksheet.write("O1", "M5 FREQ")
        worksheet.write("P1", "M5 MAG")
        worksheet.write("Q1", "M6 FREQ")
        worksheet.write("R1", "M6 MAG")
        
    





root = tk.Tk()
root.geometry("900x600")
root.title('A tool for processing .spa data')
root.resizable(height=False, width = False)
app = Application(root)

app.mainloop()

