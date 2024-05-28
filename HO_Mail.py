from os import startfile
import time
from tkinter import filedialog, messagebox, Label, StringVar, Radiobutton, Button
from datetime import datetime, timedelta
import openpyxl
from win32com.client import Dispatch
import tkinter as tk
from pygetwindow import getWindowsWithTitle, getAllTitles

class EmailApp:
    def __init__(self, master):
        self.master = master
        master.title("HO Email Automation")

        self.shift_var = StringVar()
        self.shift_var.set("First Shift")

        self.attachment = None

        self.shift_label = Label(master, text="Select Shift:")
        self.shift_label.grid(row=0, column=0)

        self.first_shift_button = Radiobutton(master, text="First Shift", variable=self.shift_var, value="First Shift")
        self.first_shift_button.grid(row=0, column=1)

        self.second_shift_button = Radiobutton(master, text="Second Shift", variable=self.shift_var, value="Second Shift")
        self.second_shift_button.grid(row=0, column=2)

        self.night_shift_button = Radiobutton(master, text="Night Shift", variable=self.shift_var, value="Night Shift")
        self.night_shift_button.grid(row=0, column=3)

        self.attachment_label = Label(master, text="No file attached", wraplength= 350)
        self.attachment_label.grid(row=1, column=0, columnspan=4, rowspan=2)

        self.attachment_button = Button(master, text="Attach File", command=self.attach_file)
        self.attachment_button.grid(row=3, column=0, columnspan=2)

        self.send_button = Button(master, text="Send Email", command=self.send_email)
        self.send_button.grid(row=3, column=2, columnspan=2)


    def attach_file(self):
        self.attachment = filedialog.askopenfilename()
        if self.attachment:
            self.attachment_label.config(text=(self.attachment))

    def get_date(self):
        shift = self.shift_var.get()
        if shift == "Night Shift":
            # Calculate previous day date
            date = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
        else:
            date = datetime.now().strftime('%d/%m/%Y')
        
        n = int(date[0:2])
        suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(n if n < 20 else n % 10, 'th')
        return date, str(n) + suffix
    
    def edit_excel(self, shift, dom):
        # Load the Excel file
        if self.attachment.endswith('.xlsx'):
            wb = openpyxl.load_workbook(self.attachment)
            sheet = wb.active

            # Edit the 2nd row text
            sheet['A2'] = f"{shift} Handover for {dom} {datetime.now().strftime('%B %Y')}"

            # Edit the 3rd row and column cell text
            sheet.cell(row=3, column=7, value=f"{shift} Updates")

            # Change the worksheet name
            sheet.title = f"{shift} Handover"

            # Save the changes
            wb.save(self.attachment)

            wb.close()

    def send_email(self):
        if not self.attachment:
            messagebox.showerror("Error", "Please attach a file before sending the email.")
            return

        shift = self.shift_var.get()
        date, dom = self.get_date()
        self.edit_excel(shift, dom)

        # Open the attached file after edits
        startfile(self.attachment)

        outlook = Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        subject = f'HO - {shift} Handover from "STORAGE" | {date}'
        body = (
            "<p style='font-size:11.5pt;'>Hi Team,</p>"
            f"<p style='font-size:11.5pt;'>Please find the attached {shift} HO for {date} from Storage Team.</p>"
        )

        mail.Subject = subject

        # Get the default signature
        inspector = mail.GetInspector
        word_editor = inspector.WordEditor
        signature = mail.HTMLBody

        # Add recipient email addresses here
        mail.To = "gidcind_vpc_storage@dxc.com"
        cc_list = ["anoop.chandrahasa@dxc.com", "ifthikhar-ali.khan@dxc.com", "r.abhilash@dxc.com", "gunasekaran.s2@dxc.com"]
        for cc_email in cc_list:
            recipient = mail.Recipients.Add(cc_email)
            recipient.Type = 2  # 2 represents the value for CC recipient
            recipient.Resolve()

        self.master.destroy()

         # Wait for Excel window to open and bring it to focus
        excel_window = None
        
        time.sleep(10)
          # Adjust the time as needed

        # Get the titles of all visible windows
        windows = getAllTitles()
        # Check if any window title contains "Excel"
        for window_title in windows:
            if "Excel" in window_title:
                excel_window = getWindowsWithTitle(window_title)

                if excel_window:
                    time.sleep(1)
                    excel_window[0].activate()
        
        errcount = 0
        # Check if any window title contains "Excel"
        while excel_window:
            time.sleep(10)
            # Get the titles of all visible windows
            windows = getAllTitles()
            for window_title in windows:
                if "Excel" in window_title:
                    excel_window = getWindowsWithTitle(window_title)
                    errcount = 0
                    break
                else: 
                    continue
            errcount+=1
            if errcount>2:
                excel_window = False
        
        # Append signature to the body
        mail.HTMLBody = body + signature
        time.sleep(2)
        mail.Attachments.Add(self.attachment)
        mail.Display()


def main():
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
