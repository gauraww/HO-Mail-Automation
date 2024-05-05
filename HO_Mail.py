from tkinter import filedialog, messagebox, Label, StringVar, Radiobutton, Button
from datetime import datetime, timedelta
from win32com.client import Dispatch
import tkinter as tk

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

        self.attachment_label = Label(master, text="No file attached")
        self.attachment_label.grid(row=1, column=0, columnspan=4)

        self.attachment_button = Button(master, text="Attach File", command=self.attach_file)
        self.attachment_button.grid(row=2, column=0, columnspan=2)

        self.send_button = Button(master, text="Send Email", command=self.send_email)
        self.send_button.grid(row=2, column=2, columnspan=2)


    def attach_file(self):
        self.attachment = filedialog.askopenfilename()
        if self.attachment:
            self.attachment_label.config(text=(self.attachment))

    def send_email(self):
        if not self.attachment:
            messagebox.showerror("Error", "Please attach a file before sending the email.")
            return

        outlook = Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        shift = self.shift_var.get()
        if shift == "Night Shift":
            # Calculate previous day date
            date = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
        else:
            date = datetime.now().strftime('%d/%m/%Y')

        subject = f"HO - {shift} Handover from STORAGE | {date}"
        body = (
            "<p>Hi Team,</p>"
            f"<p>Please find the attached {shift} HO for {date} from Storage Team.</p>"
        )

        mail.Subject = subject

        # Get the default signature
        inspector = mail.GetInspector
        word_editor = inspector.WordEditor
        signature = mail.HTMLBody

        # Append signature to the body
        mail.HTMLBody = body + signature
        mail.Attachments.Add(self.attachment)

        # Add recipient email addresses here
        mail.To = "gidcind_vpc_storage@dxc.com"
        cc_list = ["anoop.chandrahasa@dxc.com", "ifthikhar-ali.khan@dxc.com", "r.abhilash@dxc.com", "gunasekaran.s2@dxc.com"]
        for cc_email in cc_list:
            recipient = mail.Recipients.Add(cc_email)
            recipient.Type = 2  # 2 represents the value for CC recipient
            recipient.Resolve()
        
        mail.Display()

        self.master.destroy()

def main():
    root = tk.Tk()
    app = EmailApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
