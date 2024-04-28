Windows_Mode = True

from kivymd.app import MDApp
from kivymd.uix.filemanager import MDFileManager
from kivymd.uix.dialog import MDDialog
from kivymd.uix.button import MDRaisedButton
from kivy.core.window import Window
from kivy.lang import Builder
from kivy.clock import Clock
import os
import threading

import xlrd 
import xlwt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if (Windows_Mode):
    Window.size = (1000, 740)
    Window.top = 0
    Window.left = 200

UI = '''
MDScreenManager:

    MDScreen:

        name : "home"

        MDBoxLayout:

            orientation : "vertical"
            spacing : "0dp"
            padding : "0dp"
            # md_bg_color : [1, 0, 0, 1]

            MDTopAppBar:

                id : home_screen_top_app_bar
                title : "Email Sender"
                right_action_items : [["close", lambda x : app.close()]]
                md_bg_color : app.theme_color
                elevation : 3

            MDBoxLayout:

                orientation : "horizontal"
                spacing : "0dp"
                padding : "20dp"
                size_hint : (1, 0.7)
                # md_bg_color : [1, 0, 0, 1]

                MDBoxLayout:
                    orientation : "vertical"
                    spacing : "10dp"
                    padding : "0dp"
                    size_hint : (0.7, 1)
                    # md_bg_color : [1, 0, 0, 1]

                    MDTextField:

                        id : email
                        text : "tarun.kumar@curaj.ac.in"
                        hint_text: "Email"
                        mode: "rectangle"
                        line_color_focus : app.theme_color
                        hint_text_color_focus : app.theme_color
                        text_color_focus: "black"

                        size_hint_x : 0.9
                        adaptive_height : True
                        # pos_hint : {"center_x" : 0.5, "center_y" : 0.9}
                
                    MDTextField:

                        id : password
                        hint_text: "Enter your password"
                        mode: "rectangle"
                        line_color_focus : app.theme_color
                        hint_text_color_focus : app.theme_color
                        text_color_focus: "black"
                        password : True

                        size_hint_x : 0.9
                        # pos_hint : {"center_x" : 0.2, "center_y" : 0.5}
                    

                    MDBoxLayout:

                        orientation : "horizontal"
                        spacing : "0dp"
                        padding : "0dp"
                        # adaptive_height : True
                        # md_bg_color : [0.5, 1, 1, 1]

                        MDCheckbox:
                            id : password_checkbox
                            size_hint: None, None
                            size: "48dp", "48dp"
                            on_release : app.show_hide_password()
                            color_active : app.theme_color
                            pos_hint: {'center_x': .1, 'center_y': .8}
                        
                        MDLabel:
                            text : "Show Password"
                            pos_hint : {'center_x': .5, 'center_y': .8}
                    
                MDBoxLayout:

                    orientation : "vertical"
                    spacing : "0dp"
                    padding : "0dp"
                    # size_hint : (1, 0.5)
                    # md_bg_color : [1, 0, 0, 1]
                    

                    MDBoxLayout:

                        orientation : "vertical"
                        spacing : "0dp"
                        padding : "0dp"
                        # md_bg_color : [0, 1, 0, 1]
                        size_hint : (1, 0.8)


                        MDRaisedButton:

                            id : select_file
                            text : "Select File"
                            font_size : "18sp"
                            on_release : app.open_file_manager()
                            md_bg_color : app.theme_color
                            elevation : 3

                        MDLabel:
                            id : file_path

                    MDBoxLayout:

                        orientation : "vertical"
                        spacing : "0dp"
                        padding : "0dp"
                        # md_bg_color : [0, 0, 1, 1]
                        size_hint : (1, 0.4)

                        MDRaisedButton:

                            id : download_format
                            text : "Download Format"
                            font_size : "18sp"
                            on_release : app.download_format()
                            md_bg_color : app.theme_color
                            elevation : 3

                    MDBoxLayout:

                        orientation : "horizontal"
                        spacing : "0dp"
                        padding : "0dp"
                        # md_bg_color : [1, 1, 0, 1]

                        MDLabel:
                            text : "Send marks of : "
                            size_hint : (0.4, 1)
                            pos_hint : {'center_x' : 0, 'center_y' : 0.3}
                            bold : True

                        MDCheckbox:
                            id : cia_1
                            size_hint: None, None
                            size: "48dp", "48dp"
                            on_release : app.cia_selection()
                            color_active : app.theme_color
                            active : True
                            pos_hint : {'center_x' : 0, 'center_y' : 0.3}

                        MDLabel:
                            text : "CIA - I"
                            size_hint : (0.2, 1)
                            pos_hint : {'center_x' : 0, 'center_y' : 0.3}
                
                        MDCheckbox:
                            id : cia_2
                            size_hint: None, None
                            size: "48dp", "48dp"
                            on_release : app.cia_selection()
                            color_active : app.theme_color
                            active : True
                            pos_hint : {'center_x' : 0, 'center_y' : 0.3}

                        MDLabel:
                            text : "CIA - II"
                            pos_hint : {'center_x' : 0, 'center_y' : 0.3}

            MDBoxLayout:

                orientation : "vertical"
                spacing : "0dp"
                padding : "20dp"
                

                MDBoxLayout:

                    orientation : "vertical"
                    spacing : "20dp"
                    padding : "0dp"
                    # md_bg_color : [0, 1, 1, 1]


                    MDTextField:

                        id : subject
                        hint_text: "Subject"
                        mode: "rectangle"
                        line_color_focus : app.theme_color
                        hint_text_color_focus : app.theme_color
                        text_color_focus: "black"

                        text : "CSE214-OOP Lab Marks"

                    MDTextField:

                        id : message
                        hint_text: "Message"
                        mode: "rectangle"
                        line_color_focus : app.theme_color
                        hint_text_color_focus : app.theme_color
                        text_color_focus: "black"
                        multiline : True
                        size_hint : (1, 1)

                        text : "Dear {}\\n\\nYour CIA marks in CSE214-OOP Lab Marks are as follows: \\nCIA-I = {}/20 \\nCIA-II = {}/20 \\nIn case of any discrepancy, kindly contact me before 5:00PM today. \\n\\nRegards \\nDr. Tarun Kumar"
                    
                    MDBoxLayout:

                        orientation : "horizontal"
                        spacing : "20dp"
                        padding : "0dp"
                        adaptive_height : True

                        MDRaisedButton:

                            id : send
                            text : "Send"
                            font_size : "18sp"
                            pos_hint : {'center_x' : 0.5, 'center_y' : 0.5}
                            on_release : app.send()
                            md_bg_color : app.theme_color
                            
                            elevation : 3
                            # shadow_softness : 80
                            # shadow_softness_size : 2

                            pos_hint : {"center_x" : 0.04, "center_y" : 0.5}

                        MDSpinner:
                            id : progress_spinner
                            size_hint: None, None
                            size: dp(32), dp(32)
                            pos_hint: {'center_x': 1, 'center_y': .5}
                            active: False
                            color : app.theme_color
'''

class App(MDApp):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        self.theme_color = [123/255, 2/255, 144/255, 255/255]
        self.file_manager = MDFileManager(select_path = self.select_path, exit_manager = self.exit_file_manager)
        

        self.file_manager.background_color_toolbar = self.theme_color
        self.file_manager.background_color_selection_button = self.theme_color
        self.file_manager.icon_color = self.theme_color
        self.file_manager.select_directory_on_press_button = lambda x : self.file_manager.close()
        
        self.config_dict = {"error_dialog" : False, "mails_sent_dialog" : False}
        Clock.schedule_interval(self.handle_dialogs, 0.1)

    def handle_dialogs(self, dt):
  
        if(self.config_dict["error_dialog"]):
            self.config_dict["error_dialog"] = False

            dialog = MDDialog(
                title = "Error!",
                text = "Something went wrong.",
                buttons=[
                    MDRaisedButton(
                        text = "OK",
                        md_bg_color = self.theme_color,
                        on_release = lambda button : dialog.dismiss()
                    )
                ],
            )
            dialog.open()
        
        elif(self.config_dict["mails_sent_dialog"]):
            
            self.config_dict["mails_sent_dialog"] = False

            dialog = MDDialog(
                title = "Success!",
                text = "All mails sent successfully.",
                buttons=[
                    MDRaisedButton(
                        text = "OK",
                        md_bg_color = self.theme_color,
                        on_release = lambda button : dialog.dismiss()
                    )
                ],
            )
            dialog.open()

            
    def show_hide_password(self):
        
        if self.root.ids.password_checkbox.active:
            self.root.ids.password.password = False
        else:
            self.root.ids.password.password = True

    def open_file_manager(self):
        self.file_manager.show(os.path.expanduser("~"))


    def select_path(self, path):
        

        file_path = path
        try:
            workbook = xlrd.open_workbook(file_path)
            self.sheet = workbook.sheet_by_index(0)
            self.root.ids.file_path.text = path
            self.exit_file_manager()
        except:
            dialog = MDDialog(
                title = "Warning!",
                text = "Please select an .xls file.",
                buttons=[
                    MDRaisedButton(
                        text = "OK",
                        md_bg_color = self.theme_color,
                        on_release = lambda button : dialog.dismiss()
                    )
                ],
            )
            dialog.open()



    def exit_file_manager(self, *args):
        self.file_manager.close()
    
    def send_mails(self):

        self.root.ids.progress_spinner.active = True

        from_address = self.root.ids.email.text

        try:
            i = 1
            while(True):

                try:
                    to_address = str(self.sheet.cell_value(i, 0))
                    name = str(self.sheet.cell_value(i, 1))
                    
                    # CIA - I and CIA - II
                    if(self.root.ids.cia_1.active and self.root.ids.cia_2.active):
                        cia1_marks = str(self.sheet.cell_value(i, 2))
                        cia2_marks = str(self.sheet.cell_value(i, 3))

                    # CIA - I
                    elif(self.root.ids.cia_1.active and not self.root.ids.cia_2.active):
                        cia1_marks = str(self.sheet.cell_value(i, 2))

                    # CIA - II
                    elif(not self.root.ids.cia_1.active and self.root.ids.cia_2.active):
                        cia2_marks = str(self.sheet.cell_value(i, 3))
                    
                except:
                    break
                    
                # instance of MIMEMultipart
                msg = MIMEMultipart()
            
                # storing the senders email address  
                msg['From'] = self.root.ids.email.text
            
                # storing the receivers email address 
                msg['To'] = to_address
            
                # storing the subject 
                msg['Subject'] = self.root.ids.subject.text

                # string to store the body of the mail
                
                # CIA - I and CIA - II
                if(self.root.ids.cia_1.active and self.root.ids.cia_2.active):
                    body = self.root.ids.message.text.format(name, cia1_marks, cia2_marks)

                # CIA - I
                elif(self.root.ids.cia_1.active and not self.root.ids.cia_2.active):
                    body = self.root.ids.message.text.format(name, cia1_marks)

                # CIA - II
                elif(not self.root.ids.cia_1.active and self.root.ids.cia_2.active):
                    body = self.root.ids.message.text.format(name, cia2_marks)
                
                elif(not self.root.ids.cia_1.active and not self.root.ids.cia_2.active):
                    body = self.root.ids.message.text.format(name)
                
                # attach the body with the msg instance
                msg.attach(MIMEText(body, 'plain'))
            
                # creates SMTP session
                session = smtplib.SMTP('smtp.gmail.com', 587)
            
                # start TLS for security
                session.starttls()
            
                # Authentication
                session.login(from_address, self.root.ids.password.text)
            
                # Converts the Multipart msg into a string
                text = msg.as_string()
            
                # sending the mail
                session.sendmail(from_address, to_address, text)
            
                # terminating the session
                session.quit()

                i += 1
            
            self.config_dict["mails_sent_dialog"] = True

        except:
            self.config_dict["error_dialog"] = True

        self.root.ids.progress_spinner.active = False
        
    def send(self):
        
        if(len(self.root.ids.email.text) == 0 or len(self.root.ids.password.text) == 0 or len(self.root.ids.file_path.text) == 0):

            if(len(self.root.ids.email.text) == 0):
                message = "Please enter the email address."
            elif(len(self.root.ids.password.text) == 0):
                message = "Please enter the password."
            elif(len(self.root.ids.file_path.text) == 0):
                message = "Please select a file."

            dialog = MDDialog(
                title = "Warning!",
                text = message,
                buttons=[
                    MDRaisedButton(
                        text = "OK",
                        md_bg_color = self.theme_color,
                        on_release = lambda button : dialog.dismiss()
                    )
                ],
            )
            dialog.open()
            return
        
        threading.Thread(target=self.send_mails).start()

    def cia_selection(self):
        
        # CIA - I and CIA - II
        if(self.root.ids.cia_1.active and self.root.ids.cia_2.active):
            self.root.ids.message.text = "Dear {}\n\nYour CIA marks in CSE214-OOP Lab Marks are as follows:\nCIA-I = {}/20 \nCIA-II = {}/20 \nIn case of any discrepancy, kindly contact me before 5:00PM today. \n\nRegards \nDr. Tarun Kumar"

        # CIA - I
        elif(self.root.ids.cia_1.active and not self.root.ids.cia_2.active):
            self.root.ids.message.text = "Dear {}\n\nYour CIA marks in CSE214-OOP Lab Marks are as follows:\nCIA-I = {}/20\nIn case of any discrepancy, kindly contact me before 5:00PM today. \n\nRegards \nDr. Tarun Kumar"

        # CIA - II
        elif(not self.root.ids.cia_1.active and self.root.ids.cia_2.active):
            self.root.ids.message.text = "Dear {}\n\nYour CIA marks in CSE214-OOP Lab Marks are as follows:\nCIA-II = {}/20 \nIn case of any discrepancy, kindly contact me before 5:00PM today. \n\nRegards \nDr. Tarun Kumar"
        
        elif(not self.root.ids.cia_1.active and not self.root.ids.cia_2.active):
            self.root.ids.message.text = "Dear {}\n\nRegards \nDr. Tarun Kumar"

    def download_format(self):
        
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet("Sheet 1")

        sheet.write(0, 0, "Email")
        sheet.write(0, 1, "Name")
        sheet.write(0, 2, "CIA - I")
        sheet.write(0, 3, "CIA - II")

        workbook.save(os.path.join(os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), 'marks_format.xls'))) 

        dialog = MDDialog(
                title = "Success!",
                text = f"Format successfully saved to \n{os.path.join(os.path.join(os.path.join(os.path.expanduser('~'), 'Downloads'), 'marks_format.xls'))}",
                buttons=[
                    MDRaisedButton(
                        text = "OK",
                        md_bg_color = self.theme_color,
                        on_release = lambda button : dialog.dismiss()
                    )
                ],
            )
        dialog.open()

        

    def close(self):
        self.stop()

    def build(self):
        self.app = Builder.load_string(UI)
        return self.app
    
if(__name__ == "__main__"):
    app = App()
    app.run()

# tempkivymd
# wedz iupc zhhc hgak