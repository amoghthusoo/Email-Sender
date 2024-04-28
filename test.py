# import os, sys
# from kivy.resources import resource_add_path

from kivymd.app import MDApp
from kivy.lang.builder import Builder


UI = '''
MDLabel:
    text : "Hello World"
    halign : "center"
'''

class My_App(MDApp):

    # def resource_path(relative_path):

    #     try:
    #         base_path = sys._MEIPASS
    #     except Exception:
    #         base_path = os.path.abspath('.')
        
    #     return os.path.join(base_path, relative_path)
    
    def build(self):
        return Builder.load_string(UI)
    
if(__name__ == "__main__"):

    # if hasattr(sys, '_MEIPASS'):
    #     resource_add_path((os.path.join(sys._MEIPASS)))

    My_App().run()