
from win32com import client as win32
from datetime import datetime
import os

def generate_email(request_num,origin):


    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Subject = 'Engineer for request ' + request_num
    mail.To = origin
   
    mail.HTMLBody = r"""
    Good morning,<br><br>
    Did you still require an engineer for this request?<br><br>
    Very Respectfully,<br><br>
    Gabriel Colli
    """
    
    mail.Send()