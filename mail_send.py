









# def generate_email(request_num,origin,customer,name,total_requests):






















#     for n in range(0,total_requests):
#         today = dt.now()
#         time = today.time()
#         hour = time.strftime("%H")

#         if int(hour) < 12:
#             greeting = "Good morning"
#         elif int(hour) >= 12:
#             greeting = "Good afternoon"


#         outlook = win32.Dispatch('outlook.application')
#         mail = outlook.CreateItem(0)
#         mail.Subject = 'Engineer for request ' + request_num[n]
#         mail.To = origin[n] + "795@outlook.com"
    


#         email_text = f'''{greeting},<br><br>
#         Did you still require an engineer for {customer[n]}?<br><br>
#         Very Respectfully,<br><br>
#         {name}'''


#         mail.HTMLBody = email_text
        
#         mail.Send()

