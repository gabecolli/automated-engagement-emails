
import openpyxl
from mail_send import generate_email




workbook = openpyxl.load_workbook("workbook.xlsx")
sh = workbook.active



list_of_originators = []
#min_row=1, min_col=1, max_row=10, max_col=16
list_of_rmots = []
for row in sh.iter_rows(): 
    for cell in row:  
        if cell.column  == 1:
            list_of_rmots.append(cell.value)


for row in sh.iter_rows(): 
    for cell in row: 
        if cell.column  == 15:
            list_of_originators.append(cell.value)

if len(list_of_rmots) == len(list_of_originators):
    total_requests = len(list_of_rmots)

pre_dict = {}
for n in range(0,total_requests):
     
    generate_email(list_of_rmots[n],list_of_originators[n] + "@microsoft.com")




"""note"""

       



            











   



    




        


      








'''
column 1 is the RMOT or tracing number
column 3 is the offering
column 12 is the start date format is dd-mmm-yyyy
column 8 is the country
column 15 is the originator
'''





# count = input("How many windwows would you like to review? ")
# review_count = int(count)
# num = 7

# for rmot in list_of_rmots:
    
#    while num == 7:
#     num -= 1
#     request = "https://esxp.microsoft.com/#/supportdelivery/requestdetails/"
#     webbrowser.open_new(f"{request}{rmot}")
#             # review_count -= 1
#             # if review_count == 0:
#             #     keep_going = False

        
        
        
            


