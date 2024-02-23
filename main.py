from flask import Flask, render_template, request
import openpyxl
from openpyxl.styles import Font
from openpyxl import load_workbook, Workbook
from datetime import date,datetime

import os

app = Flask(__name__)

# Initialize ws as a global variable
ws = None

#  ------------FUNCTIONS FOR COPYING AND DELETING----------------------
def copy_row(part_data,target_sheet, quantity):

    # source_row = source_sheet[row_number]
    target_max = target_sheet.max_row+1
    target_column = target_sheet.max_column+1
    count = 0
    for cell in part_data:
        count+=1
        
        target_sheet.cell(row=target_max, column=count, value=cell)
    
    target_sheet.cell(row=target_max, column=target_column-2, value=quantity)
    print("printing this")
    print(target_sheet.cell(row=target_max, column=target_column-2))



def del_row(ws, row_nummber):
    # Your existing del_row function
    ws.delete_rows(row_nummber)

    return render_template('index.html', result_message=None, sheet=ws)



#---------------------------------------CONVERTING TO WORDS----------------------------------------------------

def convert_to_words(num):  
    if num == 0:  
        return "Zero"  
    ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]  
    tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]  
    teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]  
    words = ""
    if num>= 1000:  
        words += ones[num // 1000] + " thousand "  
    num %= 1000  
    if num>= 100:  
        words += ones[num // 100] + " hundred "  
    num %= 100  
    if num>= 10 and num<= 19:  
        words += teens[num - 10] + " "  
    # num = 0  
    elif num>= 20:
        words += tens[num // 10] + " "  
    num %= 10  
    if num>= 1 and num<= 9:  
        words += ones[num] + " "  
    return words.strip() 

#--------------------------------------FUNCTION FOR CREATING A NEW QUOTE-----------------------------------------

def new_quote():
    global file
    print("inside creating quote function-------------------------------")
    d = str(date.today())
    today =  d.replace("-","")
    print(today)
    t = datetime.now()
    time_str = str(t.strftime("%H:%M"))
    time = time_str.replace(':','')
    print(time)
    f = today+time
    n = 2

    file = f[n:]+".xlsx"

    wb = Workbook()
    wb.save(file)

    sheet1 = wb['Sheet']
    sheet1.title = 'R'
    wb.save(file)
    
    ws = wb.active
    l = ['PART NUMBER', 'ITEMP DESCRIPTION', 'PRICE (USD)', 'QTY', 'TOTAL PRICE']

    for cell in range(1, 6):    
        ws.cell(row=1, column=cell,value= l[cell-1])
            
    wb.save(file)
    print(file)
    return file
    


#-----------------------------ALL APP ROUTES---------------------------------

#------------------------------REVISE APP ROUTE------------------------------ ----------- ---------- ------------ --------- REVISE PAGE ROUTES ------------ ------------- ---------
@app.route('/revise', methods=['POST'])
def revise():
    global file
    global filename
    global filename1
    f = request.form['file']
    file = f+".xlsx"
    wb = openpyxl.load_workbook(file)
    ws1 = wb.create_sheet()
    l = str(len(wb.sheetnames)-1)
    ws1.title = 'R'+l
    filename = file[:10]
    work = wb.sheetnames[-2]
    ws = wb[work]
    print("inside revise function")
    for row in range(1,ws.max_row+1):
        for cell in range(1,ws.max_column+1):
            value = ws.cell(row=row,column=cell).value
           
            ws1.cell(row=row, column=cell,value=value)
    wb.save(file)
  
    res = len(wb.sheetnames)
    

    filename1 = filename+"-" +wb.sheetnames[-1]

    return render_template('revise.html',filename=filename1,date=d, result_message=None, sheet=ws1)

#------------------------------SHOW TABLE TO REVISE---------------------------

@app.route('/indexrev', methods=['POST'])
def indexrev():
    global ws  # Use the global ws variable
    if request.method == 'POST':
        wb = load_workbook(file)
        res = len(wb.sheetnames)
        print(res)
        ws = wb.worksheets[res-1]
    return render_template("revise.html",filename=filename,date=d, sheet=ws)


#-------------------------------ADD PRODUCT FOR REVISE---------------------------

@app.route('/addrev', methods=['GET', 'POST'])
def addrev():
    global ws 
    global d
    wb = openpyxl.load_workbook(file)
    res = len(wb.sheetnames)
    ws = wb.worksheets[res-1]
    if request.method == 'POST':
        part_no = request.form['part_no']
        quantity = request.form['quantity']


        # Load the source workbook
        source_workbook = load_workbook("price2.xlsx", read_only=True)
        source_sheet = source_workbook.active


        part_data = get_part_data(part_no)
        # Find the row with the specified Part ID and copy it to the target sheet

        
                # Copy the row and update quantity
        copy_row(part_data, ws, quantity)

                # Calculate total price and update in the target sheet
        price_column = 'C'
        result_column = 'E'
        quan_column = 'D'

                # Ensure that the values are not None before converting to float

        quan_value = ws[quan_column + str(ws.max_row)].value
        print("this is quan value", quan_value)
        price_value = ws[price_column + str(ws.max_row)].value
        print("this is price value", price_value)

        if quan_value is not None and price_value is not None:
            ws[result_column + str(ws.max_row)] = float(price_value) * float(quan_value)
        else:
            # Handle the case where either quantity or price is None
            result_message = "Error: Quantity or Price is None."
            return render_template('index.html', result_message=result_message)       


            # Specify the output file path
        filename1 = file

            # Save the target workbook to a new file
        wb.save(filename1)
        result_message = f"Part details for ID {part_no} copied successfully."

        filename1 = filename+"-" +wb.sheetnames[-1]

        return render_template('revise.html',filename=filename1, result_message=None, sheet=ws, part_data = part_data,date= d)
    return render_template('revise.html',filename=filename1,  sheet=ws, part_data = part_data,date=d)

#------------------------------CANCEL FOR REVISION------------------------------

@app.route('/delrev', methods=['POST'])
def cancelrev():
    wb = openpyxl.load_workbook(file)
    print(file)
    sheet = wb.sheetnames[-1]
    print(sheet)
    print(wb.sheetnames[-1])
    sheetname = wb.get_sheet_by_name(sheet)
    wb.remove_sheet(sheetname)
    wb.save(file)
    # if len(wb.sheetnames) <=2:
    #     filename = file[:10] + "-" + wb.sheetnames[-1]
    # else:
    filename = file[:10] + "-" + wb.sheetnames[-1]
    return render_template('revise.html',filename=filename,date=d, result_message=None, sheet=ws)

#-------------------------------UPDATE FOR REVISION------------------------------
@app.route('/update', methods=['POST'])
def update():
    wb = openpyxl.load_workbook(file)
    res = len(wb.sheetnames)
    ws = wb.worksheets[res-1]

    part_no = request.form['part_no']
    quant = request.form['quantity']

    for cell in range(1,ws.max_row+1):
        if ws.cell(row=cell,column=1).value == part_no:
            ws.cell(row=cell,column=4,value=quant)
    wb.save(file)
    return render_template("revise.html",filename=filename1,date=d, result_message=None, sheet=ws)


#---------------------------------DELETE FOR REVISON------------------------------

@app.route('/deleterev' , methods=['POST'])
def deleterev():
    del_id = request.form['del_id']

    wb = openpyxl.load_workbook(file)
    res = len(wb.sheetnames)
    ws = wb.worksheets[res-1]

    for row_number in range(2, ws.max_row +1):
        if ws.cell(row=row_number, column=1).value == del_id:
            del_row(ws, row_number)
            wb.save(file)
            break
    return render_template('revise.html', date=d,filename = filename1,result_message=None,sheet=ws)


#------------------------------TOTAL FOR REVISION-----------------------------------
@app.route('/totalrev', methods=['POST'])
def totalrev():
    global ws
    if request.method == 'POST':
        wb = load_workbook(file)
        res = len(wb.sheetnames)
        ws = wb.worksheets[res-1]
        
        total = 0.00
        result_column = 'E'
        for cell in range(2,ws.max_row+1):
            t = float(ws.cell(row=cell,column=5).value)
            total += t

        ftotal = round(total,2)
        ws[result_column + str(ws.max_row+1)].value = ftotal

    for cell in range(1,ws.max_column):
        ws.cell(row=ws.max_row, column=cell, value=" ")
        
    ws.cell(row=ws.max_row,column=3, value="TOTAL PRICE")

    
    return render_template('revise.html',filename=filename1,date=d, sheet=ws)



#--------------------------------ADD NEW QUOTATION----------------------------- ---------------------- ------------------ ------------------- NEW QUOATATION------------- ----------------- -----

@app.route('/createquote', methods=['POST'])
def createquote():
    global file
    print("inside creating quote function-------------------------------")
    file = new_quote()
    wb = load_workbook(file)
    res= len(wb.sheetnames)
    ws = wb.worksheets[res-1]
    filename = file[:10]

    print("THIS IS FILE NAME============================",file)
    return render_template("add.html", result_message=None,date=d,filename= filename,sheet=ws)

#---------------------------------ADD NEW PRODUCT FOR QUOTE--------------------------

#----- FUNCTION FOR SEARCHING PART DESC -------

def get_part_data(part_id):
    wb = load_workbook('price2.xlsx')
    ws = wb.active
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == part_id:
            return row
    return None
    

# ----- SEARCH ROUTE FOR PART DESC -------
@app.route('/search', methods=['POST'])
def addnew():
    global part_data
    print("in search")
    part_id = request.form.get('part_no')
    print(part_id)
    part_data = get_part_data(part_id)
    print(part_data)
    return {'part_data':part_data}




#--------------------------------DELETE APP ROUTE------------------------------  -------------- ----------- ------------- MAIN PAGE ROUTES -------- ------------- ----------- -----------
@app.route('/delete', methods=['POST'])
def delete():
    global ws  # Use the global ws variable
    if request.method == 'POST':
        del_id = request.form['del_id']

        wb = openpyxl.load_workbook(file)
        ws = wb.active

        for row_number in range(2, ws.max_row + 1):
            if ws.cell(row=row_number, column=1).value == del_id:
                del_row(ws, row_number)
                wb.save(file)
                break
        return render_template('index.html',filename=filename,date=d, result_message=None,sheet=ws)
    

#----------------------------VIEWING QUOTE-----------------------------------
@app.route('/view', methods=['POST'])
def view():
    global ws  # Use the global ws variable
    global filename
    print("in view")
    print(filename)
    if request.method == 'POST':
        res = 1
        try:
            f = request.form['file']
            file = f+'.xlsx'
            wb = load_workbook(file)
            res = len(wb.sheetnames)
            ws = wb.worksheets[res-1]
            if res==1:
                filename1 = filename
            else:
                filename1 = filename+"-" +wb.sheetnames[-1]
        except:
            result_message = "No File Found"
            return render_template("index.html",result_message=result_message,filename=filename,date=d, sheet=ws)
        print("this is file", file)

    return render_template("view.html",filename=filename1,date=d, sheet=ws)


#-------------------------------ADDING NEW PRODUCT---------------------------------

#------ FUNCTION FOR SEARCHING PART DESC -------

def get_part_data(part_id):
    wb = load_workbook('price2.xlsx')
    ws = wb.active
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] == part_id:
            return row
    return None


#------------- SEARCH ROUTE ----------
@app.route('/search', methods=['POST'])
def search():
    global part_data
    print("in search")
    part_id = request.form.get('part_no')
    print(part_id)
    part_data = get_part_data(part_id)
    print(part_data)
    return {'part_data':part_data}


#------------------------------------ ROUTE TO DELETE LATEST QUOTE ------------------------------

@app.route('/deleted', methods=['POST'])

def cancel():
    print(file)
    os.remove(file)
    return render_template('index.html',filename=filename,date=d, sheet=ws)



#------------------------------------------ MAIN PAGE ROUTE -------------------------------------------------
@app.route('/', methods=['GET', 'POST'])

def add():
    global ws 
    global file
    global filename
    global d


    file = new_quote()
    print(file)
    filename = file[:10]

    t = date.today()
               
    d= t.strftime("%d-%m-%Y")
    now = datetime.now() # current date and time
    date_time = now.strftime("%I:%M %p")
    wb = openpyxl.load_workbook(file)
    ws = wb.active
  
    if request.method == 'POST':
        part_no = request.form['part_no']
        quantity = request.form['quantity']

        part_data = get_part_data(part_no)


        # Load the source workbook           mmmmv,v.                  
        source_workbook = load_workbook("price2.xlsx", read_only=True)
        source_sheet = source_workbook.active

        return render_template('index.html', result_message=None, date=d,filename=filename, sheet=ws,date_time=date_time, part_data=part_data)
    return render_template('index.html',filename=filename,date=d,date_time=date_time, sheet=ws)


#---------------------COPY DATA------------------------------------------------------------------


@app.route('/copy', methods=['GET', 'POST'])
def copy():
    global ws 
    global file
    global filename   
    global d

    print("in copy")
    wb = openpyxl.load_workbook(file)
    ws = wb.active 

    if request.method == 'POST':
        part_no = request.form['part_no']
        quantity = request.form['quantity']


        # Load the source workbook
        source_workbook = load_workbook("price2.xlsx", read_only=True)
        source_sheet = source_workbook.active

   
        found_Part = False

        part_data = get_part_data(part_no)

        print(part_data)
      
        copy_row(part_data,ws,quantity)

        #         # Calculate total price and update in the target sheet
        price_column = 'C'
        result_column = 'E'
        quan_column = 'D'

        #         # Ensure that the values are not None before converting to float

        quan_value = ws[quan_column + str(ws.max_row)].value
        print("this is quan value", quan_value)
        price = ws[price_column + str(ws.max_row)].value
        price2 = format(price,'.2f')
        ws[price_column + str(ws.max_row)].value = price2
        price_value = ws[price_column + str(ws.max_row)].value
        print("this is price value", price_value)

        if quan_value is not None and price_value is not None:
            ws[result_column + str(ws.max_row)] = float(price_value) * float(quan_value)
            res = ws[result_column + str(ws.max_row)].value
            print(res)
            res2 = format(res, '.2f')
            print(res2)
            ws[result_column + str(ws.max_row)] = res2
        else:
                    # Handle the case where either quantity or price is None
            result_message = "Error: Quantity or Price is None."
            return render_template('index.html',filename=filename,date=d, result_message=result_message)
        found_Part = True

        if found_Part:
            # Specify the output file path
            filename1 = file

            # Save the target workbook to a new file
            wb.save(filename1)
            result_message = f"Part details for ID {part_no} copied successfully."
        else:
            result_message = f"Part details for ID {part_no} not found."
        filename = file[:10]

        return render_template('index.html', result_message=None,filename=filename,date=d, sheet=ws)
    return render_template('index.html',filename=filename,date=d, sheet=ws)

#----------------------------------GETTING TOTAL PRICE-------------------------------

@app.route('/total', methods=['POST'])
def total():
    global ws
    if request.method == 'POST':
        wb = load_workbook(file)
        ws = wb.active
        total = 0.00
        result_column = 'E'
        for cell in range(2,ws.max_row+1):
            t = float(ws.cell(row=cell,column=5).value)
            total += t
            
       
        ftotal = round(total,2)
        ws[result_column + str(ws.max_row+1)].value = ftotal
    

    words = convert_to_words(int(total))
    print(words)
    

    for cell in range(1,ws.max_column):
        ws.cell(row=ws.max_row, column=cell, value=" ")
    
    ws.cell(row=ws.max_row,column=3,value=words)
    ws.cell(row=ws.max_row,column=2, value="TOTAL PRICE")

    
    return render_template('index.html',filename=filename,date=d, sheet=ws)

#------------------------------SHOWS TABLE-----------------------------------------

@app.route('/index', methods=['POST'])
def index():
    global ws  # Use the global ws variable
    if request.method == 'POST':
        wb = load_workbook(file)
        res = len(wb.sheetnames)
        ws = wb.worksheets[res-1]
    
    return render_template("index.html",filename=filename,date=d, sheet=ws)


#----------MAIN----------
if __name__ == '__main__':
    app.run(debug=True)
