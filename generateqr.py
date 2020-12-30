import qrcode
import xlrd
import sys
import os
def makeqr(ProuductName,weight,AgentNumber,pincode,qualitymark,temp,qron,dateofshipping,boxsize,mobilenum):
	qr=qrcode.QRCode(   # size of QR
		version=1,
		box_size=10,
		border=5
		)
   # list off fields in QR
	data= ProuductName #imput
	qr.add_data("Product name: ")
	qr.add_data(data)
	qr.add_data("\n")
	
	data= weight #imput
	qr.add_data('Weight in KG: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	data= AgentNumber #imput
	qr.add_data('Agent Number: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	data= qualitymark #imput
	qr.add_data('Quality Mark: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	data= temp #imput
	qr.add_data('Temp. Maintain: ' )
	qr.add_data(data)
	qr.add_data("\n")
	
	
	data= qron #imput
	qr.add_data('QR is placed on: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	data= dateofshipping #imput
	qr.add_data('date of shipping: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	data= boxsize #imput
	qr.add_data('Box Size: ')
	qr.add_data(data)
	qr.add_data("\n")
	

	data= mobilenum #imput
	qr.add_data('Mobile number: ')
	qr.add_data(data)
	qr.add_data("\n")
	
	
	qr.make(fit=True)
	img=qr.make_image(fill="black",back_color="white")
	img.save(ProuductName+"qr.png")      #i edit QR image name 
	
if __name__ == "__main__":
    error_list = []
    error_count = 0

    os.chdir(os.path.dirname(os.path.abspath((sys.argv[0]))))  #read later

    # Read data from an excel sheet from row 2
    Book = xlrd.open_workbook('...\\List.xlsx')  # Change the path if needed 
    WorkSheet = Book.sheet_by_name('Sheet1')
    
    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1
        # taking data from sheet 
        ProuductName = WorkSheet.cell_value( row, 0 )
        
        weight = WorkSheet.cell_value( row, 1 )
        
        AgentNumber = WorkSheet.cell_value( row, 2 )
        
        pincode = WorkSheet.cell_value( row, 3 )
        
        qualitymark = WorkSheet.cell_value( row, 4 )
        
        temp = WorkSheet.cell_value( row, 5 )
        
        qron = WorkSheet.cell_value( row, 6 )
        
        dateofshipping = WorkSheet.cell_value( row, 7 )
        
        boxsize = WorkSheet.cell_value( row, 8 )
        
        mobilenum = WorkSheet.cell_value( row, 9 )

        
       
        # Make QR 
        filename = makeqr(ProuductName,weight,AgentNumber,pincode,qualitymark,temp,qron,dateofshipping,boxsize,mobilenum )
        
        # Successfully fields in QR
        if filename != -1:
            #email_certi( filename, receiver )
            print ("Sent to " + ProuductName)
        # Add to error list
        else:
            error_list.append( ProuductName )
            error_count += 1

    # Print all failed IDs
    print (str(error_count) + " Errors- List:" + ','.join(error_list))
