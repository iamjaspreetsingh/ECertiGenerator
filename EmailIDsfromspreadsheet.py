import PIL
import openpyxl as xw

# to send SAME mails at once taking all emails from Column B and adding delimiter ','
# now just mailing via gmail.com

wbOne = xw.load_workbook('responses.xlsx')
sht1 = wbOne['Sheet1']

for registrants in range(1, sht1.max_row):
	# column B
	name = str(sht1['B' + str(registrants)].value)
	print(name+",")



