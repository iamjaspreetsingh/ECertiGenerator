import PIL
from PIL import ImageFont
from PIL import Image
from PIL import ImageDraw
import openpyxl as xw

wbOne = xw.load_workbook('responses.xlsx')
sht1 = wbOne['Sheet1']

font = ImageFont.truetype("./fonts/roboto.ttf", 25)
font1 = ImageFont.truetype("./fonts/roboto.ttf", 19)
for registrants in range(1, sht1.max_row):
	img = Image.open("sampleCertificate.png")
	draw = ImageDraw.Draw(img)
	name = str(sht1['A' + str(registrants)].value)
	print(name)
	xLocation = 412 - len(name) * 7
	draw.text((xLocation,279), name, (0,0,0), font=font)
	college = str(sht1['C' + str(registrants)].value)
	print(college)
	xLocation = 560 - len(college) * 8
	draw.text((xLocation,350), college, (0,0,0), font=font1)
	draw = ImageDraw.Draw(img)
	del draw
	img.save('./generatedCertis/'+name+'.png', 'PNG')
