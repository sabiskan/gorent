import gspread
import json
import requests
from datetime import datetime, timedelta
import time
import hashlib
import traceback
import string
import pytz
from woocommerce import API

credentials = {}
client = gspread.service_account_from_dict(credentials)
SHEET1w = client.open("Gorent").sheet1
SHEET1r = SHEET1w.get_all_values()

wcapi = API()

def next_available_row(sheet, cols_to_sample=2):
  cols = sheet.range(1, 1, sheet.row_count, cols_to_sample)
  return max([cell.row for cell in cols if cell.value]) + 1

MAXR = next_available_row(SHEET1w)

#MAXR = len(SHeet1r) - можно заменить начиная с def next_avaliable...

MAXRANGE = "A1:Z"+str(MAXR)
print(MAXR)
# корректное отображение MAXRANGE:
# def maxrange(max_len_row, num_row):		#твои атрибуты (max_col, MAXR)
#     if max_len_row <= 26:
#         return f"A1:{chr(ord('A') + max_len_row - 1)}{num_row}"
#     else:
#         max_len_row_new = max_len_row - 26
#         return "A1:" + chr(ord('A') + max_len_row//27 - 1) + chr(ord('A') + max_len_row_new - 1) + str(num_row)
# max_col = len(max(SHEET1r))
# MAXRANGE = maxrange(max_col, MAXR)
product_ids = []
available_data = {}
add_com_col = 15
id_dig_col = 2
id_ggsell_col = 3
id_wmcent_col = 4
rent_date_col = 10
prices_col = 11
account_name_col = 8
days_excel_dict = {7: prices_col, 14: prices_col+1, 21: prices_col+2, 28: prices_col+3}
ru_months = {1: "Янв", 2: "Февр", 3: "Марта", 4: "Апр", 5: "Мая", 6: "Июня", 7: "Июля", 8: "Авг", 9: "Сент", 10: "Окт", 11: "Нояб", 12: "Дек"}
transform_dict = dict(zip(range(1,27), string.ascii_uppercase))
def read_excel(r, c):
    cellv = SHEET1r[r-1][c-1]
    return cellv
def write_excel(r, c, val):
    SHEET1w.cell(r, c).value = val
def transcell(i,j):
    val = transform_dict[j]+str(i)
    return val+":"+val
def format_date(str1):
	try:
		str1 = datetime.strptime(str1, '%Y-%m-%d %H:%M:%S')
		return str1
	except:
		str1 = datetime.strptime(str1, '%d.%m.%Y %H:%M:%S')
		return str1

time_zone_ru = pytz.timezone('Europe/Moscow')
date_finish = datetime.now(time_zone_ru).strftime("%Y-%m-%d %H:%M:%S")
date_start = str(datetime.strptime(read_excel(1, 1), '%Y-%m-%d %H:%M:%S'))
write_excel(1, 1, date_finish)
print(read_excel(1, 1), date_finish)

for ind in range(2, MAXR+1):
	id_dig = read_excel(ind, id_dig_col)
	if id_dig is not None:
		if product_ids.count(id_dig) == 0:
			product_ids.append(id_dig)
	id_ggsell=read_excel(ind, id_ggsell_col)
	if id_ggsell is not None:
		if product_ids.count(id_ggsell) == 0:
			product_ids.append(id_ggsell)
	id_wmcent=read_excel(ind, id_wmcent_col)
	if id_wmcent is not None:
		if product_ids.count(id_wmcent) == 0:
			product_ids.append(id_wmcent)
	# rent_date_check = read_excel(ind, rent_date_col)
	# shop_product_id = read_excel(ind, prices_col+14)
	# prev_available_date = read_excel(ind, prices_col+15)
	# if shop_product_id is not None:
	# 	if rent_date_check is not None and rent_date_check != '' and rent_date_check != prev_available_date:
	# 		if shop_product_id not in available_data:
	# 			available_data[shop_product_id] = rent_date_check
	# 			SHEET1r[ind-1][prices_col+14] = rent_date_check
	# 		elif available_data[shop_product_id]!="DEL" and available_data[shop_product_id] > rent_date_check:
	# 			available_data[shop_product_id] = rent_date_check
	# 			SHEET1r[ind-1][prices_col+14] = rent_date_check
	# 			print("last date", rent_date_check)
	# 	else:
	# 		try:
	# 			available_data[shop_product_id] = "DEL"
	# 		except:
	# 			continue
# available_data = {key:val for key, val in available_data.items() if val != "DEL"}
# out_of_stock_data = list(available_data.items())
# print(out_of_stock_data)
	

hashIDs = ''
for product_id in product_ids:
	hashIDs += str(product_id)

finalhash1 = ()

finalhash2 = ()
apikey = hashlib.sha256(finalhash2)
apisign = requests.post()
apisign = json.loads(apisign.text)
token = apisign['token']


for i in range(2, MAXR+1):
	excel_date = read_excel(i, rent_date_col)
	if excel_date is None or excel_date == "":
		SHEET1r[i-1][prices_col+4] = ""
		for j in range(prices_col,prices_col+4):
			SHEET1r[i-1][j-1]=SHEET1r[i-1][j-1].replace("#", "")
	# elif datetime.strptime(excel_date, '%d.%m.%Y %H:%M:%S') < datetime.now():
	elif format_date(excel_date) < datetime.now():
		SHEET1r[i-1][prices_col+4] = "!CHANGE!"
		for j in range(prices_col,prices_col+4):
			SHEET1r[i-1][j-1]=SHEET1r[i-1][j-1].replace("#", "")
	else:
		SHEET1r[i-1][prices_col+4] = ""

# r = requests.post('https://api.digiseller.ru/api/seller-sells, json={
#   "id_seller": 213678,
#   "product_ids": product_ids,
#   "date_start": date_start,
#   "date_finish": str(date_finish),
#   "returned": 1,
#   "rows": 100,
#   "page": 1,
#   "sign": key.hexdigest()})
# print(r.text)
# allinfo = json.loads(r.text)
# print(allinfo)

r = requests.post('https://api.digiseller.ru/api/seller-sells/v2?token='+str(token), json={
  "product_ids": product_ids,
  "date_start": date_start,
  "date_finish": str(date_finish),
  "returned": 1,
  "page": 1,
  "rows": 100
  })
allinfo = json.loads(r.text)
miss_basket=[]
full_basket=[]
print(allinfo)

for row in allinfo['rows']:
	invoice_id = row['invoice_id']
	parsed_id = int(row['product_id'])
	account = row['product_entry'][row['product_entry'].find(': ')+2:row['product_entry'].find('Password:')].strip()
	date_pay = datetime.strptime(row['date_pay'], "%Y-%m-%d %H:%M:%S")
	info = json.loads(requests.get('https://api.digiseller.ru/api/purchase/info/'+str(invoice_id)+'?token='+str(token)).text)
	purchase_date = datetime.strptime(info["content"]['purchase_date'], "%d.%m.%Y %H:%M:%S")
	print(info)
	print(purchase_date)
	try:
		for col in range( id_dig_col, id_dig_col+3):
			for i in range(2, MAXR+1):
				try:
					excel_id = int(read_excel(i, col))
				except:
					excel_id = read_excel(i, col)
				excel_account = read_excel(i, account_name_col)
				if parsed_id == excel_id:
					if excel_account == account:
						print(excel_id, account)
						SHEET1r[i-1][prices_col+4+col] = ""
						print(excel_account, i)
						content_id = SHEET1r[i-1][add_com_col+2]
						id_number = SHEET1r[i-1][id_dig_col-1]
						deleted = True
						if content_id is not None and id_number is not None and content_id != '':
							options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(content_id)+"&productid="+str(id_number)+'&token='+str(token))
							jsonData = json.loads(options.text)
							if jsonData['retval']==0:
								SHEET1r[i-1][add_com_col+2] = ''
							else:
								SHEET1r[i-1][add_com_col+6] = options.text
								deleted = False
								print(excel_account, "Delete Error")
							print(1, content_id, id_number)
						content_id = SHEET1r[i-1][add_com_col+3]	
						id_number = SHEET1r[i-1][id_ggsell_col-1]
						if content_id is not None and id_number is not None and content_id != '':
							options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(content_id)+"&productid="+str(id_number)+'&token='+str(token))
							jsonData = json.loads(options.text)
							if jsonData['retval']==0:
								SHEET1r[i-1][add_com_col+3] = ''
							else:	
								SHEET1r[i-1][add_com_col+7] = options.text
								deleted = False
								print(excel_account, "Delete Error")
							print(2, content_id, id_number)
						content_id = SHEET1r[i-1][add_com_col+4]	
						id_number = SHEET1r[i-1][id_wmcent_col-1]
						if content_id is not None and id_number is not None and content_id != '':
							options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(content_id)+"&productid="+str(id_number)+'&token='+str(token))
							jsonData = json.loads(options.text)
							if jsonData['retval']==0:
								SHEET1r[i-1][add_com_col+4] = ''
							else:	
								SHEET1r[i-1][add_com_col+8] = options.text
								deleted = False
								print(excel_account, "Delete Error")
							print(3, content_id, id_number)
						if not deleted:
							SHEET1r[i-1][add_com_col-1] = 'd'

						if info['content']['options'] is None:
							miss_basket.append([purchase_date,i])
							print("test1: ",miss_basket)
							SHEET1r[i-1][prices_col+9] = 'DAYS??'
							try:
								for a in miss_basket:
									for b in full_basket:
										if a[0]==b[0]:
											SHEET1r[a[1]-1][rent_date_col-1] = b[1]
											print("test3: ",a,' :: ',b)
											miss_basket.remove(a)
							except:
								continue
							continue
						else:
							option_days = info['content']['options'][0]['user_data']
						option_days = int(option_days[:option_days.find(' ')])
						try:
							j = days_excel_dict[option_days]
							SHEET1r[i-1][j-1] = "#"+read_excel(i, j)
						except:
							SHEET1r[i-1][prices_col+5] = option_days
						end_date = str(date_pay + timedelta(days=option_days))
						SHEET1r[i-1][rent_date_col-1] = end_date
						full_basket.append([purchase_date, end_date])
						print("test2: ",full_basket, ":::", excel_id)
						print(rent_date_col, date_pay + timedelta(days=option_days))
						for a in miss_basket:
							for b in full_basket:
								if a[0]==b[0]:
									SHEET1r[a[1]-1][rent_date_col-1] = b[1]
									print("test3: ",a,' :: ',b)
									miss_basket.remove(a)
						break
	except Exception as e:
		print('Ошибка:\n', traceback.format_exc())
		input()
		continue

for i in range(1, MAXR):
	# Delete all p short descr
	# try:
	# 	short_descr = wcapi.get("products/"+str(SHEET1r[i][prices_col+13])).json()['short_description']
	# 	sdput_del = short_descr[:short_descr.find("<p")]
	# 	data ={"short_description": sdput_del}
	# 	wcapi.put("products/"+str(SHEET1r[i][prices_col+13]), data).json()
	# except:
	# 	print("Not in gorent.shop")
	if SHEET1r[i][add_com_col-1] is not None:
		if SHEET1r[i][add_com_col-1].find('a') !=-1:	
			add_request = requests.post('https://api.digiseller.ru/api/product/content/add/text?token='+str(token), json={
				"product_id":int(SHEET1r[i][id_dig_col-1]),
				"content":[
					{
						"value":"Login: "+str(SHEET1r[i][account_name_col-1])+" Password: "+str(SHEET1r[i][account_name_col-2])+"\n\n"+str(SHEET1r[i][id_dig_col-2])
					}
				]
			})
			add_request = json.loads(add_request.text)
			try:
				SHEET1r[i][add_com_col+2] = add_request['content'][0]['content_id']
				SHEET1r[i][add_com_col] = ''
				SHEET1r[i][rent_date_col-1] = ''
				short_descr = wcapi.get("products/"+str(SHEET1r[i][prices_col+13])).json()['short_description']
				sdput_del = short_descr[:short_descr.find("<p")]
				data ={"short_description": sdput_del}
				wcapi.put("products/"+str(SHEET1r[i][prices_col+13]), data).json()
			except:
				print(i," Not in gorent.shop")
				input()

		if SHEET1r[i][add_com_col-1].find('b') !=-1:	
			add_request = requests.post('https://api.digiseller.ru/api/product/content/add/text?token='+str(token), json={
				"product_id":int(SHEET1r[i][id_ggsell_col-1]),
				"content":[
					{
						"value":"Login: "+str(SHEET1r[i][account_name_col-1])+" Password: "+str(SHEET1r[i][account_name_col-2])+"\n\n"+str(SHEET1r[i][id_dig_col-2])
					}
				]
			})
			add_request = json.loads(add_request.text)
			SHEET1r[i][add_com_col+3] = add_request['content'][0]['content_id']
			SHEET1r[i][add_com_col] = ''
			SHEET1r[i][rent_date_col-1] = ''
		if SHEET1r[i][add_com_col-1].find('c') !=-1:	
			add_request = requests.post('https://api.digiseller.ru/api/product/content/add/text?token='+str(token), json={
				"product_id":int(SHEET1r[i][id_wmcent_col-1]),
				"content":[
					{
						"value":"Login: "+str(SHEET1r[i][account_name_col-1])+" Password: "+str(SHEET1r[i][account_name_col-2])+"\n\n"+str(SHEET1r[i][id_dig_col-2])
					}
				]
			})
			add_request = json.loads(add_request.text)
			SHEET1r[i][add_com_col+4] = add_request['content'][0]['content_id']
			SHEET1r[i][add_com_col] = ''
			SHEET1r[i][rent_date_col-1] = ''
		deleted = True
		if SHEET1r[i][add_com_col-1].find('d') !=-1:
			if SHEET1r[i][add_com_col+2] is not None and SHEET1r[i][id_dig_col-1] is not None:
				options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(SHEET1r[i][add_com_col+2])+'&productid='+str(SHEET1r[i][id_dig_col-1])+"&token="+str(token))
				jsonData = json.loads(options.text)
				if jsonData['retval']==0:	
					SHEET1r[i][add_com_col+2] = ''
				else:	
					SHEET1r[i][add_com_col+6] = options.text
					deleted = False
			if SHEET1r[i][add_com_col+3] is not None and SHEET1r[i][id_ggsell_col-1] is not None:
				options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(SHEET1r[i][add_com_col+3])+'&productid='+str(SHEET1r[i][id_ggsell_col-1])+"&token="+str(token))
				jsonData = json.loads(options.text)
				if jsonData['retval']==0:	
					SHEET1r[i][add_com_col+3] = ''
				else:	
					SHEET1r[i][add_com_col+7] = options.text
					deleted = False
			if SHEET1r[i][add_com_col+4] is not None and SHEET1r[i][id_wmcent_col-1] is not None:
				options = requests.get('https://api.digiseller.ru/api/product/content/delete?contentid='+str(SHEET1r[i][add_com_col+4])+'&productid='+str(SHEET1r[i][id_wmcent_col-1])+"&token="+str(token))
				jsonData = json.loads(options.text)
				if jsonData['retval']==0:	
					SHEET1r[i][add_com_col+4] = ''
				else:	
					SHEET1r[i][add_com_col+8] = options.text
					deleted = False
		if SHEET1r[i][add_com_col-1].find('u') !=-1:
			print("UU")
			id_dig = SHEET1r[i][id_dig_col-1]
			if id_dig is not None:
				id_dig = int(id_dig)
				if product_ids.count(id_dig) == 0:
					product_ids.append(id_dig)
					options = requests.get('https://api.digiseller.ru/api/products/info?product_id='+str(id_dig), headers={'Accept': 'application/json'})
					options = json.loads(options.text)
					base_price = int(options['product']['prices']['wmr'][:options['product']['prices']['wmr'].find(' ')])
					for col in options['product']['options'][0]['variants']:
						text = col['modify']
						option_days = col['text']
						option_days = int(option_days[:option_days.find(' ')])
						excel_days = days_excel_dict[option_days]
						if text == '':
							SHEET1r[i][excel_days-1] = base_price
							continue
						else:
							SHEET1r[i][excel_days-1] = base_price + int(text[1:text.find(' ')])
				else:
					for col in options['product']['options'][0]['variants']:
						text = col['modify']
						option_days = col['text']
						option_days = int(option_days[:option_days.find(' ')])
						excel_days = days_excel_dict[option_days]
						if text == '':
							SHEET1r[i][excel_days-1] = base_price
							continue
						else:
						 	SHEET1r[i][excel_days-1] = base_price + int(text[1:text.find(' ')])
		if SHEET1r[i][add_com_col-1].find('p') !=-1:
			print(SHEET1r[i][prices_col-1], i)
			for ind_id in range(3):
				if SHEET1r[i][id_ggsell_col-ind_id] is not None and SHEET1r[i][id_ggsell_col-ind_id] != '':
					add_request = requests.post('https://api.digiseller.ru/api/product/edit/uniquefixed/'+str(SHEET1r[i][id_ggsell_col-ind_id])+'?token='+str(token), json={
						"price": {
		        	"price": SHEET1r[i][prices_col-1],
		        	"currency": "RUB"
		    		}
					})
					add_request = json.loads(add_request.text)
			print(add_request)

	if deleted:
		SHEET1r[i][add_com_col-1] = ''
	else:
		SHEET1r[i][add_com_col-1] = 'd'
	rent_date_check = read_excel(i-1, rent_date_col)
	shop_product_id = read_excel(i-1, prices_col+14)
	prev_available_date = read_excel(i-1, prices_col+15)

	if shop_product_id is not None:
		if rent_date_check is not None and rent_date_check != '' :
			if rent_date_check == prev_available_date:
				available_data[shop_product_id] = "DEL"
			elif shop_product_id not in available_data:
				available_data[shop_product_id] = [rent_date_check, i]
				# SHEET1r[i-2][prices_col+14] = rent_date_check
			elif available_data[shop_product_id]!="DEL" and format_date(available_data[shop_product_id][0]) > format_date(rent_date_check) and rent_date_check != prev_available_date:
				# print("prev date: ", available_data[shop_product_id][0] ,"last date ", rent_date_check)
				available_data[shop_product_id] = [rent_date_check, i]
				# SHEET1r[i-2][prices_col+14] = rent_date_check
		else:
			try:
				available_data[shop_product_id] = "DEL"
			except:
				continue
available_data = {key:val for key, val in available_data.items() if val != "DEL"}
out_of_stock_data = list(available_data.items())
print(out_of_stock_data)

for duo in out_of_stock_data:
	short_descr = wcapi.get("products/"+str(duo[0])).json()['short_description']
	html_out_of_stock = "<p style='text-align:center;font-weight:900;font-size:20px'>Будет доступен: "+str(format_date(duo[1][0]).day)+' '+ru_months[format_date(duo[1][0]).month]+"</p>"
	sdput_clear = short_descr[:short_descr.find("<p")]
	sdput = sdput_clear+html_out_of_stock
	SHEET1r[duo[1][1]-2][prices_col+14] = duo[1][0] 
	if short_descr.find(html_out_of_stock) != -1:
		print("same")
		continue
	data ={"short_description": sdput}
	wcapi.put("products/"+str(duo[0]), data).json()

input()


tet = False
while tet == False:	
	try:
		SHEET1r[0][0] = str(date_finish)
		SHEET1w.update(MAXRANGE, SHEET1r, value_input_option='USER_ENTERED')
		tet = True
		print("fine")
	except:
		print("error")
		continue