import hashlib
import json
import string
import traceback
from datetime import datetime, timedelta
import asyncio
import aiohttp
import gspread
import pytz
import requests
from woocommerce import API
from contextvars import ContextVar

credentials = {}
client = gspread.service_account_from_dict(credentials)
spreadsheet = client.open("Gorent")
SHEET1w = spreadsheet.sheet1
SHEET1r = SHEET1w.get_all_values()
wcapi = API()

MAXR = len(SHEET1r)

def add_col(start_index, endIndex):
    request_body = {
        "requests": [
            {
                "insertDimension": {
                    "range": {
                        "sheetId": SHEET1w.id,
                        "dimension": "COLUMNS",
                        "startIndex": start_index,
                        "endIndex": endIndex
                    },
                    "inheritFromBefore": False
                }
            }
        ]
    }

    spreadsheet.batch_update(request_body)

def maxrange(max_len_row, num_row):  # твои атрибуты (max_col, MAXR)
    if max_len_row <= 26:
        return f"A1:{chr(ord('A') + max_len_row - 1)}{num_row}"
    else:
        max_len_row_new = max_len_row - 26
        return "A1:" + chr(ord('A') + max_len_row // 27 - 1) + chr(ord('A') + max_len_row_new - 1) + str(num_row)


def find_indices(row_tags, keyword):
    return [index + 1 for index, tag in enumerate(row_tags) if keyword in tag]


max_col = len(max(SHEET1r))

MAXRANGE = maxrange(max_col, MAXR)
product_ids = []
available_data = {}
row_tags = SHEET1w.row_values(1)

marketplace_index = find_indices(row_tags, "ID_")
first_mplace_id, last_mplace_id = min(marketplace_index), max(marketplace_index)
marketplace_cont = find_indices(row_tags, "Cont_")
first_mplace_cont, last_mplace_cont = min(marketplace_cont), max(marketplace_cont)
marketplace_error = find_indices(row_tags, "Error_")
first_mplace_error, last_mplace_error = min(marketplace_error), max(marketplace_error)
rent_days = [day[0] for day in map(lambda x: (x, 'day' in x), row_tags) if day[1]]

add_mplace_chek = [category for category in [marketplace_index, marketplace_cont, marketplace_error]]

if max(add_mplace_chek, key=len) != sorted(add_mplace_chek, reverse=True, key=len)[-1]:
    max_elem = (max(add_mplace_chek, key=len), add_mplace_chek.index(max(add_mplace_chek, key=len)))
    new_elem_index = max_elem[0][-1]
    add_mplace_chek.remove(max_elem[0])
    mplace_name = row_tags[new_elem_index - 1].split('_')[1]
    for elem in add_mplace_chek:
        elem.append(elem[-1] + 1)
    add_mplace_chek.insert(max_elem[1], max_elem[0])
    prefiks = ['ID_', 'Cont_', 'Error_']
    iter_count = 0
    for category in range(len(add_mplace_chek)):
        add_elem_index = add_mplace_chek[category][-1] - 1
        add_elem = f'{prefiks[category]}{mplace_name}'
        if add_elem not in row_tags:
            row_tags.insert(add_elem_index + iter_count, add_elem)
            add_col(add_elem_index + iter_count, add_elem_index + iter_count + 1)
            SHEET1r = SHEET1w.get_all_values()
            iter_count += 1
    marketplace_index = find_indices(row_tags, "ID_")
    first_mplace_id, last_mplace_id = min(marketplace_index), max(marketplace_index)
    marketplace_cont = find_indices(row_tags, "Cont_")
    first_mplace_cont, last_mplace_cont = min(marketplace_cont), max(marketplace_cont)
    marketplace_error = find_indices(row_tags, "Error_")
    first_mplace_error, last_mplace_error = min(marketplace_error), max(marketplace_error)
    rent_days = [day[0] for day in map(lambda x: (x, 'day' in x), row_tags) if day[1]]

    max_col = len(row_tags)
    MAXR = len(SHEET1r)
    MAXRANGE_1 = maxrange(max_col, 1)
    SHEET1r = SHEET1w.get_all_values()
    SHEET1w.update(MAXRANGE_1, [[str(cell) for cell in row_tags]])
    SHEET1r = SHEET1w.get_all_values()
    max_col = len(max(SHEET1r))
    MAXRANGE = maxrange(max_col, MAXR)

for tag in enumerate(row_tags):
    globals()[tag[1]] = tag[0] + 1

ru_months = {1: "Янв", 2: "Февр", 3: "Марта", 4: "Апр", 5: "Мая", 6: "Июня", 7: "Июля", 8: "Авг", 9: "Сент", 10: "Окт",
             11: "Нояб", 12: "Дек"}

commands = string.ascii_lowercase.replace('d', '').replace('u', '').replace('p', '')
command_mplace_dict = dict(zip(commands, marketplace_index))


def read_excel(r, c):
    cellv = SHEET1r[r - 1][c - 1]
    return cellv


def write_excel(r, c, val):
    SHEET1w.cell(r, c).value = val


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

for ind in range(2, MAXR + 1):
    for ID_marketplace in marketplace_index:
        id_market = read_excel(ind, ID_marketplace)
        if id_market is not None:
            if product_ids.count(id_market) == 0:
                product_ids.append(id_market)
# rent_date_check = read_excel(ind, Date_end)
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

r = requests.post('https://api.digiseller.ru/api/seller-sells/v2?token=' + str(token), json={
    "product_ids": product_ids,
    "date_start": date_start,
    "date_finish": str(date_finish),
    "returned": 1,
    "page": 1,
    "rows": 100
})
allinfo = json.loads(r.text)
miss_basket = []
full_basket = []
print(allinfo)

counter = ContextVar('i', default=2)
gl_excel_account = ContextVar('excel_account', default='')
deleted = True

async def get_purchase_info(invoice_id, token):
    async with aiohttp.ClientSession() as session:
        async with session.get(f'https://api.digiseller.ru/api/purchase/info/{invoice_id}?token={token}') as response:
            return await response.json()


async def get_otions(content_id, id_number, token):
    async with aiohttp.ClientSession() as session:
        async with session.get(
                f'https://api.digiseller.ru/api/product/content/delete?contentid={content_id}&productid={id_number}&token={token}') as response:
            return await response.json()


async def delete_processes():
    tasks = [delete_process(index) for index in range(first_mplace_id, last_mplace_id + 1)]
    await asyncio.gather(*tasks)


async def delete_process(index):
    excel_account = gl_excel_account.get()
    i = counter.get()
    content_id = SHEET1r[i - 1][first_mplace_cont - first_mplace_id + index - 1]
    id_number = SHEET1r[i - 1][index - 1]
    if content_id is not None and id_number is not None and content_id != '':
        options = await get_otions(content_id, id_number, token)
        if options['retval'] == 0:
            SHEET1r[i - 1][first_mplace_cont - first_mplace_id + index - 1] = ''
        else:
            SHEET1r[i - 1][first_mplace_error - first_mplace_id + index - 1] = str(options)
            global deleted
            deleted = False
            print(deleted)
            print(excel_account, "Delete Error")
        print(1, content_id, id_number)


async def process_row(row, token):
    invoice_id = row['invoice_id']
    parsed_id = int(row['product_id'])
    account = row['product_entry'][row['product_entry'].find(': ') + 2:row['product_entry'].find('Password:')].strip()
    date_pay = datetime.strptime(row['date_pay'], "%Y-%m-%d %H:%M:%S")
    info = await get_purchase_info(invoice_id, token)
    purchase_date = datetime.strptime(info["content"]['purchase_date'], "%d.%m.%Y %H:%M:%S")
    print(account)
    print('info: ', info)
    print(purchase_date)
    try:
        for col in range(first_mplace_id, last_mplace_id + 1):
            for i in range(2, MAXR + 1):
                this_counter = counter.get()
                this_counter = i
                counter.set(this_counter)
                try:
                    excel_id = int(read_excel(i, col))
                except:
                    excel_id = read_excel(i, col)
                excel_account = read_excel(i, Login)
                this_excel_account = gl_excel_account.get()
                this_excel_account = excel_account
                gl_excel_account.set(this_excel_account)
                if parsed_id == excel_id:
                    if excel_account == account:
                        SHEET1r[i - 1][first_mplace_cont - first_mplace_id + col - 1] = ''
                        global deleted
                        deleted = True
                        await delete_processes()
                        if not deleted:
                            SHEET1r[i - 1][Commandos - 1] = 'd'
                        if info['content']['options'] is None:
                            miss_basket.append([purchase_date, i])
                            print("test1: ", miss_basket)
                            SHEET1r[i - 1][Debug - 1] = 'DAYS??'
                            try:
                                for a in miss_basket:
                                    for b in full_basket:
                                        if a[0] == b[0]:
                                            SHEET1r[a[1] - 1][Date_end - 1] = b[1]
                                            print("test3: ", a, ' :: ', b)
                                            miss_basket.remove(a)
                            except:
                                continue
                            continue
                        else:
                            option_days = info['content']['options'][0]['user_data']
                        option_days = option_days[:option_days.find(' ')]
                        try:
                            for elem in rent_days:
                                if option_days in elem:
                                    j = globals()[elem]
                                    SHEET1r[i - 1][j - 1] = "#" + read_excel(i, j)
                        except:
                            SHEET1r[i - 1][Comment - 1] = option_days
                        end_date = str(date_pay + timedelta(days=int(option_days)))
                        SHEET1r[i - 1][Date_end - 1] = end_date
                        full_basket.append([purchase_date, end_date])
                        print("test2: ", full_basket, ":::", excel_id)
                        print(Date_end, date_pay + timedelta(days=int(option_days)))
                        for a in miss_basket:
                            for b in full_basket:
                                if a[0] == b[0]:
                                    SHEET1r[a[1] - 1][Date_end - 1] = b[1]
                                    print("test3: ", a, ' :: ', b)
                                    miss_basket.remove(a)
                        break
    except Exception as e:
        print('Ошибка:\n', traceback.format_exc())
        input()
        return


async def process_rows(allinfo, token):
    tasks = [process_row(row, token) for row in allinfo['rows']]
    await asyncio.gather(*tasks)


asyncio.run(process_rows(allinfo, token))

i_abce = ContextVar('i', default=1)
i_a = ContextVar('i_a', default=1)
i_d = ContextVar('i_d', default=1)
i_u = ContextVar('i_u', default=1)
i_p = ContextVar('i_p', default=1)



async def upload_abce(token, prod_id, value):
    async with aiohttp.ClientSession() as session:
        async with session.post(f'https://api.digiseller.ru/api/product/content/add/text?token={token}',
                                json={
                                    "product_id": prod_id,
                                    "content": [
                                        {
                                            "value": value
                                        }
                                    ]
                                }) as response:
            return await response.json()

async def check_abce_command(command):
    i = i_abce.get()
    if SHEET1r[i][Commandos - 1].find(command) != -1:
        prod_id = int(SHEET1r[i][command_mplace_dict[command] - 1])
        login = SHEET1r[i][Login - 1]
        password = SHEET1r[i][Pass - 1]
        store_name = SHEET1r[i][first_mplace_id - 2]  # Steam, UPlay, Epic, etc
        value = f'{store_name} \n\nLogin: {login} Password: {password}'
        add_request = await upload_abce(token, prod_id, value)
        SHEET1r[i][first_mplace_cont - first_mplace_id + command_mplace_dict[command] - 1] = \
            add_request['content'][0][
                'content_id']
        SHEET1r[i][Status - 1] = ''
        SHEET1r[i][Date_end - 1] = ''

async def run_abce():
    tasks = [check_abce_command(command) for command in command_mplace_dict]
    await asyncio.gather(*tasks)

async def d_command(cont):
    i = i_d.get()
    content = SHEET1r[i][cont[1] - 1]
    prod_id = SHEET1r[i][marketplace_index[cont[0]] - 1]
    if content is not None and content != '' and prod_id is not None:
        options = await get_otions(content, prod_id, token)
        if options['retval'] == 0:
            SHEET1r[i][cont[1] - 1] = ''
        else:
            SHEET1r[i][marketplace_error[cont[0]] - 1] = str(options)
            global deleted
            deleted = False

async def run_d_command():
    tasks = [d_command(cont) for cont in enumerate(marketplace_cont)]
    await asyncio.gather(*tasks)

async def u_command_get(id_dig):
    async with aiohttp.ClientSession() as session:
        async with session.get(
                f'https://api.digiseller.ru/api/products/info?product_id={id_dig}', headers={'Accept': 'application/json'}) as response:
            return await response.json()

def set_price(options, base_price):
    i = i_u.get()
    for col in options['product']['options'][0]['variants']:
        text = col['modify']
        option_days = col['text']
        option_days = int(option_days[:option_days.find(' ')])
        for elem in rent_days:
            if str(option_days) in elem:
                excel_days = globals()[elem]
        if text == '':
            SHEET1r[i][excel_days - 1] = base_price
            continue
        else:
            SHEET1r[i][excel_days - 1] = base_price + int(text[1:text.find(' ')])

async def run_u_command(id_dig, id_dig_dict):
    if id_dig is not None:
        id_dig = int(id_dig)
        if id_dig not in id_dig_dict:
            options = await u_command_get(id_dig)
            base_price = int(
                options['product']['prices']['wmr'][:options['product']['prices']['wmr'].find(' ')])
            id_dig_dict[id_dig] = [options, base_price]
            set_price(options, base_price)
        else:
            options, base_price = id_dig_dict[id_dig]
            set_price(options, base_price)

async def p_command_post(mplace, price, token):
    async with aiohttp.ClientSession() as session:
        async with session.post(f'https://api.digiseller.ru/api/product/edit/uniquefixed/{mplace}?token={token}', json={
            "price": {
                "price": price,
                "currency": "RUB"
            }
        }) as response:
            response_text = await response.text()
            add_request = json.loads(response_text)
            return add_request

async def p_command(ind_id):
    i = i_p.get()
    price = str(SHEET1r[i][globals()[rent_days[0]] - 1])
    mplace = SHEET1r[i][ind_id - 1]
    if mplace is not None and mplace != '':
        price = str(SHEET1r[i][globals()[rent_days[0]] - 1])
        mplace = SHEET1r[i][ind_id - 1]
        await p_command_post(mplace, price, token)

async def run_p_command():
    tasks = [p_command(ind_id) for ind_id in range(first_mplace_id, last_mplace_id + 1)]
    await asyncio.gather(*tasks)

async def all_commands_check(i):
    if SHEET1r[i][Commandos - 1] is not None:
        i_abce.set(i)
        await run_abce()
        if SHEET1r[i][Commandos - 1].find('a') != -1:
            try:
                short_descr = wcapi.get("products/" + str(SHEET1r[i][GorentID - 1])).json()['short_description']
                sdput_del = short_descr[:short_descr.find("<p")]
                data = {"short_description": sdput_del}
                wcapi.put("products/" + str(SHEET1r[i][GorentID - 1]), data).json()
            except:
                print(i + 1, " Not in gorent.shop")
        deleted = True
        if SHEET1r[i][Commandos - 1].find('d') != -1:
            i_d.set(i)
            await run_d_command()
        if SHEET1r[i][Commandos - 1].find('u') != -1:
            print("UU")
            i_u.set(i)
            id_dig_dict = {}
            id_dig = SHEET1r[i][ID_plati - 1]
            await run_u_command(id_dig, id_dig_dict)
        if SHEET1r[i][Commandos - 1].find('p') != -1:
            i_p.set(i)
            print(SHEET1r[i][Commandos - 1], i)
            await run_p_command()

    if deleted:
        SHEET1r[i][Commandos - 1] = ''
    else:
        SHEET1r[i][Commandos - 1] = 'd'
    rent_date_check = read_excel(i + 1, Date_end)
    shop_product_id = read_excel(i + 1, GorentID)
    prev_available_date = read_excel(i + 1, LastDate)

    if shop_product_id is not None:
        if rent_date_check is not None and rent_date_check != '':
            if rent_date_check == prev_available_date:
                available_data[shop_product_id] = "DEL"
            elif shop_product_id not in available_data:
                available_data[shop_product_id] = [rent_date_check, i + 1]
            # SHEET1r[i-2][prices_col+14] = rent_date_check
            elif available_data[shop_product_id] != "DEL" and format_date(
                    available_data[shop_product_id][0]) > format_date(
                rent_date_check) and rent_date_check != prev_available_date:
                # print("prev date: ", available_data[shop_product_id][0] ,"last date ", rent_date_check)
                available_data[shop_product_id] = [rent_date_check, i + 1]
            # SHEET1r[i-2][prices_col+14] = rent_date_check
        else:
            try:
                available_data[shop_product_id] = "DEL"
            except:
                return

async def run_all_commands_chek():
    tasks = [all_commands_check(i) for i in range(1, MAXR)]
    await asyncio.gather(*tasks)

asyncio.run(run_all_commands_chek())

available_data = {key: val for key, val in available_data.items() if val != "DEL"}
out_of_stock_data = list(available_data.items())
print(out_of_stock_data)

for i in range(2, MAXR + 1):
    excel_date = read_excel(i, Date_end)
    if excel_date is None or excel_date == "":
        SHEET1r[i - 1][Status - 1] = ""
        for j in range(globals()[rent_days[0]], globals()[rent_days[-1]] + 1):
            SHEET1r[i - 1][j - 1] = str(SHEET1r[i - 1][j - 1]).replace("#", "")
    # elif datetime.strptime(excel_date, '%d.%m.%Y %H:%M:%S') < datetime.now():
    elif format_date(excel_date) < datetime.now():
        SHEET1r[i - 1][Status - 1] = "!CHANGE!"
        for j in range(globals()[rent_days[0]], globals()[rent_days[-1]] + 1):
            SHEET1r[i - 1][j - 1] = str(SHEET1r[i - 1][j - 1]).replace("#", "")
    else:
        SHEET1r[i - 1][Status - 1] = ""

for duo in out_of_stock_data:
    short_descr = wcapi.get("products/" + str(duo[0])).json()['short_description']
    html_out_of_stock = "<p style='text-align:center;font-weight:900;font-size:20px'>Будет доступен: " + str(
        format_date(duo[1][0]).day) + ' ' + ru_months[format_date(duo[1][0]).month] + "</p>"
    sdput_clear = short_descr[:short_descr.find("<p")]
    sdput = sdput_clear + html_out_of_stock
    SHEET1r[duo[1][1] - 1][LastDate - 1] = duo[1][0]
    if short_descr.find(html_out_of_stock) != -1:
        print("same")
        continue
    data = {"short_description": sdput}
    wcapi.put("products/" + str(duo[0]), data).json()

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
