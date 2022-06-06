# xlrd version is important. New versions don't work with Excel files :(
import xlrd
import time
from loguru import logger
from sys import stderr
from requests import Session
from pyuseragents import random as random_useragent

# setup logging
logger.remove()
logger.add(stderr, format="<white>{time:HH:mm:ss}</white> | <level>{level: <8}</level> | <cyan>{line}</cyan> - <blue>{message}</blue>")

POST_URL = "https://wkt.io/cms/testnet-applications"

# Your data file, where you store emails + names
loc = ("./data.xls")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)


rows, cols = sheet.nrows, sheet.ncols
logger.info(f'Rows number: {rows}, Cols number: {cols}')

if rows == 0:
    logger.error('No data to use!')
else:
    for index in range(rows):
        acc_values = sheet.row_values(index)
        email, name = acc_values[0], acc_values[1]
        if email != "":
            if name != "":
                session = Session()
                session.headers.update({
                    'user-agent': random_useragent(),
                    'accept': 'application/json, text/plain, */*',
                    'accept-language': 'ru,en;q=0.9,vi;q=0.8,es;q=0.7',
                    'origin': 'https://wkt.io',
                    'referer': 'https://wkt.io/testnet-competition',
                    'content-type': 'application/json'})

                r = session.post(POST_URL,
                                 json={
                                    "email": email,
                                    "name": name
                                 })
                if r.status_code == 200:
                    logger.success(f"Account number {index+1} registered, account id: {r.text}")
                else:
                    logger.error(f"Status code - {r.status_code}, text - {r.text}")
            else:
                logger.error(f"Index {index + 1} - Name is empty")
        else:
            logger.error(f"Index {index+1} - Email is empty")
        logger.info(f"Index {index+1}, values: {acc_values}")
        time.sleep(3)