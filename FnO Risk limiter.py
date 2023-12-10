import os, sys
from unicodedata import name
import pandas as pd
# pd.set_option('max.columns', None)
import datetime, time
import xlwings as xw
from kiteconnect import KiteConnect
from kiteconnect import exceptions
import requests
import urllib3
import ssl
import traceback
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
import requests

from time import sleep
import pyotp
pd.options.mode.chained_assignment = None

# Instance PW - password

#### INPUTS ####
MAX_LOSS_SWITCH = True
STRIKE_SL_SWITCH = True
LOT_SZ_SWITCH = True

MAX_LOSS_PCT = 0.05 # Max Loss would be 5% of the total trading capital.
STRIKE_SL_PCT = 0.05 # Each individual trade should not have SL of more than 2% of the trading capital

MIN_LOT = 1
MAX_LOT = 36

excel_integration = False

name = 'Aditya'

        
def login_in_zerodha(api_key, api_secret, user_id, user_pwd, totp_key):
    driver = uc.Chrome(version_main=114)
    driver.get(f'https://kite.trade/connect/login?api_key={api_key}&v=3')
    login_id = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="userid"]'))
    login_id.send_keys(user_id)

    pwd = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="password"]'))
    pwd.send_keys(user_pwd)

    submit = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="container"]/div/div/div[2]/form/div[4]/button'))
    submit.click()

    sleep(1)
    
    # totp
    totp = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//input[@type="text"]'))
    authkey = pyotp.TOTP(totp_key)
    totp.send_keys(authkey.now())

    # continue_btn = WebDriverWait(driver, 10).until(lambda x: x.find_element(By.XPATH, '//*[@id="container"]/div/div/div[2]/form/div[3]/button'))
    # continue_btn.click()

    sleep(5)

    url = driver.current_url
    initial_token = url.split('request_token=')[1]
    request_token = initial_token.split('&')[0]

    driver.close()

    kiteobj = KiteConnect(api_key = api_key)
    #print(request_token)
    data = kiteobj.generate_session(request_token, api_secret=api_secret)
    kiteobj.set_access_token(data['access_token'])
    access_token = data['access_token']
    with open(path,'w') as login:
        login.write(access_token)
    print('Access Token Generated!')

    return kiteobj

def position_exit(symb, qty, dir, prod, reason):
    print(f"Exiting Open Position with Market Order {dir}, Reason {reason}, for: {symb}")
    try:
        order_id = kite.place_order(variety=kite.VARIETY_REGULAR,
                                        exchange=kite.EXCHANGE_NFO, tradingsymbol= symb,
                                        transaction_type=kite.TRANSACTION_TYPE_SELL if dir=="Sell" else kite.TRANSACTION_TYPE_BUY,
                                        quantity= qty, validity=kite.VALIDITY_DAY,
                                        product=kite.PRODUCT_NRML if prod=='NRML' else kite.PRODUCT_MIS,
                                        order_type=kite.ORDER_TYPE_MARKET)

        print("Success- Order placed ID is:", order_id)
    except Exception as e:
        print(f"Exception Occurred While Exiting Open Position:{str(e)}")
        sys.exit(1)
        
def cancel_all_orders():
    orders = kite.orders()
    for order in orders:
        if order['status'] in ['TRIGGER PENDING', 'OPEN']:
            try:
                cancel_id = kite.cancel_order(variety=order['variety'], order_id=order['order_id'],
                                            parent_order_id=order['parent_order_id'])
                print('Cancelled:', order['tradingsymbol'], cancel_id)
            except Exception as e:
                print("Order Cancel Exception:", e)
                sys.exit(1)
        
if __name__ == '__main__':
    try:
        usr_id = "USER_ID"
        pw = "PASSWORD"
        t_key = "TOPT_TOKEN"
        api_k = "APIKEY"  # API_key
        api_s = "SECRET"  # API_secret
        
        today = pd.to_datetime('today').strftime('%Y-%m-%d')
        path = f"{name}_access_token_{today}.txt"
        isFile = os.path.isfile(path) 
        print('Access Token Exists:', isFile)

        if not isFile:
            print('Genrating Access Token')
            kite = login_in_zerodha(api_key=api_k, api_secret=api_s, user_id=usr_id, user_pwd=pw, totp_key=t_key)
        else:
            access_token = open(path, 'r').read()
            kite = KiteConnect(api_key=api_k)
            kite.set_access_token(access_token)
            
        if excel_integration:
            if not os.path.isfile(f"{name}Algo.xlsx"):
                print('Not exists')
                wb = xw.Book()
                try:
                    wb.sheets.add(today)
                except:
                    pass
                wb.save(f"{name}Algo.xlsx")
                wb.close()
                wb = xw.Book(f"{name}Algo.xlsx")
            else:
                wb = xw.Book(f"{name}Algo.xlsx")
                try:
                    wb.sheets.add(today)
                except:
                    pass
                wb.save(f"{name}Algo.xlsx")
            sht = wb.sheets(today)
            
        condition = False
        condition_triggered = False

        lot_size = {"NIFTY": 50,
                "BANKNIFTY": 25} # Lot Size

        while datetime.time(9, 0) > datetime.datetime.now().time():
            sleep(1)

        print("------------Algo Started------------")

        margin = kite.margins()
        capital = margin['equity']['net'] + margin['equity']['utilised']['debits']
        
        MAX_LOSS = capital * MAX_LOSS_PCT
        STRIKE_SL = capital * STRIKE_SL_PCT
        
        print(f"Capital: {capital} \n MAX LOSS: {MAX_LOSS} \n STRIKE SL: {STRIKE_SL}")
        print(f"MIN Lot Size: {MIN_LOT} \n MAX Lot Size: {MAX_LOT}")
        

    except Exception as e:
        print('Algo Start Exception', e)
        print(traceback.format_exc())

    
    try:
        while datetime.time(15, 20, 10) > datetime.datetime.now().time():
            try:
                if LOT_SZ_SWITCH:
                    orders = kite.orders()
                    if len(orders) > 0:
                        for order in orders:
                            if "BANKNIFTY" in order['tradingsymbol']:
                                order['Lots'] = abs(order['quantity'])/25 # BANKNIFTY
                            elif "NIFTY" in order['tradingsymbol']:
                                order['Lots'] = abs(order['quantity'])/50 # NIFTY
                            else:
                                order['Lots'] = None
                            
                            if (
                                 order['Lots'] is not None
                                 and (
                                     (order['Lots'] < MIN_LOT)
                                     or (order['Lots'] > MAX_LOT)
                                 )
                                 and order['status']
                                 in [
                                     'TRIGGER PENDING',
                                     'OPEN',
                                     'VALIDATION PENDING',
                                     'PUT ORDER REQUEST RECEIVED',
                                     'OPEN PENDING',
                                     'MODIFY VALIDATION PENDING',
                                 ]
                             ):    
                                try:
                                    cancel_id = kite.cancel_order(variety=order['variety'], order_id=order['order_id'],
                                                                parent_order_id=order['parent_order_id'])
                                    print('Order Cancelled - LotSizeMiss:', order['tradingsymbol'], cancel_id)
                                except Exception as e:
                                    print("Order Cancel Exception - LotSizeMiss:", e)
                                    sys.exit(1)
                
                positions = pd.DataFrame.from_dict(kite.positions()['net'])
                if not positions.empty:
                    # first filter then check for len
                    positions = positions.loc[positions['exchange']=="NFO"].reset_index(drop=True)
                if len(positions) > 0:
                    instruments_list = [f"{positions['exchange'][i]}:{positions['tradingsymbol'][i]}".upper() for i in range(len(positions))]
                    instruments_list = list(set(instruments_list))
                    
                    ltp = pd.DataFrame({k[4:]: v["last_price"] for k, v in kite.ltp([instruments_list]).items()}.items(), columns=['tradingsymbol','LTP'])
                    
                    positions = pd.merge(positions, ltp, on='tradingsymbol', how='left')
                    
                    positions.loc[positions['quantity'] == 0, 'Exit'] = True
                    positions.loc[(positions['quantity'] != 0), 'Exit'] = False
                    
                    positions['Live_PnL'] = (positions['sell_value'] - positions['buy_value']) + (positions['quantity'] * positions['LTP'] * positions['multiplier'])
                    
                    positions.loc[positions['quantity'] < 0, 'direction'] = 'Buy'
                    positions.loc[positions['quantity'] > 0, 'direction'] = 'Sell'
                    
                    positions.loc[positions['tradingsymbol'].str.contains("BANKNIFTY"), 'Lots_Sz'] = 25 # BANKNIFTY
                    positions.loc[~(positions['tradingsymbol'].str.contains("BANKNIFTY")), 'Lots_Sz'] = 50 # NIFTY
                        
                    positions['Lots'] = abs(positions['quantity'])/positions['Lots_Sz']
                    
                    if MAX_LOSS_SWITCH:
                        condition = sum(positions['Live_PnL']) < -(MAX_LOSS)
                    
                    if STRIKE_SL_SWITCH:
                        strike_pos = positions[(positions['Exit'] == False) & (positions['Live_PnL'] < -(STRIKE_SL))].copy()
                    else:
                        strike_pos = pd.DataFrame()
                    
                    if LOT_SZ_SWITCH:
                        lot_pos = positions[(positions['Exit'] == False) & (positions['Lots'] < MIN_LOT) | (positions['Lots'] > MAX_LOT)].copy()
                    else:
                        lot_pos = pd.DataFrame()
                        
                    if condition and not condition_triggered:
                        positions.loc[positions['Exit'] == False].apply(lambda row : position_exit(row['tradingsymbol'], abs(row['quantity']), row['direction'], row['product'], "MaxLossHit"), axis = 1)
                        
                        print("Cancelling all orders - MaxLossHit")
                        cancel_all_orders()
                        
                        # positions.loc[positions['Exit'] == False, 'Exit_Time'] = datetime.datetime.now()
                        # positions.loc[positions['Exit'] == False, 'Exit_Reason'] = 'SL_Hit'
                        # positions.loc[positions['Exit'] == False, 'Exit'] = True
                        print('Done for the day!')
                        condition_triggered = True
                        sleep(10)
                    elif condition and condition_triggered:
                        cancel_all_orders()
                        open_pos = positions.loc[positions['Exit'] == False]
                        if not open_pos.empty:
                            open_pos.loc[open_pos['Exit'] == False].apply(lambda row : position_exit(row['tradingsymbol'], abs(row['quantity']), row['direction'], row['product'], "MaxLossHitKill"), axis = 1)
                    else:
                        if not strike_pos.empty:
                            strike_pos.loc[strike_pos['Exit'] == False].apply(lambda row : position_exit(row['tradingsymbol'], abs(row['quantity']), row['direction'], row['product'], "StrikeLossHit"), axis = 1)
                            sleep(10) # sleep for orders to execute
                        
                        if not lot_pos.empty:
                            lot_pos.loc[(lot_pos['Lots'] < MIN_LOT), 'Extra_Lots'] = lot_pos['quantity']
                            lot_pos.loc[(lot_pos['Lots'] > MAX_LOT), 'Extra_Lots'] = (lot_pos['Lots'] - MAX_LOT)*lot_pos['Lots_Sz']
                            lot_pos['Extra_Lots'] = lot_pos['Extra_Lots'].astype(int)
                            
                            lot_pos.loc[lot_pos['Exit'] == False].apply(lambda row : position_exit(row['tradingsymbol'], abs(row['Extra_Lots']), row['direction'], row['product'], "LotSizeMiss"), axis = 1)
                            sleep(10) # sleep for orders to execute
                    
                    if excel_integration:
                        try:
                            sht["A1"].options(pd.DataFrame, header=1, index=True, expand='table').value = positions
                        except Exception as e:
                            print("Excel Error", e)
                            continue
                    # if condition: # So Excel Updates!
                    #     print('Done for the day!')
                    #     send_telegram_message('Done for the day! - MaxLossHit')
                    #     break
                    sleep(0.25)
                else:
                    sleep(1)
            
            except ssl.SSLEOFError as e:
                print('SSLEOFError Exception', e)
                sleep(1)
                continue
            except ssl.SSLError as e:
                print('SSLError Exception', e)
                sleep(1)
                continue
            except requests.exceptions.SSLError as e:
                print('SSLError Exception', e)
                sleep(1)
                continue
            except urllib3.exceptions.MaxRetryError as e:
                print('MaxRetryError Exception', e)
                sleep(1)
                continue
            except urllib3.exceptions.ReadTimeoutError as e:
                print('ReadTimeoutError Exception', e)
                sleep(1)
                continue
            except requests.exceptions.ReadTimeout as e:
                print('ReadTimeout Exception', e)
                sleep(1)
                continue
            except urllib3.exceptions.ProtocolError as e:
                print('ProtocolError Exception', e)
                sleep(1)
                continue
            except requests.exceptions.ConnectionError as e:
                print('ConnectionError Exception', e)
                sleep(1)
                continue
            except requests.exceptions.BaseHTTPError as e:
                print('BaseHTTPError Exception', e)
                sleep(1)
                continue
            except requests.exceptions.RetryError as e:
                print('RetryError Exception', e)
                sleep(1)
                continue
            except urllib3.exceptions.ConnectTimeoutError as e:
                print('ConnectTimeoutError Exception', e)
                sleep(1)
                continue
            except urllib3.exceptions.ResponseError as e:
                print('ResponseError Exception', e)
                sleep(1)
                continue
            except exceptions.NetworkException as e:
                print('NetworkException Exception', e)
                sleep(1)
                continue
            except exceptions.DataException as e:
                print('NetworkException Exception', e)
                sleep(1)
                continue
            except ConnectionResetError as e:
                print('ConnectionResetError Exception', e)
                sleep(1)
                continue
        if excel_integration:
            wb.save('NiteshAlgo.xlsx')
            wb.close()
        print("------------Algo Ended------------")
    except Exception as e:
        print('Algo Exception', e)
        print(traceback.format_exc())
    
