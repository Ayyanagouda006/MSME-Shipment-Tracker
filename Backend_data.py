#!/usr/bin/env python
# coding: utf-8

# In[1]:


from pymongo import MongoClient
import pandas as pd
import ast
import json
from datetime import datetime
import logging
import os
from openpyxl import load_workbook

# Base log folder
log_folder = r'logs'
os.makedirs(log_folder, exist_ok=True)

# Helper to create a logger
def setup_logger(name, log_file, level=logging.INFO):
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Prevent adding multiple handlers if already added
    if not logger.handlers:
        handler = logging.FileHandler(log_file, mode='a')
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    
    return logger

# Create all three loggers
mongolog = setup_logger('mongolog', os.path.join(log_folder, 'mongo_data_extraction.log'))
booking_processlog = setup_logger('booking_processlog', os.path.join(log_folder, 'booking_process.log'))
comparisonlog = setup_logger('comparisonlog', os.path.join(log_folder, 'comparison.log'))

def epoch_to_date(epoch_time):
    if pd.isna(epoch_time):
        mongolog.debug("Epoch time is NaN")
        return pd.NaT
    try:
        epoch_time_sec = int(epoch_time) / 1000
        date_str = datetime.fromtimestamp(epoch_time_sec).strftime('%d-%m-%Y')
        mongolog.debug(f"Converted epoch {epoch_time} to date {date_str}")
        return date_str
    except (ValueError, TypeError) as e:
        mongolog.error(f"Error converting epoch {epoch_time}: {e}")
        return pd.NaT

def extract_date(data, target_label):
    try:
        if isinstance(data, str):
            data_list = ast.literal_eval(data)
        elif isinstance(data, list):
            data_list = data
        else:
            mongolog.debug(f"Invalid data type for extract_date: {type(data)}")
            return None

        for item in data_list:
            if isinstance(item, dict) and item.get('label') == target_label:
                mongolog.debug(f"Found label {target_label} with value {item.get('value')}")
                return item.get('value')
    except (SyntaxError, ValueError) as e:
        mongolog.error(f"Error parsing data in extract_date: {e}")
        return None
    return None

    
def extract_duty_invoice(row):
    try:
        files = row['files']
        if isinstance(files, str):
            files = ast.literal_eval(files)
        if isinstance(files, list):
            duty_invoices = [item for item in files if item.get('label') == 'Custom Duties & Taxes Invoice']
            if duty_invoices:
                invoice = duty_invoices[0]
                approved_status = invoice.get('approved', '').strip()
                mongolog.debug(f"Duty invoice status: {approved_status}")
                return pd.Series([row['createdOn'], approved_status or 'Pending'])
        return pd.Series([None, None])
    except Exception as e:
        mongolog.error(f"Error in extract_duty_invoice for row {row['_id'] if '_id' in row else 'unknown'}: {e}")
        return pd.Series([None, None])

def contains_duty_invoice(files):
    try:
        if isinstance(files, str):
            files = ast.literal_eval(files)
        result = any(item.get('label') == 'Custom Duties & Taxes Invoice' for item in files)
        mongolog.debug(f"contains_duty_invoice: {result}")
        return result
    except Exception as e:
        mongolog.error(f"Error in contains_duty_invoice: {e}")
        return False


def fetch_data():
    # MongoDB server details
    host = "65.1.22.99"  # MongoDB server IP
    port = "27017"  # MongoDB server port
    database_name = "agdb-prod2"

 
    try:
        mongolog.info("Connecting to MongoDB...")
        # Create the connection
        client = MongoClient(f'mongodb://{host}:{port}/')
        db = client[database_name]

        mongolog.info("Fetching from Bookings collection")
        bookings_collection = db["Bookings"]

        # Query to fetch only required fields
        bookings_projection = {
            "_id": 1,
            "bookingDate": 1,
            "entityId":1,
            "status": 1,
            "fba":1,
            "contract.cargoTotals.totChargeableWeight":1,
            "contract.fbaPallets":1,
            "contract.shipmentType": 1,
            "contract.shipmentScope": 1,
            "contract.origin": 1,
            "contract.finalPlaceOfDelivery":1,
            "contract.destination": 1,

        }

        # Fetch data from Bookings collection
        bookings_cursor = bookings_collection.find({}, bookings_projection)
        bookings_data = list(bookings_cursor)

        bookings = pd.DataFrame(bookings_data)
        bookings = bookings[~bookings['status'].isin(['CANCELLED', 'Cancellation Requested'])]
        bookings['bookingDate'] = pd.to_datetime(bookings['bookingDate'])
        bookings = bookings[bookings['bookingDate'] >= '2025-01-01']
        # Extract vendor IDs from nested dictionary
        bookings['shipmentType'] = bookings['contract'].apply(lambda x: x.get('shipmentType') if isinstance(x, dict) else x)
        bookings = bookings[bookings['shipmentType'] == 'LCL']
        bookings['shipmentScope'] = bookings['contract'].apply(lambda x: x.get('shipmentScope') if isinstance(x, dict) else x)
        bookings['fbaPallets'] = bookings['contract'].apply(lambda x: x.get('fbaPallets') if isinstance(x, dict) else x)
        bookings['origin'] = bookings['contract'].apply(lambda x: x.get('origin') if isinstance(x, dict) else x)
        bookings['finalPlaceOfDelivery'] = bookings['contract'].apply(lambda x: x.get('finalPlaceOfDelivery') if isinstance(x, dict) else x)
        bookings['totChargeableWeight'] = bookings['contract'].apply(lambda x: x.get('cargoTotals', {}).get('totChargeableWeight', '') if isinstance(x, dict) else '')
        bookings = bookings.drop(columns=['contract'])
        mongolog.info(f"Fetched {len(bookings)} records from Bookings")

        mongolog.info("Fetching from SHEntities collection")
        shentities_collection = db["SHEntities"]

        # Query to fetch only required fields
        shentities_projection = {
            "entityName": 1,
            "entityId":1,
            "customer.crossBorder.salesVertical":1
        }

        shentities_cursor = shentities_collection.find({}, shentities_projection)
        shentities_data = list(shentities_cursor)

        shentities = pd.DataFrame(shentities_data)
        shentities['salesVertical'] = shentities['customer'].apply(lambda x: x.get('crossBorder', {}).get('salesVertical', '') if isinstance(x, dict) else '')
        mongolog.info(f"Fetched {len(shentities)} SHEntities")

        mongolog.info("Fetching from Bookingdsr collection")
        bookingdsr_collection = db["Bookingdsr"]

        # Query to fetch only required fields
        bookingdsr_projection = {
            "_id": 1,
            "sob_pol":1,
            "gatein_pol":1,
            "hbl_number":1,
            "mbl_number":1,
            "etd_at_pol":1,
            "stuffing_confirmation":1,
            "pol_container_number":1,
            "eta_fpod":1,
            "sob_pol":1,
            "gatein_fpod":1,
            "carrier":1,
            "consolidator":1,
            "importClearance.label":1,
            "importClearance.value":1,
            "vdes.destination":1,
            "vdes.atdfrompod":1,
            "vdes.actual_delivery_date":1,
            "vdes.total_package":1,
            "last_free_date_at_fpod":1,
            "delivery_order_release":1,
            "remarks":1
        }

        bookingdsr_cursor = bookingdsr_collection.find({}, bookingdsr_projection)
        bookingdsr_data = list(bookingdsr_cursor)

        bookingdsr = pd.DataFrame(bookingdsr_data)
        bookingdsr['importClearance Date'] = bookingdsr['importClearance'].apply(lambda x: extract_date(x, 'Customs Clearance Complete'))
        bookingdsr = bookingdsr.drop(columns=['importClearance'])
        mongolog.info(f"Fetched {len(bookingdsr)} Bookingdsr records")

        mongolog.info("Fetching from Myactions collection")
        Myactions_collection = db["Myactions"]
        # Query to fetch only required fields
        Myactions_projection = {
            "_id.bookingNum": 1,
            "actionName": 1,
            "files":1,
            "createdOn":1
        }

        Myactions_cursor = Myactions_collection.find({}, Myactions_projection)
        Myactions_data = list(Myactions_cursor)

        Myactions = pd.DataFrame(Myactions_data)
        Myactions = Myactions[Myactions['actionName'] == 'Invoice Acceptance']
        Myactions['_id'] = Myactions['_id'].apply(lambda x: x.get('bookingNum') if isinstance(x, dict) else x)
        Myactions['createdOn'] = Myactions['createdOn'].apply(epoch_to_date)
        Myactions = Myactions[Myactions['files'].apply(contains_duty_invoice)]
        mongolog.info(f"Filtered to {len(Myactions)} Invoice Acceptance actions")
        
        mongolog.info("Fetching from Addressdetails collection")
        Addressdetails_collection = db["Addressdetails"]
        # Query to fetch only required fields
        Addressdetails_projection = {
            "_id": 1,
            "fbacode":1
        }

        Addressdetails_cursor = Addressdetails_collection.find({}, Addressdetails_projection)
        Addressdetails_data = list(Addressdetails_cursor)
        Addressdetails = pd.DataFrame(Addressdetails_data)
        mongolog.info(f"Fetched {len(Addressdetails)} Addressdetails records")
        
        mongolog.info("Fetching from Agusers collection")
        Agusers_collection = db["Agusers"]
        Agusers_projection = {"email":1}
        Agusers_cursor = Agusers_collection.find({}, Agusers_projection)
        Agusers_data = list(Agusers_cursor)
        Agusers = pd.DataFrame(Agusers_data)
        Agusers = Agusers[Agusers['email'].str.contains('@agraga.com', na=False)]
        mongolog.info(f"Fetched {len(Agusers)} Agusers records")

        bookings = pd.merge(bookings, shentities[['entityId', 'entityName', 'salesVertical']], on='entityId', how='left')
        bookings = bookings[bookings['salesVertical'] == 'MSME']
        bookings = pd.merge(bookings, bookingdsr, on='_id', how='left')
        bookings = pd.merge(bookings, Myactions[['_id', 'files', 'createdOn']], on='_id', how='left')
        bookings[['Duty Invoice', 'Duty Invoice Status']] = bookings.apply(extract_duty_invoice, axis=1)
        mongolog.info("Merged all datasets successfully")
        

        return bookings, shentities ,bookingdsr, Myactions ,Addressdetails, Agusers
 
    except Exception as e:
        mongolog.error(f"Error in fetch_data: {e}")
        mongolog.info('*'*100)
        return None, None, None, None, None
        
    finally:
        # Close the connection
        try:
            client.close()
            mongolog.info("MongoDB connection closed.")
            mongolog.info('*'*100)
        except:
            mongolog.info("MongoDB connection was not established, so no need to close.")
            mongolog.info('*'*100)


# In[2]:


def booking_process(bookings, Addressdetails):
    booking_processlog.info("Started booking_process function")
    result_rows = []  # Store all rows

    for i, rows in bookings.iterrows():
        try:
            booking_id = rows['_id']
            booking_date = rows['bookingDate']
            status = rows['status']
            fba = rows['fba']
            shipmentType = rows['shipmentType']
            shipmentScope = rows['shipmentScope']
            fbaPallets = rows['fbaPallets']
            origin = rows['origin']
            fpod = rows['finalPlaceOfDelivery']
            totChargeableWeight = rows['totChargeableWeight']
            entityName = rows['entityName']
            duty_invoice = rows['Duty Invoice']
            duty_invoice_status = rows['Duty Invoice Status']

            container_number = rows['pol_container_number']
            gatein_pol = rows['gatein_pol']
            sob_pol = rows['sob_pol']
            hbl_number = rows['hbl_number']
            mbl_number = rows['mbl_number']
            gatein_fpod = rows['gatein_fpod']
            delivery_order_release = rows['delivery_order_release']
            remarks = rows['remarks']
            stuffing_confirmation = rows['stuffing_confirmation']
            etd_at_pol = rows['etd_at_pol']
            eta_fpod = rows['eta_fpod']
            last_free_date_at_fpod = rows['last_free_date_at_fpod']
            carrier = rows['carrier']
            consolidator = rows['consolidator']
            importClearance = rows['importClearance Date']
            vdes = rows['vdes']

            booking_processlog.info(f"Processing Booking ID: {booking_id}")

            # Ensure vdes is a list
            if isinstance(vdes, str):
                try:
                    vdes = json.loads(vdes)
                except json.JSONDecodeError:
                    booking_processlog.warning(f"JSON parse error for Booking ID {booking_id}: {vdes}")
                    vdes = []
            elif not isinstance(vdes, list):
                booking_processlog.warning(f"Unexpected vdes format for Booking ID {booking_id}: {vdes}")
                vdes = []

            if vdes:
                for des in vdes:
                    ad = des.get('destination', '')

                    if ad:
                        subAddressdetails = Addressdetails.loc[Addressdetails['_id'] == ad, 'fbacode']
                        fbacode = subAddressdetails.iloc[0] if not subAddressdetails.empty else ''
                    else:
                        booking_processlog.warning(f"Missing destination in vdes for Booking ID {booking_id}")
                        fbacode = ''

                    row_data = {
                        'Customer Name': entityName,
                        'MBL#': mbl_number,
                        'HBL#': hbl_number,
                        'Agraga Booking #': booking_id,
                        'Booking Status': status,
                        'FBA?': fba,
                        'ISF Filing': '',
                        'Stuffing Date': stuffing_confirmation,
                        'Container #': container_number,
                        'ETD': etd_at_pol,
                        'ETA': eta_fpod,
                        'SOB': sob_pol,
                        'ATA': gatein_fpod,
                        'Carrier': carrier,
                        'Consolidator': consolidator,
                        'Origin': origin,
                        'FPOD': fpod,
                        'CFS': '',
                        'Delivery Address': ad,
                        'FBA Code': fbacode,
                        'Freight Broker': '',
                        'Transporter': '',
                        'Delivery Quote': '',
                        'Packages': des.get('total_package', ''),
                        'Pallets': fbaPallets,
                        'importClearance': importClearance,
                        'Duty Invoice': duty_invoice,
                        'Duty Invoice Status': duty_invoice_status,
                        'Actual # of Pallets': '',
                        'Ready for Pick-up Date': '',
                        'LFD': last_free_date_at_fpod,
                        'DO Release Approved?': '',
                        'HBL Released Date': '',
                        'DO Released Date': delivery_order_release,
                        'Pick-up Date': des.get('atdfrompod', ''),
                        'Pick up number': '',
                        'Delivery Appointment Date': '',
                        'Delivery Date': des.get('actual_delivery_date', ''),
                        'Vendor Delivery Invoice': '',
                        'Updated Status Remarks': remarks,
                        'PRO Number': '',
                        'Storage Incurred (Days)': '',
                    }

                    result_rows.append(row_data)
            else:
                # No vdes – create a single row
                row_data = {
                    'Customer Name': entityName,
                    'MBL#': mbl_number,
                    'HBL#': hbl_number,
                    'Agraga Booking #': booking_id,
                    'Booking Status': status,
                    'FBA?': fba,
                    'ISF Filing': '',
                    'Stuffing Date': stuffing_confirmation,
                    'Container #': container_number,
                    'ETD': etd_at_pol,
                    'ETA': eta_fpod,
                    'SOB': sob_pol,
                    'ATA': gatein_fpod,
                    'Carrier': carrier,
                    'Consolidator': consolidator,
                    'Origin': origin,
                    'FPOD': fpod,
                    'CFS': '',
                    'Delivery Address': '',
                    'FBA Code': '',
                    'Freight Broker': '',
                    'Transporter': '',
                    'Delivery Quote': '',
                    'Packages': '',
                    'Pallets': fbaPallets,
                    'importClearance': importClearance,
                    'Duty Invoice': duty_invoice,
                    'Duty Invoice Status': duty_invoice_status,
                    'Actual # of Pallets': '',
                    'Ready for Pick-up Date': '',
                    'LFD': last_free_date_at_fpod,
                    'DO Release Approved?': '',
                    'HBL Released Date': '',
                    'DO Released Date': delivery_order_release,
                    'Pick-up Date': '',
                    'Pick up number': '',
                    'Delivery Appointment Date': '',
                    'Delivery Date': '',
                    'Vendor Delivery Invoice': '',
                    'Updated Status Remarks': remarks,
                    'PRO Number': '',
                    'Storage Incurred (Days)': '',
                }

                result_rows.append(row_data)

        except Exception as e:
            booking_processlog.error(f"Error processing Booking ID {rows.get('_id', 'UNKNOWN')}: {str(e)}")

    final_df = pd.DataFrame(result_rows)
    booking_processlog.info(f"Finished booking_process with {len(result_rows)} rows created.")
    booking_processlog.info('*'*100)
    
    return final_df


# In[3]:


def process_report(existing_report, generated_report):
    key_col = 'Agraga Booking #'
    exclude_cols = ['ISF Filing', 'CFS', 'Freight Broker', 'Transporter', 'Delivery Quote',
                    'Actual # of Pallets', 'Ready for Pick-up Date', 'DO Release Approved?', 
                    'HBL Released Date', 'Pick up number', 'Delivery Appointment Date',
                    'Vendor Delivery Invoice', 'PRO Number', 'Storage Incurred (Days)']

    compare_cols = [col for col in existing_report.columns if col not in exclude_cols + [key_col]]

    comparisonlog.info(f"Total rows in existing report: {len(existing_report)}")
    comparisonlog.info(f"Total rows in generated report: {len(generated_report)}")
    comparisonlog.info(f"Columns to compare (excluding keys and excluded): {compare_cols}")

    existing_df = existing_report.set_index(key_col)
    generated_df = generated_report.set_index(key_col)

    # Step 1: Key comparison
    common_keys = existing_df.index.intersection(generated_df.index)
    new_keys = generated_df.index.difference(existing_df.index)

    comparisonlog.info(f"Common keys: {len(common_keys)}")
    comparisonlog.info(f"New keys (to be added): {len(new_keys)}")

    # Step 2: Row comparison
    updated_rows = existing_df.loc[common_keys, compare_cols] \
        .compare(generated_df.loc[common_keys, compare_cols], keep_shape=True, keep_equal=False)

    changed_keys = updated_rows.dropna(how='all').index
    comparisonlog.info(f"Changed records: {len(changed_keys)}")

    # Log actual changes (optional: can be verbose)
    if not updated_rows.empty:
        comparisonlog.info("Sample changes:")
        comparisonlog.info(updated_rows.dropna(how='all').head().to_string())

    # Step 3: Apply updates
    final_df = existing_df.copy()
    final_df.loc[changed_keys, compare_cols] = generated_df.loc[changed_keys, compare_cols]

    # Step 4: Add new records
    new_rows_df = generated_df.loc[new_keys]
    final_df = pd.concat([final_df, new_rows_df], axis=0)

    comparisonlog.info(f"Final updated report rows: {len(final_df)}")
    comparisonlog.info('*'*100)
    
    return final_df.reset_index()

def close_logger(name):
    logger = logging.getLogger(name)
    for handler in logger.handlers[:]:
        handler.close()
        logger.removeHandler(handler)


# In[4]:


bookings, shentities, bookingdsr, Myactions, Addressdetails, Agusers= fetch_data()

with pd.ExcelWriter(r"data/Users.xlsx", engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    Agusers.to_excel(writer, sheet_name='Agusers', index=False)
    
generated_report = booking_process(bookings,Addressdetails)

existing_report_path = r"data/report.xlsx"

if os.path.isfile(existing_report_path):
    existing_report = pd.read_excel(existing_report_path)
    processed_report = process_report(existing_report,generated_report)
    processed_report.to_excel(existing_report_path, index = False)
else:
    generated_report.to_excel(existing_report_path, index = False)
    comparisonlog.info(f"New Report Generated with rows: {len(generated_report)}")
    comparisonlog.info('*'*100)

close_logger('mongolog')
close_logger('booking_processlog')
close_logger('comparisonlog')

# bookings.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\bookings.xlsx")
# shentities.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\shentities.xlsx")
# bookingdsr.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\bookingdsr.xlsx")
# Myactions.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\Myactions.xlsx")
# Addressdetails.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\Addressdetails.xlsx")
# report.to_excel(r"D:\Ayyanagouda\MSME Shipment Tracker\data\report.xlsx")


# In[ ]:




