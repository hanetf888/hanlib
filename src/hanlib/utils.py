import os
import datetime
from datetime import timedelta
import win32com.client as client
import time
from openpyxl.styles import Border, Side
from pathlib import Path
from xbbg import blp
import hanconfig
import database as han_db
import inspect
import zipfile
import xlrd
import csv
import logging

hc = hanconfig.HanConfig()
db = han_db.getdatabaseobject(hc)

logger = logging.getLogger(__name__)

class HANUser(object):
    __username = os.getenv("USERNAME")

    @property
    def username(self):
        return self.__username

    @property
    def user_path(self):
        return hc.cfg["filepath"]["environment"]

    def __repr__(self):
        return self.__username


def fileexists(file_path):
    """Checks if the file exists locally"""
    file_path_obj = Path(file_path)
    return file_path_obj.exists() and file_path_obj.is_file()


def is_business_day(date_obj):
    """
    Checks if a given date is a business day (Monday-Friday).
    Does not account for public holidays.
    """
    # Monday is 0, Sunday is 6
    return date_obj.weekday() < 5


def getlastbusinessday(run_date, days_offset):
    """
    rundate must be a datetime.date
    days_offset must be an integer

    Returns the last business date according to the offset.
    days_offset = -1 will return the previous business day
    days_offset = -2 will return the business day 2 days ago

    function returns -1 if there is an error

    """

    if not isinstance(run_date, datetime.date):
        print("utils.getlastbusinessday: run_date passed is not a datetime.date")
        return -1

    if not isinstance(days_offset, int):
        print("utils.getlastbusinessday: days_offset passed is not an int")
        return -1

    run_date_with_offset = run_date + timedelta(days=days_offset)

    # if the run_date with the offset applied is a weekend then we keep
    # subtracting days until it is a weekday.
    while run_date_with_offset.weekday() > 4:
        if days_offset == -2:
            run_date_with_offset = run_date_with_offset + timedelta(days=days_offset)
        else:
            run_date_with_offset = run_date_with_offset + timedelta(
                days=days_offset - 1
            )
    return run_date_with_offset


def get_busday_x_days_from_date(start_date, x_days):
    """
    Returns a date that is 'business_days' from 'start_date', excluding weekends and optional holidays.

    :param start_date: (datetime or str) The starting date (format: 'YYYY-MM-DD' if string)
    :param business_days: (int) Number of business days to add (can be negative for past dates)
    :param holidays: (list of str) Optional list of holiday dates (format: 'YYYY-MM-DD')
    :return: (datetime) The resulting business date
    """
    if isinstance(start_date, str):
        start_date = datetime.strptime(start_date, "%Y-%m-%d")

    # holidays = set(datetime.strptime(h, "%Y-%m-%d").date() for h in (holidays or []))

    current_date = start_date
    step = 1 if x_days > 0 else -1  # Forward or backward

    while x_days != 0:
        current_date += timedelta(days=step)
        # if current_date.weekday() < 5 and current_date not in holidays:  # Monday-Friday and not a holiday
        if current_date.weekday() < 5:  # Monday-Friday
            x_days -= step

    return current_date


def get_first_business_day_of_month(year, month):
    """
    Calculates the first business day of a given month.
    """
    day = 1
    while True:
        try:
            current_date = datetime.date(year, month, day)
            if is_business_day(current_date):
                return current_date
            day += 1
        except ValueError:
            # If we try to create a date like Feb 30, it will raise ValueError.
            # This means we've gone past the end of the month, which shouldn't happen
            # if the month has at least one business day (which it always will).
            raise Exception(
                f"Could not find a business day in {month}/{year}. This should not happen."
            )


def is_weekend(date):
    """
    Checks if the given date is a weekend (Saturday or Sunday).

    :param date: (datetime.date) The date to check.
    :return: (bool) True if it's a weekend, False otherwise.
    """
    return date.weekday() >= 5  # 5 = Saturday, 6 = Sunday


def mailsubjectexists(subject, folder, email):
    outlook = client.Dispatch("outlook.application")
    mapi = outlook.GetNamespace("MAPI")

    messages = mapi.Folders(email).Folders(folder).Items

    messages = messages.Restrict(f"[Subject] = '{subject}'")

    if len(messages) == 0:
        # If the length of the list is zero then the email is not found - return false
        return False
    else:
        # If the length of the list is> 0 then the email has been found - return true
        return True


def setborder(ws, startrow, startcol, endrow, endcol):
    thin = Side(border_style="thin", color="000000")

    if endrow < startrow:
        raise ValueError("Start end row is before start row")
    if endcol < startcol:
        raise ValueError("Start end column is before start column")

    for col in range(startcol, endcol + 1):
        ws.cell(startrow, col).border = Border(top=thin)

    for col in range(startcol, endcol + 1):
        ws.cell(endrow, col).border = Border(bottom=thin)

    for rw in range(startrow, endrow):
        ws.cell(rw, startcol).border = Border(left=thin)
        ws.cell(rw, endcol).border = Border(right=thin)

    # Set corners
    ws.cell(startrow, startcol).border = Border(top=thin, left=thin)
    ws.cell(startrow, endcol).border = Border(top=thin, right=thin)
    ws.cell(endrow, startcol).border = Border(bottom=thin, left=thin)
    ws.cell(endrow, endcol).border = Border(bottom=thin, right=thin)


# getlastmonthend , months must be a positive number.  3 measn it gets 3 monthends ago
def getlastmonthend(curdate, months):

    returndate = curdate

    for i in range(months):
        returndate = datetime.date(
            returndate.year, returndate.month, 1
        ) - datetime.timedelta(days=1)

    return returndate


def getlastmonthendbd(curdate, months):
    returndate = curdate

    for i in range(months):
        returndate = datetime.date(
            returndate.year, returndate.month, 1
        ) - datetime.timedelta(days=1)

    return getlastbusinessday(returndate, 0)


# def getlastquarterend(curdate):
#
#     # curdate must always be a month end.
#
#     if curdate.month in [1, 2, 3]:
#         mth = 12
#         days = 31
#         yr = curdate.year - 1
#     elif curdate.month in [4, 5, 6]:
#         mth = 3
#         days = 31
#         yr = curdate.year
#     elif curdate.month in [7, 8, 9]:
#         mth = 6
#         days = 30
#         yr = curdate.year
#     elif curdate.month in [10, 11, 12]:
#         mth = 9
#         days = 30
#         yr = curdate.year
#
#     return datetime.date(yr,mth,days)


def getlastquarterend(curdate):
    today = curdate
    current_quarter = (today.month - 1) // 3 + 1
    last_quarter_end_month = (current_quarter - 1) * 3
    if last_quarter_end_month == 0:  # If it's Q1, go back to the previous year's Q4
        last_quarter_end_month = 12
        year = today.year - 1
    else:
        year = today.year

    return datetime(year, last_quarter_end_month, 1) + timedelta(days=-1)


def getquarterenddate(curdate, quarters):
    qtrtrenddate = getlastquarterend(curdate)
    for i in range(1, quarters):
        qtrtrenddate = getlastmonthend(qtrtrenddate, 3)
    return qtrtrenddate


def getNetFlows(ticker, curdate, date1d, date30d, date90d, date180d):

    dfnetflowdata = blp.bdh(
        ticker + "Equity",
        ["eqy_sh_out", "fund_net_asset_val"],
        date90d,
        curdate,
        Days="NON_TRADING_WEEKDAYS",
        Fill="P",
    )
    dfnetflowdata.reset_index(inplace=True)
    dfnetflowdata.columns = ["date", "eqy_sh_out", "fund_net_asset_val"]
    dfnetflowdata["eqy_sh_out1"] = dfnetflowdata["eqy_sh_out"].shift(1)
    dfnetflowdata["net_shs"] = dfnetflowdata["eqy_sh_out"] - dfnetflowdata[
        "eqy_sh_out"
    ].shift(1)
    dfnetflowdata["net_flow"] = (
        dfnetflowdata["net_shs"] * dfnetflowdata["fund_net_asset_val"]
    )

    dfnetflowdata.to_excel(
        f"C:\\Users\\{os.getenv('USERNAME')}\\HANetf\\Operations - Documents\\Projects\\PowerBI Reports"
        + f"\\Competitor Dashboard\\{ticker}.xlsx"
    )

    netflow1d = dfnetflowdata[
        (dfnetflowdata["date"] >= date1d) & (dfnetflowdata["date"] <= curdate)
    ]["net_flow"].sum()
    netflow30d = dfnetflowdata[
        (dfnetflowdata["date"] >= date30d) & (dfnetflowdata["date"] <= curdate)
    ]["net_flow"].sum()
    netflow90d = dfnetflowdata[
        (dfnetflowdata["date"] >= date90d) & (dfnetflowdata["date"] <= curdate)
    ]["net_flow"].sum()
    netflow180d = dfnetflowdata[
        (dfnetflowdata["date"] >= date180d) & (dfnetflowdata["date"] <= curdate)
    ]["net_flow"].sum()

    return [netflow1d, netflow30d, netflow90d, netflow180d]


def get30DADV(ticker, todate):

    turnover = blp.bdh(
        ticker + " Equity",
        "turnover",
        todate + timedelta(days=-30),
        todate,
        Currency="USD",
        Days="NON_TRADING_WEEKDAYS",
        Fill="P",
    )
    if turnover.empty:
        adv30d = 0
    else:
        turnover.columns = ["turnover"]
        adv30d = turnover["turnover"].mean()

    return adv30d


def getTotalADV(tickers, curdate):

    adv = 0
    for ticker in tickers:
        adv += get30DADV(ticker, curdate)

    return adv


# def getPerformance(ticker, startdate, enddate, currency):
#     # This does not currently get performance with dividends

#     navs = blp.bdh(
#         ticker + " Equity",
#         "fund_net_asset_val",
#         startdate,
#         enddate,
#         Currency=currency,
#         Days="NON_TRADING_WEEKDAYS",
#         Fill="P",
#     )
#     navs.dropna(inplace=True)
#     navs.columns = ["fund_net_asset_val"]

#     divs = blp.bdh(
#         ticker + "Equity",
#         "DVD_HIST_ALL",
#         startdate,
#         enddate,
#         Currency=currency,
#         Days="NON_TRADING_WEEKDAYS",
#         Fill="P",
#     )

#     # alldata = navs.join(divs)
#     alldata = navs

#     startnav = navs["fund_net_asset_val"].iloc[0]
#     endnav = navs["fund_net_asset_val"].iloc[-1]

#     return endnav / startnav - 1


def get5dayspread(ticker, curdate):
    spread = blp.bdh(
        ticker + " Equity",
        "PY535",
        curdate,
        curdate,
        Days="ALL_CALENDAR_DAYS",
        Fill="P",
    )
    if spread.empty:
        return 0
    else:
        spread.columns = ["PY535"]
        return spread["PY535"].iloc[0]


# Database Utilities


def copyTable(tablename, fromdb, todb):
    logger.info(f"copying {tablename}")
    sql = f"""select * from data.{tablename}"""
    dfdata = fromdb.query(sql, write=False, commit=False)
    dfdata = dfdata.drop(dfdata.columns[0], axis=1)

    # sql = f"drop table data.{tablename}"
    # todb.query(sql, write=True, commit=True)

    todb.createTable(
        dfdata, tablename, schema="data", append=False, clean_dataframe=True, log=True
    )


def refreshTestfromProd():

    # Create prod config
    hcprod = hanconfig.HanConfig()
    proddb = han_db.getdatabaseobject(hcprod)

    # Create dev config
    hcdev = hanconfig.HanConfig()
    hcdev.cfg = hanconfig.configdict["dev"]
    devdb = han_db.getdatabaseobject(hcdev)

    tables = [
        "product",
        "shareclass",
        "listing",
        "nav",
        "country_lookup",
        "country_iso",
        "region",
        "indices",
        "dividend",
        "ap",
        "im",
        "indextrackrecord",
        "index_values",
        "fx_rates",
        "security_entitlement",
        "etc_physical_holdings",
        "flows",
        "passportingmatrix",
        "share_class_asset",
    ]

    feedtables = [
        "email_data_paths",
        "sftp_data_paths",
        "fefundinfo",
        "german_equity_ratio",
        "sector",
    ]

    for table in tables:
        copyTable(table, proddb, devdb)

    for table in feedtables:
        copyTable(table, proddb, devdb)


def getAccountMapDict():
    db = han_db.getdatabaseobject(hc)
    sql = """select new_id, old_id from data.acct_id_map"""
    dfacct_id_map = db.query(sql, write=False, commit=False)
    dfacct_id_map.set_index("new_id", inplace=True)
    acct_id_map = dfacct_id_map.to_dict()
    acct_map = acct_id_map["old_id"]

    return acct_map


def sendemail(
    curdate, emailto, emailcc, emailbcc, emailsubject, emailbody, emailattachments
):

    if hc.environment != "prod":
        emailsubject += f" {hc.environment}"

    if not hc.cfg['send_emails']:
        logging.info(
            f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M")} sendemail: Email sending is turned OFF in hanconfig, not sending email with subject {emailsubject}'
        )
        return True

    # Check if email has already been sent.  If it has, then do not re-send the email
    if not (mailsubjectexists(emailsubject, "Sent Items", hc.cfg["email"]["sender"])):
        logging.info(
            f'{datetime.datetime.now().strftime("%Y-%m-%d %H:%M")} sendemail: Email not yet sent with subject {emailsubject}'
        )

        outlook = client.Dispatch("Outlook.Application")

        message = outlook.CreateItem(0)

        if hc.environment == "prod":
            message.To = emailto
            message.CC = emailcc
            message.BCC = emailbcc

        elif hc.environment == "uat":
            message.To = "uat@hanetf.com"
            message.CC = "uat@hanetf.com"
            message.BCC = "uat@hanetf.com"
        elif hc.environment == "dev":
            message.To = "dev@hanetf.com"
            message.CC = "dev@hanetf.com"
            message.BCC = "dev@hanetf.com"

        message.Subject = f"{emailsubject}"

        # Insert the comments as a dataframe html
        message.HTMLBody = emailbody

        if len(emailattachments) > 0:
            for filename in emailattachments:
                if isinstance(filename, Path):
                    filename = str(filename)
                logging.info(f"Attaching file: {filename}")
                message.Attachments.Add(Source=filename)

        message.Send()

    else:
        logging.info(
            f"{datetime.datetime.now().strftime('%Y-%m-%d %H:%M')} sendemail: Email already sent, doing nothing"
        )
    return True


def doNothing():
    pass


def getFinalTermsRegulator(filename: str):
    regulators = ["FCA", "CBI", "SFSA"]
    for regulator in regulators:
        if filename.lower().find(regulator.lower()) >= 0:
            return regulator


def findFile(
    folder: str,
    keyword: str,
    extension: str,
    how: str = "latest",
    returns: str = "filepath",
) -> str:
    """Provide a folder, keyword, and extension, along with how ('latest' or 'earliest') and what to return ('filepath' or 'date')
    and this returns the (latest or earliest) (filepath or saved date) with that keyword and extension in the specified folder, as a (string or datetime.date object).
    """

    files = os.listdir(folder)

    relevant_files = []
    relevant_dates = []

    for filename in files:

        filepath = folder + "\\" + filename
        created_date = datetime.datetime.fromtimestamp(os.path.getmtime(filepath))

        if filepath.find(keyword) >= 0 and filepath.find(extension) >= 0:

            relevant_files.append(filepath)
            relevant_dates.append(created_date)

    relevant_file_date = max(relevant_dates) if how == "latest" else min(relevant_dates)
    relevant_index = relevant_dates.index(relevant_file_date)
    relevant_filename = relevant_files[relevant_index]

    return relevant_filename if returns == "filepath" else relevant_file_date.date()


def castAsString(data):
    try:
        return str(data)
    except Exception:
        return data


def toDate(data, to="date"):
    if to == "date":
        return datetime.datetime.date(datetime.datetime.strptime(data, "%Y%m%d"))
    else:
        return datetime.datetime.strftime(data, "%Y%m%d")


def divByMillion(data):

    return data / 1000000


def percent(data):
    return data / 100


def current_function_name():
    return inspect.currentframe().f_code.co_name


class FolderNotFound(Exception):
    """Custom exception for when folder is not found."""

    pass


def zip_and_clean_folder_with_exclusions(curdate, folder_path, excludelist):
    # Extract the folder name from the folder path
    foldername = os.path.basename(os.path.abspath(folder_path))
    lastbusdate = get_busday_x_days_from_date(curdate, -1)

    # Check if the folder exists
    if not os.path.exists(folder_path):
        raise FolderNotFound(f"The folder '{folder_path}' does not exist.")

    # Path for the zip file
    zip_file_path = os.path.join(folder_path, f"{foldername}.zip")

    # Gather a list of all files in the folder excluding the excludelist and the ZIP file itself
    files_to_zip = [
        f
        for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f))
        and f not in excludelist  # Exclude files in the excludelist
        and f != f"{foldername}.zip"  # Exclude the target folder ZIP file
        and curdate.strftime("%Y%m%d")
        not in f  # Exclude files with the current date in the filename
        and curdate.strftime("%d.%m.%Y")
        not in f  # Exclude files with the current date in the filename
        and curdate.strftime("%d_%m_%Y")
        not in f  # Exclude files with the current date in the filename
        and curdate.strftime("%d%m")
        not in f  # Exclude files with the current date in the filename
        and curdate.strftime("%d %m %Y")
        not in f  # Exclude files with the current date in the filename
        and lastbusdate.strftime("%Y%m%d")
        not in f  # Exclude files with the current date in the filename
        and lastbusdate.strftime("%d.%m.%Y")
        not in f  # Exclude files with the current date in the filename
        and lastbusdate.strftime("%d_%m_%Y")
        not in f  # Exclude files with the current date in the filename
        and lastbusdate.strftime("%d%m")
        not in f  # Exclude files with the current date in the filename
        and lastbusdate.strftime("%d %m %Y")
        not in f  # Exclude files with the current date in the filename
    ]

    # If the zip file does not exist, create it
    if not os.path.exists(zip_file_path):
        with zipfile.ZipFile(zip_file_path, "w") as zf:
            for file in files_to_zip:
                zf.write(os.path.join(folder_path, file), arcname=file)
        logger.info(f"Created new zip file: {zip_file_path}")
    else:
        # If the zip file exists, add any files not already in it and not in the excludelist
        with zipfile.ZipFile(zip_file_path, "a") as zf:
            existing_files = set(zf.namelist())  # Files already in the zip
            for file in files_to_zip:
                if file not in existing_files:
                    zf.write(os.path.join(folder_path, file), arcname=file)
                    file_path = os.path.join(folder_path, file)

                    logger.info(
                        f"Zipped file {file_path} to existing zip file {zip_file_path}"
                    )

                    # If the file is in the zip file, not on the exclude list and exits
                    # in the folder, delete it
                    if (
                        file not in excludelist
                        and file in zf.namelist()
                        and os.path.exists(file_path)
                    ):
                        os.remove(file_path)
                        logger.info(
                            f"Deleted file: {file_path} as in zipfile {zip_file_path}"
                        )

        logger.info(f"Updated existing zip file: {zip_file_path}")

    # # Delete files that are now zipped, except files listed in the
    # with zipfile.ZipFile(zip_file_path, 'r') as zf:
    #     zipped_files = zf.namelist()  # Files present in the zip
    #     for file in zipped_files:
    #         if file not in excludelist:  # Skip files in the excludelist
    #             file_path = os.path.join(folder_path, file)


def checkholdingscomplete(curdate: datetime.date) -> bool:
    """
    Checks if all holdings are complete for the given date.

    Args:
        curdate (datetime.date): Current date for evaluation.

    Returns:
        bool: False if the DataFrame is not empty, True if it's empty.

    Raises:
        Exception: If an unexpected error occurs.
    """
    # Adjust curdate to the last business day (1 day before)

    # Execute the query using db.sql and assign the result to a DataFrame
    hc = hanconfig.HanConfig()
    db = han_db.getdatabaseobject(hc)

    # SQL query to find missing holdings
    # The SQL looks for isins that are live but are not in the holdings table for today
    sql = f"""
        select 	sc.gsp_account_number
        , 		share_class_name
        ,		share_class_inception_date
        from 	data.product pr
        ,		data.shareclass sc
        where	pr.live
        and		sc.share_class_live
        and		pr.id = sc.product_id
        and		pr.product_class = 'ETF'
        and     sc.share_class_inception_date <= '{curdate.strftime("%Y-%m-%d")}'
        and		sc.gsp_account_number not in ( 
            select distinct account_number
            from data.holdings
            where effective_date = '{curdate.strftime("%Y-%m-%d")}'
            )
    """

    df = db.query(sql)

    # Check the DataFrame's contents
    if not df.empty:
        # If the DataFrame is not empty, return False
        return False
    elif df.empty:
        # If the DataFrame is empty, return True
        return True
    else:
        # This condition should never occur, but raise an exception as a safeguard
        raise Exception("Error in checkholdingscomplete")


def checkholdingsexist(curdate: datetime.date) -> bool:
    """
    Checks if all holdings are complete for the given date.

    Args:
        curdate (datetime.date): Current date for evaluation.

    Returns:
        bool: False if the DataFrame is not empty, True if it's empty.

    Raises:
        Exception: If an unexpected error occurs.
    """
    # Adjust curdate to the last business day (1 day before)

    # Execute the query using db.sql and assign the result to a DataFrame
    hc = hanconfig.HanConfig()
    db = han_db.getdatabaseobject(hc)

    # SQL query to find missing holdings
    sql = f"""
    SELECT * 
    FROM data.holdings h, data.product pr
    WHERE effective_date = '{curdate.strftime("%Y-%m-%d")}'
      AND pr.live
      AND pr.product_class = 'ETF'
      AND pr.gsp_account_number = h.account_number
    """

    df = db.query(sql)

    # Check the DataFrame's contents
    if not df.empty:
        # If the DataFrame is not empty, holdings exist, return True
        return True
    elif df.empty:
        # If the DataFrame is empty, holdings do not exist, return False
        return False
    else:
        # This condition should never occur, but raise an exception as a safeguard
        raise Exception("Error in checkholdingsExist")


def convert_xls_to_csv_with_formats(xls_filename, csv_filename):
    # Open the .xls file
    workbook = xlrd.open_workbook(xls_filename)
    sheet = workbook.sheet_by_index(
        0
    )  # Use the first sheet (can be adjusted dynamically)

    # Open the .csv file for writing
    with open(csv_filename, "w", newline="") as csv_file:
        csv_writer = csv.writer(csv_file)

        # Loop through rows in the Excel sheet
        for row_idx in range(sheet.nrows):
            row = []
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                cell_type = cell.ctype  # Get the type of the cell
                cell_value = cell.value

                if cell_type == xlrd.XL_CELL_DATE:  # Handle date types
                    # Convert Excel serial date to a Python datetime object
                    date_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
                    cell_value = datetime.datetime(*date_tuple).strftime(
                        "%Y-%m-%d"
                    )  # Format the date as YYYY-MM-DD
                elif cell_type == xlrd.XL_CELL_NUMBER:  # Handle numbers
                    cell_value = (
                        "%.15g" % cell_value
                    )  # Handle large numbers and preserve precision

                row.append(cell_value)  # Append the formatted cell value to the row
            logger.info(f"{row}")
            csv_writer.writerow(row)  # Write the row to the CSV

    logger.info(
        f"Converted '{xls_filename}' to '{csv_filename}' with formats preserved!"
    )

def excel_open_force_recalc(filename, excel_tab, cell_update):

    try: 
        excel = client.DispatchEx('Excel.Application')
        # Display the excel application
        excel.Visible = True

        #turn off all dialog boxes and notificatioins
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow

        # 1. open the file
        workbook = excel.Workbooks.Open(filename, UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)
        logger.info(rf'Workbook "{filename}" opened in Excel.')
        
        # 2. move to the tab to update 
        sheet = workbook.Sheets(excel_tab)
        sheet.Activate()
        logger.info(f"Activated sheet: {excel_tab}")
        time.sleep(5)  # give Excel time to move to sheet

        # 3. Perform an Action (e.g., updating cell to force a recalculation and refresh cache)
        sheet.Range(cell_update).Value = " "   
        logger.info(f"Updated cell {cell_update} on tab {excel_tab} value to force recalculation.")

        # 3.1 Force Excel to recalculate all formulas as well
        workbook.Application.CalculateFull()
        logger.info("Forced full recalculation of all formulas in workbook.")

        # Give Excel a moment to recalculate (especially useful for complex workbooks)
        time.sleep(5) 
        
        # 4. Save the Workbook
        workbook.Save()
        logger.info(f"Workbook {filename} saved successfully.")

    except Exception as e:
        logger.error(f"An error occurred while processing the Excel file: {filename} {e}")
        
    finally:    
        workbook.Close(SaveChanges=False)
        excel.Quit()    
        logger.info(f"Excel closed successfully '{filename}'.")
        time.sleep(5) 

