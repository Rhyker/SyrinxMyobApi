__author__ = 'Jade McKenzie', 'Tyler McCamley'

import MyobAuth
import csv
import datetime
import pyodbc
import configparser
import os
import sys

DEBUG = False

# Included Tasks:
# > Clear old information from SyrinxEH database: syrinx_clear()
# > Get overdue information from MYOB: get_overdue(days_over, min_total)
# > Set overdue customers in Syrinx: set_overdue(overdue_customers, days_over)

config = configparser.ConfigParser()
curpath = os.path.dirname(os.path.realpath(sys.argv[0]))
cfgpath = os.path.join(curpath, 'settings.ini')
config.read(cfgpath)

folder_location = config.get("DEFAULT", "folder_location")

# MYOB Authorisation
app = MyobAuth.EagleMyobApi()

# Required data for logs
from_syrinx = ['Alert Notes', 'Prop:ExcludeFromDebtCollecting']
from_json = ['Number', 'Date', 'DisplayID', 'Name', 'BalanceDueAmount',
             'Subtotal', 'TotalTax', 'TotalAmount', 'JournalMemo']

# Connect to SyrinxEH Database
connection = pyodbc.connect(r"Driver={SQL Server};"
                            r"Server=" + config.get("SQL", "server") + r";"
                            r"Database=" + config.get("SQL", "db") + r";"
                            r"UID=" + config.get("SQL", "uid") + r";"
                            r"PWD=" + config.get("SQL", "pwd") + r";"
                            r"Trusted_Connection="
                            + config.get("SQL", "trusted") + r";")
cursor = connection.cursor()


# Clear old information from SyrinxEH database
def syrinx_clear():
    """
    Removes automated notes entered prior to now from Syrinx Customer Alert
    """

    if DEBUG:
        print("-------- DEBUG ON --------")

    to_clear = list()
    cursor.execute('SELECT '
                   'CST_ACCOUNT_NUMBER, '
                   'CST_ALERT_TEXT '
                   'FROM SyrinxEH.dbo.TH_CUSTOMERS '
                   'WHERE CST_CURRENT_FLAG = 1 '
                   'AND CST_ACCOUNT_NUMBER IS NOT NULL '
                   'AND (CST_ALERT_TEXT IS NOT NULL '
                   'OR CST_PROP2 IS NOT NULL '
                   'OR CST_ON_HOLD = 1)')

    # Build list of customers to clear
    for row in cursor:
        if row[1] is None:
            to_clear.append((row[0], ''))
            continue
        pos = row[1].find('~~')
        if (pos > 0) and (pos != len(row[1])):
            to_clear.append((row[0], row[1][pos + 3:(len(row[1]))]))
        elif pos <= 0:
            to_clear.append((row[0], row[1]))

    # Clear notes and remove on hold flag from listed customers
    for customer in to_clear:

        query = ("UPDATE SyrinxEH.dbo.TH_CUSTOMERS SET "
                 "CST_ALERT_CHG_DATE = NULL, "
                 "CST_ALERT_CHG_USERID = NULL, "
                 "CST_ALERT_TEXT = "
                 + ("NULL" if customer[1] == '' else ("'" + customer[1] + "'"))
                 + ", "
                 "CST_ON_HOLD = 0 "
                 "WHERE CST_ACCOUNT_NUMBER = '" + customer[0] + "'")

        if DEBUG:
            print("customer, query: ", customer, query)

        else:
            cursor.execute(query)
            connection.commit()

    print("Outdated overdue data cleared from Syrinx. Ready for update.")


# Get overdue information from MYOB
def get_overdue(days_over, min_total):
    """
    Gets JSON file from MYOB with invoice details and processes each line entry,
    logging important details and accumulating a total balance for all *OVERDUE*
    invoices.
    """

    # Get JSON file and check validity
    results = app.myob_request('overdue', min_total, days_over)
    if not results:
        return 'No JSON file returned.'

    # Create detailed log with information from individual invoices
    log_dict_sum = dict()
    log_dict_det = dict()
    now = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
    filename_det = "log_" + now + "_detailed.txt"

    with open(folder_location + filename_det, 'w', newline='') \
            as log_detailed:
        f_write = csv.writer(log_detailed, delimiter='\t',
                             quotechar='|', quoting=csv.QUOTE_MINIMAL)
        f_write.writerow(from_json)
        for result in results:
            important_info = list()
            for field in from_json:
                if field in ['DisplayID', 'Name']:
                    important_info.append(result['Customer'][field])
                    continue
                if field == 'Date':
                    # Reformat date from y/m/d to d/m/y
                    dt_ymd = result[field]
                    dt_dmy = f"{dt_ymd[8:10]}/{dt_ymd[5:7]}/{dt_ymd[0:4]}"
                    important_info.append(dt_dmy)
                    continue
                important_info.append(result[field])
            f_write.writerow(important_info)
            log_dict_det[important_info[0]] = important_info[1:]

            # Build dictionary with invoice information summarised by customer
            if result['Customer']['DisplayID'] not in log_dict_sum:
                log_dict_sum[result['Customer']['DisplayID']] = [0, 0.0, None]

            log_dict_sum[result['Customer']['DisplayID']][1] += \
                result['BalanceDueAmount']

    # Query Syrinx Customers for existing notes and HoldOverdue Status
    cursor.execute('SELECT '
                   'CST_ACCOUNT_NUMBER, '
                   'CST_ALERT_TEXT, '
                   'CST_ON_HOLD, '
                   'CASE WHEN CST_PROP2 IS NULL THEN 0 ELSE 1 END AS CST_PROP2 '
                   'FROM SyrinxEH.dbo.TH_CUSTOMERS '
                   'WHERE CST_CURRENT_FLAG = 1 '
                   'AND CST_ACCOUNT_NUMBER IS NOT NULL '
                   'AND (CST_ALERT_TEXT IS NOT NULL '
                   'OR CST_PROP2 IS NOT NULL '
                   'OR CST_ON_HOLD = 1)')

    # Update log_dict with results from query
    for row in cursor:
        if row[0] not in log_dict_sum:
            log_dict_sum[row[0]] = [row[3], 0.0, row[1]]
            continue
        log_dict_sum[row[0]][0] = row[3]
        log_dict_sum[row[0]][2] = row[1]

    if DEBUG:
        print("log_dict_sum: ", log_dict_sum)
        print("log_dict_det: ", log_dict_det)

    print("Overdue information retrieved. Details stored in file '"
          + filename_det + "' located in folder '" + folder_location)

    import OverdueLog
    OverdueLog.convert_log_to_excel(log_dict_sum, log_dict_det, now)

    return [log_dict_sum, log_dict_det]


# Set overdue customers & update Alert Notes
def set_overdue(overdue_customers, days_over):
    """
    Accepts a dictionary of customers with a list of their exclusion status (a
    Prop field in Misc tab of CRM in Syrinx), accumulated overdue balance and
    manually entered alert notes. Constructs a new Alert Note in Syrinx, places
    customers on hold if required and creates a log.
    """

    # For every item in given dictionary, update Syrinx with new information.
    now = datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")
    filename_sum = "log_" + now + "_summary.txt"

    with open(folder_location + filename_sum, 'w', newline='') as log_summary:
        f_write = csv.writer(log_summary, delimiter='\t',
                             quotechar='|', quoting=csv.QUOTE_MINIMAL)
        f_write.writerow([from_json[3], from_syrinx[1], 'TotalOwing',
                          from_syrinx[0]])
        for key, value in overdue_customers.items():
            f_write.writerow([key, value[0], value[1], value[2]])
            if value[1] > 0:

                total_owing = ("%.2f" % value[1])
                days_string = str(days_over)
                kept_notes = ((' ' + value[2]) if value[2] else '')
                customer_code = str(key)

                if value[0] == 1:

                    query = "UPDATE SyrinxEH.dbo.TH_CUSTOMERS SET " \
                            "CST_ALERT_CHG_DATE = GETDATE(), " \
                            "CST_ALERT_CHG_USERID = 16, " \
                            "CST_ALERT_TEXT = 'OVERDUE ACCOUNT: $" \
                            + total_owing + " is the total of all invoices " \
                            "due over " + days_string + " days ago. For more " \
                            "info, see MYOB. This customer is excluded from " \
                            "debt collecting. To stop excluding this " \
                            "customer from debt collection and place this " \
                            "customer on hold, go to CRM, tick on hold and " \
                            "remove N from Misc > Hold Overdue. This message " \
                            "and status will disappear when the account is " \
                            "up to date. ~~" + kept_notes + "' " \
                            "WHERE CST_ACCOUNT_NUMBER = '" + customer_code + "'"

                elif value[0] == 0:
                    query = "UPDATE SyrinxEH.dbo.TH_CUSTOMERS " \
                            "SET CST_ON_HOLD = 1, " \
                            "CST_ALERT_CHG_DATE = GETDATE(), " \
                            "CST_ALERT_CHG_USERID = 16, " \
                            "CST_ALERT_TEXT = 'ON HOLD DUE TO OVERDUE " \
                            "ACCOUNT: $" + total_owing + " to be paid " \
                            "immediately. This balance reflects the total " \
                            "of invoices due over " + days_string + " days " \
                            "ago. For more info, see MYOB. To bypass this " \
                            "hold, go to CRM and remove the on hold status " \
                            "temporarily. This message and status will " \
                            "disappear when the account is up to date. ~~" \
                            + kept_notes + "' " \
                            "WHERE CST_ACCOUNT_NUMBER = '" + customer_code + "'"

                if DEBUG:
                    print("Update query to be executed: ", query)

                else:
                    cursor.execute(query)
                    connection.commit()

    connection.close()

    print("Customers updated. Summary stored in '" + filename_sum
          + "' located in folder " + folder_location)

    if DEBUG:
        print("----- END OF PROGRAM -----")
