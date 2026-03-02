#==========================================
# IMPORTS
#==========================================
import gspread #interact with the sheet
from google.oauth2.service_account import Credentials 

scopes = [ # accessing using google apis
    "--insert google api link--"
]

#==========================================
# CONFIG
#==========================================
creds_file = "credentials.json" # refrencing credential file from google api
sheet_id = "--insert sheet ID--"

# SHEET TABS===============================
master_sheet_name = "MEB AV Equipment" # master sheet
vlan_prefix_sheets = "AV DHCP" # vlan sheets prefix

# mastersheet columns
room_col = "313"
MAC_col = "MAC"
IP_address_col = "IP Address"
host_name_col = "Host Name"


#==========================================
# AUTHORIZATION
#==========================================
def get_client():
    creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
    return gspread.authorize(creds)


#==========================================
# LOAD MASTER INVENTORY
#==========================================
def load_inventory(sheet):
    master = sheet.worksheet(master_sheet_name)

    # get ALL data
    data = master.get_all_values()

    headers = data[0]   # first row = headers
    rows = data[1:]     # everything else = data

    inventory_by_ip = {}

    for row in rows:
        row_dict = dict(zip(headers, row))

        ip = str(row_dict.get(IP_address_col, "")).strip()

        if not ip:
            continue

        inventory_by_ip[ip] = {
            "313": row_dict.get(room_col, ""),
            "MAC": row_dict.get(MAC_col, ""),
            "Host Name": row_dict.get(host_name_col, ""),
        }

    return inventory_by_ip


#==========================================
# POPULATE AV DHCP Tabs
#==========================================
def populate_vlan_sheets(sheet, inventory_by_ip):
    network_tabs = [
        ws for ws in sheet.worksheets()
        if ws.title.startswith(vlan_prefix_sheets)
    ]

    for network_sheet in network_tabs:
        print(f"Processing {network_sheet.title}")

        ips = network_sheet.col_values(1)  # column A = IPs
        updates = []  # collect all updates for batch

        for row_index, ip in enumerate(ips, start=1):
            if row_index <= 5 or not ip:
                continue

            ip = ip.strip()

            if ip in inventory_by_ip:
                device = inventory_by_ip[ip]

                # append the update instead of writing immediately
                updates.append({
                    "range": f"B{row_index}:D{row_index}",
                    "values": [[
                        device["MAC"],        # Column B
                        device["Host Name"],  # Column C
                        device["313"]         # Column D
                    ]]
                })

        # send all updates at once
        if updates:
            network_sheet.batch_update(updates)


#==========================================
# MAIN
#==========================================
def main():
    client = get_client()
    sheet = client.open_by_key(sheet_id)

    inventory_by_ip = load_inventory(sheet)
    populate_vlan_sheets(sheet, inventory_by_ip)

# execute
if __name__ == "__main__":
    main()

