#==========================================
# IMPORTS
#==========================================
import gspread # Read/write in Google Sheet 
from google.oauth2.service_account import Credentials # Authenication via Google API account credentials

scopes = [ # Defined Google API permissions
    "--insert google api link--"
]

#==========================================
# CONFIG
#==========================================
creds_file = "--insert credentials json--" # Google serviec account credentials json
sheet_id = "--insert sheet ID--" # ID of the google Sheet to access

# SHEET TABS===============================
master_sheet_name = "MEB AV Equipment" # Master sheet inventory name
vlan_prefix_sheets = "AV DHCP" # Prefix for all VLAN sheets

# Mastersheet column headers
room_col = "313"
MAC_col = "MAC"
IP_address_col = "IP Address"
host_name_col = "Host Name"


#==========================================
# AUTHORIZATION
#==========================================
""" Parameter(s): This function doesn't require any parameters
    Function: Authenticating Google Sheets from service accoumt.
    Returns: A gspread client object to interact with sheets. """
def get_client():
    creds = Credentials.from_service_account_file(creds_file, scopes=scopes)
    return gspread.authorize(creds) # Authorize and return client


#==========================================
# LOAD MASTER INVENTORY
#==========================================
""" Parameter(s): Needs sheet to open and retrieve IP data.
    Function: Read all rows from the master sheet and creates a dictionary mapping IP
    addresses to device info.
    Returns: Dictionary of IP's."""
def load_inventory(sheet):
    master = sheet.worksheet(master_sheet_name) # Open master sheet

    data = master.get_all_values() # Get all data
    headers = data[0]   # First row = headers
    rows = data[1:]     # Everything else = data

    inventory_by_ip = {} # Dictionary to store info by IP

    for row in rows:
        row_dict = dict(zip(headers, row)) # Map each row into dictionary {header: value}

        ip = str(row_dict.get(IP_address_col, "")).strip() # Get IP address for the row

        if not ip: # Skip rows with no IP
            continue

        inventory_by_ip[ip] = { # Store device info by IP
            "313": row_dict.get(room_col, ""), # Room
            "MAC": row_dict.get(MAC_col, ""), # MAC
            "Host Name": row_dict.get(host_name_col, ""), # Hostname
        }

    return inventory_by_ip # Return dictionary of IP's


#==========================================
# POPULATE AV DHCP Tabs
#==========================================
""" Parameter(s): Sheet, and Dictionary of IP's from load_inventory
    Function: Goes through all VLAN sheets, reads the IP's, and updates
    the corresponding columns with MAC, Host Name, and Room info.
    Uses batch updates to reduce API calls.
    Return: This function doesn't return. """
def populate_vlan_sheets(sheet, inventory_by_ip):
    network_tabs = [ # Get all sheets with the VLAN prefix
        ws for ws in sheet.worksheets()
        if ws.title.startswith(vlan_prefix_sheets)
    ]

    for network_sheet in network_tabs:
        print(f"Processing {network_sheet.title}") # Print VLAN sheet that is being proccessed

        ips = network_sheet.col_values(1)  # Column A = IP addresses
        updates = []  # List collecting bactch updates

        for row_index, ip in enumerate(ips, start=1): # Loop through al rows
            if row_index <= 5 or not ip: # Skip header rows (row <= 5) or empty cells
                continue

            ip = ip.strip() # Clean whitespace

            if ip in inventory_by_ip: # Update only if IP exists in master inventory
                device = inventory_by_ip[ip] # Get device info

                # Append the update instead of sending indiviudal updates
                updates.append({
                    "range": f"B{row_index}:D{row_index}", # Columns B-D                  "values": [[
                        device["MAC"],        # Column B
                        device["Host Name"],  # Column C
                        device["313"]         # Column D
                    ]]
                })

        # Send all updates in a single batch to reduce API quota usage
        if updates:
            network_sheet.batch_update(updates)


#==========================================
# MAIN
#==========================================
""" Parameter(s): No parameters needed.
    Function: Main entry point: authenticate, load master inventory,
    and updates all VLAN sheets.
    Return: This function doesn't return. """
def main():
    client = get_client() # Authenticate gpsread client
    sheet = client.open_by_key(sheet_id) # Open Google Sheet by its ID

    inventory_by_ip = load_inventory(sheet) # Load master inventory into dictionary
    populate_vlan_sheets(sheet, inventory_by_ip) # Update VLAN sheets


#==========================================
# EXECUTE SCRIPT
#==========================================
if __name__ == "__main__":
    main() # Run main function

