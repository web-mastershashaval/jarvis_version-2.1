import os

def display_available_networks():
    """Display available WiFi networks."""
    print("Scanning available WiFi networks...")
    os.system('cmd /c "netsh wlan show networks interface=Wi-Fi"')

def create_new_connection(name, SSID, password):
    """Create a new WiFi connection profile."""
    config = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{name}</name>
    <SSIDConfig>
        <SSID>
            <name>{SSID}</name>
        </SSID>
    </SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM>
        <security>
            <authEncryption>
                <authentication>WPA2PSK</authentication>
                <encryption>AES</encryption>
                <useOneX>false</useOneX>
            </authEncryption>
            <sharedKey>
                <keyType>passPhrase</keyType>
                <protected>false</protected>
                <keyMaterial>{password}</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>"""
    
    # Write the configuration to an XML file
    xml_file = f"{name}.xml"
    with open(xml_file, 'w') as file:
        file.write(config)
    
    # Add the profile to the system
    command = f"netsh wlan add profile filename=\"{xml_file}\" interface=Wi-Fi"
    os.system(command)

def connect(name, SSID):
    """Connect to the specified WiFi network."""
    command = f"netsh wlan connect name=\"{name}\" ssid=\"{SSID}\" interface=Wi-Fi"
    os.system(command)

def main():
    """Main function to execute WiFi operations."""
    display_available_networks()
    
    name = input("Enter the name of the Wi-Fi network you want to connect to: ")
    password = input("Enter the password for the Wi-Fi network: ")
    
    create_new_connection(name, name, password)
    connect(name, name)
    
    print("Attempting to connect. If you aren't connected, ensure the credentials are correct and try again.")

if __name__ == "__main__":
    main()
