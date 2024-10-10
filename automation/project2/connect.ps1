# Define the SSID and Password
$SSID = "GWJ3404C_50411C39BBB6"
$Password = "YjhlNmM5M2"

# Check if the profile exists
$profile = netsh wlan show profiles | Select-String $SSID

if ($profile) {
    Write-Host "Profile for $SSID already exists."
} else {
    # Create a new profile XML file
    $xmlProfile = @"
<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>$SSID</name>
    <SSIDConfig>
        <SSID>
            <name>$SSID</name>
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
                <keyMaterial>$Password</keyMaterial>
            </sharedKey>
        </security>
    </MSM>
</WLANProfile>
"@

    # Save the XML to a temporary file
    $profileFile = "$env:TEMP\wifiprofile.xml"
    $xmlProfile | Out-File -Encoding utf8 -FilePath $profileFile

    # Add the profile to the system
    netsh wlan add profile filename=$profileFile

    # Clean up the temporary file
    Remove-Item $profileFile
}

# Connect to the Wi-Fi network
netsh wlan connect name=$SSID
