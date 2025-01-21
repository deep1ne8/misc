# Import the CSV file containing MAC addresses and vendor information
$Records = Import-CSV "C:\Oui.csv"

# Define the target PC MAC address to search for
$PCMAC = '00-01-C8'

# Iterate through each record in the CSV
ForEach ($Record in $Records) {
    # Check if the MAC address matches the target PC MAC address
    if ($Record.MAC -eq $PCMAC) {
        # Retrieve and output the vendor associated with the MAC address
        $Vendor = $Record.Vendor
        Write-Host $Vendor
    }
}
