# Variables
$driverDownloadPath = "https://downloads.canon.com/bicg2024/drivers/Generic_Plus_UFRII_v3.00_Set-up.exe"
$portAddress = "10.0.7.10" # Printer IP
$printerName = "Library Copier BW" 
$driverName = "Canon Generic Plus UFR II V300" 

# Do not modify these variables
$driverDownloaded = "C:\Technology\Temp\PrinterDriver.exe"
$extractPath = "C:\Technology\Temp\PrinterDriver"
$driverPath = "C:\Technology\Temp\x64\etc"
$portName = "IP_$portAddress"
$portExists = Get-Printerport -Name $portName -ErrorAction SilentlyContinue
$printerExists = Get-Printer -Name $printerName -ErrorAction SilentlyContinue

# Remove old printer if it exists
if ($printerExists) {
    Remove-Printer -Name $printerName
}

# Create local storage folder
New-Item -ItemType Directory -Force -Path C:\Technology\Temp\PrinterDriver

# Download the driver
Invoke-WebRequest $driverDownloadPath -OutFile $driverDownloaded

# Run the self-extracting archive
Start-Process -FilePath $driverDownloaded -ArgumentList "/S -InstallPath=$extractPath -y" -Wait

# Add to Windows Driver Store
Get-ChildItem $driverPath -Recurse -Filter "*.inf" -Force | ForEach-Object { PNPUtil.exe /add-driver $_.FullName /install }

# Add Driver
Add-PrinterDriver -Name $driverName

# Add Printer Port
if (-not $portExists) {
    Add-PrinterPort -Name $portName -PrinterHostAddress $portAddress
}

# Install Printer
if (-not $printerExists) {
    Add-Printer -Name $printerName -PortName $portName -DriverName $driverName
}

# Set Default Color to Grey Scale
Set-PrintConfiguration -PrinterName $printerName -Color $false

# Set as default printer
(Get-WmiObject -Class Win32_Printer | Where-Object -Property Name -EQ $printerName).SetDefaultPrinter()

# Delete downloaded files
Remove-Item -Path $driverDownloaded -Force
Remove-Item -Path $extractPath -Force -Recurse
