<center><h1>AG Disk Imager 🌄</h1></center>

<center>The AG - all good - disk imager is a PowerShell script which makes it simple to image Windows machines
with a <b>PowerShell enabled</b> Windows Preinstallation Environment (WinPE), and is handy for when you don't need
complicated solutions and infrastructure that cost lots of money. Simple grab a USB hardrive to get down to business.</center>
  
# Table of Contents
1. [Usage](#Usage)
2. [Introduction](#Introduction)
3. [Docs](#Docs)

# Usage
Drop ag-disk-imager.ps1 into a folder with with a **manifest.json** (see below for details), **wim** and **drivers** folder and then simply run the script.

# Introduction
AG Uses a JSON file as its manifest where machine models are defined. AG uses WMIC to detect the manufacturer model from the machine's firmware and then match it with the model defined in the manifest file. This is then used to generate a menu pointing to the defined WIM files and also ensures correct drivers are injected.
  
**AG Imager will:**
* Clean and partition the disk for you on the fly (supports UEFI and BIOS)
  * Ensuring not to overwrite your external boot drive for machines which enumerate their disks in and odd order
* Inject drivers from a local repository on the USB drive after imaging
* Detect and write the correct bootloader (UEFI/BIOS)

# Docs
ag-disk-imager.ps1 needs to be in a directory with a file named *manifest.json* and two folders: a *wim* folder, and a *drivers* folder. On the first run AG will check for, and if needed, create the folders required for you. A folder with the machine's model name (found in the firmware) will be created in the drivers folder. You will then be asked to place exported drivers into this folder.
  
## manifest.json
AG needs the computer model(s) listed in a manifest.json as key values, the value for each model is an object which can contain a key for an image name and value for its wim file name (located in the wim folder), or just set the key name to 'config' and point it to 'defaults' see below for an example. 

```json
{
  "defaults": {
    "Win10-Home": "1909-home-soe.wim",
    "Win10-Enterprise": "1909-ent-soe.wim"
  },
  "20DAS0L00": {
    "config":"defaults"
  },
  "Acer TravelMate P449-M": {
    "config": "append",
    "Win-10-1709": "1709-ent-soe.wim"
  }, 
  "Surface Laptop 2": {
    "config": "replace",
    "Win-10-2004": "nameme.wim"
  }
}
```
The 'config' attribute can also be set to *append* or *replace*. If set to *append* AG will list the 'defaults' wims and also machine specific WIMS specified under the machine's model name, while *replace* will only list the machine specific WIM options. 
## Drivers
Place exported driver inf files into their respective model folder names in the 'drivers'f folder. See Microsoft's [Export-WindowsDriver cmdlet](https://docs.microsoft.com/en-us/powershell/module/dism/export-windowsdriver?view=win10-ps) for more information. Alternatively you can just use DISM.  
    
So for the 'Acer TravelMate P449-M' model listen in the aforementioned exmaple manifest you would have a folder named 'Acer TravelMate P449-M' in the *drivers* folder with your exported drives inside.