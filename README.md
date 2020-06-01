# AG Disk Imager 🌄

The AG - all good - disk imager is a PowerShell script which makes it simple to image Windows machines
with a PowerShell enabled Windows Preinstallation Environment (WinPE), and is handy for when you don't want
complicated solutions that cost money or require lots of infrastructure. Simply grab a USB hardrive and get down to business.
  
## Table of Contents

1. [Usage](#Usage)
2. [Introduction](#Introduction)
3. [Docs](#Docs)

## Usage

Drop ag-disk-imager.ps1 into a folder with with a **manifest.json** (see below for details), **wim**, **drivers** and **unattend** folders then simply run the script. AG will generate these folder and guide the user on its first run depending on how the manifest.json is configured. You just need to provide the Windows images.

## Introduction

AG Uses a JSON file as its manifest where machine models are defined. WMIC is used to detect the manufacturer model from the machine's firmware and then match it with the model defined in the manifest file. This is then used to generate a menu pointing to defined image tasks within the manifest for the user to select when imaging devices.
  
**AG Imager will:**

* Clean and partition the disk for you on the fly (supports UEFI and BIOS)
  * Ensuring never to overwrite your external boot drive for machines which enumerate their disks in and odd order
* If desired, inject drivers from a local repository on the USB drive after imaging
* Also inject an unattend.xml file of your choice into the completed image if required.
* Detect and write the correct bootloader (UEFI/BIOS)

## Docs

ag-disk-imager.ps1 needs to be in a directory with a file named 'manifest.json' and two folders: a 'wim' folder, a 'drivers' folder and an 'unattend' folder. On the first run AG will check for, and if needed, create the folders required for you. A folder with the machine's model name (found in the firmware) will be created in the drivers folder. You will then be asked to place exported drivers into this folder.
  
## manifest.json

AG needs the computer model(s) listed in a manifest.json as key values, this value will be matched to the machines model number in its firmware (see JSON example below). It's from this information that AG will generate a menu for the user to select from when the script is first run.  

Each machine model listed will contain an object which must contain a key for the device's full name, named 'name', a key called 'config' which defines the tasks which get dynamically added to the program menu, and a 'tasks' key which lists the different imaging tasks and their options. The 'tasks' key then needs another nested object with 'wim', 'drivers' and 'unattend' keys. See example below for a clearer understanding.  

**Note:** the defaults key in the manifest does not need to be nested under 'tasks' as it doesn't require additional values.

```json
{
  "defaults": {
    "Win10-1709-Enterprise": {
      "wim": "win10-ent-1709-soe.wim",
      "drivers": true,
      "unattend": "win10-unattend.xml"
    },
    "Factory Image": {
      "wim": "",
      "drivers": false,
      "unattend": ""
    }
  },
  "20DAS0L00": {
    "name": "Lenovo ThinkPad 11e Gen2",
    "config": "defaults"
  },
  "10M8S3CF00": {
    "name": "Lenovo ThinkCentre m710s",
    "config": "replace",
    "tasks": {
      "Win10-Desktop-SOE": {
        "wim": "win10-ent-1909-soe.wim",
        "drivers": false,
        "unattend": ""
      }
    }
  }
}
```

The 'config' attribute can be set to *defaults*, *append* or *replace*. If set to *append* AG will list the 'defaults' tasks and also machine specific tasks specified under the machine's model name, while *replace* will only list the machine specific tasks. If the config key is set to 'defaults' only the tasks from the defaults list will be shown. 

## Drivers

Place exported driver inf files into their respective model folder names in the 'drivers' folder. See Microsoft's [Export-WindowsDriver cmdlet](https://docs.microsoft.com/en-us/powershell/module/dism/export-windowsdriver?view=win10-ps) for more information. Alternatively you can just use DISM.  
  
So for the 'Acer TravelMate P449-M' model listed in the aforementioned example manifest you would have a folder named 'Acer TravelMate P449-M' in the *drivers* folder with your exported drivers inside. If you don't wish to inject drivers simply set the 'driver' key to **false** in its task.

## Unattend

Place unattend.xml files into the unattend folder and then define them in your manifest.json for each task. You can name the files anything you like while in the unattend folder and AG will copy the defined file (in the manifest) to the windows/panther directory after imaging renaming it to unattend.xml. If you leave this value blank AG will not try to inject an unattend file.