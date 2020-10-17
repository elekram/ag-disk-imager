# AG Disk Imager 🌄

The AG - all good - disk imager is a PowerShell script which makes it simple to image Windows machines
with a PowerShell enabled Windows Preinstallation Environment (WinPE), and is handy for when you don't want
complicated solutions that cost money or require lots of infrastructure. Simply grab a USB hardrive and get down to business.
  
## Table of Contents

1. [Usage](#Usage)
2. [Introduction](#Introduction)
3. [Docs](#Docs)

## Usage

Drop ag-disk-imager.ps1 into a folder with with a **manifest.json** (see below for details), **wim**, **drivers** and **unattend** folders then simply run the script. AG will generate required folders as required and guide the user on its first run depending on how the manifest.json is configured. At a minimum, you will need to provide a Windows image and reference it in the **manifest.json**. Drivers and unattend files are optional. 

You can also specify a custom diskpart script file (diskpart /s [filename]) using the optional **disklayout** key for a task in the **manifest.json**. Script files need to be placed in a folder called 'custom-disk-layouts'. AG will check the folder for the specified file and if found apply the custom disk/partition layout.

## Introduction

AG Uses a JSON file as its manifest where tasks and machine models are defined. WMIC is used to detect the manufacturer model from the machine's firmware and then match it with the model defined in the manifest file. This is then used to generate a menu pointing to defined image tasks within the manifest for the user to select when imaging devices. If AG doesn't find a reference to the detected machine model under the 'models' key in the manifest it will generate a menu with all available tasks under the 'tasks' key in the manifest.
  
**AG Imager will:**

* Clean and partition the disk for you on the fly (supports UEFI and BIOS)
  * Ensuring never to overwrite your external boot drive for machines which enumerate their disks in an odd order
* If desired, inject drivers from a local repository on the USB drive after imaging
* Also inject an unattend.xml file of your choice into the completed image if required
* Cleanup stale windows bootloader entries
* Detect and write the correct bootloader (UEFI/BIOS)

## Docs

ag-disk-imager.ps1 needs to be in a directory with a file named **manifest.json** and three folders: a **wim** folder, a **drivers** folder and an **unattend** folder. On the first run AG will check for, and if needed, create the folders required for you. If the task contains a 'drivers' key set with a boolean value set 'true' a folder with the machine's model name (found in the devices firmware) will be created in the drivers folder under the windows version you define using the 'version' key in the manifest (this allows driver sets for different vesions of Windows to be organised). You will then be asked to place exported drivers (or SCCM driver packs) into this folder, though you can option not to inject drivers if you so choose by settings the 'drivers' key to false or by simply not adding these keys to the task object in the manifest. At a minimum the 'wim' key is required for a task to run.
  
## manifest.json (see example below)

AG needs the computers you want to image model(s) listed in a **manifest.json** as keys under the 'models' key in order to generate a custom tasks menu for that device. The models key will be matched to the machine's model number found in its firmware. Each model needs an array with the task names defined matching the 'tasks' object in the **manifest.json** (see JSON example below). If AG can't match the detected machine model with an entry in the in the **manifest.json** under 'models' then it will generate you a menu with all tasks listed under the 'tasks' key from the **manifest.json**.

```json
{
  "tasks": {
    "SOE-Win10-1709-Enterprise": {
      "wim": "win10-ent-1709-soe.wim",
      "drivers": "1709",
      "unattend": "unattend.xml"
    },
    "Win10-1909": {
      "wim": "l420_7856_4cm_naplan.wim",
      "unattend": "unattend.xml"
    },
    "Win10 ThinkCenter-M710 Image": {
      "wim": "tc_m710s.wim",
      "disklayout": "Custom-Diskpart-Script.txt"
    }
  },
  "models": {
    "20DAS02L00": [
      "SOE-Win10-1709-Enterprise"
    ],
    "20G90003AU": [
      "SOE-Win10-1709-Enterprise",
      "Win10 ThinkCenter-M710"
    ]
  }
}
```

## Drivers

Place exported driver inf files into their respective model folder names in the drivers folder matching the version folder from the **manifest.json** folder. See Microsoft's [Export-WindowsDriver cmdlet](https://docs.microsoft.com/en-us/powershell/module/dism/export-windowsdriver?view=win10-ps) for more information. Alternatively you can just use DISM or download SCCM packages from your devices vendor. Note: a 'drivers' key is required in the manifest.json under the task in order to inject drivers.

## Unattend

Place unattend.xml files into the unattend folder and then define them in your manifest.json for each task. You can name the files anything you like while in the unattend folder and AG will copy the defined file (in the manifest) to the windows/panther directory after imaging renaming it to unattend.xml. If you leave this value blank or don't define the key at all AG will skip this step all together.

## Custom Disk Layouts (diskpart)

Place diskpart script files (diskpart /s [filename]) into the 'custom-disk-layouts' folder and specify them using the **disklayout** key in the **manifest.json** see manfiest.json example above. If the specified sciprt is found AG will override the default disk/partition layout with the one found in the script.