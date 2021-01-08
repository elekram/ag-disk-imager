# AG Disk Imager 🌄

The AG - all good - disk imager is a PowerShell script which makes it simple to image Windows machines
with a PowerShell enabled Windows Preinstallation Environment (WinPE), and is handy for when you don't want
complicated solutions that cost money or require lots of infrastructure. Simply grab a USB hardrive and get down to business.
  
## Table of Contents

1. [Usage](#Usage)
2. [Introduction](#Introduction)
3. [Docs](#Docs)

## Usage

Drop ag-disk-imager.ps1 into a folder with with a **manifest.json** (see below for details), **wim** and optional **drivers**, **unattend** and **custom-disk-layout** folders then simply run the script. AG will generate required folders as required and guide the user depending on how the manifest.json is configured. At a minimum, you will need to provide a Windows image and reference it in the **manifest.json** within a task (see manifest.json example). Drivers, unattend, and custom diskpart script files are optional.

## Introduction

AG uses a JSON file as its manifest where tasks and machine models are defined. WMIC is used to detect the manufacturer model from the machine's firmware and then match it with the model defined in the manifest file. A menu for all defined tasks will be generated from all the named tasks under the 'tasks' key in the manifest.json. You can then define machine specific tasks for machine models under the 'models' key in the manifest.json, the first task listed here will be the default task, and all tasks associated with a machine model will then be highlighted for convenience when AG runs.
  
**AG Imager will:**

* Clean and partition the disk for you on the fly (supports UEFI and BIOS)
  * Ensuring never to overwrite your external boot drive for machines which enumerate their disks in an odd order
* If desired, inject drivers from a local repository on the USB drive after imaging
* Also inject an unattend.xml file of your choice into the completed image if required
* Cleanup stale windows bootloader entries
* Detect and write the correct bootloader (UEFI/BIOS)

## Docs

ag-disk-imager.ps1 needs to be in a directory with a file named **manifest.json** and three folders: a **wim** folder, a **drivers** folder and an **unattend** folder. On the first run AG will check for, and if needed, create the folders required for you. If the task defined in the manifest.json contains a 'drivers' key, the value of that key will be the name of the folder AG looks for inside the drivers folder. AG will look for a folder that has the name of the detected machine model under the folder you defined with the 'drivers' key in the manifest. It's recommended that you define windows version for the 'drivers' key example 1903, 1909 etc. (this allows driver sets for different vesions of Windows to be organised). You will then be asked to place exported drivers (or SCCM driver packs) into this folder, though you can option not to inject drivers if you so choose by omitting the 'drivers' key in the task object in the manifest. At a minimum the 'wim' key is required for a task to run.
  
## manifest.json (see example below)

For convenience AG will attempt to generate a custom menu highlighting recommended tasks for a specific machine model. It does this by comparing the detected machine model (from the devices firmware) against model names listed under the 'models' key in the **manifest.json**. Each model name is a key that expects an array. The array should be populated with task names from the 'tasks' key in the **manifest.json**. The first task name in the array will be the default task. All tasks in the array will be highlighted in the AG task menu when the script first runs, making it easy for the user to identify tasks for that device. Below is an example of a valid **manifest.json**

```json
{
  "tasks": {
    "SOE-Win10-1909-Enterprise": {
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
      "SOE-Win10-1909-Enterprise",
      "Win10 ThinkCenter-M710"
    ]
  }
}
```

## Drivers

Place exported driver inf files into their respective model folder names in the drivers folder matching the folder defined using the 'drivers' key in the **manifest.json** folder. See Microsoft's [Export-WindowsDriver cmdlet](https://docs.microsoft.com/en-us/powershell/module/dism/export-windowsdriver?view=win10-ps) for more information. Alternatively you can just use DISM or download SCCM packages from your devices vendor. Note: a 'drivers' key is required in the manifest.json under the task in order to inject drivers.

## Unattend

Place unattend.xml files into the unattend folder and then define them in your manifest.json for each task. You can name the files anything you like while in the unattend folder and AG will copy the defined file (in the manifest) to the windows/panther directory after imaging renaming it to unattend.xml. If you leave this value blank or don't define the key at all AG will skip this step all together.

## Custom Disk Layouts (diskpart)

Place diskpart script files (diskpart /s [filename]) into the 'custom-disk-layouts' folder and specify them using the **disklayout** key in the **manifest.json** see manfiest.json example above. If the specified sciprt is found AG will override the default disk/partition layout with the one found in the script.