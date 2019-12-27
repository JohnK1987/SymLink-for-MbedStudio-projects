# SymLink-for-MbedStudio-projects
This script can manage the MbedOs library for possibility make a new project Offline and also make a project without a additional download of the library approx 1GB for each new project. That is achieved thank to use of Symbolic or hard links in the Windows applied to the MbedOs library folder and its mbed-os.lib file.

# How to use:
* In the Mbed Studio make a new project (Empty or Bare metal or Blinky) with name SOURCE and uncheck the checkbox about "Set project as active"
* Create a new folder. For example C:\MbedSource.
* Move this vbscript (SLfMS.vbs) and the SOURCE project from the MbedStudio Workspace to your new folder what you made in previous step.
* Run the vbscript and follow instructions in boxes.

This is my first VBS script and certainly not faultless but for my personal use it looks functional.
