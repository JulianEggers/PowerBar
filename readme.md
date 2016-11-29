# PowerBar
PowerBar is a Visual Basic for Applications (VBA) Powerpoint-Macro for including a progress bar into a Powerpoint presentation.
The macro creates a progress bar at the bottom of each slide. Additionally the progress of the presentation is printed as percentage on the progress bar.

## Preparation of Powerpoint

1. Save your presentation as .pptm (Powerpointpresentation with Macros) (Otherwise the macro will bot be saved.)
2. [Activate Developer Options](https://support.office.com/en-us/article/Show-the-Developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45#ID0EAABAAA=2016,_2013,_2010)

## Create the Macro
1. Go into the Developer Options
2. Click on ``Visual Basic`` to open the Visual Basic Editor
3. Import the [PowerBar.bas](PowerBar.bas) file or create a new Module ``Insert -> Module`` and paste the code of the PowerBar.bas file into into the Editor. If you copy and paste the code you have to remove the first line ``Attribute VB_Name = "PowerBar"``
4. Save the module and exit the editor


## Options
Some changes can be made in the code of the add-on:
- change the color
- skip the first pages
- change position and size of the progress bar

## Run Macro
1. Go into the Developer Options
2. Click on ``Macros``
3. Run ``RefreshPowerbar`` to remove the Powerbar (if already created) and recreate it
