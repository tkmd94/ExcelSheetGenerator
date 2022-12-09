# ExcelSheetGenerator
 
ExcelSheetGenerator" is an ESAPI script that outputs data to a specified cell in an Excel template file and saves it.

# Features

* The following data can be output to the specified cell of an Excel template file and saved.
  * Patient ID
  * Patient name
  * Course ID
  * Plan ID
  * Dose per fraction
  * No. of fraction
  * Total dose
  * PlanApproval status
  * Isocenter position
  * Structure volume
* Folder name and file name can be specified based on patient/plan information.
* Settings can be written in XML format and multiple conditions can be selected and executed in a dialog box.
* Can be executed even on a terminal without Excel installed.

* Applications
  * Automatic input of patient information into check sheets created for each patient.

# Demo

! [Screen capture of planCompare UI](https://github.com/tkmd94/ExcelSheetGenerator/blob/master/demo.gif)

# Requirement

* Eclipse 15.6 or higher (older versions are not checked)

# Installation
1. Create an Excel template file. 
2. Place [ExcelSheetGeneratorPreference.xml] and change the path at line 35 of [MainWindow.xaml.cs].

    ````PreferencePath = @"D:\Dshare\ExcelSheetGenerator\ExcelSheetGeneratorPreference.xml";```
    
3. modify the settings in [ExcelSheetGeneratorPreference.xml].
    * templatename: Name to be displayed in the dialog
    * templatePath: Template file path
    * exportPath: directory path to save
    * filename: Name of the file to save
    * type: Data type (info/DVH)
    * data: Name of the data
    * structure: Contour name to output data (type:DVH,data:Volume only)
    * address: Location of the cell where the data will be output

4. Build the project and generate the DLL file [ExcelSheetGenerator.esapi.dll]. 
5. Copy the generated DLL file to the folder specified in [Tools] -> [Script] on the Eclipse toolbar.
6. place the Excel template file in the template file path set in [2].

# Usage

***Use this source code at your own risk. **

1. run a script from [Tools] -> [Script] on the Eclipse toolbar 
2. Select a condition from the dialog window and click [Generate] button.
3. When the script execution completes, the destination directory will be displayed, and confirm the created file.
 
# Author
 
* Takashi Kodama
 
# License
 
"ScreenCapture_eDoc" is under [MIT License](https://en.wikipedia.org/wiki/MIT_License).
