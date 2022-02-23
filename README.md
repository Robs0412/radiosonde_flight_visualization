## radiosonde_flight_visualization
Excel VBA macro for visualization of [radiosondy.info](https://radiosondy.info) weather balloon flight data

### Introduction:
I started this project to get first insights into Excel VBA. Some 'real life data' and a target to achieve are always a good propulsion – so I started to build a workbook to visualize weather balloon flight data. 
For sure, it is not a masterpiece in coding, but for beginners in data analysis it might be a good starting point. The charts generated in the workbook may give you a better idea about what is going on up in the air or you may just find some helpful lines in the macro for your own projects.

As a source, the workbook is dealing with archived data from [radiosondy.info](https://radiosondy.info). This site is showing crowd-sourced data from weather balloons up in the air. 
The macro pulls apart the text strings of the recorded balloon data, does some juggling and formatting steps and finally it prints charts for visualization.

The post-processed data could been used now as an input for other tools, models or coding projects, like flight or weather prediction. The charts might be interesting for your projects or presentations.

I tested the macro with hundreds of flights. Most of the time it worked fine, but sometimes errors came up especially when generating the charts. Excel occasionally struggles to copy & paste named objects. A second try usually works. Just download the Excel file, get an idea how it works, feel free to improve and adjust the macro or the charts to suite your needs. Most of the steps in the code have comments or links to know where it is coming from.

### Let's get started:
Just download the latest Excel file and some CSV example. Later you can also download CSVs with flight data via the macro or via [radiosondy.info](https://radiosondy.info). As you may have guessed the example CSVs are from chases I participated successfully together with other seekers or with my wife. All unlucky chases are not included of course ;-)


##### Warnings:

![Activate_Macros.PNG](__used_asset__/Activate_Macros.PNG)

![Activate_Macros_2.PNG](__used_asset__/Activate_Macros_2.PNG)

For running the macro, please ensure your Excel can execute it. You may get one of those warnings which you can accept.

---

##### Information:

![Information_splash_screen.PNG](__used_asset__/Information_splash_screen.PNG)

When starting up the macro it will already execute some actions. It will set the 'decimal separator' to '.' (dot). In case you are used to it anyway this will not bother you. For countries like Germnay which are used to ',' (comma) this affect your typing actions.
The change is just affecting this particular visualization workbook. All other workbooks opened in parallel will use normal system settings.
You can edit the start-up features in VBA editor (Alt+F11) in the general workbook layer (deactivate the splash screen, force other actions, ...).

Attention: the macro will definitely change your calculation setting. While using the functions it will switch between automatic/manual mode but will stay in automatic at the end. Check 'Formula' ribbon and calculation options to change back to your prefered setup.

---

##### The main window:
![Main_window.png](__used_asset__/Main_window.png)

The elements of the main screen in this 'Import' sheet will we explained below.

---

##### The result after some clicks:
![Example_output.png](__used_asset__/Example_output.png)

We'll get there in a minute... :-)

---

### Elements:

![Overview_Files.PNG](__used_asset__/Overview_Files.PNG)

File paths and CSV file name section - it gives you an idea which flights will be imported into separate sheets. It can contain files on the local system or from [radiosondy.info](https://radiosondy.info). The drive location or the link will be filled automatically by the 'CSVs from drive' or 'Web-Load CSV' functions.

---

![Buttons_I.PNG](__used_asset__/Buttons_I.PNG)

- 'Select CSVs from drive' (multiple) for adding to the file paths / name section, Excel will open first the folder where the macro is located.
- 'Delete file paths' will kick out all elements in file path / name section.
- 'Web-Load CSV' will generate a link according to the flight identifier, please stick to the spelling rules.
- 'Download CSV' will store the CSV of a specific flight (same folder as macro), the CSV will not be added automatically to the file paths list.
- 'Open radiosondy.info' will open [radiosondy.info](https://radiosondy.info) in your standard browser.

Hints:
- Avoid gaps between files in the file paths / links section as the macro will bring up an error later while running the import.
- When deleting files / links use the 'Delete file paths' button.
- When using the 'Web-Load' or 'Download' function, please ensure that Excel can access the internet (check firewall,...).

---

![Buttons_II.PNG](__used_asset__/Buttons_II.PNG)

- 'Import CSVs' will import all CSVs which are listed in the file paths section. The import is in raw format with a seperate sheet for each flight.
- 'Delete sheets' will kick out all generated sheets except the 'Import' main sheet.

Hints:
- If a flight is already imported it will not be imported once more (no duplicated sheets).
- The main 'Import' sheet should always stay at first tab position as some actions are depending on that. Please don't move it to another tab location later.

---

![Buttons_III.PNG](__used_asset__/Buttons_III.PNG)

- 'Process CSVs' will do the juggling and formatting of the raw data, some extra collumns and information will be added. Will be applied to all imported flights
- 'Draw charts' will generate charts in each imported & processed sheet based on the preconfigured charts in main 'Import' sheet
- 'Search sheet' will help you to find a specific sheet in those many parallel imported flights (optional: use small arrows left of the 'Import' tab)

Hints:
- Information will pop up if a sheet is empty (no data to be process) or if you want to draw charts without processing first
- Already processed sheets or sheets with charts inserted can't be processed again - the function is checking a flag inside the sheets
- Please keep the preconfigured charts in main 'Import' sheet unchanged, simple formatting is possible (lines, axis,...)
- Avoid changes of chart object names or size except you already figured out how the macro is working
- 'Draw charts' causes sometimes errors due to internal objects copy/paste flaws. Try the function a second time or just scroll once through whole preconfigured charts section to help Excel preload the objects.

---

![Overview_ExcelUI.PNG](__used_asset__/Overview_ExcelUI.PNG)

If you like the workbook in a 'cleaned' app style try the buttons to toggle between full or reduced UI. The option can also be activated during start-up of the macro, see VBA editor (Alt+F11) in the general workbook layer section.

Hint:
- Make sure you switch back to 'Show' in the end to get back ribbons and the rest

---
