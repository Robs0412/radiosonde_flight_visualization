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
First download the latest Excel file and some CSV example. Later you can also download CSV flight data via the macro or via [radiosondy.info](https://radiosondy.info). As you may have guessed the example CSV fligh data file are from chases I participated successfully with others or together with my wife. All the unlucky chases are not included of course ;-)

For running the macro, please ensure your Excel can execute it. You may get one of the following warnings which you can accept:

#### Warnings:
![Activate_Macros.PNG](__used_asset__/Activate_Macros.PNG)

![Activate_Macros_2.PNG](__used_asset__/Activate_Macros_2.PNG)


#### The main window:
![Main_window.png](__used_asset__/Main_window.png)

#### The result after some clicks:
![Example_output.png](__used_asset__/Example_output.png)

We'll get there in a minute...

### Elements:
---
![Overview_Files.PNG](__used_asset__/Overview_Files.PNG)

File paths and CSV file name section - it gives you an idea which flights will be imported into separate sheets. It can contain files on the local system or from [radiosondy.info](https://radiosondy.info). The drive location or the link will be filled automatically by the 'CSVs from drive' or 'Web-Load CSV' functions.

---
![Buttons_I.PNG](__used_asset__/Buttons_I.PNG)
- Select local CSVs (multiple) for adding to the file paths / name section, Excel will open the folder where the macro is located
- 'Delete file paths' will kick out all elements in file path / name section
- 'Web-Load CSV' will generate a link according to the flight identifier, please stick to the spelling rules
- 'Download CSV' will store the CSV of a specific flight (same folder as macro), the CSV will not be added automatically to the file paths list
- 'Open radiosondy.info' will open [radiosondy.info](https://radiosondy.info) in your standard browser

Hint:
- Avoid gaps between files in the file paths / links section as the macro will bring up an error later while running the import
- When deleting files / links use the 'Delete file paths' button
- When using the 'Web-Load' or 'Download' function, please ensure that Excel can access the internet (check firewall,...)



