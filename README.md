# radiosonde_flight_visualization
Excel VBA macro for visualization of [radiosondy.info](https://radiosondy.info) weather balloon flight data

# Introduction
I started this project to get first insights into Excel VBA. Some 'real life data' and a target to achieve are always a good propulsion – so I started to build a workbook to visualize weather balloon flight data. 
For sure, it is not a masterpiece in coding, but for beginners in data analysis it might be a good starting point. The charts generated in the workbook may give you a better idea about what is going on up in the air or you may just find some helpful lines in the macro for your own projects.

As a source, the workbook is dealing with archived data from [radiosondy.info](https://radiosondy.info). This site is showing crowd-sourced data from weather balloons up in the air. 
The macro pulls apart the text strings of the recorded balloon data, does some juggling and formatting steps and finally it prints charts for visualization.

The post-processed data could been used now as an input for other tools, models or coding projects, like flight or weather prediction. The charts might be interesting for your projects or presentations.

I tested the macro with hundreds of flights. Most of the time it worked fine, but sometimes errors came up especially when generating the charts. Excel occasionally struggles to copy & paste named objects. A second try usually works. Just download the Excel file, get an idea how it works, feel free to improve and adjust the macro or the charts to suite your needs. Most of the steps in the code have comments or links to know where it is coming from.
