# Python-Excel-Scraper

This file is a simple excel scraper that pulls in data from a directory that contains Excel files and stores it in an array for future use. 

Currently, this scraper is being used to pull in routing decision ratios from intersections on the Purple Route for Tiger Transit on the campus of Clemson University. The data is pulled into an array that separates it by file and sheet (thus leaving the data mostly in the same format that it is in Excel). 

The scraper also includes code to separate out the data based on keywords in the file name. Right now, this functionality is being used to differentiate between data collected on Monday, Wednesday, and Friday or on Tuesday and Thursday.
