GroundHog Day
=============

## Overview
An attempt to illustrate the progression of stories across a LeanKit board over time. The program considers which updates to do each "day". (A day is really only an increment of a counter.)

## Setup

The program can be run as a long term command line program, or as one that is called periodically as a cron job. The program requires a spreadsheet to list the changes that you want to make.

If you just list the changes as a list of 'Create' entries, then effectively you have a card importer (make sure you use the '-o' option). Bear in mind that this was not written for that purpose and might be a little awkward in use.

The spreadsheet requires three sheets: 'Config', 'Changes', and a third one named whatever you like to help you organise the change info. In my example file (Scrum Team.xlsx), the third sheet is named 'Level 0-0'

The program will attempt to prompt you on how to build the right info in the spreadsheet, but if it can't work it out, then you will need to refer to the example one in this repo. The logs are sent out using Sytem.out from Java - wherever that might go in your system.....

## Command Line Options

-f      the XLSX spreadsheet to read from

-c      use this if the program is to be repetitively called by a cron job (needs -s option)

-s      the name of the status file to use to keep track of what day the program is on when using the -c option

-o      run through the list once

-u      'day' update rate in seconds. Must be used without '-c' option. A rate of zero seconds will make the code wait for user input before going to a days updates

(-m and -d not implemented yet)

java -jar target\Groundhog-1.0-SNAPSHOT-spring-boot.jar -f "Scrum Teams.xlsx" -c -s a.txt

If you do not use the -c option then the program will attempt to sit in a sleep loop being woken up every (24 hours at) 3 in the morning.

If you do, then the program will attempt to make (and initialise) the file in the directory it is running in with the (-s) name you supply. If it cannot, then it will fail completely

## Config data

The 'Config' sheet must contain  4 columns with titles: url, username, password, apikey, cyclelength in Row 1. Usage depends on whether you are using an apiKey or a username/password pair. You don't have to supply both. The next valid row in the sheet will be used as the required config.

## Changes

The program looks on the Changes sheet for the column titles: 'Day Delta', 'Item Sheet', 'Item Row', 'Action', 'Value1' and 'Value2' in Row 1.
The Day Delta is the day from the start of the program that you want the change to occur, the 'Item Sheet' and 'Item Row' identify the artifact that needs updating. The Value1 column is used for plain fields, but those updates that require more info make use of Value1 and Value2. Hopefully, there are some explanations in the example spreadsheet.

## Artifact Info

The other sheets that require artifact definitions can have any number of plain fields defined for the creation event. The first two columns must have titles ID and 'Board Name'. The ID is blank at the start and when the artifact is created, the program wil write the id of the one created back into the ID field. The 'Board Name' is obviously where the artifact resides. The ID can be used for further updates (e.g. add Parent). If you want to re-run the create, then just delete the ID field contents for that row and run the command to go through for just the 'day' that is set for that row. You can set the row to 99 and then run with the options "-f <filename> -u 0 -b 99 " to run it manually.

If you put a field in that is not part of the standard Card field set, then it is checked in the list of Custom Fields for the board. At the moment, the type of Custom field is not checked, so might barf at some point. Log an issue if something doesn't work.


