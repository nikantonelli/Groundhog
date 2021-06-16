GroundHog Day
=============

## Overview
An attempt to illustrate the progression of stories across a LeanKit board over time.

## Setup

The program can be run as a long term command line program, or as one that is called periodically 
as a cron job. The program requires a spreadsheet to list the changes that you want to make.

If you just list the changes as a list of 'Create' entries, then effectively you have a card importer.

The spreadsheet requires three sheets: 'Config', 'Changes', and a third one named whatever you like to 
help you organise the change info. In my example file (Scrum Team.xlsx), the third sheet is named 'Level 0-0'

The program will attempt to prompt you on how to build the right info in the spreadsheet, but if it can't work it out,
then you will need to refer to the example one in this repo. The logs are sent out using Sytem.out from Java - wherever
that might go in your system.....

## Command Line Options

-f      the XLSX spreadsheet to read from

-c      use this if the program is to be repetitively called by a cron job (needs -s option)

-s      the name of the status file to use to keep track of what day the program is option

(-m and -d not implemented yet)

java -jar target\Groundhog-1.0-SNAPSHOT-spring-boot.jar -f "Scrum Teams.xlsx" -c -s a.txt

If you do not use the -c option then the program will attempt to sit in a sleep loop being woken up
every (24 hours at) 3 in the morning.

If you do, then the program will attempt to make (and initialise) the file in the directory it is running in with the name you 
supply. If it cannot, then it will fail completely

## Config data

The 'Config' sheet must contain  4 columns (with the first row as the header). Usage depends on whether you are using an apiKey 
or a username/password pair. You don't have to supply both

    url, username, password, apikey, cyclelength
