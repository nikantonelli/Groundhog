GroundHog Day
=============

## Overview
An attempt to illustrate the progression of stories across a LeanKit board over time. The program 
considers which updates to do each "day". (A day is really only an increment of a counter.)

## Setup

The program can be run as a command line program, or as one that is called periodically as a cron job. 
The program requires a spreadsheet to list the changes that you want to make.

If you just list the changes as a list of 'Create' entries, then effectively you have a card importer
(make sure you use the '-o' option). Bear in mind that this was not written for that purpose and 
might be a little awkward in use.

The spreadsheet requires three sheets: 'Config', 'Changes', and a third one named whatever you like 
to help you organise the change info. In my example file (Scrum Team.xlsx), the third sheet is named 
'Level 0-0'

The program will attempt to prompt you on how to build the right info in the spreadsheet, but if 
it can't work it out, then you will need to refer to the example one in this repo. The logs are 
sent out using Sytem.out from Java - wherever that might go in your system.....

## Command Line Options

-f \<file\>     

the XLSX spreadsheet to read from

-c            

use this if the program is to be repetitively called by a cron job (needs -s option)

-s \<file\>     

the name of the status file to use to keep track of what day the program is on when using the -c option

-o           

run through the list once

-u \<rate\>           

'day' update rate in seconds. Must be used without '-c' option. A rate of zero seconds will 
              make the code wait for user input before going to a days updates

-d \<mode\>    

delete all user created cards (not from spreadsheet) where mode is 'day' (at the start
              of every day processing), 'cycle' (at the end of the cycle)

-m \<lane\>     

move all cards to this lane at end of cycle. This also clears out the ID field in
              the spreadsheet so that the next cycle creates new items
              DO NOT move cards to a lane that has a WIP limit as that doesn't make any sense here.
              If you try to, the program will put in a wipOverrideComment of
              "Archiving boards from GroundHog cycle" for you.


## Example command line
java -jar target\Groundhog-2.0-spring-boot.jar -f "Scrum Teams.xlsx" -c -s a.txt

If you do not use the -c or -u options then the program will attempt to sit in a sleep loop being 
woken up every (24 hours at) 3 in the morning.

If you use -c, then the program will attempt to make (and initialise) the file in the directory it 
is running in with the (-s) name you supply. If it cannot, then it will fail completely

## Config data

The 'Config' sheet must contain  4 columns with titles: url, username, password, apikey, cyclelength 
in Row 1. Usage depends on whether you are using an apiKey or a username/password pair. You don't 
have to supply both. The next valid row in the sheet will be used as the required config.

## Changes

The program looks on the Changes sheet for the column titles: 
  
'Day Delta', 'Item Sheet', 'Item Row', 'Action', 'Value1' and 'Value2' in Row 1.
  
The Day Delta is the day from the start of the program that you want the change to occur, the 
'Item Sheet' and 'Item Row' identify the artifact that needs updating. The Value1 column is used 
for plain fields, but those updates that require more info make use of Value1 and Value2. Hopefully, 
there are some explanations in the example spreadsheet.
  
Changes are made sequentially. This program is not built for speed! The reason for sequential is 
that after a 'Create', the Id must be written back into the spreadsheet so that a subsequent change
can use it as a Parent. This will not work with parallel creations.

Changes are written back after every one to protect against network or program failures. In this way, 
re-running the program should pick up from where it left off without re-creating a whole bunch of stuff.

## Artifact Info

The other sheets that require artifact definitions can have any number of plain fields defined
for the creation event. The first two columns must have titles ID and 'Board Name'. The ID 
is blank at the start and when the artifact is created, the program wil write the id of the one
created back into the ID field. The 'Board Name' is obviously where the artifact resides. The ID 
can be used for further updates (e.g. add Parent). If you want to re-run the create, then just 
delete the ID field contents for that row and run the command to go through for just the 'day'
that is set for that row. You can set the row to 99 and then run with the options:
  "-f <filename> -u 0 -b 99 " to run it manually.

If you put a field in that is not part of the standard Card field set, then it is checked in the 
list of Custom Fields for the board. At the moment, the type of Custom field is not checked, so 
might barf at some point. 

If at any time you see a persistent card on your board which is of the default type and has a title of 
"dummy title", then you can delete it. It was due to a failure during the two-stage creation that the
app uses. You probably had a network error or used Ctrl-C at the wrong time.
  
Log an issue if something doesn't work.

## LICENCE
This code is supplied as-is with no support, implied or otherwise. You are free to use it as you see 
fit as far as I am concerned. If you use it, you are responsible for the consequences. I will 
do best endeavours to fix stuff, but I do have a day job and that takes precedence.

The code makes use of the API of Leankit so there is intellectual property involved there. Make sure you 
are covered as it is your responsibility.
