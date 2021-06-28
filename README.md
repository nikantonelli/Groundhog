Importer
=============

## Overview
An attempt to import groups of changes listed in an Excel spreadsheet.  For simplistic import of items,
thisis straightforward. The problem comes when you want to add connections between items. The reference
to another item is done through its ID. The ID must exist before you can use it.

To get around this problem you can use the formula capability of a cell to point to the ID cell of the 
parent record. Every time an item is created, the program writes the ID back in and updates all the 
formulas

All you have to do, as a user, is to make sure that the formulas are correct..... good luck!

## Setup

The program can be run as a command line program and requires a spreadsheet to list the changes that 
you want to make.

The spreadsheet requires three sheets: 'Config', 'Changes', and a third one named whatever you like 
to help you organise the change info. In my example file (Scrum Team.xlsx), the third sheet is named 
'Level 0-0'

The program will attempt to prompt you on how to build the right info in the spreadsheet, but if 
it can't work it out, then you will need to refer to the example one in this repo. The logs are 
sent out using Sytem.out from Java - wherever that might go in your system.....

## Command Line Options

-f \<file\>     

the XLSX spreadsheet to read from

-g \<group\>  Allows you to group the changes into blocks

## Example command line
java -jar target\Groundhog-2.0-spring-boot.jar -f "Scrum Teams.xlsx" -g 2

## Config data

The 'Config' sheet must contain  4 columns with titles: url, username, password, apikey, cyclelength 
in Row 1. Usage depends on whether you are using an apiKey or a username/password pair. You don't 
have to supply both. The next valid row in the sheet will be used as the required config.

## Changes

The program looks on the Changes sheet for the column titles: 
  
'Group', 'Item Sheet', 'Item Row', 'Action', 'Value1' and 'Value2' in Row 1.
  
The Group identifier is used to select row of changes. The 'Item Sheet' and 'Item Row' identify the artifact 
that needs updating. 

NORE: 'Create' entries only allow for fields that require a simple data entry. Those fields that need two values
to be entered (e.g. externalLink), currently, need to be separated with a comma ",". The externalLink field therefore
must not have any comma characters in its label. The Lane field may require a overrideWipComment. If so, add this to
the end of the Lane, once again, separated by a comma ",". See the example spreadsheet.

Value1 and Value2 are only used if you want to add subsequent 'Modify' rows to the Changes sheet. 
For example, you might want to create a set of items and then at a later date add some updates to the
same items. As the IDs are already stored, you can stream a set of updates directly into the items

DO NOT import into a board with a lane WIP setting that will be exceeded by your creates - the 
program will barf at you. If you want to add to Lanes, but add a overrideWipComment, use the Modify 
capability of the changes sheet and add the Lane into Value1 and the commment into Value2

You cannot put the externalLink into the Create, it must be done with a Modify change entry. This
is because it requires two values as the same time.
  
Changes are made sequentially. This program is not built for speed! The reason for sequential is 
that after a 'Create', the Id must be written back into the spreadsheet so that a subsequent change
can use it as a Parent. This will not work with parallel creations.

Any updates are written back after every one to protect against network or program failures. In this way, 
re-running the program should pick up from where it left off without re-creating a whole bunch of stuff.

## Artifact Info

The other sheets that require artifact definitions can have any number of plain fields defined
for the creation event. The first two columns must have titles ID and 'Board Name'. The ID 
is blank at the start and when the artifact is created, the program wil write the id of the one
created back into the ID field. The 'Board Name' is obviously where the artifact resides. The ID 
can be used for further updates (e.g. add Parent). If you want to re-run the create, then just 
delete the ID field contents for that row and run the command to go through for just the 'group'
that is set for that row. You can set the row to 99 and then run with the options:
  "-f <filename> -g 99 " to run it manually.

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
