<div align="center">

## EZ INI Module to read / write INI Files


</div>

### Description

Offers 6 ini functions that are extremely useful. Handy, because it has a "setup" function that specifies the filename and the number of characters you'ld like to retrieve (max) so that you don't have to re-specify them every time you write to your INI file (great if you only have one or two INI files per project). Someone else did a "complete INI module" that actually hard-code edited the ini file to find out the number of sections, and so on, but this one uses in-built features of GetPrivateProfileString. Have fun!

<br><br>

Read_Ini -- Returns a string with the appropriate INI key in it

<br><br>

Write_Ini -- Writes to a specific ini key

<br><br>

Read_Sections -- Returns a Chr(0) delimited string containing all the [Sections] of the INI file

<br><br>

Read_Keys -- Same as above, but returns all the keys=values under the appropriate section

<br><br>

Delete_Key and Delete_Section do just what they say =)

<br><br>Note:<br>

vbCrLf's are turned into a funky string because I noticed when writing vbcrlf's to ini files it has the tendency to screw up and actually place them on separate lines so you can't retrieve the whole thing again. The funky string is recognized by this module when you read the values in again, and converts em back to vbCrLf
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-01-15 00:44:00
**By**             |[Shroom](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shroom.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[CODE\_UPLOAD137871152001\.zip](https://github.com/Planet-Source-Code/shroom-ez-ini-module-to-read-write-ini-files__1-14430/archive/master.zip)








