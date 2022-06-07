README-NOTES.txt

Trying to record some notes on the project/code/game

Build Notes:

- Needs Visual Basic Power pak: http://go.microsoft.com/fwlink/?LinkID=145727&clcid=0x804
Severity	Code	Description	Project	File	Line	Suppression State
Error	BC30002	Type 'Microsoft.VisualBasic.PowerPacks.LineShape' is not defined.	innovation	\innovation\main.Designer.vb	228	Active

- and of course that doesn't work needs version 3.0
https://stackoverflow.com/questions/67288928/are-there-alternatives-to-the-microsoft-visual-basic-power-packs

- Tried to add an icon file and that didn't work....
Running environment notes:

The apps needs

Dim filename As String = ".\Innovation.txt" ' Holds all the card data



Issue#5 Created fix for missing data file Innovation.txt

See the end of this message for details on invoking 
just-in-time (JIT) debugging instead of this dialog box.

************** Exception Text **************
System.IndexOutOfRangeException: Index was outside the bounds of the array.
   at Innovation.Main_Renamed.launch_game(Object& players)
   
See the end of this message for details on invoking 
just-in-time (JIT) debugging instead of this dialog box.

************** Exception Text **************
System.IndexOutOfRangeException: Index was outside the bounds of the array.
   at Innovation.Main_Renamed.launch_game(Object& players)
   
           Dim Msg, Style, Title, Response, MyString
        Msg = "Do you want to continue ?"    ' Define message.
        Style = vbYesNo Or vbCritical Or vbDefaultButton2    ' Define buttons.
        Title = "MsgBox Demonstration"    ' Define title. 
        ' Display message.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then    ' User chose Yes.
            MyString = "Yes"    ' Perform some action.
        Else    ' User chose No.
            MyString = "No"    ' Perform some action.
        End If
