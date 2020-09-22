<div align="center">

## VB Tech Tips 2


</div>

### Description

Just some VB Tech Tips, nothing fancy, but informative.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Clint LaFever](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/clint-lafever.md)
**Level**          |Beginner
**User Rating**    |4.8 (76 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/clint-lafever-vb-tech-tips-2__1-31492/archive/master.zip)





### Source Code

<p><b><small><font face="Verdana">Take Advantage of Related Documents Area In Project Window</font></small></b></p>
<p><small><font face="Verdana">If you use a resource file in your application, you can see the RES file
appear in the project window under &quot;Related Documents.&quot;&nbsp; This is
the only type of file that VB automatically adds to this node of the project
tree.&nbsp; You can add any type of file you like to this area manually,
though.&nbsp; From the Project menu, selected Add File, or right click on the
project window and select Add File from the context menu.&nbsp; In the dialog
box, select All Files for the file type and check the Add As Related Document
option.&nbsp; Adding additional related files here helps organize your project
and gives you quick access to useful items, including design documents,
databases, resource scripts, help project files, and so on.&nbsp; Once a file
has been added, double-click on it in the project window to open it with the
appropriate application.</font></small></p>
<hr>
<p><small><b><font face="Verdana">Browse VB Command as You Type</font></b></small></p>
<p><small><font face="Verdana">When you refer to an object in VB, you get a dropdown list of that object's
properties and methods.&nbsp; But, did you know that the statements and functions
of the VB language can be pulled up in the same way.&nbsp; You can view the list
as you type in one of two ways.&nbsp; One (which just shows how it all works) is
to type VBA. then the list will appear.&nbsp; There you can see the list off all
VB functions and have it filter down as you type.&nbsp; The quicker way is to
just press CTRL+SPACE prior to typing your VB function/command.&nbsp; i.e.;&nbsp;
On a blank line press CTRL+SPACE then type ms&nbsp; At this point it should be
at MsgBox.&nbsp; While yes most commands are short enough that the CTRL+SPACE
does not really save you any time, but one you will not having any typos and
two, it will help you remember/find a call that you do not use much.</font></small></p>
<hr>
<p><small><b><font face="Verdana">Use the Watch Window to Drill Down into Objects/Collections During Debug</font></b></small></p>
<p><small><font face="Verdana">All of know about the immediate window, but I find very few developers who
know of the Watch Window.&nbsp; The watch window is a very nice tool to drill
down any any variable whether it is a standard type or an object or a
collection.&nbsp; For an example, open on of your database project where you
open a recordset.&nbsp; Set a break point after you open your recordset.&nbsp;
Run your code.&nbsp; When you hit your break point, highlight your recordset
variable right there on that line (double click on it too for quicker
highlighting).&nbsp; Now right click your highlighted variable and choose Add
Watch.&nbsp; On then next window click Ok.&nbsp; Presto your recordset variable
show now be displayed in the Watch Window.&nbsp; You can expand it out and drill
down to all the properties within.&nbsp; Not only is this a good tool to use to
inspect your objects and collections at runtime, but also a good teaching tool
to help those trying to understand objects and how they are constructed.</font></small></p>
<hr>
<p><small><b><font face="Verdana">Show the Standard File Properties Dialog</font></b></small></p>
<p><small><font face="Verdana">If your program has an Explorer shell-style interface, you probably want to
supply the standard File/Properties dialog.&nbsp; Do this by using the
ShellExecuteEX API function:</font></small></p>
<p><small><font face="Courier New">Private Type SHELLEXECUTEINFO<br>
&nbsp;&nbsp;&nbsp; cbSize as Long<br>
&nbsp;&nbsp;&nbsp; fMask as Long<br>
&nbsp;&nbsp;&nbsp; hWnd as Long<br>
&nbsp;&nbsp;&nbsp; lpVerb as String<br>
&nbsp;&nbsp;&nbsp; lpFile as String<br>
&nbsp;&nbsp;&nbsp; lpParameters as String<br>
&nbsp;&nbsp;&nbsp; lpDirectory as String<br>
&nbsp;&nbsp;&nbsp; nShow as Long<br>
&nbsp;&nbsp;&nbsp; hInstApp as Long<br>
&nbsp;&nbsp;&nbsp; lpIDList as Long<br>
&nbsp;&nbsp;&nbsp; lpClass as String<br>
&nbsp;&nbsp;&nbsp; dwHotKey as Long<br>
&nbsp;&nbsp;&nbsp; hIcon as Long<br>
&nbsp;&nbsp;&nbsp; hProcess as Long<br>
End Type</font></small></p>
<p><small><font face="Courier New">Private Declare Function ShellExecuteEX Lib _<br>
&nbsp;&nbsp;&nbsp; &quot;shell32&quot; (lpSEIAs SHELLEXECUTEINFO) As Long<br>
Private Const SEE_MASK_INVOKELIST=&amp;HC</font></small></p>
<p><small><font face="Courier New">Private Sub ShowFileProperties(ByVal aFile as
String,h as Long)<br>
&nbsp;&nbsp;&nbsp; Dim sei as SHELLEXECUTEINFO<br>
&nbsp;&nbsp;&nbsp; sei.hWnd=h<br>
&nbsp;&nbsp;&nbsp; sei.lpVerb=&quot;properties&quot;<br>
&nbsp;&nbsp;&nbsp; sei.lpFile=aFile<br>
&nbsp;&nbsp;&nbsp; sei.fMask=SEE_MASK_INVOKEIDLIST<br>
&nbsp;&nbsp;&nbsp; sei.cbSize=len(sei)<br>
&nbsp;&nbsp;&nbsp; ShellExecuteEX sei<br>
End Sub</font></small></p>
<p><small><font face="Verdana">Please note I typed this directly in here and not in the IDE so there may be
typos.</font></small></p>
<hr>
<p><small><b><font face="Verdana">Start Up in Your Code Folder</font></b></small></p>
<p><small><font face="Verdana">For the shortcut you use to open VB, change the Start In in the Properties of
the shortcut to point to the folder you prefer to have as the default Open and
Save to start at.&nbsp;</font></small> </p>
<hr>
<p><small><b><font face="Verdana">Trick the P&amp;D Wizard</font></b></small></p>
<p><small><font face="Verdana">Do you have external files to your application that you want to make sure are
always included with your application when you go and create a new setup for
it.&nbsp; Just use this little trick:</font></small></p>
<p><small><font face="Courier New">#If False Then<br>
&nbsp;&nbsp;&nbsp; Private Declare Sub Foo Lib &quot;VIDEO.AVI&quot; ()<br>
#End If</font></small></p>
<p><small><font face="Verdana">VB will ignore this statement but the P&amp;D Wizard will not.&nbsp; The
P&amp;D Wizard will pick up this line and also remember to add this file to the
list of files required for your application.</font></small></p>
<p><small><font face="Verdana">I know this is not really all that useful, but it is a nice trick.</font></small></p>
<hr>

