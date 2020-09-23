VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAO Example"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert new record"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear ListView"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read records"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Width           =   7455
      Begin MSComctlLib.ListView ListView1 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":0442
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   145
         Width           =   7215
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7455
      Begin VB.Label lblDBRS 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'you have to declare the variables you use
Public ItemIndex As Integer 'index of a row in listview

' Sample Access DAO Project - by Rob t.H. - ottooliebol@hotmail.com
'               DAO = Data Access Objects
' This example explains how to use DAO and a listview control
'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-
' This Application uses an Access 2000 Database. To work with this database, you
' need to use the Microsoft DAO 3.6 Object Library DLL!
' You can find it here: Menu Project, References
' --
' For Microsoft Office 2000 (Access, this example) you need version 3.6 of the DAO DLL,
' for Microsoft Office 97 (Access) you need version 3.51 of the DAO DLL.
' Version 3.6 will NOT work with Office 97 and below!!
' So with Project References, only use ONE DAO DLL!!!

Private Sub cmdRead_Click()
' if there is data in the listview, clear it first before populating it.
' we just simulate a click on the Clear ListView1 button
' (never type 2 times the same code if 1 time can also do the job :-))
If ListView1.ListItems.Count > 0 Then cmdClear_Click


' now we declare the variables we use with the database stuff.
' DB is the database itself, RS is the recordset (in this case: table People)
' --
' for this stuff to work, we need the Microsoft DAO 3.6 Object Library DLL!!
Dim DB As Database
Dim RS As Recordset

' Open database DAOtest.mdb (in the same path as the application files are)
Set DB = OpenDatabase(App.Path + "\DAOtest.mdb")
' Open recordset (table: People) in database DB
Set RS = DB.OpenRecordset("People")


' If there are no records in the table: (EOF = End Of File)
If RS.EOF Then
' display text in the label (The Name Property is the name of the table (People)
    lblDBRS.Caption = "There were no records found in table " & RS.Name
Else
' We found some records!!
' display in the label how many records we found
    lblDBRS.Caption = "There are " & RS.RecordCount & " records found in table " & RS.Name

    ' Now we are going to read the records from the table.
    ' We use a loop for this, read until End Of File
    
    ' We use the integer 'a', so we know in which row we are writing (listview)
    Dim a As Integer
    a = 1
    
    Do Until RS.EOF
        
        ' The first column (name) is always the columnheader as we call this,
        ' the rest (address, telephone) are subitems of the columnheader
        ' --
        ' So we got the first column, this is ListView1.ListItems.
        ' The others are ListView1.ListItems.SubItems(index)
        ' We must know were to write the subitems, that's where we use 'a' for.
        ' Everytime we go trough this loop we increase 'a' with 1.
        ' That means, if you want, the next row.
        ' --
        ' A field from e.g. a table can be used with RS!Fieldname.
        ' If a field contains a space, you must use RS("Field Name")
        ' With RS we mean the recordset, in this case table: People.
        ListView1.ListItems.Add , , RS!Name
        ListView1.ListItems(a).ListSubItems.Add , , RS!Address
        ListView1.ListItems(a).ListSubItems.Add , , RS!Telephone
        
        ' We must know which row contains the data from which row from the table People.
        ' Therefore we copy the ID from the table to the Tag of the row which contains
        ' the data from that row in the table. Sorry if this sounds complicated,
        ' but I'm not that good in teaching LOL :-)
        ListView1.ListItems(a).Tag = RS!ID
        
        ' While we are copying, we use a progressbar (pb) to show the progress
        pb.Value = Int(RS.PercentPosition)
        
        ' In the form caption, show the progress of reading the table in percentages
        Form1.Caption = "Busy with reading from table " & RS.Name & ": " & Int(RS.PercentPosition) & "%"
        ' We use Int(RS.PercentPosition), Int means a number without decimals.
                
        ' Increase 'a' (our progress counter so we know in which row we are writing)
        ' with 1
        a = a + 1
        
        ' Move to the next row in the recordset (table: People)
        RS.MoveNext
        
    ' Do the same stuff again, until we reached the End Of File
    Loop

' Ok, we completed reading from the table, now we reset the progressbar
pb.Value = 0

' And we reset the form caption
Form1.Caption = "DAO Example - Table: " & RS.Name

End If

'we are done with the recordset and the database, so we close them now
RS.Close
DB.Close

End Sub

Private Sub cmdInsert_Click()
' Now we want to insert a new record to the table People.
' We are going to show another form to do this.
frmInsert.Show

'we have to reset ItemIndex, otherwise in some cases the ItemIndex will be remembered
ItemIndex = 0

End Sub

Private Sub cmdDelete_Click()
' the user wants to delete the selected record.
' if there are no records in the listview: exit sub
If ListView1.ListItems.Count = 0 Then Exit Sub

'if itemindex = 0 then there is nothing selected!
If ItemIndex <> 0 Then
    'ok, there is something selected!
    'now we will get the name from the items name, and ask for delete confirmation
    Dim Ask As String
    Ask = MsgBox("Are you sure that you want to delete '" & ListView1.ListItems.Item(ItemIndex).Text & "'?", vbYesNo + vbInformation, "Delete record")
        
    If Ask = vbYes Then
        'user had pressed yes, please delete ;-)
        'now we have to get the ID of the row in de database (item's tag!) and delete record from table
    
        Dim DB As Database
        Dim RS As Recordset
        
        ' Open database DAOtest.mdb (in the same path as the application files are)
        Set DB = OpenDatabase(App.Path + "\DAOtest.mdb")
        ' Open recordset (table: People) in database DB
        Set RS = DB.OpenRecordset("People")
        
        'we will seek the record with ID equal to item's tag
        'first we set the table's index (this is the ID)
        'we need this because otherwise we can't use the seek function
        RS.Index = "ID"
         
        'we move the recordpointer to the first record, this way we can seek the whole table
        RS.MoveFirst
            
        'here we check if the seek functions result is equal to the item's tag.
        'the item's tag contains the ID from the table.
        'we stored it when we were reading from the table
        RS.Seek "=", ListView1.ListItems.Item(ItemIndex).Tag
    
        If RS.NoMatch Then
            'if there was no match (the ID couldn't be found in the table)
            MsgBox ("The record can't be found in the table!"), vbOKOnly + vbCritical, "Error"
            
            
            'we reset itemindex to 0, this means that nothing is selected in listview1
            ItemIndex = 0
            
            'close recordset and database
            RS.Close
            DB.Close
            
            Exit Sub
        Else
            'there was a match! we will now delete the record from the database
            RS.Delete
            
            'and we will delete the row from the listview
            ListView1.ListItems.Remove (ItemIndex)
        End If
           
        'close recordset and database
        RS.Close
        DB.Close
    
    End If

'we reset itemindex to 0, this means that nothing is selected in listview1
ItemIndex = 0
    
End If


            
End Sub

Private Sub cmdClear_Click()
'clear the listview
ListView1.ListItems.Clear

'we have to reset ItemIndex, otherwise in some cases the ItemIndex will be remembered
ItemIndex = 0

'set caption to text no table opened
lblDBRS.Caption = "No table opened at the moment"
'set form caption
Form1.Caption = "DAO Example"


End Sub

Private Sub Form_Load()
'build the listviews column headers
Dim Header As ColumnHeader

'build header Name
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Name"
    Header.Width = ListView1.Width * 0.29
    ' width of the header is 29% of the total width of listview1.
    ' I choose 29%, not 30%, because with 29% the horizontal scrollbar will disappear.
    ' So in total I use 99% and not 100%

'build header Address
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Address"
    Header.Width = ListView1.Width * 0.4
    ' width of the header is 40% of the total width of listview1
    
'build header Telephone
Set Header = ListView1.ColumnHeaders.Add()
    Header.Text = "Telephone"
    Header.Width = ListView1.Width * 0.3
    ' width of the header is 30% of the total width of listview1
    
' Set the caption of the label
lblDBRS.Caption = "No table opened at the moment"

'we have to reset ItemIndex, so there is nothing selected in listview1
ItemIndex = 0

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
ItemIndex = Item.Index

End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'if you right click on a row on the populated listview, the contents must
'appear in a messagebox on the screen.
'--
'if you left click, then the item you clicked on must be selected
'(this goes automaticlly)
'--

If Button = vbRightButton Then
    'Right mouse button is clicked
    '--
    'first we check if there is data in the listview, otherwise our application will crash.
    'if there is no data, exit this sub
    If ListView1.ListItems.Count < 1 Then Exit Sub
    
    'here we get all the items and even the tag which contains the ID from the table
    '--
    'btw, vbCrLf is a carriage retun + line feed. The same as enter or return.
    'It means a new line :-)
    MsgBox ("Name: " & ListView1.SelectedItem.Text & vbCrLf & _
            "Address: " & ListView1.SelectedItem.SubItems(1) & vbCrLf & _
            "Telephone: " & ListView1.SelectedItem.SubItems(2) & vbCrLf & _
            "ID: " & ListView1.SelectedItem.Tag) _
            , vbOKOnly + vbInformation, "Details"
End If

'we reset itemindex to 0, this means that nothing is selected in listview1
ItemIndex = 0
    
End Sub
