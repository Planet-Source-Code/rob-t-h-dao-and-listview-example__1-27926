VERSION 5.00
Begin VB.Form frmInsert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAO Example - Insert new record"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3645
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
      Begin VB.TextBox txtTelephone 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtAddress 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Telephone"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
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
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
' The user wants to cancel, so we simply unload this form from memory
Unload Me

End Sub

Private Sub cmdInsert_Click()
' OK, the user wants to insert the data to the table People
' But before we do that, we want to check if everything is filled in!

If txtName.Text = "" Then MsgBox ("Please fill in the name first!"), vbOKOnly + vbCritical, "Error" _
: Exit Sub

If txtAddress.Text = "" Then MsgBox ("Please fill in the address first!"), vbOKOnly + vbCritical, "Error" _
: Exit Sub

If txtTelephone.Text = "" Then MsgBox ("Please fill in the telephone number first!"), vbOKOnly + vbCritical, "Error" _
: Exit Sub

' OK, everything seems ok! Now we open a connection to the database

' Open database DAOtest.mdb (in the same path as the application files are)
Set DB = OpenDatabase(App.Path + "\DAOtest.mdb")
' Open recordset (table: People) in database DB
Set RS = DB.OpenRecordset("People")

'Now we are going to add a new record to the recordset! :-)
With RS 'use recordset
    
    .AddNew 'add new record
    
    !Name = txtName.Text 'field name must contain content of txtname
    !Address = txtAddress.Text 'field address must contain content of txtaddress
    !Telephone = txtTelephone.Text 'field telephone must contain content of txttelephone
    
    .Update 'As soon as we use RS.Update, then the data is written to the table.

End With

' now we show a little messagebox, showing the user that the new record is added
MsgBox ("The new record is added to the database!"), vbOKOnly + vbInformation, "Added"

' Close recordset and database
RS.Close
DB.Close

' and now close the form
Unload Me

End Sub

Private Sub Form_Load()
' we will set the max characters allowed for the textboxes, so the user can't type
' more characters then we can save to the table
txtName.MaxLength = 20
txtAddress.MaxLength = 35
txtTelephone.MaxLength = 10

End Sub

Private Sub txtTelephone_KeyPress(KeyAscii As Integer)
' we want the user to enter a phone number, so no other characters then numbers are
' allowed! The ascii range for numbers is 47..57.
' If a user presses a key different then a number, we just not write it in the textbox!
' Only exception is Ascii Char 8, this is backspace :-)

If KeyAscii <> 8 And KeyAscii < 47 Or KeyAscii > 57 Then KeyAscii = 0: Exit Sub

End Sub
