VERSION 5.00
Begin VB.Form RegBook 
   Caption         =   "My PhoneBook"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   5370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Backup 
      Caption         =   "Backup Phonebook"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   4560
      Width           =   5055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Find User Entry"
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   5175
      Begin VB.CommandButton Search 
         Caption         =   "Find Name"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox FindSurName 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox FindName 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   $"RegBook.frx":0000
         Height          =   855
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label7 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Surname:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Phone/address"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton SaveInfomation 
         Caption         =   "Save New"
         Height          =   255
         Left            =   4080
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox Phone2 
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Text            =   "Incomplete"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Address 
         Height          =   285
         Left            =   840
         TabIndex        =   5
         Text            =   "Incomplete"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.TextBox Phone1 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Text            =   "Incomplete"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox SurName 
         Height          =   285
         Left            =   3240
         TabIndex        =   2
         Text            =   "Incomplete"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Name1 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "Incomplete"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Surname:"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Phone 2:"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Phone 1:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "RegBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub SaveInfomation_Click()
WriteReg "HKLM\Software\PhoneBook\" & SurName.Text & "\" & Name1.Text & "\" & "Name1", Name1.Text
WriteReg "HKLM\Software\PhoneBook\" & SurName.Text & "\" & Name1.Text & "\" & "Name2", SurName.Text
WriteReg "HKLM\Software\PhoneBook\" & SurName.Text & "\" & Name1.Text & "\" & "Phone1", Phone1.Text
WriteReg "HKLM\Software\PhoneBook\" & SurName.Text & "\" & Name1.Text & "\" & "Phone2", Phone2.Text
WriteReg "HKLM\Software\PhoneBook\" & SurName.Text & "\" & Name1.Text & "\" & "Address", Address.Text

End Sub

Private Sub Search_Click()
If FindSurName.Text = "" Then If FindName.Text = "" Then MsgBox "er...try entering the names you wish to search for"

            If ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Name1") = FindName.Text Then
  Name1.Text = ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Name1")
SurName.Text = ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Name2")
 Phone1.Text = ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Phone1")
 Phone2.Text = ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Phone2")
Address.Text = ReadReg("HKLM\Software\PhoneBook\" & FindSurName.Text & "\" & FindName.Text & "\" & "Address")
Else
MsgBox "entry not found"
End If
End Sub

Private Sub Backup_Click()
MsgBox "To Backup your phonebook open regedit and Export the reg entry 'HKEY_LOCAL_MACHINE\Software\PhoneBook' "
End Sub

Public Sub WriteReg(Folder As String, Value As String)
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    b.RegWrite Folder, Value

End Sub

Public Function ReadReg(Value As String) As String
    Dim b As Object, R As String
    R = ""
    On Error GoTo 1
    Set b = CreateObject("wscript.shell")
    R = b.RegRead(Value)
1
    ReadReg = R
End Function
