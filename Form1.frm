VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder Master!!!"
   ClientHeight    =   5505
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4260
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
      Height          =   3255
      Left            =   2280
      TabIndex        =   12
      Top             =   720
      Width           =   1815
      Begin VB.OptionButton Option6 
         Caption         =   "HTML File"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "RecycleBin"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Network"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Printer"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Control Panel"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Windows Icon"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   2055
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   720
      TabIndex        =   5
      Top             =   3960
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "About"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Lock"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Unlock"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Text            =   "\"
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4080
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0ECA
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " General Corporation Bangladesh"
      Height          =   195
      Left            =   840
      TabIndex        =   21
      Top             =   240
      Width           =   2340
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   300
      Left            =   1200
      TabIndex        =   20
      Top             =   15
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   19
      Top             =   0
      Width           =   1680
   End
   Begin VB.Menu User 
      Caption         =   "Option"
      Begin VB.Menu change 
         Caption         =   "Change Password"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About1 
      Caption         =   "About"
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                 Folder Master
'   Author: Didarul Alam ( Bangladesh )
' Senior Programmer: General Corporation Bangladesh


' This is a Simple Folder Security
'You can change the outlook of the folder by 6 options..

'This code is dedicated to Tahmina Nur Chowdhury
'She is my Love..



Dim FileName As String
Dim filename2 As String
Dim data As String





Private Sub About_Click()
Command3_Click
End Sub

Private Sub change_Click()
didar$ = InputBox("Please Enter The New Password", "New Password", "gsi911")
   If didar = "" Then
   MsgBox "You didn't enter any Password", 16, "Error"
   Exit Sub
   End If
    r% = WritePrivateProfileString("password", "pass", didar, iniPath$)
    If r% <> 1 Then MsgBox "An error occurred while writing SerialNumber."


End Sub

Private Sub Command1_Click()
 Dim count As String
 Dim X As Long
 Dim ext As String
 Dim filename3 As String
 Dim count2 As Variant
Dim i As Variant
On Error Resume Next



For i = 0 To Len(Dir1.Path)
startposition = i
startposition = InStr(startposition + 1, Dir1.Path, "\", vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text3.SetFocus
Text5.Text = Text3.SelStart
 
 count2 = Len(Dir1.Path) - Text5.Text - 1
  data = Right(Dir1.Path, count2)
file$ = Left(Dir1.Path, Len(Dir1.Path) - count2)
End If
Next i





If Option1.Value = True Then ext = ".{00021401-0000-0000-C000-000000000046}"
If Option2.Value = True Then ext = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
If Option3.Value = True Then ext = ".{2227A280-3AEA-1069-A2DE-08002B30309D}"
If Option4.Value = True Then ext = ".{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"
If Option5.Value = True Then ext = ".{645FF040-5081-101B-9F08-00AA002F954E}"
If Option6.Value = True Then ext = ".{25336920-03F9-11CF-8FD0-00AA00686F13}"



 FileName = file & data & ext
Name Dir1.Path As FileName

Dir1.Path = file

End Sub

Private Sub Command2_Click()
Dim count As String
Dim count2 As String
Dim filename7 As String
Dim filename3 As String
Dim i As Variant
On Error Resume Next

Text3.Text = Dir1.Path
For i = 0 To Len(Text3.Text)
startposition = i
startposition = InStr(startposition + 1, Text3.Text, Text4.Text, vbTextCompare)
If startposition > 0 Then
Text3.SelStart = startposition - 1
Text5.Text = Text3.SelStart
 count2 = Len(Text3.Text) - Text5.Text - 1
  data = Right(Text3.Text, count2)
 file$ = Left(Text3.Text, Len(Text3.Text) - count2)
 End If
Next i

FileName = Dir1.Path
  Text1.Text = FileName
  filename7 = FileName

count = (Len(data) - 39)
Text1.Text = Left(data, count)
 FileName = file & Right(Text1.Text, count)
 
Text1.Text = "ren " & filename7 & " " & FileName
Name filename7 As FileName
Dir1.Path = file

End Sub


Private Sub Command3_Click()
MsgBox "Folder Master. Allrights Reserved.© Copyright By General Corporation Bangladesh.", 32, "About"
End Sub

Private Sub Dir1_Change()
If LCase(Dir1.Path) = LCase("c:\Program Files") Then
MsgBox "Folder Security Cannot Protect 'Program Files'..", 16, "TNC Warning!!"
Dir1.Path = "c:\"
End If
If LCase(Dir1.Path) = LCase("c:\Windows") Then
MsgBox "Folder Security Cannot Protect 'Windows'..", 16, "TNC Warning!!"
Dir1.Path = "c:\"
End If
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
iniPath$ = App.Path + "\rwini.ini"
  didar$ = GetFromINI("password", "pass", iniPath$)
  
    
If didar = "" Then
didar = InputBox("Please Enter The New Password", "New Password", "gsi911")
  
   If didar = "" Then
   MsgBox "You didn't enter any Password", 16, "Error"
   Exit Sub
   End If

    r% = WritePrivateProfileString("password", "pass", didar, iniPath$)
    If r% <> 1 Then MsgBox "An error occurred while writing SerialNumber."
End If


  check = InputBox("Please Enter Passord", "Enter Password")
  If check = didar Then
  Else
  MsgBox "Not a valid user", 16, "Error"
  End
  Exit Sub
  End If

Dir1.Path = "c:\"
End Sub

Private Sub Help_Click()
MsgBox "            General HELP" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "      Select any folder and choose option which way you want to protect your folder." & Chr(13) & Chr(10) & "          Click Lock button to protect." & Chr(13) & Chr(10) & "   NB: Don't try to protect 'Windows', 'Program Files' or any other system folders..", 32, "HELP"

End Sub
