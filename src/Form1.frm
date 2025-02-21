VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GUI VBScript messagebox creator"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   12495
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   12495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "Preview messagebox"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      TabIndex        =   23
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Use custom icon and button code. This will allow you to type in your custom icon and button code into the textbox on the left"
      Height          =   975
      Left            =   8040
      TabIndex        =   22
      Top             =   4920
      Width           =   3975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Retry and Cancel buttons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   21
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Yes and No buttons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   20
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Yes, No, and Cancel buttons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   19
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Abort, Retry, and Ignore buttons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   18
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "OK and Cancel buttons"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      TabIndex        =   17
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OK button only"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   16
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Information Message icon"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   15
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Warning Message icon"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   14
      Top             =   2400
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Warning Query icon"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   13
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Error"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      MaskColor       =   &H8000000F&
      TabIndex        =   12
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate VBScript messagebox"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   11.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1200
      TabIndex        =   5
      Top             =   2520
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Icons and buttons:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   120
      Width           =   2535
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   7200
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   12480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   7200
      Y1              =   0
      Y2              =   6000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Button code:"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Icon code:"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message:"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Title:"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "You can make your own messagebox by using VBScript, except with this program it is more user-friendly"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GUI VBScript messagebox creator"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
Form2.Show
End Sub

Private Sub Command1_Click()
If Text3.Text = "" Then
c = MsgBox("You need to choose an icon on the right panel!", 0 + 64, "GUI VBScript messagebox creator")
Else
If Text4.Text = "" Then
c = MsgBox("You need to choose a button on the right panel!", 0 + 64, "GUI VBScript messagebox creator")
Else
FileName = InputBox("Please type your file name below. You can also specify your own directory to save the file, default is program's root directory. There is no need to add the "".vbs"" file extension.")
If Trim(FileName) = "" Then
    b = MsgBox("It looks like you have not specified a file name or clicked cancel.", 0 + 64, "GUI VBScript messagebox creator")
Else
Open FileName + ".vbs" For Output As #1
Print #1, "rem Made with GUI VBScript messagebox creator"
Print #1, "x=msgbox(" & """" & Text2.Text & """" & ", " & Text4.Text & "+" & Text3.Text & ", " & """" & Text1.Text & """" & ")"
Close
a = MsgBox("Your messagebox has been saved. Default location of the file is in this program's root directory. ", 0 + 64, "GUI VBScript messagebox creator")
End If
End If
End If
End Sub

Private Sub Command10_Click()
Text4.Text = "4"
End Sub

Private Sub Command11_Click()
Text4.Text = "5"
End Sub

Private Sub Command12_Click()
enablecustom = MsgBox("You're about to enable the ""Use custom icon and button code"" option. This will allow you to enter your custom icon and button code into the textbox. It is recommended to click ""No"" on this message box if you're unsure what you're doing, as entering an invalid icon or button code may cause your script to malfunction.", 4 + 48, "GUI VBScript messagebox creator")

If enablecustom = vbYes Then

Text3.Enabled = True
Text4.Enabled = True

Else


End If


End Sub

Private Sub Command13_Click()
If Text3.Text = "" Then
c = MsgBox("You need to choose an icon on the right panel!", 0 + 64, "GUI VBScript messagebox creator")
Else
If Text4.Text = "" Then
c = MsgBox("You need to choose a button on the right panel!", 0 + 64, "GUI VBScript messagebox creator")
Else
Dim buttonPreview As Integer
buttonPreview = CInt(Text4.Text) + CInt(Text3.Text)
PREVIEW = MsgBox(Text2.Text, buttonPreview, Text1.Text)
End If
End If
End Sub

Private Sub Command2_Click()
Text3.Text = "16"
End Sub

Private Sub Command3_Click()
Text3.Text = "32"
End Sub

Private Sub Command4_Click()
Text3.Text = "48"
End Sub

Private Sub Command5_Click()
Text3.Text = "64"
End Sub

Private Sub Command6_Click()
Text4.Text = "0"
End Sub

Private Sub Command7_Click()
Text4.Text = "1"
End Sub

Private Sub Command8_Click()
Text4.Text = "2"
End Sub

Private Sub Command9_Click()
Text4.Text = "3"
End Sub

