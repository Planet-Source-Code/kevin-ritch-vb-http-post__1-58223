VERSION 5.00
Begin VB.Form Form1 
   Caption         =   """Ask ARIN"" -  Written by Kevin Ritch - 2005 - V8Software.com"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6510
   Icon            =   "pform.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   480
      Left            =   120
      Picture         =   "pform.frx":0442
      ScaleHeight     =   420
      ScaleWidth      =   6120
      TabIndex        =   24
      Top             =   120
      Width           =   6180
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   6255
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   2400
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   5775
      End
      Begin VB.TextBox ResultText 
         Height          =   285
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   5775
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "Organization ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1605
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   885
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   885
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "State or Province"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "Zip / Postal Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   240
         TabIndex        =   18
         Top             =   3360
         Width           =   795
      End
      Begin VB.Label DataLabel 
         AutoSize        =   -1  'True
         Caption         =   "Referral Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   240
         TabIndex        =   17
         Top             =   3960
         Width           =   1605
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   6255
      Begin VB.CommandButton Command3 
         Caption         =   "Visit ARIN's Home Page"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox fielddata 
         Height          =   330
         Left            =   600
         TabIndex        =   0
         Top             =   600
         Width           =   1770
      End
      Begin VB.CommandButton Command2 
         Caption         =   "?"
         Height          =   255
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enter IP Address to query ( Or any ARIN ID )"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   600
         TabIndex        =   15
         Top             =   240
         Width           =   4545
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "POST REQUEST VIA THE INTERNET"
      Height          =   480
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   3555
   End
   Begin VB.TextBox server 
      Height          =   345
      Left            =   1200
      TabIndex        =   11
      Text            =   "ws.arin.net/cgi-bin/whois.pl"
      Top             =   8880
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Label Label3 
      Caption         =   "ACTCalendar.com/BinaryRead.asp"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   8760
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter data to send to the server "
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   8640
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Screen.MousePointer = 11
 Me.Refresh
 Command1.Enabled = False
 For i = 0 To 6
  ResultText(i).Text = ""
 Next i
 On Error GoTo Done:
 m_cPostBuffer$ = ""
 AddPostKey "queryinput", Me.fielddata.Text
 Dim script As String
 Dim host As String
 host = Left$(server.Text, InStr(server.Text, "/") - 1)
 script = Mid$(server.Text, InStr(server.Text, "/") + 1)
 A$ = PostForm(host, script)
 Call ExtractData(A$, "OrgName:", 0)
 Call ExtractData(A$, "Address:", 1)
 Call ExtractData(A$, "City:", 2)
 Call ExtractData(A$, "StateProv:", 3)
 Call ExtractData(A$, "PostalCode:", 4)
 Call ExtractData(A$, "Country:", 5)
 Call ExtractData(A$, "ReferralServer:", 6)
Done:
 On Error GoTo 0
 Screen.MousePointer = Default
 Me.Refresh
 Command1.Enabled = True
End Sub
Private Sub Command2_Click()
 MsgBox "ARIN's WHOIS service provides a mechanism for finding contact and registration information for resources registered with ARIN. ARIN's database contains IP addresses, autonomous system (AS) numbers, organizations or customers that are associated with these resources, and related Points of Contact (POCs).", vbInformation, "About ARIN Queries"
End Sub
Private Sub Command3_Click()
 Shell "Explorer http://www.arin.net/index.html", vbMaximizedFocus
End Sub
Private Sub fielddata_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Command1.SetFocus
  Call Command1_Click
 End If
End Sub
Sub ExtractData(SourceStr As String, Findstr As String, TextIndex As Integer)
 S = InStr(SourceStr, Findstr)
 If S Then
  B$ = Mid$(SourceStr, S, 200)
  Mid$(B$, 1, Len(Findstr)) = Space$(Len(Findstr))
  S = InStr(B$, Chr$(10))
  If S Then
   B$ = Left$(B$, S - 1)
   ResultText(TextIndex).Text = Trim$(B$)
   ResultText(TextIndex).SelStart = 0
   ResultText(TextIndex).SelLength = Len(ResultText(TextIndex).Text)
  End If
 End If
End Sub
