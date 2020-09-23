VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   4530
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Parse"
      Height          =   540
      Left            =   105
      TabIndex        =   1
      Top             =   1260
      Width           =   4530
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Text            =   "Art_Pencils_Games"
      Top             =   210
      Width           =   4530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Parse index:"
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   630
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "String to parse:"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   0
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

MsgBox (Parse(Text1.Text, "_", Val(Text2.Text)))

End Sub

Private Function Parse(ByVal SrcString As String, ByVal Parser As String, ByVal ReturnString As Integer)

'String Parser by James J. Kelly Jr.
'You may use this function without my consent however
'i would appreciate it if you were to tell me or give me
'credits but do as you like.

'This function returns the desired parsed string

'I've declared nessary values incase Option Base 0 or Option Explicit is included
'in the user program.

'To use this function do it like this:
'Text1.Text = Parse(Text1.Text, ",", 1)
'That should do it

'btw these functions for convience use a 1 based index
'rather then a 0 based index.

'if you are confused use 1 insted of 0. That should be
'easy for you to figure out.

Dim i As Long
Dim CrntVal As Long
Dim ParseStr() As String

ReDim ParseStr(0) As String
CrntVal = 0

For i = 1 To Len(SrcString)
If Mid(SrcString, i, 1) <> Parser Then ParseStr(CrntVal) = ParseStr(CrntVal) + Mid(SrcString, i, 1)
If Mid(SrcString, i, 1) = Parser Then ReDim Preserve ParseStr(CrntVal + 1) As String: CrntVal = CrntVal + 1
If Mid(SrcString, i, 1) = Len(SrcString) Then ReDim Preserve ParseStr(CrntVal + 1) As String: CrntVal = CrntVal + 1
Next i

Parse = ParseStr(ReturnString - 1)

ReDim ParseStr(0) As String

End Function

Private Function ParseLen(ByVal SrcString As String, ByVal Parser As String)

'This function returns the number of parsable strings

'I've declared nessary values incase Option Base 0 or Option Explicit is included
'in the user program.

'To use this function do it like this:
'Text1.Text = ParseLen(Text1.Text, ",")
'That should do it

Dim i As Long
Dim CrntVal As Long

CrntVal = 0

For i = 1 To Len(SrcString)
If Mid(SrcString, i, 1) = Parser Then CrntVal = CrntVal + 1
If Mid(SrcString, i, 1) = Len(SrcString) Then CrntVal = CrntVal + 1
Next i

ParseLen = Str(CrntVal + 1)

End Function
