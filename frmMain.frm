VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set object properties with scripting"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   870
      Left            =   3840
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   3840
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Change Script"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   255
      Left            =   2280
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Execute"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtScript 
      Height          =   2415
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain.frx":000C
      Top             =   2280
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtScript 
      Height          =   2415
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Text            =   "frmMain.frx":006B
      Top             =   2280
      Width           =   6375
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sScript As String

Private Sub Command2_Click()
    Dim EmptyArray()
    Dim i As Integer
    Dim TextArray() As String
    Dim sLine As String
    Dim obj1 As Object
    
    TextArray = Split(sScript, vbCrLf)
    
    For i = LBound(TextArray) To UBound(TextArray)
        sLine = TextArray(i)
        If sLine Like "*:=*" Then
            If Mid(sLine, 1, InStr(1, sLine, ":=")) Like "*.*" Then
                If LCase(Mid(sLine, 1, InStr(1, sLine, ".") - 1)) <> "me" Then
                    Set obj1 = CallByName(Me, Mid(sLine, 1, InStr(1, sLine, ".") - 1), VbGet)
                    CallByName obj1, Trim(Mid(sLine, InStr(1, sLine, ".") + 1, InStr(1, sLine, ":=") - InStr(1, sLine, ".") + 1 - 2)), VbLet, Trim(Mid(sLine, InStr(1, sLine, ":=") + 2))
                Else
                    CallByName Me, Trim(Mid(sLine, InStr(1, sLine, ".") + 1, InStr(1, sLine, ":=") - InStr(1, sLine, ".") + 1 - 2)), VbLet, Trim(Mid(sLine, InStr(1, sLine, ":=") + 2))
                End If
            End If
            
        End If
        
        
    Next i
End Sub

Private Sub Command4_Click()
    txtScript(0).Visible = Not txtScript(0).Visible
    txtScript(1).Visible = Not txtScript(1).Visible
    If txtScript(0).Visible Then sScript = txtScript(0) Else sScript = txtScript(1)
End Sub

Private Sub Form_Load()
    sScript = txtScript(0)
End Sub
