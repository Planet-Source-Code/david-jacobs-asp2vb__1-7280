VERSION 5.00
Begin VB.Form frmCodeParser 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ASP to VB - Version: Beta"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Convert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   6480
      Width           =   1575
   End
   Begin VB.TextBox txtParsed 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   3360
      Width           =   10095
   End
   Begin VB.TextBox txtCode 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   720
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmCodeParser.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   9975
   End
End
Attribute VB_Name = "frmCodeParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodeSplit() As String
Dim HTMLString() As String
Dim codeString As String
Dim ASPString() As String
Dim ASPSplitCount
Dim HTMLSplitCount
Dim CodeSplitCount

Private Sub Convert_Click()
    CodeSplitCount = 0
    codeString = txtCode.Text
    codeString = Replace(codeString, "<%=", "<%Response.Write ")
    codeString = Replace(codeString, "<% =", "<%Response.Write ")
    codeString = Replace(codeString, "%>", "<%")
    codeString = Replace(codeString, "<%=", "<%" & (Chr(34) & " " & Chr(38) & " "))
    If Left(codeString, 2) <> "<%" Then
        CodeSplit() = Split(vbCrLf & codeString, "<%")
    Else
        CodeSplit() = Split(codeString, "<%")
    End If
    If UBound(CodeSplit()) > 0 Then
        Do Until CodeSplitCount >= UBound(CodeSplit()) + 1
            ParseHTML
            ParseASP
        Loop
    Else
        ParseHTML
    End If
    txtParsed.Text = Join(CodeSplit, "")
End Sub
Private Sub ParseHTML()
    HTMLSplitCount = 0
    HTMLString() = Split(CodeSplit(CodeSplitCount), vbCrLf)
    Do Until HTMLSplitCount > UBound(HTMLString())
        HTMLString(HTMLSplitCount) = Replace(HTMLString(HTMLSplitCount), Chr(34), Chr(34) & Chr(34))
        If EmptyString(HTMLString(HTMLSplitCount)) = False Then
            HTMLString(HTMLSplitCount) = "Response.Write " & Chr(34) & HTMLString(HTMLSplitCount) & Chr(34)
        End If
        HTMLSplitCount = HTMLSplitCount + 1
    Loop
    CodeSplit(CodeSplitCount) = Join(HTMLString, vbCrLf) & vbCrLf
    CodeSplitCount = CodeSplitCount + 1
End Sub
Private Sub ParseASP()
    ASPSplitCount = 0
    If (CodeSplitCount < UBound(CodeSplit()) + 1) Then
        ASPString() = Split(CodeSplit(CodeSplitCount), vbCrLf)
        ASPSplitCount = 0
        Do Until ASPSplitCount >= UBound(ASPString())
            'additional code can go here
            ASPSplitCount = ASPSplitCount + 1
        Loop
        CodeSplit(CodeSplitCount) = Join(ASPString, vbCrLf) & vbCrLf
        
        CodeSplitCount = CodeSplitCount + 1
    End If
End Sub
Private Function EmptyString(TestString As String) As Boolean
    If Trim(Replace(TestString, vbTab, "")) = vbCrLf Then
        EmptyString = True
    ElseIf Trim(Replace(TestString, vbTab, "")) = "" Then
        EmptyString = True
    End If
End Function
