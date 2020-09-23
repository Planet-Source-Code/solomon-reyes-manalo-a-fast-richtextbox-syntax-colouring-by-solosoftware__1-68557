VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Syntax Highlighting Control (RichTextBox)"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Visit my Website (Clik Here)"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   5970
      Width           =   6645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   6870
      TabIndex        =   4
      Top             =   5970
      Width           =   1215
   End
   Begin Project1.SOLO_RTBSyntax SOLO_RTBSyntax1 
      Height          =   5235
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   9234
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Form1.frx":0000
      RightMargin     =   1.00000e5
      ColorOf_ProceduresORFunctions=   33023
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Vote!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1380
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.solosoftware.co.nr"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5970
      TabIndex        =   3
      Top             =   120
      Width           =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "solo_sevensix@yahoo.com"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6150
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A Syntax Highlighting Control by: Solomonn R. Manalo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4680
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================
' An RTB Syntax Highlighting Control.
' Made by solomon manalo
'==========================================================

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim ex As String

Function ReadTextFileContents(filename As String) As String
    Dim fnum As Integer, isOpen As Boolean
    On Error GoTo Error_Handler
    ' Get the next free file number.
    fnum = FreeFile()
    Open filename For Input As #fnum
    ' If execution flow got here, the file has been open without error.
    isOpen = True
    ' Read the entire contents in one single operation.
    ReadTextFileContents = Input(LOF(fnum), fnum)
    ' Intentionally flow into the error handler to close the file.
Error_Handler:
    ' Raise the error (if any), but first close the file.
    If isOpen Then Close #fnum
    If Err Then Err.Raise Err.Number, , Err.Description
End Function

Private Sub Command1_Click()
Unload Me: End
End Sub

Private Sub Command2_Click()
OpenURL "http://www.solosoftware.co.nr", Me.hWnd
End Sub

Private Sub Form_Load()
Dim nL As String
Call InitSyntaxEditor
SOLO_RTBSyntax1.Text = ReadTextFileContents(App.Path & "\Sample.txt")
Call InitSyntaxEditor
End Sub

Public Sub InitSyntaxEditor()
'You can change the Deafult syntax colors of this control by its
'properties in design time, or in coding style.
With SOLO_RTBSyntax1
'Try to Uncomment this code...
'===================================================
'     .ColorOf_ReservedWords = vbRed
'     .ColorOf_ProceduresORFunctions = vbBlue
'     .ColorOf_Comments = vbYellow
'     .ColorOf_Strings = vbGreen
'===================================================
     .Syntax_CommentChar = CommentCharacter
     .Syntax_StringChar = StringCharacter
     .Syntax_Delimiter = SplittingCharacter
     .Syntax_Operators = Operators
     .Syntax_LogicalOperators = LOperators
     .Syntax_ReservedWords = RESERVED
     .Syntax_ProceduresORFunctions = FUNC_OBJ
     'Refreshes the syntax coloring always call this property
     'when there is change in Syntaxes or Colors
     .ReColorize
     End With
End Sub

Public Function OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Function
