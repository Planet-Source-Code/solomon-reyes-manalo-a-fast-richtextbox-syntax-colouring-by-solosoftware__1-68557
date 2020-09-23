VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl SOLO_RTBSyntax 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ScaleHeight     =   2655
   ScaleWidth      =   3675
   ToolboxBitmap   =   "Syntax-RTB.ctx":0000
   Begin RichTextLib.RichTextBox rtb 
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4154
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      RightMargin     =   1.00000e5
      TextRTF         =   $"Syntax-RTB.ctx":0312
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "SOLO_RTBSyntax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'CREDITS:
'Special Thanks To:------------------------------
'   rtbSyntax Control.
'   Aaron Bennear
'   Adds syntax highlighting to the RichTextBox
'   control.
'------------------------------------------------
'================================================
'Major Modifications/Enhancements/Original Codig Made By:
'   Solomon R. Manalo
'   SOLOSoftware
'   solo_sevensix@yahoo.com
'   http://www.solosoftware.co.nr
'   "making it more Beginner Friendly Control
'   and Easy to use"
'   'Add This Usercontrol to any Project'
' ===============================================
'Note
'   If there is any chance that you can Upgrade
'   this Control Please send me a copy. and
'   DO NOT REMOVE CREDITS.......

Option Explicit
'Private Constant Declarations
Private Const WM_SETREDRAW = &HB
Private Const WM_COMMAND = &H111
Private Const EM_SETMARGINS = &HD3
Private Const EM_SETREADONLY = &HCF
Private Const EM_SETSEL = &HB1
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_GETLINE = &HC4
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_LINESCROLL = &HB6
Private Const EM_HIDESELECTION = &H43F
Private Const EC_LEFTMARGIN = 1
'API Declarations
'One WinAPI call. Used to suppress repainting during parsing.
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageVal Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Const Delimiter = vbTab & " "
Const RGB_NORMAL As String = "0,0,0" 'Black

'===================================================
'SyntaxColorTypes: this enumerated lists detects  '|
'                  what type of syntax is to be   '|
'                  colored                        '|
'===================================================
Enum SyntaxColorTypes
    ColorOf_Syntax_Comment = 0
    ColorOf_Syntax_String = 1
    ColorOf_Syntax_ReservedWords = 2
    ColorOf_Syntax_ProcedureORFunctions = 3
    ColorOf_Syntax_Normal = 4
    ColorOf_Syntax_Operators = 5
    ColorOf_Syntax_LogicalOperators = 6
    ColorOf_Syntax_Constants = 7
    ColorOf_Syntax_Objects = 8
End Enum
'===================================================

' Global variable used to suppress parsing until the end of a series of
' changes. Or, in the Change event itself to prevent cascaded Change events.
Private mbInChange As Boolean

' Keeping track of current and previous insertion point. Used to determine
' what portion of text has changed.
Private mlPrevSelStart As Long
Private mlCurSelStart As Long

'Default Property Values:
'===================================================
'Default ForeColor Properties             '|
'===================================================
Const m_def_ColorOf_Comments = &HC000&               '|
Const m_def_ColorOf_ReservedWords = &HC00000         '|
Const m_def_ColorOf_ProceduresORFunctions = &HFF8080 '|
Const m_def_ColorOf_LogicalOperators = &HC0&         '|
Const m_def_ColorOf_Operators = &HC0&                '|
Const m_def_ColorOf_Objects = &HC0C000               '|
Const m_def_ColorOf_Constants = &H80FF&              '|
Const m_def_ColorOf_Strings = &HFF00FF               '|
Const m_def_ColorOf_Normal = &H0&                    '|
'===================================================
'===================================================
'Default Properties for Syntaxes // NULL or Empty '|
'===================================================
'Const m_def_Syntax_CommentChar = ""              '|
Const m_def_Syntax_StringChar = """"              '|
Const m_def_Syntax_CommentChar = "'"              '|
Const m_def_Syntax_ReservedWords = ""             '|
Const m_def_Syntax_ProceduresORFunctions = ""     '|
Const m_def_Syntax_Delimiter = vbCrLf             '|
Const m_def_Syntax_LogicalOperators = ""          '|
Const m_def_Syntax_Operators = ""                 '|
Const m_def_Syntax_Objects = ""                   '|
Const m_def_Syntax_Constants = ""                 '|
'===================================================
Const m_def_ForeColor = 0
Const m_def_hWnd = 0
'===================================================
'Property Variables:                              '|
'===================================================
Dim m_ColorOf_Comments As OLE_COLOR               '|
Dim m_ColorOf_ReservedWords As OLE_COLOR          '|
Dim m_ColorOf_ProceduresORFunctions As OLE_COLOR  '|
Dim m_ColorOf_LogicalOperators As OLE_COLOR       '|
Dim m_ColorOf_Operators As OLE_COLOR              '|
Dim m_ColorOf_Objects As OLE_COLOR                '|
Dim m_ColorOf_Constants As OLE_COLOR              '|
Dim m_ColorOf_Strings As OLE_COLOR                '|
Dim m_ColorOf_Normal As OLE_COLOR                 '|
Dim m_Syntax_ReservedWords As String              '|
Dim m_Syntax_ProceduresORFunctions As String      '|
Dim m_Syntax_Delimiter As String                  '|
Dim m_Syntax_LogicalOperators As Variant          '|
Dim m_Syntax_Operators As Variant                 '|
Dim m_Syntax_Objects As Variant                   '|
Dim m_Syntax_Constants As String                  '|
Dim m_Syntax_StringChar As String                 '|
Dim m_Syntax_CommentChar As String                '|
Dim m_ForeColor As Long                           '|
Dim m_hWnd As Long                                '|
'===================================================
Dim CommentCharCount As Integer                   '|
Dim StringCharCount As Integer                    '|
'===================================================

Event Declarations():
Event Change() 'MappingInfo=rtb,rtb,-1,Change
Event Click() 'MappingInfo=rtb,rtb,-1,Click
Event DblClick() 'MappingInfo=rtb,rtb,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyDown
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=rtb,rtb,-1,KeyUp
Event KeyPress(KeyAscii As Integer) 'MappingInfo=rtb,rtb,-1,KeyPress
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=rtb,rtb,-1,MouseUp
Event SelChange() 'MappingInfo=rtb,rtb,-1,SelChange

' Sub UserControl_Initialize
' Position constituate control, call initialization.
Private Sub UserControl_Initialize()
    rtb.Top = 0
    rtb.Left = 0
    mlPrevSelStart = 0
End Sub

' Sub rtb_Change
' Determine the changed region and feed to the parser.
Private Sub rtb_Change()
    RaiseEvent Change

    If mbInChange = True Then
        ' change is being blocked or deferred
        GoTo ExitSub
    End If
    ' suppress change events generated during this change event
    mbInChange = "True"

    Dim srtb As String      ' working string
    ' add final cariage return so last line is processed
    srtb = rtb.Text & vbCrLf

    ' preserve selection and restore at end
    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long

    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength

    Dim lStartPos As Long
    Dim lEndPos As Long

    If mlPrevSelStart < rtb.SelStart Then
        lStartPos = mlPrevSelStart
        lEndPos = rtb.SelStart
    Else
        lStartPos = rtb.SelStart
        lEndPos = mlPrevSelStart
    End If

    If lStartPos > 1 Then
        ' set start position to beginning of line
        If InStrRev(srtb, vbCrLf, lStartPos - 1) > 0 Then
            lStartPos = InStrRev(srtb, vbCrLf, lStartPos - 1) + Len(vbCrLf) - 1
        Else
            lStartPos = 0
        End If
    Else
        lStartPos = 0
    End If

    ' set end position to end of line
    If InStr(lEndPos + 1, srtb, vbCrLf) > 0 Then
        lEndPos = InStr(rtb.SelStart + 1, srtb, vbCrLf) - 1
    Else
        lEndPos = Len(srtb) - 1
    End If

    Dim x As Long

    'prevent textbox from repainting
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)
    ' send affected text to the parser, along with its position in the
    ' RichTextBox
    If lStartPos <> lEndPos Then
        Split_ToLines Mid(srtb, lStartPos + 1, lEndPos - lStartPos), rtb, lStartPos
    End If
    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelLength
    'allow texbox to repaint
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh
    mbInChange = False
ExitSub:

End Sub

' Sub rtb_SelChange
' Keep track of previous SelStart to allow determination of
' affected region.
Private Sub rtb_SelChange()
    RaiseEvent SelChange
    mlPrevSelStart = mlCurSelStart
    mlCurSelStart = rtb.SelStart
End Sub

' Sub rtb_KeyDown
' Normally, tabbing leaves the control, but instead, we want to insert
' tab into edited text.
Private Sub rtb_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)

    If KeyCode = Asc(vbTab) Then  ' TAB key was pressed.
      ' Ignore the TAB key, so focus doesn't leave the control
      KeyCode = 0
      ' Replace selected text with the tab character
      rtb.SelText = vbTab
    End If
End Sub

' Sub ReColorize
' Manipulate tracked previous selection and current selection
' to force reparsing of entire text.
Public Sub ReColorize()
    'prevent textbox from repainting
    Dim x As Long
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 0, 0)

    Dim lOrigSelStart As Long
    Dim lOrigSelLength As Long
    
    lOrigSelStart = rtb.SelStart
    lOrigSelLength = rtb.SelLength

    mlPrevSelStart = 0
    rtb.SelStart = Len(rtb.Text)
    rtb_Change
    rtb.SelStart = lOrigSelStart
    rtb.SelLength = lOrigSelStart

    'allow texbox to repaint
    x = SendMessage(rtb.hWnd, WM_SETREDRAW, 1, 0)
    'force repaint
    rtb.Refresh
End Sub

' Sub Split_ToLines
' Feed text, line by line, to the parser.
Private Sub Split_ToLines(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)
    Dim lStartPos As Long
    Dim lEndPos As Long
    lStartPos = 1
    s = s & vbCrLf
    lEndPos = InStr(lStartPos, s, vbCrLf)
    Do While lEndPos > 0
        Parse_SplittedLines Mid(s, lStartPos, lEndPos - lStartPos), rtb, RTBPos + lStartPos - 1
        lStartPos = lEndPos + Len(vbCrLf)
        lEndPos = InStr(lStartPos, s, vbCrLf)
    Loop
End Sub

' Sub Parse_SplittedLines
' Lines are treated independently. Parse_SplittedLines is the main parsing code. Scan
' line from left to right, emitting text to be colored.
Private Sub Parse_SplittedLines(ByVal s As String, rtb As RichTextBox, ByVal RTBPos As Long)
    'Debug.Print s
    
    Dim bInString As Boolean    ' are we in a quoted string?
    bInString = False
    
    Dim bInWord As Boolean      ' are we in a word? (not a string, Syntax_Comment,
                                ' or delimiter)
    bInWord = False
    
    Dim sCurString As String        ' the current set of characters
    Dim lCurStringStart As Long     '   - where it starts
    Dim sCurChar As String          ' the current character
    
    Dim i As Long
    
    CommentCharCount = Len(Syntax_CommentChar)
    StringCharCount = Len(Syntax_StringChar)
    For i = 1 To Len(s)

        sCurChar = Mid(s, i, CommentCharCount)
        If sCurChar = Syntax_CommentChar Then
            ' if Syntax_Comment character occurs within a quoted string, it doesn't
            ' count
            If Not bInString Then
                ' this is a Syntax_Comment. we are done with the line
                If bInWord Then
                    ' before we encounterd the Syntax_Comment we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - CommentCharCount + 1
                    sCurString = ""
                    bInWord = False
                End If
            
                Highlight rtb, ColorOf_Syntax_Comment, i + RTBPos - 1, Len(s) - i + CommentCharCount
                GoTo ExitSub    ' rest of line is Syntax_Comment
            End If
        End If
        
        sCurChar = Mid(s, i, StringCharCount)
        If sCurChar = Syntax_StringChar Then
            ' if not already in a string, then this quote begins a string
            ' otherwise, we are in a string, and this quote ends it
            If bInString Then
                sCurString = sCurString & sCurChar
                Highlight rtb, ColorOf_Syntax_String, lCurStringStart + RTBPos - 1, i - StringCharCount + 1
                sCurString = ""
                bInString = False
            Else
                If bInWord Then
                    ' before we encounterd the string we were processing a word
                    Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - StringCharCount + 1
                    sCurString = ""
                    bInWord = False
                End If
                bInString = True
                sCurString = sCurChar
                lCurStringStart = i
            End If
            
            GoTo Next_i ' get next character
        End If
                
        If InStr(1, Delimiter, sCurChar) > 0 Then
            If bInWord Then
                ' before we encounterd the delimiter we were processing a word
                Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
                sCurString = ""
                bInWord = False
            End If
            
            Highlight rtb, ColorOf_Syntax_Normal, i + RTBPos - 1, 1
            GoTo Next_i
        End If
            
        If (Not bInWord) And (Not bInString) Then
            bInWord = True
            sCurString = sCurChar
            lCurStringStart = i
            
            GoTo Next_i ' get next character
        End If
            
        ' add current character to the "word" we are in the middle of
        sCurString = sCurString & sCurChar
Next_i:     ' VB style continue
    Next
    
    If bInString Then
        ' before we encounterd the end of the line we were processing a string
        Highlight rtb, ColorOf_Syntax_String, lCurStringStart + RTBPos - 1, i - lCurStringStart
    ElseIf bInWord Then
        ' before we encounterd the end of the line we were processing a word
        Highlight rtb, ParseWord(sCurString), lCurStringStart + RTBPos - 1, i - lCurStringStart
    End If

ExitSub:
    Exit Sub
End Sub

' Function ParseWord
' Determine color for this word by checking for its existence in the keyword
' lists. The word being checked it padded with spaces to prevent matches
' with substrings of keywords.
Private Function ParseWord(ByVal Word As String) As SyntaxColorTypes
Dim WordToFind As String

    WordToFind = Syntax_Delimiter & Word & Syntax_Delimiter
    
    If InStr(1, Syntax_CommentChar, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_Comment
    ElseIf InStr(1, Syntax_ReservedWords, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_ReservedWords
    ElseIf InStr(1, Syntax_ProceduresORFunctions, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_ProcedureORFunctions
    'ElseIf InStr(1, Syntax_Delimiter, WordToFind, vbTextCompare) > 0 Then
    '    ParseWord = ColorOf_Syntax_Delimiter
    ElseIf InStr(1, Syntax_LogicalOperators, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_LogicalOperators
    ElseIf InStr(1, Syntax_Operators, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_Operators
    ElseIf InStr(1, Syntax_Objects, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_Objects
    ElseIf InStr(1, Syntax_Constants, WordToFind, vbTextCompare) > 0 Then
        ParseWord = ColorOf_Syntax_Constants
    Else
        ParseWord = ColorOf_Syntax_Normal
    End If
End Function

' Sub Highlight
' Color this range in the RichTextBox. Note that you could also apply bold,
' italic, etc. to the selection at the same time.
Private Sub Highlight(rtb As RichTextBox, SyntaxType As SyntaxColorTypes, StartPos As Long, Length As Long)
        rtb.SelStart = StartPos
        rtb.SelLength = Length

    Select Case SyntaxType
        Case SyntaxColorTypes.ColorOf_Syntax_Comment
            rtb.SelColor = ColorOf_Comments
        Case SyntaxColorTypes.ColorOf_Syntax_String
            rtb.SelColor = ColorOf_Strings
        Case SyntaxColorTypes.ColorOf_Syntax_ReservedWords
            rtb.SelColor = ColorOf_ReservedWords
        Case SyntaxColorTypes.ColorOf_Syntax_ProcedureORFunctions
            rtb.SelColor = ColorOf_ProceduresORFunctions
        Case SyntaxColorTypes.ColorOf_Syntax_Operators
            rtb.SelColor = ColorOf_Operators
        Case SyntaxColorTypes.ColorOf_Syntax_LogicalOperators
            rtb.SelColor = ColorOf_LogicalOperators
        Case SyntaxColorTypes.ColorOf_Syntax_Constants
            rtb.SelColor = ColorOf_Constants
        Case SyntaxColorTypes.ColorOf_Syntax_Objects
            rtb.SelColor = ColorOf_Objects
        Case SyntaxColorTypes.ColorOf_Syntax_Normal
            rtb.SelColor = ColorOf_Normal
        Case Else
            rtb.SelColor = ColorOf_Normal
    End Select

End Sub

' Sub UserControl_Resize
' Constituate control is always same size as user control.
Private Sub UserControl_Resize()
    rtb.Width = UserControl.ScaleWidth
    rtb.Height = UserControl.ScaleHeight
End Sub

' *****************************************************************************
' Properties
' For the most part this code is generated by the VB ActiveX Control Wizard
' *****************************************************************************

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,AutoVerbMenu
Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = rtb.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
    rtb.AutoVerbMenu() = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = rtb.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    rtb.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,BorderStyle
Public Property Get BorderStyle() As RichTextLib.BorderStyleConstants
    BorderStyle = rtb.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As RichTextLib.BorderStyleConstants)
    rtb.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,BulletIndent
Public Property Get BulletIndent() As Single
    BulletIndent = rtb.BulletIndent
End Property

Public Property Let BulletIndent(ByVal New_BulletIndent As Single)
    rtb.BulletIndent() = New_BulletIndent
    PropertyChanged "BulletIndent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = rtb.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    rtb.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,FileName
Public Property Get filename() As String
    filename = rtb.filename
End Property

Public Property Let filename(ByVal New_FileName As String)
    rtb.filename() = New_FileName
    PropertyChanged "FileName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Font
Public Property Get Font() As Font
    Set Font = rtb.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set rtb.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,HideSelection
Public Property Get HideSelection() As Boolean
    HideSelection = rtb.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    rtb.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MemberInfo=8,0,0,0
Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let hWnd(ByVal New_hWnd As Long)
    m_hWnd = New_hWnd
    PropertyChanged "hWnd"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Locked
Public Property Get Locked() As Boolean
    Locked = rtb.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    rtb.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = rtb.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    rtb.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = rtb.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set rtb.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = rtb.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    rtb.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,RightMargin
Public Property Get RightMargin() As Single
    RightMargin = rtb.RightMargin
End Property

Public Property Let RightMargin(ByVal New_RightMargin As Single)
    rtb.RightMargin() = New_RightMargin
    PropertyChanged "RightMargin"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Text
Public Property Get Text() As String
    Text = rtb.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    mbInChange = True
    rtb.Text() = New_Text
    mbInChange = False
    ReColorize

    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Find
Public Function Find(ByVal bstrString As String, Optional ByVal vStart As Variant, Optional ByVal vEnd As Variant, Optional ByVal vOptions As Variant) As Long
    Find = rtb.Find(bstrString, vStart, vEnd, vOptions)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,GetLineFromChar
Public Function GetLineFromChar(ByVal lChar As Long) As Long
    GetLineFromChar = rtb.GetLineFromChar(lChar)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,LoadFile
Public Sub LoadFile(ByVal bstrFilename As String, Optional ByVal vFileType As Variant)
    rtb.LoadFile bstrFilename, vFileType
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Refresh
Public Sub Refresh()
    rtb.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,SaveFile
Public Sub SaveFile(ByVal bstrFilename As String, Optional ByVal vFlags As Variant)
    rtb.SaveFile bstrFilename, vFlags
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,SelPrint
Public Sub SelPrint(ByVal lHDC As Long, Optional ByVal vStartDoc As Variant)
    rtb.SelPrint lHDC, vStartDoc
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,Span
Public Sub Span(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
    rtb.Span bstrCharacterSet, vForward, vNegate
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING Syntax_CommentED LINES!
'MappingInfo=rtb,rtb,-1,UpTo
Public Sub UpTo(ByVal bstrCharacterSet As String, Optional ByVal vForward As Variant, Optional ByVal vNegate As Variant)
    rtb.UpTo bstrCharacterSet, vForward, vNegate
End Sub

Private Sub rtb_Click()
    RaiseEvent Click
End Sub

Private Sub rtb_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtb_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub rtb_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub rtb_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub rtb_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ForeColor = m_def_ForeColor
    m_hWnd = m_def_hWnd
'    m_Syntax_CommentChar = m_def_Syntax_CommentChar
    m_Syntax_ReservedWords = m_def_Syntax_ReservedWords
    m_Syntax_ProceduresORFunctions = m_def_Syntax_ProceduresORFunctions
    m_Syntax_Delimiter = m_def_Syntax_Delimiter
    m_Syntax_LogicalOperators = m_def_Syntax_LogicalOperators
    m_Syntax_Operators = m_def_Syntax_Operators
    m_Syntax_Objects = m_def_Syntax_Objects
    m_Syntax_Constants = m_def_Syntax_Constants
    m_ColorOf_Comments = m_def_ColorOf_Comments
    m_ColorOf_ReservedWords = m_def_ColorOf_ReservedWords
    m_ColorOf_ProceduresORFunctions = m_def_ColorOf_ProceduresORFunctions
    m_ColorOf_LogicalOperators = m_def_ColorOf_LogicalOperators
    m_ColorOf_Operators = m_def_ColorOf_Operators
    m_ColorOf_Objects = m_def_ColorOf_Objects
    m_ColorOf_Constants = m_def_ColorOf_Constants
    m_ColorOf_Strings = m_def_ColorOf_Strings
    m_ColorOf_Normal = m_def_ColorOf_Normal
    m_Syntax_StringChar = m_def_Syntax_StringChar
    m_Syntax_CommentChar = m_def_Syntax_CommentChar
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' prevent parsing while file is loading
        
    rtb.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    rtb.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    rtb.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    rtb.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    rtb.Enabled = PropBag.ReadProperty("Enabled", True)
    rtb.filename = PropBag.ReadProperty("FileName", "")
    rtb.HideSelection = PropBag.ReadProperty("HideSelection", True)
    rtb.Locked = PropBag.ReadProperty("Locked", False)
    rtb.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    rtb.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    rtb.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    rtb.Text = PropBag.ReadProperty("Text", "")
        
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set rtb.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    mbInChange = True
    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    mbInChange = False
    ReColorize

    m_Syntax_ReservedWords = PropBag.ReadProperty("Syntax_ReservedWords", m_def_Syntax_ReservedWords)
    m_Syntax_ProceduresORFunctions = PropBag.ReadProperty("Syntax_ProceduresORFunctions", m_def_Syntax_ProceduresORFunctions)
    m_Syntax_Delimiter = PropBag.ReadProperty("Syntax_Delimiter", m_def_Syntax_Delimiter)
    m_Syntax_LogicalOperators = PropBag.ReadProperty("Syntax_LogicalOperators", m_def_Syntax_LogicalOperators)
    m_Syntax_Operators = PropBag.ReadProperty("Syntax_Operators", m_def_Syntax_Operators)
    m_Syntax_Objects = PropBag.ReadProperty("Syntax_Objects", m_def_Syntax_Objects)
    m_Syntax_Constants = PropBag.ReadProperty("Syntax_Constants", m_def_Syntax_Constants)
    
    m_ColorOf_Comments = PropBag.ReadProperty("ColorOf_Comments", m_def_ColorOf_Comments)
    m_ColorOf_ReservedWords = PropBag.ReadProperty("ColorOf_ReservedWords", m_def_ColorOf_ReservedWords)
    m_ColorOf_ProceduresORFunctions = PropBag.ReadProperty("ColorOf_ProceduresORFunctions", m_def_ColorOf_ProceduresORFunctions)
    m_ColorOf_LogicalOperators = PropBag.ReadProperty("ColorOf_LogicalOperators", m_def_ColorOf_LogicalOperators)
    m_ColorOf_Operators = PropBag.ReadProperty("ColorOf_Operators", m_def_ColorOf_Operators)
    m_ColorOf_Objects = PropBag.ReadProperty("ColorOf_Objects", m_def_ColorOf_Objects)
    m_ColorOf_Constants = PropBag.ReadProperty("ColorOf_Constants", m_def_ColorOf_Constants)
    m_ColorOf_Strings = PropBag.ReadProperty("ColorOf_Strings", m_def_ColorOf_Strings)
    m_ColorOf_Normal = PropBag.ReadProperty("ColorOf_Normal", m_def_ColorOf_Normal)
    m_Syntax_StringChar = PropBag.ReadProperty("Syntax_StringChar", m_def_Syntax_StringChar)
    m_Syntax_CommentChar = PropBag.ReadProperty("Syntax_CommentChar", m_def_Syntax_CommentChar)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AutoVerbMenu", rtb.AutoVerbMenu, False)
    Call PropBag.WriteProperty("BackColor", rtb.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("BorderStyle", rtb.BorderStyle, 1)
    Call PropBag.WriteProperty("BulletIndent", rtb.BulletIndent, 0)
    Call PropBag.WriteProperty("Enabled", rtb.Enabled, True)
    Call PropBag.WriteProperty("FileName", rtb.filename, "")
    Call PropBag.WriteProperty("Font", rtb.Font, Ambient.Font)
    Call PropBag.WriteProperty("HideSelection", rtb.HideSelection, True)
    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("Locked", rtb.Locked, False)
    Call PropBag.WriteProperty("MaxLength", rtb.MaxLength, 0)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", rtb.MousePointer, 0)
    Call PropBag.WriteProperty("RightMargin", rtb.RightMargin, 0)
    Call PropBag.WriteProperty("Text", rtb.Text, "")
    Call PropBag.WriteProperty("Syntax_ReservedWords", m_Syntax_ReservedWords, m_def_Syntax_ReservedWords)
    Call PropBag.WriteProperty("Syntax_ProceduresORFunctions", m_Syntax_ProceduresORFunctions, m_def_Syntax_ProceduresORFunctions)
    Call PropBag.WriteProperty("Syntax_Delimiter", m_Syntax_Delimiter, m_def_Syntax_Delimiter)
    Call PropBag.WriteProperty("Syntax_LogicalOperators", m_Syntax_LogicalOperators, m_def_Syntax_LogicalOperators)
    Call PropBag.WriteProperty("Syntax_Operators", m_Syntax_Operators, m_def_Syntax_Operators)
    Call PropBag.WriteProperty("Syntax_Objects", m_Syntax_Objects, m_def_Syntax_Objects)
    Call PropBag.WriteProperty("Syntax_Constants", m_Syntax_Constants, m_def_Syntax_Constants)
    Call PropBag.WriteProperty("ColorOf_Comments", m_ColorOf_Comments, m_def_ColorOf_Comments)
    Call PropBag.WriteProperty("ColorOf_ReservedWords", m_ColorOf_ReservedWords, m_def_ColorOf_ReservedWords)
    Call PropBag.WriteProperty("ColorOf_ProceduresORFunctions", m_ColorOf_ProceduresORFunctions, m_def_ColorOf_ProceduresORFunctions)
    Call PropBag.WriteProperty("ColorOf_LogicalOperators", m_ColorOf_LogicalOperators, m_def_ColorOf_LogicalOperators)
    Call PropBag.WriteProperty("ColorOf_Operators", m_ColorOf_Operators, m_def_ColorOf_Operators)
    Call PropBag.WriteProperty("ColorOf_Objects", m_ColorOf_Objects, m_def_ColorOf_Objects)
    Call PropBag.WriteProperty("ColorOf_Constants", m_ColorOf_Constants, m_def_ColorOf_Constants)
    Call PropBag.WriteProperty("ColorOf_Strings", m_ColorOf_Strings, m_def_ColorOf_Strings)
    Call PropBag.WriteProperty("ColorOf_Normal", m_ColorOf_Normal, m_def_ColorOf_Normal)
    Call PropBag.WriteProperty("Syntax_StringChar", m_Syntax_StringChar, m_def_Syntax_StringChar)
    Call PropBag.WriteProperty("Syntax_CommentChar", m_Syntax_CommentChar, m_def_Syntax_CommentChar)
End Sub

' *****************************************************************************
' Run Time Only Properties
' NOT generated by the ActiveX Control Wizard. Each of these procedures has
' its Procedure Attribute "Don't Show In Property Browser" set to true.
' *****************************************************************************

Public Property Get SelAlignment() As SelAlignmentConstants
Attribute SelAlignment.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelAlignment = rtb.SelAlignment
End Property

Public Property Let SelAlignment(ByVal New_SelAlignment As SelAlignmentConstants)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelAlignment = New_SelAlignment
End Property

Public Property Get SelBold() As Boolean
Attribute SelBold.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelBold = rtb.SelBold
End Property

Public Property Let SelBold(ByVal New_SelBold As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelBold = New_SelBold
End Property

Public Property Get SelItalic() As Boolean
Attribute SelItalic.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelItalic = rtb.SelItalic
End Property

Public Property Let SelItalic(ByVal New_SelItalic As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelItalic = New_SelItalic
End Property

Public Property Get SelStrikethru() As Boolean
Attribute SelStrikethru.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelStrikethru = rtb.SelStrikethru
End Property

Public Property Let SelStrikethru(ByVal New_SelStrikethru As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    rtb.SelStrikethru = New_SelStrikethru
End Property

Public Property Get SelUnderline() As Boolean
Attribute SelUnderline.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    SelUnderline = rtb.SelUnderline
End Property

Public Property Let SelUnderline(ByVal New_SelUnderline As Boolean)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    rtb.SelUnderline = New_SelUnderline
End Property

Public Property Get SelBullet() As Variant
Attribute SelBullet.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    SelBullet = rtb.SelBullet
End Property

Public Property Let SelBullet(ByVal New_SelBullet As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelBullet = New_SelBullet
End Property

Public Property Get SelCharOffset() As Variant
Attribute SelCharOffset.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelCharOffset = rtb.SelCharOffset
End Property

Public Property Let SelCharOffset(ByVal New_SelCharOffset As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelCharOffset = New_SelCharOffset
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelRTF = rtb.SelRTF
End Property

Public Property Let SelRTF(ByVal New_SelRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelRTF = New_SelRTF
End Property

Public Property Get SelTabCount() As Integer
Attribute SelTabCount.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelTabCount = rtb.SelTabCount
End Property

Public Property Let SelTabCount(ByVal New_SelTabCount As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelTabCount = New_SelTabCount
End Property

Public Property Get SelTabs(Index As Integer) As Integer
Attribute SelTabs.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelTabs = rtb.SelTabs(Index)
End Property

Public Property Let SelTabs(Index As Integer, ByVal New_SelTabs As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelTabs(Index) = New_SelTabs
End Property

Public Property Get SelColor() As Variant
Attribute SelColor.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelColor = rtb.SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelColor = New_SelColor
End Property

Public Property Get SelHangingIndent() As Integer
Attribute SelHangingIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelHangingIndent = rtb.SelHangingIndent
End Property

Public Property Let SelHangingIndent(ByVal New_SelHangingIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelHangingIndent = New_SelHangingIndent
End Property

Public Property Get SelIndent() As Integer
Attribute SelIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelIndent = rtb.SelIndent
End Property

Public Property Let SelIndent(ByVal New_SelIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelIndent = New_SelIndent
End Property

Public Property Get SelRightIndent() As Integer
Attribute SelRightIndent.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelRightIndent = rtb.SelRightIndent
End Property

Public Property Let SelRightIndent(ByVal New_SelRightIndent As Integer)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelRightIndent = New_SelRightIndent
End Property

Public Property Get SelLength() As Long
Attribute SelLength.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelLength = rtb.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelLength = New_SelLength
End Property

Public Property Get SelStart() As Long
Attribute SelStart.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelStart = rtb.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelStart = New_SelStart
End Property

Public Property Get SelText() As String
Attribute SelText.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelText = rtb.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelText = New_SelText
End Property

Public Property Get SelProtected() As Variant
Attribute SelProtected.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    
    SelProtected = rtb.SelProtected
End Property

Public Property Let SelProtected(ByVal New_SelProtected As Variant)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
    
    rtb.SelProtected = New_SelProtected
End Property

Public Property Get TextRTF() As String
Attribute TextRTF.VB_MemberFlags = "400"
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 394
    End If
    TextRTF = rtb.TextRTF
End Property

Public Property Let TextRTF(ByVal New_TextRTF As String)
    ' prevent display in property browser
    If Ambient.UserMode = False Then
        Err.Raise 383
    End If
        
    mbInChange = True
    rtb.TextRTF = New_TextRTF
    mbInChange = False
    
    ReColorize

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Syntax_ReservedWords() As String
Attribute Syntax_ReservedWords.VB_Description = "List of Syntax for Reserved Words"
    Syntax_ReservedWords = m_Syntax_ReservedWords
End Property

Public Property Let Syntax_ReservedWords(ByVal New_Syntax_ReservedWords As String)
    m_Syntax_ReservedWords = New_Syntax_ReservedWords
    PropertyChanged "Syntax_ReservedWords"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Syntax_ProceduresORFunctions() As String
Attribute Syntax_ProceduresORFunctions.VB_Description = "List of Syntax for Procedures Or Functions"
    Syntax_ProceduresORFunctions = m_Syntax_ProceduresORFunctions
End Property

Public Property Let Syntax_ProceduresORFunctions(ByVal New_Syntax_ProceduresORFunctions As String)
    m_Syntax_ProceduresORFunctions = New_Syntax_ProceduresORFunctions
    PropertyChanged "Syntax_ProceduresORFunctions"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Syntax_Delimiter() As String
Attribute Syntax_Delimiter.VB_Description = "Charcter(s) that separates syntaxes"
    Syntax_Delimiter = m_Syntax_Delimiter
End Property

Public Property Let Syntax_Delimiter(ByVal New_Syntax_Delimiter As String)
    m_Syntax_Delimiter = New_Syntax_Delimiter
    PropertyChanged "Syntax_Delimiter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,
Public Property Get Syntax_LogicalOperators() As Variant
    Syntax_LogicalOperators = m_Syntax_LogicalOperators
End Property

Public Property Let Syntax_LogicalOperators(ByVal New_Syntax_LogicalOperators As Variant)
    m_Syntax_LogicalOperators = New_Syntax_LogicalOperators
    PropertyChanged "Syntax_LogicalOperators"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,
Public Property Get Syntax_Operators() As Variant
    Syntax_Operators = m_Syntax_Operators
End Property

Public Property Let Syntax_Operators(ByVal New_Syntax_Operators As Variant)
    m_Syntax_Operators = New_Syntax_Operators
    PropertyChanged "Syntax_Operators"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,
Public Property Get Syntax_Objects() As Variant
    Syntax_Objects = m_Syntax_Objects
End Property

Public Property Let Syntax_Objects(ByVal New_Syntax_Objects As Variant)
    m_Syntax_Objects = New_Syntax_Objects
    PropertyChanged "Syntax_Objects"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Syntax_Constants() As String
Attribute Syntax_Constants.VB_Description = "List of Syntax for Constants"
    Syntax_Constants = m_Syntax_Constants
End Property

Public Property Let Syntax_Constants(ByVal New_Syntax_Constants As String)
    m_Syntax_Constants = New_Syntax_Constants
    PropertyChanged "Syntax_Constants"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Comments() As OLE_COLOR
Attribute ColorOf_Comments.VB_Description = "Sets the ForeColor of Syntax Comments"
    ColorOf_Comments = m_ColorOf_Comments
End Property

Public Property Let ColorOf_Comments(ByVal New_ColorOf_Comments As OLE_COLOR)
    m_ColorOf_Comments = New_ColorOf_Comments
    PropertyChanged "ColorOf_Comments"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_ReservedWords() As OLE_COLOR
Attribute ColorOf_ReservedWords.VB_Description = "Sets the ForeColor of Syntax ReservedWords"
    ColorOf_ReservedWords = m_ColorOf_ReservedWords
End Property

Public Property Let ColorOf_ReservedWords(ByVal New_ColorOf_ReservedWords As OLE_COLOR)
    m_ColorOf_ReservedWords = New_ColorOf_ReservedWords
    PropertyChanged "ColorOf_ReservedWords"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_ProceduresORFunctions() As OLE_COLOR
Attribute ColorOf_ProceduresORFunctions.VB_Description = "Sets the ForeColor of Syntax Procedures Or Functions"
    ColorOf_ProceduresORFunctions = m_ColorOf_ProceduresORFunctions
End Property

Public Property Let ColorOf_ProceduresORFunctions(ByVal New_ColorOf_ProceduresORFunctions As OLE_COLOR)
    m_ColorOf_ProceduresORFunctions = New_ColorOf_ProceduresORFunctions
    PropertyChanged "ColorOf_ProceduresORFunctions"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_LogicalOperators() As OLE_COLOR
Attribute ColorOf_LogicalOperators.VB_Description = "Sets the ForeColor of Syntax Logical Operators"
    ColorOf_LogicalOperators = m_ColorOf_LogicalOperators
End Property

Public Property Let ColorOf_LogicalOperators(ByVal New_ColorOf_LogicalOperators As OLE_COLOR)
    m_ColorOf_LogicalOperators = New_ColorOf_LogicalOperators
    PropertyChanged "ColorOf_LogicalOperators"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Operators() As OLE_COLOR
Attribute ColorOf_Operators.VB_Description = "Sets the ForeColor of Syntax Operators"
    ColorOf_Operators = m_ColorOf_Operators
End Property

Public Property Let ColorOf_Operators(ByVal New_ColorOf_Operators As OLE_COLOR)
    m_ColorOf_Operators = New_ColorOf_Operators
    PropertyChanged "ColorOf_Operators"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Objects() As OLE_COLOR
Attribute ColorOf_Objects.VB_Description = "Sets the ForeColor of Syntax Objects"
    ColorOf_Objects = m_ColorOf_Objects
End Property

Public Property Let ColorOf_Objects(ByVal New_ColorOf_Objects As OLE_COLOR)
    m_ColorOf_Objects = New_ColorOf_Objects
    PropertyChanged "ColorOf_Objects"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Constants() As OLE_COLOR
Attribute ColorOf_Constants.VB_Description = "Sets the ForeColor of Syntax Contants"
    ColorOf_Constants = m_ColorOf_Constants
End Property

Public Property Let ColorOf_Constants(ByVal New_ColorOf_Constants As OLE_COLOR)
    m_ColorOf_Constants = New_ColorOf_Constants
    PropertyChanged "ColorOf_Constants"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Strings() As OLE_COLOR
Attribute ColorOf_Strings.VB_Description = "Sets the ForeColor of Syntax Strings "
    ColorOf_Strings = m_ColorOf_Strings
End Property

Public Property Let ColorOf_Strings(ByVal New_ColorOf_Strings As OLE_COLOR)
    m_ColorOf_Strings = New_ColorOf_Strings
    PropertyChanged "ColorOf_Strings"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ColorOf_Normal() As OLE_COLOR
Attribute ColorOf_Normal.VB_Description = "The Deafult Text Color"
    ColorOf_Normal = m_ColorOf_Normal
End Property

Public Property Let ColorOf_Normal(ByVal New_ColorOf_Normal As OLE_COLOR)
    m_ColorOf_Normal = New_ColorOf_Normal
    PropertyChanged "ColorOf_Normal"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Syntax_StringChar() As String
Attribute Syntax_StringChar.VB_Description = "Character(s) that indicates the start of a string value"
    Syntax_StringChar = m_Syntax_StringChar
End Property

Public Property Let Syntax_StringChar(ByVal New_Syntax_StringChar As String)
    m_Syntax_StringChar = New_Syntax_StringChar
    PropertyChanged "Syntax_StringChar"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Syntax_CommentChar() As String
Attribute Syntax_CommentChar.VB_Description = "Character(s) that indicates the start of a string comment"
    Syntax_CommentChar = m_Syntax_CommentChar
End Property

Public Property Let Syntax_CommentChar(ByVal New_Syntax_CommentChar As String)
    m_Syntax_CommentChar = New_Syntax_CommentChar
    PropertyChanged "Syntax_CommentChar"
End Property


'===============================================================================
Public Function GetLineCount() As Long
    ' return the total line count of the code window
    GetLineCount = SendMessage(rtb.hWnd, EM_GETLINECOUNT, 0, 0)
End Function

Public Sub LineScroll(ByVal Lines As Long)
   SendMessageVal rtb.hWnd, EM_LINESCROLL, 0&, Lines
End Sub

Public Function GetCurrentLine() As Long
    GetCurrentLine = SendMessage(rtb.hWnd, EM_LINEFROMCHAR, rtb.SelStart, 0) + 1
End Function

Public Function GetLineText() As String
    On Error Resume Next
    Dim line As Long, lngStart As Long
    Dim start As Long
    
    line = GetCurrentLine
    lngStart = SendMessage(rtb.hWnd, EM_LINEINDEX, line - 1, 0&)
    start = lngStart
    line = line + 1
    lngStart = SendMessage(rtb.hWnd, EM_LINEINDEX, line - 1, 0&)
    If lngStart = -1 Then lngStart = Len(rtb.Text) + 2
    GetLineText = Mid$(rtb.Text, start + 1, lngStart - start - 2)
End Function

Public Function GetColumn() As Integer
   Dim lLine As Long
   Dim cCol As Long, lChar As Long, i As Long
   
   lChar = rtb.SelStart + 1
   cCol = SendMessageLong(rtb.hWnd, EM_LINELENGTH, lChar - 1, 0&)
   lLine = 1 + SendMessageLong(rtb.hWnd, EM_LINEFROMCHAR, rtb.SelStart, 0&)
   i = SendMessageLong(rtb.hWnd, EM_LINEINDEX, lLine - 1, 0&)
   GetColumn = lChar - i - 1
End Function
