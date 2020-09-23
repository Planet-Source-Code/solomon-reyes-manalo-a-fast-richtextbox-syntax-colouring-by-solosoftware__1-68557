Attribute VB_Name = "modSyntax"
Option Explicit

' Space surrounding each word is significant. It allows searching on whole
' words. Note that these constant declares are long and could reach the line
' length limit of 1023 characters. If so, simply split to 2 constants and
' combine into a third constant with the appropriate name.

'========================================================================================
'Modify this as you want....
' these constants are actually the syntaxes of VB/VBscript
' You Can Change its syntaxes in Any languages Like C/C++,
' Pascal, JAVA, Delphi, etc.
' EXCEPT FOR HTML TAGS
'Tip:
'if you know how to manipulate text files it is better to save this
'syntax references in a text file. and then Call/Open/Read/Write it
'using FileSystemObject or an API FSO
Public Const RESERVED As String = " As Call Case Const Dim Do Each Else ElseIf Empty" & _
                           " End Eqv Erase Error Exit Explicit False For" & _
                           " Function If Imp In Is Loop Mod Next Not Nothing" & _
                           " Null On Private Public Randomize ReDim" & _
                           " Resume Select Set Step Sub Then To True Until Wend" & _
                           " While Implicit String Integer" & _
                           " Double Option Long Variant String"
Public Const FUNC_OBJ As String = " Anchor Array Asc Atn CBool CByte CCur CDate CDbl" & _
                           " Chr CInt CLng Cos CreateObject CSng CStr Date" & _
                           " DateAdd DateDiff DatePart DateSerial DateValue" & _
                           " Day Dictionary Document Element Err Exp" & _
                           " FileSystemObject  Filter Fix Int Form" & _
                           " FormatCurrency FormatDateTime FormatNumber" & _
                           " FormatPercent GetObject Hex History Hour" & _
                           " InputBox InStr InstrRev IsArray IsDate IsEmpty" & _
                           " IsNull IsNumeric IsObject Join LBound LCase Left" & _
                           " Len Link LoadPicture Location Log LTrim RTrim" & _
                           " Trim Mid Minute Month MonthName MsgBox Navigator" & _
                           " Now Oct Replace Right Rnd Round ScriptEngine" & _
                           " ScriptEngineBuildVersion ScriptEngineMajorVersion" & _
                           " ScriptEngineMinorVersion Second Sgn Sin Space Split" & _
                           " Sqr StrComp String StrReverse Tan Time TextStream" & _
                           " TimeSerial TimeValue TypeName UBound UCase VarType" & _
                           " Weekday WeekDayName Window Year "
Public Const Operators As String = " + - * / \ = % ^ # "
Public Const LOperators As String = " And Or Xor True False Not Is Like "

'========================================================================================

' This Variable symbolizes a char which splits every word
' Ex. " And Call Case If For "
' The Splitting Character is a space

' Actualy Any character is Valid or vbCrlf
' Ex: "/And/Call/Case/If/For/"
' The Splitting Character is "/"
Public Const SplittingCharacter As String = " "

  'Specifies/Indicates a Comment "'" (VB,VBScript)
Public Const CommentCharacter As String = "'"
  'Specifies/Indicates a String """" (VB,VBScript)
Public Const StringCharacter As String = """"



