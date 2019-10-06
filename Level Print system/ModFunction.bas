Attribute VB_Name = "ModFunction"

Public Sub StartUpPosition(ByRef MyForm As Form)
    
'    MyForm.Left = ((Screen.Width - MyForm.Width) / 2) + (frmMDI_Main.dxSideBar1.Width / 2)
    MyForm.Left = ((Screen.Width - MyForm.Width) / 2)
    MyForm.Top = ((Screen.Height - MyForm.Height) / 2) + 350
End Sub

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Sub Transparency(Frm As Form, Level%, Optional KeepControls As Boolean = False)
    
    Dim MSG As Long
    
    MSG = GetWindowLong(Frm.hWnd, (-20))
    MSG = MSG Or &H80000
    SetWindowLong Frm.hWnd, (-20), MSG
    SetLayeredWindowAttributes Frm.hWnd, 0, Level, &H2
    
End Sub


Public Sub TextEnable(MForm As Form, bEnable As Boolean)
    Dim i As Integer
    Dim Ocontrol As Control
     
    With MForm
    For Each Ocontrol In MForm
        If UCase(TypeName(Ocontrol)) = "TEXTBOX" Then
            If Ocontrol.Tag <> CONTROL_PERMANENT_DISABLE Then Ocontrol.Enabled = bEnable
        End If
     Next Ocontrol
    End With
End Sub


Public Sub GridCount(objGrid As Control)
        Dim i As Integer
    
   For i = 1 To objGrid.Rows - 1
        objGrid.TextMatrix(i, 0) = i
        Next
End Sub
'End Sub


Public Sub GridCellCustomize(Grid1 As Control, KeyAscii As Integer, Optional afterdot As Integer)
    If Chr(KeyAscii) = "." And afterdot = 0 Then KeyAscii = 0: Exit Sub
    
    Select Case KeyAscii
    
        Case 46:
                If InStr(Grid1.EditText, ".") Or (Len(Grid1.EditText) - Grid1.EditSelStart) > afterdot Then KeyAscii = 0
                Exit Sub
    End Select
    If (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) And (Not IsNumeric(Chr$(KeyAscii))) Then
            KeyAscii = 0
            Exit Sub
    ElseIf KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub
    End If
     
    Dim l%
      l = InStr(Grid1.EditText, ".")
      If Grid1.EditSelStart < l Then Exit Sub
      If l > 0 Then
       If (Len(Grid1.EditText) - l) >= afterdot Then KeyAscii = 0
      End If
End Sub

Public Sub textboxcustomize(text As TextBox, KeyAscii As Integer, Optional afterdot As Integer)
    If Chr(KeyAscii) = "." And afterdot = 0 Then KeyAscii = 0: Exit Sub
    
    Select Case KeyAscii
    
        Case 46:
                If InStr(text, ".") Or (Len(text) - text.SelStart) > afterdot Then KeyAscii = 0
                Exit Sub
    End Select
    If (KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack) And (Not IsNumeric(Chr$(KeyAscii))) Then
            KeyAscii = 0
            Exit Sub
    ElseIf KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub
    End If
     
    Dim l%
      l = InStr(text, ".")
      If text.SelStart < l Then Exit Sub
      If l > 0 Then
       If (Len(text) - l) >= afterdot Then KeyAscii = 0
      End If
End Sub

Public Sub GridLength(Grid1 As Control, KeyAscii As Integer, MaxLength)
    If KeyAscii <> 8 Then
        If Len(Grid1.EditText) > MaxLength Then KeyAscii = 0
    End If
End Sub

Public Function GridSum(objGrid As Control, Col As Integer)
    Dim i As Integer
    Dim sum As Double
   For i = 1 To objGrid.Rows - 1
          sum = sum + Val(IIf(objGrid.TextMatrix(i, Col) = "", "0", objGrid.TextMatrix(i, Col)))
   Next
         GridSum = sum
 End Function
 
'----------------------------------------------------------------------------
'                                             Help Page
'----------------------------------------------------------------------------
'Public Sub ShowWinHelp(HelpFile As String, _
'                    Optional View As String = "TOC")
'
'   Const HELP_TAB = &HF&
'   Const TB_CONTENTS = 0
'   Const TB_INDEX = -2
'   Const TB_FIND = -1
'   Dim ViewConst As Long
'
'   Select Case View
'      Case "TOC"
'         ViewConst = TB_CONTENTS
'      Case "INDEX"
'         ViewConst = TB_INDEX
'      Case "FIND"
'         ViewConst = TB_FIND
'      Case Else
'         Exit Sub
'   End Select
'   WinHelp 0, HelpFile, HELP_TAB, ViewConst
'
'End Sub
 
 

'---------------------------------------------------------------------------------------------------------------------
'This function is used only for numaric value convert word Don't touch this function this function is 100% accurate---
'---------------------------------------------------------------------------------------------------------------------
Function InWords(ByVal GetAmount As Variant) As String
On Error GoTo Kick_Errors 'if error goto Kick_Errors Labels
'Declare some necessary variable
Dim tempNum As Integer
Dim getTaka As Variant
Dim getPaisa As Integer, getPaisaainWords As String
Dim AmountinWords As String
Dim Arrindex As Integer
'Dim NumInWord1 As Double

'Check whether getAmount contain valid number
If Not IsNumeric(GetAmount) Then Exit Function

'Check whether getAmount>999999999.99
If GetAmount > 999999999.99 Then Exit Function

'array for thousand and million only that calculate here
NumInWord1 = Array(" ", "Thousand", "Million")

GetAmount = Abs(GetAmount)              'make positive
getTaka = Int(GetAmount)                'get taka part
getPaisa = (GetAmount - getTaka) * 100  'get taka part
If getTaka > 0 Then         'if there is taka,
                            'the following Loop use to get
                            'hundreds,thousands, then millions.

Do
    tempNum = getTaka Mod 1000
    getTaka = Int(getTaka / 1000)
'Set output
If tempNum <> 0 Then
    AmountinWords = GetAmWords(tempNum) & " " & _
                    NumInWord1(Arrindex) & " " & AmountinWords
    End If
    Arrindex = Arrindex + 1
Loop While getTaka > 0
If getPaisa > 0 Then
    getPaisaainWords = GetAmWords(getPaisa)
    AmountinWords = RTrim(AmountinWords) & _
                    "Taka and" & getPaisaainWords & " Paisa"
Else
    AmountinWords = RTrim(AmountinWords) & " Taka Only"
    End If
End If
getOut: 'label getOut
    InWords = AmountinWords
    Exit Function

Kick_Errors: 'label Kick_Errors
    'If text box contain wrong data just return empty string
    AmountinWords = " "
    Resume getOut
End Function

Function GetAmWords(ByVal GetAmount As Integer) As String
    Static UnitOnes As Variant
    Static UnitTens As Variant
    Dim AmountinWords As String
    Dim getNumDigit As Integer
    'Set UnitOnes if have no elements
If IsEmpty(UnitOnes) Then
    UnitOnes = Array(" ", "One", "Two", "Three", "Four", _
                     "Five", "Six", "Seven", "Eight", "Nine", "Ten", _
                     "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", _
                     "Sixteen", "Seventeen", "Eighteen", "Nineteen", "Twenty")

End If

'What about others
If IsEmpty(UnitTens) Then
    UnitTens = Array(" ", " ", "Twenty", "Thirty", "Forty", "Fifty", _
                     "Sixty", "Seventy", "Eighty", "Ninety")

End If

'Calculate hundreds and rest value
getNumDigit = GetAmount \ 100
GetAmount = GetAmount Mod 100
'If hundred found
If getNumDigit > 0 Then
    AmountinWords = UnitOnes(getNumDigit) & "  Hundred"
End If
'Select Word for Unit Ones and Tens
Select Case GetAmount
    Case 1 To 20 'get from UnitOnes array
            AmountinWords = AmountinWords & _
                            " " & UnitOnes(GetAmount)
 
    Case 21 To 99 'get from UnitOnes array
            getNumDigit = GetAmount \ 10
            GetAmount = GetAmount Mod 10
 
 If getNumDigit > 0 Then
    AmountinWords = AmountinWords & _
                    " " & UnitTens(getNumDigit)
 
 End If
 
 If GetAmount > 0 Then
    AmountinWords = AmountinWords & _
                    " " & UnitOnes(GetAmount)
    End If
 End Select
    GetAmWords = AmountinWords
 End Function
 
'End value convert in word function ---------------------------------------



