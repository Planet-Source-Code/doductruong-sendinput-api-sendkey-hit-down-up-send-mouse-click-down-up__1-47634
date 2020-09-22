Attribute VB_Name = "mdlMain"
'(C) 30030812 DoDucTruong, Truong2D@Yahoo.com
'Special thanks to Robert N. for His code at: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=39798&lngWId=1
'Copy this code to Single-Module Project then hit F5 to run
Option Explicit

Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2

Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_ABSOLUTE = &H8000

Private Type MOUSEINPUT
    dx As Long
    dy As Long
    mouseData As Long
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
    End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
    End Type

Private Type HARDWAREINPUT
    uMsg As Long
    wParamL As Integer
    wParamH As Integer
    End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
    End Type

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Enum ControlKey
    Ctrl = 1
    Alt = 2
    Shift = 4
    Caps = 8
    Win = 16
    PrintScr = 32
    SysPopup = 64
    NumLock = 128
End Enum

Private Const SWP_NOSIDE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOW = 5
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWDEFAULT = 10
Private Const SW_RESTORE = 9
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1
Private Const SW_HIDE = 0

Private Sub SetWindowTopMost(lngHWND As Long)
  Call SetWindowPos(lngHWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIDE)
End Sub

Private Function ActiveWindow(Optional ByVal ClsName As String = vbNull, Optional ByVal WinCaption As String = vbNull, Optional ByVal CMD_SHOW As Long = SW_SHOWNORMAL) As Boolean
    Dim hw As Long, timeout As Long
    hw = FindWindow(ClsName, WinCaption)
    timeout = 0
    While (hw <= 0) And (timeout < 5)
        Wait 100
        hw = FindWindow(ClsName, WinCaption)
        timeout = timeout + 1
    Wend
    If hw > 0 Then
        Debug.Print "Found: " & ClsName & vbTab & WinCaption
        SetWindowTopMost hw
        ShowWindow hw, CMD_SHOW
        ActiveWindow = True
    Else
        Debug.Print "Not Found: " & ClsName & vbTab & WinCaption
        ActiveWindow = False
    End If
End Function

Private Sub SendKey(ByVal vKey As Integer, Optional booDown As Boolean = False)
    Dim GInput(0) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = vKey
    If Not booDown Then
        KInput.dwFlags = KEYEVENTF_KEYUP
    End If
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

Private Sub TypeText(ByVal inTxt As String, Optional intDelay As Integer = 0) 'intDelay x 10ms
    Dim L As Long, i As Long, tmp As String, j As Long
    Dim txt As String, vKey As Integer, booShift As Boolean
    
    txt = UCase(inTxt)
    L = Len(txt)
    For i = 0 To L - 1 Step 1
        tmp = Mid(inTxt, i + 1, 1)
        booShift = False
        vKey = Asc(UCase(tmp))
        Select Case tmp
            Case "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "+", "|", ":", "<", ">", """", "{", "}", "?": booShift = True
            Case "A" To "Z": booShift = True: vKey = Asc(UCase(tmp))
            Case Else: vKey = Asc(UCase(tmp))
        End Select
        
        Dim ExtraKey, strExtraKey
        strExtraKey = Array("!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "~", "_", "+", "|", ":", "<", ">", """", "`", "\", ";", "'", ",", ".", "/", "-", "=", "{", "}", "[", "]", "?")
        ExtraKey = Array(49, 50, 51, 52, 53, 54, 55, 56, 57, 48, 192, 189, 187, 220, 186, 188, 190, 222, 192, 220, 186, 222, 188, 190, 191, 189, 187, 219, 221, 219, 221, 191)
        For j = LBound(ExtraKey) To UBound(ExtraKey) Step 1
            If tmp = strExtraKey(j) Then
                vKey = ExtraKey(j)
                Exit For
            End If
        Next j
        
        
        If booShift Then
            SendKey vbKeyShift, True
        End If
        
        Wait intDelay
        
        'Down
        SendKey vKey, True
        
        'Up
        SendKey vKey, False
        
        If booShift Then
            SendKey vbKeyShift, False
        End If
    Next i
End Sub

Private Sub SendString(ByVal txt As String, Optional booDown As Boolean = False, Optional ByVal enumCtrl As ControlKey = 0)
    Dim GInput() As GENERALINPUT
    Dim KInput As KEYBDINPUT
    Dim L As Long, i As Long, tmp As String
    txt = UCase(txt)
    L = Len(txt)
    ReDim GInput(0 To L - 1) As GENERALINPUT
    For i = 0 To L - 1 Step 1
        tmp = Mid(txt, i + 1, 1)
        Select Case tmp
            Case "*": KInput.wVk = vbKeyMultiply
            Case "+": KInput.wVk = vbKeyAdd
            Case "-": KInput.wVk = vbKeySubtract
            Case "/": KInput.wVk = vbKeyDivide
            Case ".": KInput.wVk = vbKeyDecimal
            Case "?": KInput.wVk = 191
            Case Else: KInput.wVk = Asc(tmp)
        End Select
        If Not booDown Then
            KInput.dwFlags = KEYEVENTF_KEYUP
        End If
        GInput(i).dwType = INPUT_KEYBOARD
        CopyMemory GInput(i).xi(0), KInput, Len(KInput)
    Next i
    If (enumCtrl And Ctrl) Then SendKey vbKeyControl, booDown
    If (enumCtrl And Alt) Then SendKey vbKeyMenu, booDown
    If (enumCtrl And Caps) Then SendKey vbKeyCapital, booDown
    If (enumCtrl And NumLock) Then SendKey vbKeyNumlock, booDown
    If (enumCtrl And PrintScr) Then SendKey vbKeyPrint, booDown
    If (enumCtrl And Shift) Then SendKey vbKeyShift, booDown
    If (enumCtrl And SysPopup) Then SendKey 93, booDown
    If (enumCtrl And Win) Then SendKey 91, booDown
    Call SendInput(L, GInput(0), Len(GInput(0)))
End Sub


'RightDown, RightUp is the same
Private Sub LeftDown()
    Dim GInput(0 To 0) As GENERALINPUT
    Dim KInput As MOUSEINPUT
    KInput.dwFlags = MOUSEEVENTF_LEFTDOWN
    GInput(0).dwType = INPUT_MOUSE
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

Private Sub LeftUp()
    Dim GInput(0 To 0) As GENERALINPUT
    Dim KInput As MOUSEINPUT
    KInput.dwFlags = MOUSEEVENTF_LEFTUP
    GInput(0).dwType = INPUT_MOUSE
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

Private Sub Wait(x10ms)
  Dim t As Long
  t = Timer * 100 + x10ms
  Do
    DoEvents
  Loop While Timer * 100 < t
End Sub

Private Sub Ping(IP As String)
'    SendKey vbKeyEscape, True
'    SendKey vbKeyEscape, False
'    SendString "R", True, Win
'    SendString "R", False, Win
'    ActiveWindow "Run", "MsoCommandBarPopup" '"#32770"
'    TypeText "Cmd", 5
'    SendKey vbKeyReturn, True
'    SendKey vbKeyReturn, False
    
    On Error Resume Next
    Shell "cmd.exe"
    
    If Not Err Then 'OS>Win2000
        If ActiveWindow("ConsoleWindowClass", "C:\WINNT\system32\cmd.exe", SW_SHOWMAXIMIZED) Then
            TypeText "Ping " & IP, 2
            SendKey vbKeyReturn, True
            SendKey vbKeyReturn, False
            Wait 300
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub NotepadHello()
'    SendKey vbKeyEscape, True
'    SendKey vbKeyEscape, False
'    SendString "R", True, Win
'    SendString "R", False, Win
'    ActiveWindow "Run", "MsoCommandBarPopup" '"#32770"
'    TypeText "Notepad", 2
'    SendKey vbKeyReturn, True
'    SendKey vbKeyReturn, False

    Shell "Notepad.exe"
    
    If ActiveWindow("Notepad", "Untitled - Notepad", SW_SHOWMAXIMIZED) Then
        TypeText "Hello, how are you today?", 2
        SendKey vbKeyReturn, True
        SendKey vbKeyReturn, False
        TypeText "!@#$%^&*()~_+|:<>`\;',./-={}[]?"""
        SendKey vbKeyReturn, True
        SendKey vbKeyReturn, False
        TypeText "Do you think about this Code, Please vote it !", 4
        SendString "O", True, Alt
        SendString "O", False, Alt
        SendKey vbKeyF, True
        SendKey vbKeyF, False
        SendString "S", True, Alt
        SendString "S", False, Alt
        TypeText "36", 2
        SendKey vbKeyReturn, True
        SendKey vbKeyReturn, False
        Wait 100
    End If
End Sub

Private Function RunMSPaint() As Boolean
'    SendKey vbKeyEscape, True
'    SendKey vbKeyEscape, False
'    SendKey vbKeyEscape, True
'    SendKey vbKeyEscape, False
'
'    SendString "R", True, Win
'    SendString "R", False, Win
'    ActiveWindow "Run", "MsoCommandBarPopup" '"#32770"
'    TypeText "MSPaint", 2
'    SendKey vbKeyReturn, True
'    SendKey vbKeyReturn, False
    'SendString " ", True, Alt
    'SendString " ", False, Alt
    'SendKey vbKeyX, True
    'SendKey vbKeyX, False
    
    Shell "MSPaint.exe"
    If ActiveWindow("MSPaintApp", "untitled - Paint", SW_SHOWMAXIMIZED) Then
        RunMSPaint = True
    Else
        RunMSPaint = False
    End If
End Function

Private Sub DrawText(ByVal txt As String)
    Dim tmp() As String, i As Long, j As Long
    Dim xyArr() As String
    tmp = Split(txt, ";")
    For i = LBound(tmp) To UBound(tmp) Step 1
        xyArr = Split(tmp(i), ",")
        SetCursorPos xyArr(0) + 200, xyArr(1) + 200
        LeftDown
        For j = LBound(xyArr) + 2 To UBound(xyArr) Step 2
            Wait 10
            'Debug.Print "ij: " & i & vbTab & j
            SetCursorPos xyArr(j) + 200, xyArr(j + 1) + 200
        Next j
        LeftUp
    Next i
End Sub

Private Sub DrawNline()
    If RunMSPaint Then
        SendString "E", True, Ctrl
        SendString "E", False, Ctrl
        TypeText "640"
        SendKey vbKeyTab, True
        SendKey vbKeyTab, False
        TypeText "480"
        SendKey vbKeyReturn, True
        SendKey vbKeyReturn, False
        'DrawText "200,200,500,400"
        DrawText "40,98,40,31,84,98,84,31;100,72,125,72;142,31,142,95,178,95;194,50,194,98;194,32,194,35;216,50,216,97,220,58,228,51,239,52,246,58,245,97;265,72,302,72,299,61,289,52,273,52,263,70,267,87,278,95,293,95,300,85"
    End If
End Sub

Private Sub Main()
    'ActiveWindow "Progman", "Program Manager"
    DrawNline
    'ActiveWindow "Progman", "Program Manager"
    Ping "127.0.0.1"
    'ActiveWindow "Progman", "Program Manager"
    NotepadHello
    
    Shell "Explorer.exe Http://N-Line.co.kr/~dotruong", vbMaximizedFocus
End Sub

