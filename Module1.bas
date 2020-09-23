Attribute VB_Name = "Module1"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Enum ProgLong
    mVB = 1
    mCPlus = 2
    mDelphi = 3
End Enum

Private Type T_RGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Public T_RGB As T_RGB
Public WebColour As String
Public KeepCol As Long
Public Pallet(32) As Long

Public Function InvertColour(C As Long) As Long
Dim R As Long, G As Long, B As Long: R = C
    If R < 0 Then R = -R
    If R > 16777216 Then B = R \ 16777216: R = R - (B * 16777216)
    If R > 65535 Then B = R \ 65536: R = R - (B * 65536)
    If R > 255 Then G = R \ 256: R = R - (G * 256)
    InvertColour = RGB(-(R - 255), -(G - 255), -(B - 255))
End Function

Public Sub Gradient(TheObject As Object, Redval As Integer, Greenval As Integer, Blueval As Integer, TopToBottom As Boolean)
' I did not write this part found on the net somewere can't remmber'
' just like to say thanks for whoever did.

    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / 512)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step / 4
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 0 To 63
        'This draws the colored bar.
        Dim NewR As Integer, NewG As Integer, NewB As Integer
        
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        
        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub


Function RgbtoLong(StrRgb As String) As Long
Dim I As Integer, mColVal As Variant, iCol As Long
    
    mColVal = Split(StrRgb, ",")
    
    For I = LBound(mColVal) To UBound(mColVal)
        iCol = iCol + mColVal(I) * 256 ^ n
        n = n + 1
    Next
    
    Erase mColVal
    RgbtoLong = iCol
    n = 0
    
End Function

Public Function Dec2Web(hDecCol As Long) As String
Dim StrHex As String
    StrHex = Hex(hDecCol)
    Do While Len(StrHex) < 6
        StrHex = "0" & StrHex
        DoEvents
    Loop
    Dec2Web = "#" & Right(StrHex, 2) & Mid(StrHex, 3, 2) & Left(StrHex, 2)
    StrHex = ""
    
End Function

Public Function DoHex(hDecCol As Long, ProgLan As ProgLong) As String
Dim StrHex As String
    StrHex = Hex(hDecCol)
    Do While Len(StrHex) < 6
        StrHex = "0" & StrHex
        DoEvents
    Loop
    
    Select Case ProgLan
        Case mDelphi
            DoHex = "$00" & StrHex
        Case mVB
            DoHex = "&H" & StrHex
        Case mCPlus
            DoHex = "0x00" & StrHex
    End Select
    
End Function

Public Sub LongToRgb(lngCol As Long)
    T_RGB.Red = lngCol And (Not &HFFFFFF00)
    T_RGB.Green = (lngCol And (Not &HFFFF00FF)) \ &H100&
    T_RGB.Blue = (lngCol And Not (&HFF00FFFF)) \ &HFFFF&
End Sub

Function FixPath(lzpath As String)
    If Right(lzpath, 1) = "\" Then FixPath = lzpath Else FixPath = lzpath & "\"
End Function

Function FindFile(lzFileName As String) As Boolean
    If Dir(lzFileName) = "" Then FindFile = False Else FindFile = True
End Function

Public Function SavePallet(lzFileName As String, PalData As String)
Dim tFile As Long, I As Long
Dim sHead As String

    tFile = FreeFile
    sHead = "PAL " & Chr(6 - 2)

    For I = frmnew.Picture1.LBound To frmnew.Picture1.UBound
        Pallet(I) = frmnew.Picture1(I).BackColor
    Next

    Open lzFileName For Binary As #1
        Put #1, , sHead
        Put #1, , Pallet
    Close #1
    
    sHead = ""
    Erase Pallet
    
End Function
