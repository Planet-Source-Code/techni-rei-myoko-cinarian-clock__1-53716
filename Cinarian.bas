Attribute VB_Name = "Cinarian"
Option Explicit
Public Const pi = 3.14159265358979
Public Const constanants = " Th U M Sh J C W L G Ri N K Ch V Y R D P Z Ki T F B H" 'Dont edit
Public Const vowels = " i a e o" 'theres only room for 4, so U was moved to consanants

Private Const CircleSize As Single = 0.3 ' The circle's radius is this * the radius of the character
Private Const StartSize As Single = 0.4 'This is where the character part starts along y(curvature) if its not tall
Private Const WidthSize As Single = 0.3 'This is where the character part starts along x(curvature) if its not wide

'Translation
Public Function cin2eng_total(Optional text As String) As String
Dim temp As String
If Len(text) > 13 Then text = left(text, 13)
If Len(text) < 13 Then text = text & String(13 - Len(text), "0")

temp = cin2eng_consanant(mid(text, 2, 2))
temp = temp & cin2eng_vowel(mid(text, 4, 1))

temp = temp & cin2eng_consanant(mid(text, 5, 2))
temp = temp & cin2eng_vowel(mid(text, 7, 1))

temp = temp & cin2eng_consanant(mid(text, 8, 2))
temp = temp & cin2eng_vowel(mid(text, 10, 1))

temp = temp & cin2eng_consanant(mid(text, 11, 2))
temp = temp & cin2eng_vowel(mid(text, 13, 1))

If left(text, 1) = "1" Then temp = temp & "."
cin2eng_total = temp
End Function
Public Function cin2eng_consanant(Optional text As String) As String
Dim strarray() As String, temp As Long
strarray = Split(constanants, " ")
If Len(text) < 2 Then text = String(2 - Len(text), "0") & text
temp = Val(left(text, 1)) * 5 + Val(right(text, 1))
cin2eng_consanant = strarray(temp)
End Function
Public Function cin2eng_vowel(Optional text As String) As String
Dim strarray() As String
strarray = Split(vowels, " ")
cin2eng_vowel = strarray(Val(text))
End Function
Public Function dec2cin(number As Long) As String
    If number > 8 Then
        dec2cin = number - 8 & "44"
    Else
        If number > 4 Then
            dec2cin = "0" & number - 4 & "4"
        Else
            dec2cin = "00" & number
        End If
    End If
End Function
Public Function dec2cin2(ByVal number As Long) As String
    Dim temp As String

    If number < 21 Then temp = "0"
    If number > 40 Then temp = "2": number = number - 40
    If number > 20 Then temp = "1": number = number - 20
    
    If number < 5 Then temp = temp & "0"
    If number > 16 Then temp = temp & "4": number = number - 16
    If number > 12 Then temp = temp & "3": number = number - 12
    If number > 8 Then temp = temp & "2": number = number - 8
    If number > 4 Then temp = temp & "1": number = number - 4
    
    temp = temp & number
    dec2cin2 = temp
End Function
Public Function Time2Cin(theTime As Date) As String
    Dim hour As Long, minuteten As Long, minuteone As Long, second As Long, ampm As Long
    hour = Format(theTime, "hh")
    If hour > 12 Or right(theTime, 2) = "PM" Then
        hour = hour - 12
        ampm = 1
    End If
    minuteone = Format(theTime, "nn")
    minuteten = minuteone \ 10
    minuteone = minuteone Mod 10
    second = Format(theTime, "ss")
    Time2Cin = ampm & dec2cin(hour) & dec2cin(minuteten) & dec2cin(minuteone) & dec2cin2(second)
End Function
'Math routines
Public Function findXY(X As Single, Y As Single, distance As Single, angle As Double, Optional isx As Boolean = True) As Single
    If isx Then findXY = X + Sin(angle) * distance Else findXY = Y + Cos(angle) * distance
End Function
Public Function DegreesToRadians(Degrees As Double) As Double 'Converts Degrees to Radians.
    DegreesToRadians = Degrees * (pi / 180)
End Function
Public Function RadiansToDegrees(Radians As Double) As Double
    RadiansToDegrees = Radians / (pi / 180)
End Function
Public Function minmax(left, mid, right) As Long
    minmax = mid
    If mid < left Then minmax = left
    If mid > right Then minmax = right
End Function
Public Function Color(step As Byte, max As Byte, Optional rw As Long = 2, Optional gw As Long = 0, Optional bw As Long = 3) As Long 'creates a gradient of colors
Dim temp As Byte, R As Byte, g As Byte, b As Byte, a As Byte
'rw/gw/bw are the weight values, dont try anything over 3 or under 0, i doubt it'll work
temp = minmax(0, (step / max) * 10, 10)

    a = minmax(0, temp * 5, 255)
    R = 255 - minmax(0, a * rw, 255)
    g = 255 - minmax(0, a * gw, 255)
    b = 255 - minmax(0, a * bw, 255)
    
    'Debug.Print "R: " & r & " G: " & g & " B: " & b & " step: " & step & " max: " & max
    Color = RGB(R, g, b)
End Function

'Graphics routines
Public Sub drawcircle(Main As Object, X As Single, Y As Single, ByVal Height As Single, Optional edgecolor As Long = vbBlack, Optional FillColor As Long = vbGreen, Optional fillmode As Long = 1)
    If fillmode > 0 Then
        Main.FillStyle = 0
        If fillmode = 2 Then FillColor = Color(0, Height * 3)
        Main.FillColor = FillColor
    End If
    
    Main.DrawWidth = 2
    Main.Circle (X, Y), (Height * CircleSize) - 2, edgecolor
    Main.FillStyle = 1
End Sub
Public Sub drawCinarian(Main As Object, X As Single, Y As Single, Height As Single, Optional edgecolor As Long = vbBlack, Optional FillColor As Long = vbGreen, Optional text As String, Optional fillmode As Long = 1, Optional pbottom As Single = -1)
    Dim count As Long, iswide As Boolean, istall As Boolean, strarray() As String, cando As Boolean
    strarray = Split("5 4 3 2 1 0 11 10 9 8 7 6", " ")
    If Len(text) > 13 Then text = left(text, 13)
    If Len(text) < 13 Then text = text & String(13 - Len(text), "0")
    If left(text, 1) = "1" Then drawcircle Main, X, Y, Height, edgecolor, FillColor, fillmode
    For count = 0 To 11
        cando = True
        Select Case mid(text, count + 2, 1)
            Case "1": istall = True: iswide = False
            Case "2": istall = False: iswide = False
            Case "3": istall = True: iswide = True
            Case "4": istall = False: iswide = True
            Case Else: cando = False
        End Select
        If cando Then DrawPart Main, X, Y, Height, edgecolor, FillColor, Val(strarray(count)), iswide, istall, fillmode, pbottom
    Next
End Sub
Private Sub DrawPart(Main As Object, X As Single, Y As Single, ByVal Height As Single, Optional edgecolor As Long = vbBlack, Optional FillColor As Long = vbGreen, Optional number As Long = 0, Optional iswide As Boolean = True, Optional istall As Boolean = True, Optional fillmode As Long = 1, Optional pbottom As Single = -1)
    Const onepart As Long = 360 / 12
    Dim Start As Single, Lft As Double, Width As Long
    
    Start = Height * StartSize
    Width = onepart
    Lft = onepart * number

    If Not istall Then
        Start = Start * 2
        Height = Height * (1 - StartSize)
    End If
    If Not iswide Then
        Width = onepart - onepart * WidthSize
        Lft = Lft + (onepart * WidthSize) / 2
    End If
    
    DrawSemiCircle Main, X, Y, Start, Height, Lft, Width, edgecolor, FillColor, fillmode, 2
End Sub

'Required a complete remaking
Public Sub DrawSemiCircle(Main As Object, X As Single, Y As Single, Start As Single, Radius As Single, angle As Double, Width As Long, Optional edgecolor As Long = vbBlack, Optional FillColor As Long = vbGreen, Optional fillmode As Long = 1, Optional DrawWidth As Long = 1)
    Dim pdegree As Double, L As Double, R As Double
    L = DegreesToRadians(angle)
    R = DegreesToRadians(angle + Width - 1)
    
    If Start > Radius Then
        pdegree = Start
        Start = Radius
        Radius = pdegree
    End If
    
    Main.DrawWidth = 2
    If fillmode > 0 Then
        For pdegree = Start + 1 To Radius - 1 Step 2
            If fillmode = 2 Then FillColor = Color(pdegree - (Start + 1), Radius - Start - 2)  'gradient
            Main.Circle (X, Y), pdegree, FillColor, L, R
            Main.Circle (X, Y), pdegree - 1, FillColor, L, R
        Next
    End If
    
    Main.DrawWidth = DrawWidth
    Main.Circle (X, Y), Radius, edgecolor, L, R
    Main.Circle (X, Y), Start, edgecolor, L, R
    
    pdegree = DegreesToRadians(90)
    L = L + pdegree
    R = R + pdegree
    
    Main.Line (findXY(X, Y, Start, L), findXY(X, Y, Start, L, False))-(findXY(X, Y, Radius, L), findXY(X, Y, Radius, L, False))
    Main.Line (findXY(X, Y, Start, R), findXY(X, Y, Start, R, False))-(findXY(X, Y, Radius, R), findXY(X, Y, Radius, R, False))
End Sub
