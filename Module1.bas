Attribute VB_Name = "Module1"
'Program written by Amol A. Ambardekar
Global ImagePixels(0 To 2, 0 To 600, 0 To 600) As Integer
Global x As Integer, y As Integer
Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Sub col3(ByRef Picture1 As PictureBox, Optional threshold As Integer = 200)
Dim i As Integer, j As Integer
Dim red As Integer, green As Integer, blue As Integer
Dim pixel As Long
'Form1.Refresh
On Error Resume Next
x = Picture1.ScaleWidth
y = Picture1.ScaleHeight
If x > 600 Or y > 600 Then
    MsgBox "Image too large to process. Please try loading a smaller image."
    x = 0
    y = 0
    Exit Sub
End If

'Form1.Width = Form1.ScaleX(Picture1.Width + 6, vbPixels, vbTwips)
'Form1.Height = Form1.ScaleY(Picture1.Height + 30, vbPixels, vbTwips)
'Form1.Refresh

'Form3.Show
'Form3.Refresh
'Dim asdf As Long
'Dim meanred As Double, meangreen As Double, meanblue As Double
    Dim intensity As Double
    asdf = 0
    For i = 0 To y - 1
        For j = 0 To x - 1
            pixel = Picture1.Point(j, i)
            red = pixel& Mod 256
            green = ((pixel And &HFF00) / 256&) Mod 256&
            blue = (pixel And &HFF0000) / 65536
            'ImagePixels(0, i, j) = red
            'ImagePixels(1, i, j) = green
            'ImagePixels(2, i, j) = blue
            'meanred = meanred + red
            'meangreen = meangreen + green
            'meanblue = meanblue + blue
            'asdf = asdf + 1
            intensity = 0.3 * red + 0.59 * green + 0.11 * blue
            If intensity > threshold Then
            SetPixelV Picture1.hdc, j, i, vbWhite
            Else
            SetPixelV Picture1.hdc, j, i, vbBlack
            End If
        Next j
DoEvents
        'Form3.ProgressBar1.Value = i * 100 / (Y - 1)
    Next i
End Sub



Function Strip_Vert(Picture1 As PictureBox)
    Dim startChar, endChar
    Dim charStarted As Boolean
    Dim pixel
    Dim isBlank As Boolean
    
    For x = 0 To (Picture1.Width - 1)
    
        isBlank = True
        
        For y = 0 To (Picture1.Height - 1)
            pixel = Hex(Picture1.Point(x, y))
            If pixel = 0 Then isBlank = False
        Next
        If x = (Picture1.Width - 1) Then
            If Not isBlank Then
                endChar = x
                Call Get_Code(Picture1, startChar, endChar)
                Debug.Print "Char at " & startChar & ", " & endChar
            End If
        End If
        
        If isBlank Then
            If charStarted Then
                endChar = x - 1
                charStarted = False
                Call Get_Code(Picture1, startChar, endChar)
                Debug.Print "Char at " & startChar & ", " & endChar
            End If
        Else
            If Not charStarted Then
                startChar = x
                charStarted = True
            End If
        End If
    Next

End Function

Function Get_Code(Picture1 As PictureBox, cStart, cEnd)
    Dim strTemp
    Dim strCode
    Dim tmpPixel
    Dim strBlank
    
    For w = 0 To (cEnd - cStart)
        strBlank = strBlank & "0"
    Next
    
    For cy = 0 To (Picture1.Height - 1)
        For cx = cStart To cEnd
            tmpPixel = Hex(Picture1.Point(cx, cy))
            
            If tmpPixel = "FFFFFF" Then
                strTemp = strTemp & "0"
            Else
                strTemp = strTemp & "1"
            End If
        Next
        If strTemp = strBlank Then
        
        Else
            strCode = strCode & strTemp
        End If
        
        strTemp = ""
    Next



    'Send data to be guessed and stored.
    Call Check_it(strCode)
End Function

Function Check_it(ckcCode)
    
    Dim strBuffer
    
    FileSystem.ChDir (App.Path)
    Filename_Database = "System.ipf"
    
    Open Filename_Database For Binary As #1
            While Not EOF(1)
                Line Input #1, strBuffer
                strBuffer = Split(strBuffer, ",")
                If strBuffer(0) = ckcCode Then
                    Form1.Text1.Text = Form1.Text1.Text & strBuffer(1)
                    Form1.List1.AddItem ckcCode
                    Close #1
                    Exit Function
                End If
            Wend
            Close #1
    Close #1
    
    Form1.Text1.Text = Form1.Text1.Text & "*"
    Form1.List1.AddItem ckcCode
End Function

Function add_dialog(sndCode)

    addChar.Text2.Text = sndCode
    addChar.Show
End Function
