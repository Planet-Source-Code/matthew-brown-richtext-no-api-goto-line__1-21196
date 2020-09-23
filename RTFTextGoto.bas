Attribute VB_Name = "Module1"
' FAST RICHTEXTBOX LINE GOTO
' Author: Matthew Brown, MMComputers
'
' This function allows you to GOTO a line in a rich
'  text box, without using API
' It is very fast. IT DOES NOT just scan through each
' line in the box 'till it finds the right one, but
' finds it by Trial and Improvement.
' This is very fast for very long lines-textboxes!
'
' Note this takes the Sensible Line Number, ie
' First Line = line 1 (Not line 0 as the RichTextBox
' uses!)
'
' If WhichLine = 0 then it puts you on the first line
' If WhichLine is greater than the total number of lines
' it puts you at the last character!
'
' ENJOY!


Public Sub SetCursorAtLine(WhichLine As Long, WhichRTFText As RichTextBox)

Dim Estimate As Long, StartP As Long, EndP As Long
Dim NumChars As Long

' nb: I've modified this from my actual app ready for posting.
' if it doesn't work, you may have to change 'WhichRTFText' (which is
' used below in the WITH...END WITH to a more direct link
' ie. I actually had: mdiMain.ActiveForm.rtfText
' I havn't been able to test this version, BUT IT SHOULD STILL WORK FINE!!!!

With WhichRTFText
    ' Maximum the estimate can be!
    NumChars = Len(.Text)

    ' Its already going to be on the right line!
    If NumChars = 0 Then
        Exit Sub
    End If
    
    ' Check if the given line is out of bounds, or Line 1
    If WhichLine <= 1 Then
        .SelStart = 0
        .SelLength = 0
        Exit Sub
    ElseIf WhichLine > (.GetLineFromChar(NumChars) + 1) Then
        .SelStart = NumChars
        .SelLength = 0
        Exit Sub
    End If
        
    ' Make first estimate
    Estimate = Int(NumChars / 2)
    StartP = 1
    EndP = NumChars

    Dim Finalised As Long ' This is not important - see later

    Do
        If WhichLine < (.GetLineFromChar(Estimate) + 1) Then
            ' estimate too big, refine...
            StartP = StartP
            EndP = Estimate
            Estimate = StartP + Int((EndP - StartP) / 2)
        ElseIf WhichLine > (.GetLineFromChar(Estimate) + 1) Then
            ' estimate too small, refine...
            StartP = Estimate
            EndP = EndP
            Estimate = StartP + Int((EndP - StartP) / 2)
        Else ' is equal! We've found the line
            Finalised = Estimate
            ' Although we know a character IN the line,
            ' this Do...Loop finds the first character on the line
            Do
                Finalised = Finalised - 1
                If Finalised = 0 Then
                    'Finalised = 1
                    .SelStart = Finalised
                    .SelLength = 0
                    Exit Do
                Else
                    If (.GetLineFromChar(Finalised) + 1) < WhichLine Then
                        Finalised = Finalised + 1
                        .SelStart = Finalised
                        .SelLength = 0
                        Exit Do
                    End If
                End If
            Loop
            Exit Do
        End If
    Loop
End With
End Sub
