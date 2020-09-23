Attribute VB_Name = "ScriptMod"
Option Explicit
Private Enum OP_CODE
    OPadd = 1
    OPsub
    OPmul
    OPdiv
    OPass 'assign:)
    OPint
    OPnum
    OPpri
    OPstr
    OPcstr
    OPinp
    OPiff
    OPthn 'then
    OPndi 'end if
    OPlss 'less than
    OPgrt 'greater than
End Enum
Private Type Node
    OP As OP_CODE
    Pointer As Long
End Type
Private MyNodes() As Node
Private Ints() As Integer
Private IntAlias() As String
Private Strs() As String
Private StrAlias() As String
Private ConstStr() As String
Public Sub CleanCode(AllCode As String)
'This sub cuts all the carraige returns out of the string
Dim CodeSplit As String
Dim CodeSplit2 As String
Do
If InStr(1, AllCode, Chr(13)) >= 1 Then
CodeSplit = Left$(AllCode, InStr(1, AllCode, Chr(13)) - 1)
CodeSplit2 = Right$(AllCode, Len(AllCode) - InStr(1, AllCode, Chr(13)) - 1)
AllCode = CodeSplit & " " & CodeSplit2
End If
Loop Until InStr(1, AllCode, Chr(13)) = 0
'Here we go!
AllCode = AllCode & " "
Compile (AllCode)
End Sub
Private Sub Compile(AllCode As String)
Dim CurCode As String
Dim Temp As Long
Dim Temp2 As Long
Dim TempVar As Integer
Dim i As Integer
ReDim IntAlias(0 To 0)
ReDim MyNodes(0 To 0)
ReDim StrAlias(0 To 0)
ReDim ConstStr(0 To 0)
ScriptForm.OutBox.Text = ""
Do
    'Get one alphanumeric value at a time
CurCode = GetNextString(AllCode)
If CurCode <> "" Then
    ReDim Preserve MyNodes(0 To (UBound(MyNodes, 1) + 1))
    Temp2 = UBound(MyNodes, 1)
    'the big ass select case that turns our code into tokens
    'and other fun stuff
    If IsNumeric(CurCode) = False Then
        Select Case CurCode
        Case "if":
            MyNodes(Temp2).OP = OPiff
            MyNodes(Temp2).Pointer = 0
        Case "<":
            MyNodes(Temp2).OP = OPlss
            MyNodes(Temp2).Pointer = Temp
        Case ">":
            MyNodes(Temp2).OP = OPgrt
            MyNodes(Temp2).Pointer = Temp
        Case "then":
            MyNodes(Temp2).OP = OPthn
            MyNodes(Temp2).Pointer = 0
        Case "endif":
            MyNodes(Temp2).OP = OPndi
            MyNodes(Temp2).Pointer = 0
        Case "input":
            MyNodes(Temp2).OP = OPinp
            MyNodes(Temp2).Pointer = 0
        Case "int":
            CurCode = GetNextString(AllCode)
            Temp = UBound(IntAlias, 1)
            IntAlias(Temp) = CurCode
            MyNodes(Temp2).OP = OPint
            MyNodes(Temp2).Pointer = Temp
            ReDim Preserve IntAlias(0 To (UBound(IntAlias, 1) + 1))
        Case "str":
            CurCode = GetNextString(AllCode)
            Temp = UBound(StrAlias, 1)
            StrAlias(Temp) = CurCode
            MyNodes(Temp2).OP = OPstr
            MyNodes(Temp2).Pointer = Temp
            ReDim Preserve StrAlias(0 To (UBound(StrAlias, 1) + 1))
        Case "=":
            MyNodes(Temp2).OP = OPass
            MyNodes(Temp2).Pointer = Temp
        Case "+":
            MyNodes(Temp2).OP = OPadd
            MyNodes(Temp2).Pointer = 0
        Case "-":
            MyNodes(Temp2).OP = OPsub
            MyNodes(Temp2).Pointer = 0
        Case "*":
            MyNodes(Temp2).OP = OPmul
            MyNodes(Temp2).Pointer = 0
        Case "/":
            MyNodes(Temp2).OP = OPdiv
            MyNodes(Temp2).Pointer = 0
        Case "print":
            MyNodes(Temp2).OP = OPpri
            MyNodes(Temp2).Pointer = 0
        Case Else:
            Temp = -1
            For i = 0 To UBound(IntAlias, 1)
                If IntAlias(i) = CurCode Then
                    MyNodes(Temp2).OP = OPint
                    MyNodes(Temp2).Pointer = i
                    Temp = i
                    Exit For
                End If
            Next i
            For i = 0 To UBound(StrAlias, 1)
                If StrAlias(i) = CurCode Then
                    MyNodes(Temp2).OP = OPstr
                    MyNodes(Temp2).Pointer = i
                    Temp = i
                    Exit For
                End If
            Next i
            If Temp = -1 Then
                MyNodes(Temp2).OP = OPcstr
                MyNodes(Temp2).Pointer = -1
                If InStr(1, CurCode, Chr(34)) = 1 Then
                    CurCode = Right(CurCode, Len(CurCode) - 1)
                    If InStr(1, CurCode, Chr(34)) > 0 Then
                        CurCode = Left(CurCode, Len(CurCode) - 1)
                    Else
                        Do
                            CurCode = CurCode & " " & GetNextString(AllCode)
                            TempVar = TempVar + 1
                            If TempVar = 40 Then
                                TempVar = MsgBox("Either you have a very long string or you forgot both quotes, halt operation?", vbYesNo)
                                If TempVar = 6 Then Exit Do
                            End If
                        Loop Until InStr(1, CurCode, Chr(34)) > 0
                        TempVar = 0
                        CurCode = Left(CurCode, InStr(1, CurCode, Chr(34)) - 1)
                    End If
                End If
                For i = 0 To UBound(ConstStr, 1)
                    If CurCode = ConstStr(i) Then
                    'if the constant already exits why make
                    'another one?
                    MyNodes(Temp2).Pointer = i
                    End If
                Next i
                If MyNodes(Temp2).Pointer = -1 Then
                    MyNodes(Temp2).Pointer = UBound(ConstStr, 1)
                    ConstStr(UBound(ConstStr, 1)) = CurCode
                    ReDim Preserve ConstStr(0 To UBound(ConstStr, 1) + 1)
                End If
            End If
        End Select
    Else
            'this runs if the string is a const number
            MyNodes(Temp2).OP = OPnum
            MyNodes(Temp2).Pointer = Val(CurCode)
    End If
    If ScriptForm.Check1.Value = 1 Then ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & ("(" & Temp2 & ") " & Str$(MyNodes(Temp2).OP) & ":" & Str$(MyNodes(Temp2).Pointer) & vbCrLf)
End If
Loop Until AllCode = ""
End Sub
Private Function GetNextString(ByRef AllCode As String)
'This is our "Parsing" device
If InStr(1, AllCode, " ") >= 1 Then
    Do
        GetNextString = Left$(AllCode, InStr(1, AllCode, " ") - 1)
        AllCode = Right$(AllCode, Len(AllCode) - InStr(1, AllCode, " "))
    Loop Until GetNextString <> "" Or AllCode = ""
End If
End Function
Public Sub Execute()
Dim i As Integer
Dim TempIndex As Integer
Dim TempOP As Byte
Dim TempOP2 As Integer
ReDim Ints(0 To 0)
ReDim Strs(0 To 0)
ScriptForm.OutBox.Text = ""
TempIndex = -1
For i = 0 To UBound(MyNodes, 1)
    If TempOP2 <= -1 Then
        If MyNodes(i).OP = OPndi Then
        TempOP2 = TempOP2 + 1
        End If
        If MyNodes(i).OP = OPiff Then
        TempOP2 = TempOP2 - 1
        End If
    End If
If TempOP2 > -1 Then
    Select Case MyNodes(i).OP
    Case OP_CODE.OPint:
        If UBound(Ints, 1) < MyNodes(i).Pointer + 1 Then
            ReDim Preserve Ints(0 To MyNodes(i).Pointer + 1)
            TempIndex = -1
        Else
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Then
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(i).Pointer)
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Ints(MyNodes(TempIndex).Pointer) <> Ints(MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPlss Then
                If Ints(MyNodes(TempIndex).Pointer) > Ints(MyNodes(i).Pointer) - 1 Then TempOP2 = -1
            End If
            If TempOP = OP_CODE.OPgrt Then
                If Ints(MyNodes(TempIndex).Pointer) < Ints(MyNodes(i).Pointer) + 1 Then TempOP2 = -1
            End If
            If TempOP = OP_CODE.OPadd Then
                If CheckOrd(i + 1) = False Then
                    If MyNodes(i).OP = OPnum Then
                        Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + MyNodes(i).Pointer
                    Else
                        Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + Ints(MyNodes(i).Pointer)
                    End If
                Else
                    i = i + 1
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + DoOrd(i)
                End If
            End If
            If TempOP = OP_CODE.OPsub Then
                If CheckOrd(i + 1) = False Then
                    If MyNodes(i).OP = OPnum Then
                        Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - MyNodes(i).Pointer
                    Else
                        Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - Ints(MyNodes(i).Pointer)
                    End If
                Else
                    i = i + 1
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - DoOrd(i)
                End If
            End If
            If TempOP = OP_CODE.OPdiv Then
                Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) / Ints(MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPmul Then
                Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) * Ints(MyNodes(i).Pointer)
            End If
            If TempOP = OP_CODE.OPpri Then
                ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & Ints(MyNodes(i).Pointer) & vbCrLf
            End If
            If TempOP = OP_CODE.OPinp Then
                Ints(MyNodes(i).Pointer) = InputBox("Enter integer")
            End If
            TempOP = 0
        End If
    Case OP_CODE.OPstr:
        If UBound(Strs, 1) < MyNodes(i).Pointer + 1 Then
            ReDim Preserve Strs(0 To MyNodes(i).Pointer + 1)
            TempIndex = -1
        Else
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Then
                    Strs(MyNodes(TempIndex).Pointer) = Strs(MyNodes(i).Pointer)
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Strs(MyNodes(TempIndex).Pointer) <> Strs(MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPpri Then
                ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & Strs(MyNodes(i).Pointer) & vbCrLf
            End If
            If TempOP = OP_CODE.OPinp Then
                Strs(MyNodes(i).Pointer) = InputBox("Enter string")
            End If
            TempOP = 0
        End If
    Case OP_CODE.OPcstr:
            If TempIndex > -1 And TempOP = OP_CODE.OPass Then
                If TempOP2 = 0 Then
                    Strs(MyNodes(TempIndex).Pointer) = ConstStr(MyNodes(i).Pointer)
                ElseIf TempOP2 = OP_CODE.OPiff Then
                    If Strs(MyNodes(TempIndex).Pointer) <> ConstStr(MyNodes(i).Pointer) Then
                        TempOP2 = -1
                    Else
                        TempOP2 = 0
                    End If
                End If
            End If
            If TempOP = OP_CODE.OPpri Then
                ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & ConstStr(MyNodes(i).Pointer) & vbCrLf
            End If
            TempOP = 0
    Case OP_CODE.OPass:
        TempIndex = i
        TempOP = MyNodes(i).OP
    Case OP_CODE.OPnum:
        If TempOP = OP_CODE.OPpri Then
            ScriptForm.OutBox.Text = ScriptForm.OutBox.Text & MyNodes(i).Pointer & vbCrLf
        End If
        If TempIndex > -1 And TempOP = OP_CODE.OPass Then
            If TempOP2 = 0 Then
                Ints(MyNodes(TempIndex).Pointer) = MyNodes(i).Pointer
            ElseIf TempOP2 = OP_CODE.OPiff Then
                If Ints(MyNodes(TempIndex).Pointer) <> MyNodes(i).Pointer Then
                    TempOP2 = -1
                Else
                    TempOP2 = 0
                End If
            End If
        End If
        If TempOP = OP_CODE.OPlss Then
            If Ints(MyNodes(TempIndex).Pointer) > MyNodes(i).Pointer - 1 Then TempOP2 = -1
        End If
        If TempOP = OP_CODE.OPgrt Then
            If Ints(MyNodes(TempIndex).Pointer) < MyNodes(i).Pointer + 1 Then TempOP2 = -1
        End If
        If TempOP = OP_CODE.OPadd Then
            If CheckOrd(i + 1) = False Then
                If MyNodes(i).OP = OPnum Then
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + MyNodes(i).Pointer
                Else
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + Ints(MyNodes(i).Pointer)
                End If
            Else
                i = i + 1
                Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) + DoOrd(i)
            End If
        End If
        If TempOP = OP_CODE.OPsub Then
            If CheckOrd(i + 1) = False Then
                If MyNodes(i).OP = OPnum Then
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - MyNodes(i).Pointer
                Else
                    Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - Ints(MyNodes(i).Pointer)
                End If
            Else
                i = i + 1
                Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) - DoOrd(i)
            End If
        End If
        If TempOP = OP_CODE.OPmul Then
            Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) * MyNodes(i).Pointer
        End If
        If TempOP = OP_CODE.OPdiv Then
            Ints(MyNodes(TempIndex).Pointer) = Ints(MyNodes(TempIndex).Pointer) / MyNodes(i).Pointer
        End If
        TempOP = 0
    Case OP_CODE.OPpri: TempOP = OP_CODE.OPpri
    Case OP_CODE.OPadd: TempOP = OP_CODE.OPadd
    Case OP_CODE.OPsub: TempOP = OP_CODE.OPsub
    Case OP_CODE.OPmul: TempOP = OP_CODE.OPmul
    Case OP_CODE.OPdiv: TempOP = OP_CODE.OPdiv
    Case OP_CODE.OPinp: TempOP = OP_CODE.OPinp
    Case OP_CODE.OPlss:
    TempOP = OP_CODE.OPlss
    TempIndex = i
    Case OP_CODE.OPgrt:
    TempOP = OP_CODE.OPgrt
    TempIndex = i
    Case OP_CODE.OPiff: TempOP2 = OP_CODE.OPiff
    Case OP_CODE.OPthn: If TempOP2 = OP_CODE.OPiff Then TempOP2 = OP_CODE.OPthn
    Case OP_CODE.OPndi: TempOP2 = 0
    End Select
End If
Next i
End Sub
Private Function CheckOrd(Index As Integer) As Boolean
CheckOrd = False
If MyNodes(Index).OP = OPmul Or MyNodes(Index).OP = OPdiv Then CheckOrd = True
End Function
Private Function DoOrd(ByRef Index As Integer) As Long
'this function does order of Z OPS
Dim Flago As Boolean
Flago = False
Reset:
If Flago = False Then
    If MyNodes(Index - 1).OP = OPnum Then
        DoOrd = MyNodes(Index - 1).Pointer
    Else
        DoOrd = Ints(MyNodes(Index - 1).Pointer)
    End If
End If
Select Case MyNodes(Index).OP
    Case OP_CODE.OPmul:
    If MyNodes(Index + 1).OP = OPnum Then
        DoOrd = DoOrd * MyNodes(Index + 1).Pointer
    Else
        DoOrd = DoOrd * Ints(MyNodes(Index + 1).Pointer)
    End If
    Case OP_CODE.OPdiv:
    If MyNodes(Index + 1).OP = OPnum Then
        DoOrd = DoOrd / MyNodes(Index + 1).Pointer
    Else
        DoOrd = DoOrd / Ints(MyNodes(Index + 1).Pointer)
    End If
End Select
If CheckOrd(Index + 2) Then
    Index = Index + 2
    Flago = True
    GoTo Reset
End If
End Function
