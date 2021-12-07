Attribute VB_Name = "Slowdown"
Option Explicit

Private Slowdowns As Object
Public Running As Boolean

Private Sub Auto_Open()
    Running = False
    Set Slowdowns = CreateObject("Scripting.Dictionary")
    Dim i As Integer
    i = 5
    Do While Not Cells(i, "A") = ""
        If Not Cells(i, "I") = "" Then
            Slowdowns(CStr(Cells(i, "A"))) = Cells(i, "I")
        End If
        i = i + 1
    Loop
End Sub

Public Sub UpdateSlowdowns(where As String)
    If Slowdowns Is Nothing Then
        Set Slowdowns = CreateObject("Scripting.Dictionary")
    End If
    Dim Intersection As Range
    Set Intersection = intersect(Range(where), Range("I:I"))
    If Not Intersection Is Nothing Then
        Dim Cell As Range
        For Each Cell In Range(where).Rows
            Slowdowns(CStr(Cells(Cell.Row, "A"))) = CStr(Cell)
        Next
    Else
        Set Intersection = intersect(Range(where), Range("D:D"))
        If Not Intersection Is Nothing Then
            Dim CurrentDate As String
            Dim CurrentSlowdown As String
            Dim i As Integer
            i = Range(where).Row - 1
            Running = True
            Do While Not Cells(i, "A") = ""
                CurrentDate = Cells(i, "A")
                If Slowdowns.Exists(CurrentDate) And Not Cells(i, "I") = Slowdowns(CurrentDate) Then
                    If Not CurrentDate = Cells(i - 1, "A") Then
                        Cells(i, "I") = Slowdowns(CurrentDate)
                    Else
                        Cells(i, "I") = ""
                    End If
                End If
                i = i + 1
            Loop
            Running = False
        End If
    End If
End Sub
