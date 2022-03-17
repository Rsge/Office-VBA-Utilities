Attribute VB_Name = "Slowdown"
Attribute VB_Description = "Module for methods related to updating the slowdown."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Module for methods related to updating the slowdown."
Option Explicit

'Variables
'@VariableDescription "Dicionary of slowdowns keyed by their dates."
Private Slowdowns As Object
Attribute Slowdowns.VB_VarDescription = "Dicionary of slowdowns keyed by their dates."
'@VariableDescription "Check if slowdown realignment is running atm."
Private Running As Boolean
Attribute Running.VB_VarDescription = "Check if slowdown realignment is running atm."


'@Description "Get if the slowdown change is running atm."
Public Property Get IsRunning() As Boolean
Attribute IsRunning.VB_Description = "Get if the slowdown change is running atm."
    IsRunning = Running
End Property
'@Description "Set if the slowdown change is running atm."
Public Property Let IsRunning(ByVal setRunning As Boolean)
Attribute IsRunning.VB_Description = "Set if the slowdown change is running atm."
    Running = setRunning
End Property

'@Description "On worksheet opening, initialize slowdown change tracking."
Private Sub Auto_Open()
Attribute Auto_Open.VB_Description = "On worksheet opening, initialize slowdown change tracking."
    IsRunning = False
    Set Slowdowns = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = StartingRow
    'Get all slowdowns with their respective dates
    Do While LenB(ActiveSheet.Cells(i, DateColumn)) > 0
        If LenB(ActiveSheet.Cells(i, SlowdownsColumn)) > 0 Then
            Slowdowns(CStr(ActiveSheet.Cells(i, DateColumn))) = ActiveSheet.Cells(i, SlowdownsColumn)
        End If
        i = i + 1
    Loop
End Sub

'@Description "Update slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
Public Sub UpdateSlowdowns(ByVal where As String)
Attribute UpdateSlowdowns.VB_Description = "Update slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
    'Init slowdowns dictionary if it's not yet created
    If Slowdowns Is Nothing Then
        Set Slowdowns = CreateObject("Scripting.Dictionary")
    End If
    'See if the update occured in the slowdown column
    Dim Intersection As Range
    Set Intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Range(Chr$(SlowdownsColumn + ColumnLetterAscii) & Colon & Chr$(SlowdownsColumn + ColumnLetterAscii)))
    If Not Intersection Is Nothing Then
        'If it did, change the slowdown dict accordingly...
        Dim Cell As Range
        For Each Cell In ActiveSheet.Range(where).Rows
            Slowdowns(CStr(ActiveSheet.Cells(Cell.Row, DateColumn))) = CStr(Cell)
        Next
    Else
        'else see, if update occured in amount column
        Set Intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Range(Chr$(AmountColumn + ColumnLetterAscii) & Colon & Chr$(AmountColumn + ColumnLetterAscii)))
        If Not Intersection Is Nothing Then
            'If it did, update the slowdown column according to (possibly) changed dates
            Dim CurrentDate As String
            Dim i As Long
            i = ActiveSheet.Range(where).Row - 1
            IsRunning = True
            Do While LenB(ActiveSheet.Cells(i, DateColumn)) > 0
                CurrentDate = ActiveSheet.Cells(i, DateColumn)
                'Doing this, ignore cells already containing the correct values
                If Slowdowns.Exists(CurrentDate) And Not ActiveSheet.Cells(i, SlowdownsColumn) = Slowdowns(CurrentDate) Then
                    If Not CurrentDate = ActiveSheet.Cells(i - 1, DateColumn) Then
                        ActiveSheet.Cells(i, SlowdownsColumn) = Slowdowns(CurrentDate)
                    Else
                        ActiveSheet.Cells(i, SlowdownsColumn) = vbNullString
                    End If
                End If
                i = i + 1
            Loop
            IsRunning = False
        End If
    End If
End Sub
