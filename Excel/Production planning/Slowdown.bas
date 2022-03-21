Attribute VB_Name = "Slowdown"
Attribute VB_Description = "Slowdown updating."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Slowdown updating."
Option Explicit

'Variables
'@VariableDescription "Dicionary of slowdowns keyed by their dates."
Private Slowdowns As Object
Attribute Slowdowns.VB_VarDescription = "Dicionary of slowdowns keyed by their dates."
'@VariableDescription "Check if slowdown realignment is running atm."
Private Running As Boolean
Attribute Running.VB_VarDescription = "Check if slowdown realignment is running atm."


'@Description "Gets if the slowdown change is running atm."
Public Property Get IsRunning() As Boolean
Attribute IsRunning.VB_Description = "Gets if the slowdown change is running atm."
    IsRunning = Running
End Property
'@Description "Sets if the slowdown change is running atm."
Public Property Let IsRunning(ByVal setRunning As Boolean)
Attribute IsRunning.VB_Description = "Sets if the slowdown change is running atm."
    Running = setRunning
End Property

'@Description "On worksheet opening, initializes slowdown change tracking."
Private Sub Auto_Open()
Attribute Auto_Open.VB_Description = "On worksheet opening, initializes slowdown change tracking."
    IsRunning = False
    Set Slowdowns = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = StartingRow
    'Getting all slowdowns with their respective dates
    Do While LenB(ActiveSheet.Cells(i, DateColumn).Value) > 0
        If LenB(ActiveSheet.Cells(i, SlowdownsColumn).Value) > 0 Then
            Slowdowns(ActiveSheet.Cells(i, DateColumn).Value) = ActiveSheet.Cells(i, SlowdownsColumn).Value
        End If
        i = i + 1
    Loop
End Sub

'@Description "Updates slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
Public Sub UpdateSlowdowns(ByVal where As String)
Attribute UpdateSlowdowns.VB_Description = "Updates slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
    'Initing slowdowns dictionary if it's not yet created
    If Slowdowns Is Nothing Then
        Set Slowdowns = CreateObject("Scripting.Dictionary")
    End If
    'Evaluating if the update occured in the slowdown column
    Dim Intersection As Range
    Set Intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Columns(SlowdownsColumn))
    If Not Intersection Is Nothing Then
        'If it did, change the slowdown dict accordingly...
        Dim SlowdownCell As Range
        Dim DateValue As String
        For Each SlowdownCell In ActiveSheet.Range(where).Rows
            DateValue = ActiveSheet.Cells(SlowdownCell.Row, DateColumn).Value
            'Add new entries and delete old ones
            If Len(SlowdownCell.Value) > 0 Then
                Slowdowns(DateValue) = SlowdownCell.Value
            ElseIf Slowdowns.Exists(DateValue) Then
                Slowdowns.Remove DateValue
            End If
        Next
    Else
        'else evaluating if update occured in amount column
        Set Intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Columns(AmountColumn))
        If Not Intersection Is Nothing Then
            'If it did, updating the slowdown column according to (possibly) changed dates
            Dim CurrentDate As String
            Dim CurrentSlowdownCell As Range
            Dim i As Long
            i = ActiveSheet.Range(where).Row - 1
            If i <= StartingRow Then
                i = StartingRow
            End If
            IsRunning = True
            Do While LenB(ActiveSheet.Cells(i, DateColumn)) > 0
                CurrentDate = ActiveSheet.Cells(i, DateColumn)
                Set CurrentSlowdownCell = ActiveSheet.Cells(i, SlowdownsColumn)
                'If date has a slowdown, enforcing it, otherwise clearing the cell
                If Slowdowns.Exists(CurrentDate) Then
                    'If previous cell's date is same as current's, clearing current cell
                    'Else, setting slowdown to correct value
                    '(Slowdown is only needed once per date)
                    If CurrentDate = ActiveSheet.Cells(i - 1, DateColumn) Then
                        CurrentSlowdownCell.Value = vbNullString
                    ElseIf Not CurrentSlowdownCell.Value = Slowdowns(CurrentDate) Then
                        CurrentSlowdownCell.Value = Slowdowns(CurrentDate)
                    End If
                ElseIf Not LenB(CurrentSlowdownCell) = 0 Then
                    CurrentSlowdownCell.Value = vbNullString
                End If
                i = i + 1
            Loop
            IsRunning = False
        End If
    End If
End Sub
