Attribute VB_Name = "Slowdown"
Attribute VB_Description = "Slowdown updating."
'@IgnoreModule IndexedUnboundDefaultMemberAccess
'@Folder "Production planning"
'@ModuleDescription "Slowdown updating."
Option Explicit

'Variables
'@VariableDescription "Dicionary of slowdowns keyed by their dates."
Private m_slowdowns As Object
Attribute m_slowdowns.VB_VarDescription = "Dicionary of slowdowns keyed by their dates."
'@VariableDescription "Check if slowdown realignment is running atm."
Private m_running As Boolean
Attribute m_running.VB_VarDescription = "Check if slowdown realignment is running atm."


'@Description "Gets if the slowdown change is running atm."
Public Property Get IsRunning() As Boolean
Attribute IsRunning.VB_Description = "Gets if the slowdown change is running atm."
    IsRunning = m_running
End Property
'@Description "Sets if the slowdown change is running atm."
Public Property Let IsRunning(ByVal setRunning As Boolean)
Attribute IsRunning.VB_Description = "Sets if the slowdown change is running atm."
    m_running = setRunning
End Property

'@Description "On worksheet opening, initializes slowdown change tracking."
Private Sub Auto_Open()
Attribute Auto_Open.VB_Description = "On worksheet opening, initializes slowdown change tracking."
    IsRunning = False
    Set m_slowdowns = CreateObject("Scripting.Dictionary")
    Dim i As Long
    i = StartingRow
    'Getting all slowdowns with their respective dates
    Do While LenB(ActiveSheet.Cells(i, DateColumn).Value) > 0
        If LenB(ActiveSheet.Cells(i, SlowdownsColumn).Value) > 0 Then
            m_slowdowns(ActiveSheet.Cells(i, DateColumn).Value) = ActiveSheet.Cells(i, SlowdownsColumn).Value
        End If
        i = i + 1
    Loop
End Sub

'@Description "Updates slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
Public Sub UpdateSlowdowns(ByVal where As String)
Attribute UpdateSlowdowns.VB_Description = "Updates slowdowns so each one is applied at it's correct date. Gets a Range object where an update happened."
    'Initing slowdowns dictionary if it's not yet created
    If m_slowdowns Is Nothing Then
        Set m_slowdowns = CreateObject("Scripting.Dictionary")
    End If
    'Evaluating if the update occured in the slowdown column
    Dim intersection As Range
    Set intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Columns(SlowdownsColumn))
    If Not intersection Is Nothing Then
        'If it did, change the slowdown dict accordingly...
        Dim slowdownCell As Range
        Dim dateValue As String
        For Each slowdownCell In ActiveSheet.Range(where).Rows
            dateValue = ActiveSheet.Cells(slowdownCell.Row, DateColumn).Value
            'Add new entries and delete old ones
            If Len(slowdownCell.Value) > 0 Then
                m_slowdowns(dateValue) = slowdownCell.Value
            ElseIf m_slowdowns.Exists(dateValue) Then
                m_slowdowns.Remove dateValue
            End If
        Next
    Else
        'else evaluating if update occured in amount column
        Set intersection = intersect(ActiveSheet.Range(where), ActiveSheet.Columns(AmountColumn))
        If Not intersection Is Nothing Then
            'If it did, updating the slowdown column according to (possibly) changed dates
            Dim currentDate As String
            Dim currentSlowdownCell As Range
            Dim i As Long
            i = ActiveSheet.Range(where).Row - 1
            If i <= StartingRow Then
                i = StartingRow
            End If
            IsRunning = True
            Do While LenB(ActiveSheet.Cells(i, DateColumn)) > 0
                currentDate = ActiveSheet.Cells(i, DateColumn)
                Set currentSlowdownCell = ActiveSheet.Cells(i, SlowdownsColumn)
                'If date has a slowdown, enforcing it, otherwise clearing the cell
                If m_slowdowns.Exists(currentDate) Then
                    'If previous cell's date is same as current's, clearing current cell
                    'Else, setting slowdown to correct value
                    '(Slowdown is only needed once per date)
                    If currentDate = ActiveSheet.Cells(i - 1, DateColumn) Then
                        currentSlowdownCell.Value = vbNullString
                    ElseIf Not currentSlowdownCell.Value = m_slowdowns(currentDate) Then
                        currentSlowdownCell.Value = m_slowdowns(currentDate)
                    End If
                ElseIf Not LenB(currentSlowdownCell) = 0 Then
                    currentSlowdownCell.Value = vbNullString
                End If
                i = i + 1
            Loop
            IsRunning = False
        End If
    End If
End Sub
