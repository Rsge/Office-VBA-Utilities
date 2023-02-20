Attribute VB_Name = "Reference"
Attribute VB_Description = "Base constants for all speed tests."
'@Folder("Base")
'@ModuleDescription("Base constants for all speed tests.")
Option Explicit

' Constants
'@VariableDescription("Label text for iteration count output.")
Public Const IterationCountLabel As String = "Number of iterations: "
Attribute IterationCountLabel.VB_VarDescription = "Label text for iteration count output."
'@VariableDescription("Format of decimal number string output.")
Public Const NumberFormat As String = "0.0###"
Attribute NumberFormat.VB_VarDescription = "Format of decimal number string output."
'@VariableDescription("Unit of measured time.")
Public Const Unit As String = " s"
Attribute Unit.VB_VarDescription = "Unit of measured time."
'@VariableDescription("Exponent of 10 for amount of iterations to do in testing.")
Public Const IterationsExponent As Long = 8
Attribute IterationsExponent.VB_VarDescription = "Exponent of 10 for amount of iterations to do in testing."
