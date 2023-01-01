Attribute VB_Name = "MNew"
Option Explicit

Public Function Nullable(ByVal vt As EVbVarType) As Nullable
    Set Nullable = New Nullable: Nullable.New_ vt
End Function
