VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Nullable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Value
Private m_vt As EVbVarType

Friend Sub New_(ByVal vt As EVbVarType)
    m_vt = vt
End Sub

Public Property Let VarType(ByVal vt As EVbVarType)
    GetMem2 vt, ByVal Me.Ptr
End Property
Public Property Get VarType() As EVbVarType
    GetMem2 ByVal Me.Ptr, VarType
End Property

Public Function VarTypeToStr() As String
    VarTypeToStr = MEVbVarType.EVbVarType_ToStr(m_vt)
End Function

Public Property Get Ptr() As LongPtr
    Ptr = VarPtr(m_Value)
End Property

Public Property Get Value()
    Value = m_Value
End Property

Public Property Let Value(aValue)
    m_Value = aValue
End Property

Public Property Set Value(aValue)
    Set m_Value = aValue
End Property

Public Property Get HasValue() As Boolean
    If TypeOf m_Value Is Object  Then
        HasValue = Not m_Value Is Nothing
    Else
        HasValue = Not IsEmpty(m_Value)
    End If
End Property

