VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Nullable"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text2 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton BtnCheckInput 
      Caption         =   "Check the last Input"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton BtnTakeInput 
      Caption         =   "Take my input"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mandatory number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Optional number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Optional As Nullable
Private m_ManValue As Double

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_Optional = MNew.Nullable(vbDouble)
End Sub

Private Sub UpdateView()
    If m_Optional.HasValue Then Text1.Text = Format(m_Optional.Value, "0.000")
    Text2.Text = Format(m_ManValue, "0.000")
End Sub

Private Sub BtnTakeInput_Click()
    Dim d As Double, s As String
    s = Text1.Text
    If (Len(s) Or IsNumeric(s)) Then
        If MString.Double_TryParse(s, d) Then m_Optional.Value = d
    Else
         m_Optional.VarType = vbEmpty
    End If
    s = Text2.Text
    If MString.Double_TryParse(s, d) Then m_ManValue = d
    UpdateView
End Sub

Private Sub BtnCheckInput_Click()
'    Dim mess As String
'    mess = "The first value of datatype " & m_Optional.VarTypeToStr & " is " & Optional_ToStr(m_Optional) & vbCrLf
'    mess = mess & "the User has given "
'    If m_Optional.HasValue Then
'        mess = mess & "the value: " & m_Optional.Value & vbCrLf
'    Else
'        mess = mess & "no value." & vbCrLf
'    End If
'    mess = mess & "The second value of datatype " & MEVbVarType.EVbVarType_ToStr(VarType(m_ManValue)) & " is " & Optional_ToStr(m_ManValue) & vbCrLf
'    mess = mess & "the User has given the value " & m_ManValue
'    MsgBox mess
    Dim mess As String: mess = Check(m_Optional) & vbCrLf & Check(m_ManValue)
    
    MsgBox mess
End Sub

Function Check(v) As String
    Dim s As String
    s = "The variable " & CheckDataType(v) & " is " & CheckOptional(v)
    If TypeOf v Is Nullable Then
        Dim o As Nullable: Set o = v
        
    End If
    '    If m_Optional.HasValue Then
'        mess = mess & "the value: " & m_Optional.Value & vbCrLf
'    Else
'        mess = mess & "no value." & vbCrLf
'    End If

End Function
    
Function CheckDataType(v) As String
    Dim s As String: s = s & "of datatype "
    If TypeOf v Is Nullable Then
        Dim o As Nullable: Set o = v
        s = s & o.VarTypeToStr
    Else
        s = s & MEVbVarType.EVbVarType_ToStr(VarType(v))
    End If
End Function

Function CheckOptional(v) As String
    CheckOptional = IIf(TypeOf v Is Nullable, "optional", "mandatory")
    'If TypeOf v Is Nullable Then
    '    CheckOptional = "optional"
    'Else
    '    CheckOptional = "mandatory"
    'End If
End Function

Function CheckValue(v) As String
    Dim s As String
    If TypeOf v Is Nullable Then
        Dim o As Nullable: Set o = v
        If o.HasValue Then
            
        Else
            
        End If
    '    If m_Optional.HasValue Then
'        mess = mess & "the value: " & m_Optional.Value & vbCrLf
'    Else
'        mess = mess & "no value." & vbCrLf
'    End If

    
End Function
