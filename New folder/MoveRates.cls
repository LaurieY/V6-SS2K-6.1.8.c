VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MoveRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

'Default values
Private Const DEFAULT_MOVERATE As Integer = 0

'Registry location
Private Const REGSEC_RATE As String = "Rate"
Private Const REGVAL_MOVERATE As String = "Move Rate Index"
'Private Const REGVAL_SPEED As String = "Move Rate Description"

Private m_sMoveRate As String
Private m_nMoveRate As Integer

Property Let MoveSpeed(sRate As String)
m_sMoveRate = sRate
g_Tracing.Trace (" in Let MoveSpeed with Value " & sRate)
'This is the Code to send to the scope Index 0= RG, 1=RC,
End Property
Property Get MoveSpeed() As String
'This is the Code to send to the scope Index 0= RG, 1=RC,

MoveSpeed = m_sMoveRate
End Property

Property Let MoveRate(nRate As Integer)
    m_nMoveRate = nRate
End Property

Property Get MoveRate() As Integer
'This is the index value
    MoveRate = m_nMoveRate
End Property
Private Sub Class_Initialize()
LoadSettings
End Sub


Public Sub LoadSettings()
    Dim buf As String
 g_Tracing.Trace ("Loading MoveRates settings  6.1.8c trace")
   
'Allow empty value or missing key in the registry
On Error Resume Next
    
    buf = g_Profile.GetValue(g_ID, REGVAL_MOVERATE, REGSEC_RATE)
    If buf <> "" Then
        m_nMoveRate = CInt(buf)
    Else
    g_Tracing.Trace ("saving MoveRates defaults  6.1.8c trace")
    m_nMoveRate = DEFAULT_MOVERATE
        'Value is empty or key is missing. Create key/value using default value
        g_Profile.WriteValue g_ID, REGVAL_MOVERATE, DEFAULT_MOVERATE, REGSEC_RATE
    End If
 
End Sub
Public Sub SaveSettings()
    g_Profile.WriteValue g_ID, REGVAL_MOVERATE, m_nMoveRate, REGSEC_RATE
    'g_Profile.WriteValue g_ID, REGVAL_SPEED, m_sMoveRate, REGSEC_RATE
End Sub
