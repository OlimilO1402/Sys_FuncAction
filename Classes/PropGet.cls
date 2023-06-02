VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Obj As Object
Private m_PropName As String
Private m_PropIsObj As Boolean

Friend Sub New_(Obj As Object, PropName As String, Optional propisobj As Boolean = False)
    Set m_Obj = Obj: m_PropName = PropName: m_PropIsObj = propisobj
End Sub
'Public Function PropGet(Obj As Object, ByVal PropName As String) As PropGet
'    Set PropGet = New PropGet: PropGet.New_ Obj, PropName
'End Function

Public Function Invoke()
    If m_PropIsObj Then
        Set Invoke = CallByName(m_Obj, m_PropName, VbCallType.VbGet)
    Else
        Invoke = CallByName(m_Obj, m_PropName, VbCallType.VbGet)
    End If
End Function


