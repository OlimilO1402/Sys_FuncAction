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
Implements Func1
'Kapselt eine Property Get, die einen Wert zurückliefert.
Private m_Obj As Object
Private m_PropName As String

Friend Sub New_(Obj As Object, PropName As String)
    Set m_Obj = Obj: m_PropName = PropName
End Sub
'Public Function PropGet(Obj As Object, ByVal PropName As String) As PropGet
'    Set PropGet = New PropGet: PropGet.New_ Obj, PropName
'End Function

Public Function Invoke()
    Invoke = CallByName(m_Obj, m_PropName, VbCallType.VbGet)
End Function

Private Function Func1_Invoke(Value As Variant) As Variant
    Func1_Invoke = Me.Invoke
End Function
