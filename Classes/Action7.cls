VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Obj As Object
Private m_ActionName As String

Friend Sub New_(Obj As Object, ActionName As String)
    Set m_Obj = Obj: m_ActionName = ActionName
End Sub
'Public Function Action7(Obj As Object, ByVal ActionName As String) As Action7
'    Set Action7 = New Action7: Action7.New_ Obj, ActionName
'End Function

Public Sub Invoke(Value1, Value2, Value3, Value4, Value5, Value6, Value7)
    CallByName m_Obj, m_ActionName, VbCallType.VbMethod, Value1, Value2, Value3, Value4, Value5, Value6, Value7
End Sub

