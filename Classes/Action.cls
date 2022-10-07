VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Kapselt eine Methode, die über keine Parameter verfügt und keinen Wert zurückgibt.
Private m_Obj As Object
Private m_ActionName As String

Friend Sub New_(Obj As Object, ActionName As String)
    Set m_Obj = Obj: m_ActionName = ActionName
End Sub
'Public Function Action(Obj As Object, ByVal ActionName As String) As Action
'    Set Action = New Action: Action.New_ Obj, ActionName
'End Function

Public Sub Invoke()
    CallByName m_Obj, m_ActionName, VbCallType.VbMethod
End Sub

