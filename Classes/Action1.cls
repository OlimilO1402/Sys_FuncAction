VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Action1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Func1
'Kapselt eine Methode, die �ber einen einzelnen Parameter verf�gt und keinen Wert zur�ckgibt.
Private m_Obj As Object
Private m_ActionName As String

Friend Sub New_(Obj As Object, ActionName As String)
    Set m_Obj = Obj: m_ActionName = ActionName
End Sub
'Public Function Action1(Obj As Object, ByVal ActionName As String) As Action1
'    Set Action1 = New Action1: Action1.New_ Obj, ActionName
'End Function

Public Sub Invoke(Value)
    CallByName m_Obj, m_ActionName, VbCallType.VbMethod, Value
End Sub

Private Function Func1_Invoke(Value As Variant) As Variant
    Me.Invoke Value
End Function
