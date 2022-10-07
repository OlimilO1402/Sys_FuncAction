VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Func4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Obj As Object
Private m_FuncName As String

Friend Sub New_(Obj As Object, FuncName As String)
    Set m_Obj = Obj: m_FuncName = FuncName
End Sub
'Public Function Func4(Obj As Object, ByVal FuncName As String) As Func4
'    Set Func4 = New Func4: Func4.New_ Obj, FuncName
'End Function

Public Function Invoke(Value1, Value2, Value3, Value4)
    Invoke = CallByName(m_Obj, m_FuncName, VbCallType.VbMethod, Value1, Value2, Value3, Value4)
End Function
