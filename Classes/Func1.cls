VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Func1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Class Func1 a wrapper for CallByName, working as a delegate for a function with 1 argument and a return value
Private m_Obj As Object
Private m_FuncName As String

Friend Sub New_(Obj As Object, FuncName As String)
    Set m_Obj = Obj: m_FuncName = FuncName
End Sub
'Public Function Func1(Obj As Object, ByVal FuncName As String) As Func1
'    Set Func1 = New Func1: Func1.New_ Obj, FuncName
'End Function

Public Function Invoke(Value)
    Invoke = CallByName(m_Obj, m_FuncName, VbCallType.VbMethod, Value)
End Function

