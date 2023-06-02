VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TesterFuncAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Value As Variant
Private m_List  As Collection

Public Sub TestAction()
    MsgBox "Hi from " & TypeName(Me) & ".TestAction"
End Sub

Public Function TestFunc() As String
    TestFunc = "Hi from " & TypeName(Me) & ".TestFunc"
End Function


Public Sub TestAction1(ByVal Value As Byte)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction1(" & CStr(Value) & ")"
End Sub
Public Function TestFunc1(ByVal Value As Byte) As String
    TestFunc1 = "Hi from " & TypeName(Me) & ".TestFunc1(" & CStr(Value) & ")"
End Function


Public Sub TestAction2(Value1, ByVal Value2 As Integer)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction2(" & CStr(Value1) & ", " & CStr(Value2) & ")"
    Value1 = 101
End Sub
Public Function TestFunc2(Value1, ByVal Value2 As Integer) As String
    TestFunc2 = "Hi from " & TypeName(Me) & ".TestFunc2(" & CStr(Value1) & ", " & CStr(Value2) & ")"
    Value1 = 255
End Function


Public Sub TestAction3(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction3(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ")"
End Sub
Public Function TestFunc3(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long) As String
    TestFunc3 = "Hi from " & TypeName(Me) & ".TestFunc3(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ")"
End Function


Public Sub TestAction4(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Double)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction4(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ")"
End Sub
Public Function TestFunc4(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Double) As String
    TestFunc4 = "Hi from " & TypeName(Me) & ".TestFunc4(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ")"
End Function


Public Sub TestAction5(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ")"
End Sub
Public Function TestFunc5(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double) As String
    TestFunc5 = "Hi from " & TypeName(Me) & ".TestFunc5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ")"
End Function


Public Sub TestAction6(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double, ByVal Value6 As Boolean)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ", " & CStr(Value6) & ")"
End Sub
Public Function TestFunc6(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double, ByVal Value6 As Boolean) As String
    TestFunc6 = "Hi from " & TypeName(Me) & ".TestFunc5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ", " & CStr(Value6) & ")"
End Function


Public Sub TestAction7(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double, ByVal Value6 As Boolean, ByVal Value7 As String)
    MsgBox "Hi from " & TypeName(Me) & ".TestAction5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ", " & CStr(Value6) & ", " & CStr(Value7) & ")"
End Sub
Public Function TestFunc7(ByVal Value1 As Byte, ByVal Value2 As Integer, ByVal Value3 As Long, ByVal Value4 As Single, ByVal Value5 As Double, ByVal Value6 As Boolean, ByVal Value7 As String) As String
    TestFunc7 = "Hi from " & TypeName(Me) & ".TestFunc5(" & CStr(Value1) & ", " & CStr(Value2) & ", " & CStr(Value3) & ", " & CStr(Value4) & ", " & CStr(Value5) & ", " & CStr(Value6) & ", " & CStr(Value7) & ")"
End Function


Public Property Get MyValue() As Variant
    MyValue = m_Value
End Property
Public Property Let MyValue(ByVal Value As Variant)
    m_Value = Value
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property
Public Property Set List(ByVal Value As Collection)
    Set m_List = Value
End Property

