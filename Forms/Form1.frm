VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command0 
      Caption         =   "Command0"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mTester As TesterFuncAction


Private Sub Form_Load()
    Me.Caption = App.EXEName & " " & App.Major & "." & App.Minor & "." & App.Revision
    Set mTester = New TesterFuncAction
End Sub

Private Sub Command0_Click()
    
    'Testing Calling a Sub and a Function without any parameters
    Dim a As Action: Set a = MNew.Action(mTester, "TestAction")
    a.Invoke
    
    Dim f As Func: Set f = MNew.Func(mTester, "TestFunc")
    MsgBox f.Invoke
    
End Sub

Private Sub Command1_Click()
    'Testing Calling a Sub and a Function with one ByVal-parameter
    'The function-signature may have every datatype
    Dim a As Action1: Set a = MNew.Action1(mTester, "TestAction1")
    a.Invoke 123
    
    Dim f As Func1: Set f = MNew.Func1(mTester, "TestFunc1")
    MsgBox f.Invoke(213)
    
End Sub

Private Sub Command2_Click()
    'Testing Calling a Sub and a Function with one ByRef-parameter v1 and one ByVal-parameter v2
    Dim v1: v1 = 123
    Dim v2 As Integer: v2 = 32156
    
    Dim a As Action2: Set a = MNew.Action2(mTester, "TestAction2")
    a.Invoke v1, v2
    MsgBox v1
    
    Dim f As Func2: Set f = MNew.Func2(mTester, "TestFunc2")
    MsgBox f.Invoke(v1, v2)
    MsgBox v1
    
End Sub

Private Sub Command3_Click()
    'Testing Calling a Sub and a Function with 3 ByVal-parameters
    Dim a As Action3: Set a = MNew.Action3(mTester, "TestAction3")
    a.Invoke 123, 32165, 214748364
    
    Dim f As Func3: Set f = MNew.Func3(mTester, "TestFunc3")
    MsgBox f.Invoke(123, 32165, 214748364)
    
End Sub

Private Sub Command4_Click()
    'Testing Calling a Sub and a Function with 4 ByVal-parameters
    Dim a As Action4: Set a = MNew.Action4(mTester, "TestAction4")
    a.Invoke 123, 32165, 214748364, 1234.5678
    
    Dim f As Func4: Set f = MNew.Func4(mTester, "TestFunc4")
    MsgBox f.Invoke(123, 32165, 214748364, 1234.5678)
    
End Sub

Private Sub Command5_Click()
    'Testing Calling a Sub and a Function with 5 ByVal-parameters
    Dim a As Action5: Set a = MNew.Action5(mTester, "TestAction5")
    a.Invoke 123, 32165, 214748364, 1234.5678, 12345678.12345
    
    Dim f As Func5: Set f = MNew.Func5(mTester, "TestFunc5")
    MsgBox f.Invoke(123, 32165, 214748364, 1234.5678, 12345678.12345)
    
End Sub

Private Sub Command6_Click()
    'Testing Calling a Sub and a Function with 5 ByVal-parameters
    Dim a As Action6: Set a = MNew.Action6(mTester, "TestAction6")
    a.Invoke 123, 32165, 214748364, 1234.5678, 12345678.12345, True
    
    Dim f As Func6: Set f = MNew.Func6(mTester, "TestFunc6")
    MsgBox f.Invoke(123, 32165, 214748364, 1234.5678, 12345678.12345, False)
    
End Sub



Private Sub Command7_Click()
    'Testing Calling a Sub and a Function with 5 ByVal-parameters
    Dim a As Action7: Set a = MNew.Action7(mTester, "TestAction7")
    a.Invoke 123, 32165, 214748364, 1234.5678, 12345678.12345, True, "Dings"
    
    Dim f As Func7: Set f = MNew.Func7(mTester, "TestFunc7")
    MsgBox f.Invoke(123, 32165, 214748364, 1234.5678, 12345678.12345, False, "Dongs")
    
End Sub



Private Sub Command8_Click()
    Dim pl As PropLet: Set pl = MNew.PropLet(mTester, "PropValue")
    Dim pg As PropGet: Set pg = MNew.PropGet(mTester, "PropValue")
    pl.Invoke = 123456789
    Dim v: v = pg.Invoke
    MsgBox v
End Sub
Private Sub Command9_Click()
    Dim pl As PropLet: Set pl = MNew.PropLet(mTester, "PropValue")
    Dim pg As PropGet: Set pg = MNew.PropGet(mTester, "PropValue")
    pl.Invoke = 12345.67890123
    Dim v: v = pg.Invoke
    MsgBox v
End Sub
Private Sub Command10_Click()
    Dim col0 As New Collection: col0.Add 123: col0.Add 456: col0.Add 789
    Dim ps As PropSet: Set ps = MNew.PropSet(mTester, "PropValue")
    Dim pg As PropGet: Set pg = MNew.PropGet(mTester, "PropValue", True)
    Set ps.Invoke = col0
    Dim col1 As Collection: Set col1 = pg.Invoke
    MsgBox col1.Item(1)
End Sub

