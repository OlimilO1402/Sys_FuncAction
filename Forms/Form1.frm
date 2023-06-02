VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command10 
      Caption         =   "Test Prop Set+Get: Object"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Test Prop Let+Get: Double"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Test Prop Let+Get: Long"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Test Func+Action: 7 Params"
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Test Func+Action: 6 Params"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Test Func+Action: 5 Params"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Test Func+Action: 4 Params"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command0 
      Caption         =   "Test Func+Action: 0 Params"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test Func+Action: 3 Params"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Func+Action: 2 Params"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Func+Action: 1 Param"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2415
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
    Dim pl As PropLet: Set pl = MNew.PropLet(mTester, "MyValue")
    Dim pg As PropGet: Set pg = MNew.PropGet(mTester, "MyValue")
    pl.Invoke = 123456789
    Dim v As Long: v = pg.Invoke
    MsgBox v
End Sub
Private Sub Command9_Click()
    Dim pl As PropLet: Set pl = MNew.PropLet(mTester, "MyValue")
    Dim pg As PropGet: Set pg = MNew.PropGet(mTester, "MyValue")
    pl.Invoke = 12345.67890123
    Dim v As Double: v = pg.Invoke
    MsgBox v
End Sub
Private Sub Command10_Click()
    Dim col0 As New Collection: col0.Add "eins": col0.Add "zwei": col0.Add "drei"
    Dim ps As PropSet:    Set ps = MNew.PropSet(mTester, "List")
    Dim pg As PropGetObj: Set pg = MNew.PropGetObj(mTester, "List")
    Set ps.Invoke = col0
    Dim col1 As Collection: Set col1 = pg.Invoke
    MsgBox col1.Item(1) & " " & col1.Item(2) & " " & col1.Item(3)
End Sub

