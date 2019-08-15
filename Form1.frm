VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1695
      Left            =   960
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Repro1
    Repro2
End Sub

Private Sub Repro1()
    Dim arr() As Variant
    
    ' Calling Array() with no parameters triggers "Invalid procedure call or argument"
    ' This works fine before KB4512508 (or corresponding update for earlier Windows versions) is installed.
    arr = Array()
    
    SubWithArrayArg arr
End Sub

Private Sub Repro2(ParamArray params() As Variant)
    Dim arr() As Variant
    
    ' Setting arr to params fails with "Invalid procedure call or argument" if no arguments were passed to the ParamArray
    ' This works fine before KB4512508 (or corresponding update for earlier Windows versions) is installed.
    arr = params
    
    SubWithArrayArg arr
End Sub

Private Sub SubWithArrayArg(ByRef Parameters() As Variant)
    MsgBox "It ran OK"
End Sub
