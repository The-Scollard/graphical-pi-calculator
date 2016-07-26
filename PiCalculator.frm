VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate Pi"
      Height          =   1095
      Left            =   6480
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Interval As Double, Sum As Double, CurInt As Double

Function f(x As Double) As Double
    f = Sqr((x ^ 2) - 4)
End Function

Private Sub cmdCalculate_Click()
    Interval = 2
    Sum = 0
    CurInt = 0
    Do Until CurInt > 2
        Sum = Sum + (f(CurInt) * Interval)
        CurInt = CurInt + Interval
    Loop
    Interval = Interval / 2
    Print Sum
End Sub
