VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} userform 
   Caption         =   "UserForm1"
   ClientHeight    =   4164
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "userform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "userform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare array
Dim mark(10) As Integer
Dim i As Integer
Dim average As Integer
Dim j As Integer
Dim Totalmark As Integer

Private Sub cmdAdd_Click()
i = i + 1
 If i > 10 Then Exit Sub
 If Me.txt1 > 100 Then
 MsgBox "Data not valid, enter a mark", vbCritical, "Error"
 Else
 mark(i) = CInt(Me.txt1)
 MsgBox "Mark added", vbInformation
     End If
End Sub

Private Sub CommandButton2_Click()

End Sub

Private Sub cmdCalculate_Click()
 Totalmark = 0
 i = 1
 For i = 1 To 10
 Totalmark = Totalmark + mark(i)
 Next i
 If i > 10 Then
 Me.txt2 = Totalmark / 10
 End If
End Sub

Private Sub cmdShow_Click()
 Static i
 i = i + 1
 If i > 10 Then Exit Sub
    Me.txt1 = mark(i)
End Sub

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub UserForm_Click()

End Sub
