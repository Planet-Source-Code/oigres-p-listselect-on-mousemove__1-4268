VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim DragIndex As Integer



    Private Sub Form_Load()
        List1.Clear
        List1.AddItem "Adam"
        List1.AddItem "Bob"
        List1.AddItem "Charles"
        List1.AddItem "David"
        List1.AddItem "Eric"
        List1.AddItem "Frank"
        List1.AddItem "George"
    End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Deselect item in listbox
List1.ListIndex = -1
End Sub

    Private Sub List1_DragDrop(Source As Control, X As Single, Y As Single)
        ListRowMove Source, DragIndex, ListRowCalc(Source, Y)
    End Sub

    Private Sub List1_MouseDown(Button As Integer, Shift As Integer, _
                               X As Single, Y As Single)
        If Button = vbRightButton Then
            DragIndex = ListRowCalc(List1, Y)
            List1.Drag
        End If
    End Sub


Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label1.Caption = "X= " & X & ":" & "Y= " & Y & ":" & "Itemindex= " & ListRowCalc(List1, Y)
List1.ListIndex = ListRowCalc(List1, Y)



End Sub
