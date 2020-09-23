Attribute VB_Name = "Module1"
    #If Win16 Then
      Declare Function SendMessage& Lib "User" (ByVal hWnd%, ByVal _
                                  wMsg%, ByVal wParam%, lParam As Any)
    #Else
      Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                     (ByVal hWnd As Long, ByVal wMsg As Long, _
                      ByVal wParam As Long, lParam As Long) As Long
    #End If

    Function ListRowCalc(lstTemp As Control, ByVal Y As Single) As Integer
       #If Win16 Then
           Const WM_USER = &H400
           Const LB_GETITEMHEIGHT = (WM_USER + 34)
       #Else
           Const LB_GETITEMHEIGHT = &H1A1
           'Determines the height of each item in ListBox control in pixels
       #End If
       Dim ItemHeight As Integer
       ItemHeight = SendMessage(lstTemp.hWnd, LB_GETITEMHEIGHT, 0, 0)
       ListRowCalc = min(((Y / Screen.TwipsPerPixelY) \ ItemHeight) + _
                     lstTemp.TopIndex, lstTemp.ListCount - 1)
    End Function

    Function min(X As Integer, Y As Integer) As Integer
         If X > Y Then min = Y Else min = X
    End Function

    Sub ListRowMove(lstTemp As Control, ByVal OldRow As Integer, _
                    ByVal NewRow As Integer)
        Dim SaveList As String, i As Integer

        If OldRow = NewRow Then Exit Sub
        SaveList = lstTemp.List(OldRow)
        If OldRow > NewRow Then
           For i = OldRow To NewRow + 1 Step -1
               lstTemp.List(i) = lstTemp.List(i - 1)
           Next i
        Else
           For i = OldRow To NewRow - 1
               lstTemp.List(i) = lstTemp.List(i + 1)
           Next i
        End If
        lstTemp.List(NewRow) = SaveList
    End Sub


