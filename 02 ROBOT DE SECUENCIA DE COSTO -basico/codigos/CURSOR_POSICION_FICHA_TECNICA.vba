Attribute VB_Name = "CURSOR_POSICION_FICHA_TECNICA"
' Access the GetCursorPos function in user32.dll
      Declare Function GetCursorPos Lib "user32" _
      (lpPoint As POINTAPI) As Long
      ' Access the GetCursorPos function in user32.dll
      Declare Function SetCursorPos Lib "user32" _
      (ByVal X As Long, ByVal y As Long) As Long

      ' GetCursorPos requires a variable declared as a custom data type
      ' that will hold two integers, one for x value and one for y value
      Type POINTAPI
         X_Pos As Long
         Y_Pos As Long
      End Type

      ' Main routine to dimension variables, retrieve cursor position,
      ' and display coordinates
      Sub Get_Cursor_Pos()

      ' Dimension the variable that will hold the x and y cursor positions
      Dim Hold As POINTAPI

      ' Place the cursor positions in variable Hold
      GetCursorPos Hold

      ' Display the cursor position coordinates
      MsgBox "X Position is : " & Hold.X_Pos & Chr(10) & _
         "Y Position is : " & Hold.Y_Pos
      End Sub

      ' Routine to set cursor position
      Sub Set_Cursor_Pos()

      ' Looping routine that positions the cursor
         For X = 1 To 480 Step 20
            SetCursorPos X, X
            For y = 1 To 40000: Next
         Next X
      End Sub

Sub modelo_cursor()
Application.Wait (Now + TimeValue("0:00:40"))

Call Get_Cursor_Pos

End Sub

