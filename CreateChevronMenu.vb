
Sub buttons_CreateMenu_MAIN()
    Debug.Print buttons_CreateMenu(Range("C6").Left, Range("c6").Top, Range("C6:D6").Width, Range("C6:C7").Height, msoShapeChevron)
End Sub
Function buttons_CreateMenu(leftI As Long, topI As Long, btnWidthI As Long, btnHeightI As Long, btnType As MsoAutoShapeType) As Boolean
    Dim dTarget As New Scripting.Dictionary, shpI As Shape, SH As Variant
    
    buttons_CreateMenu = False
    'On Error GoTo endWithError
    
    dTarget.Add shHome, ""
    dTarget.Add shFiles, ""
    
    For Each SH In dTarget
        If shpI Is Nothing Then
            Set shpI = buttons_CreateMenu_addButton(leftI, topI, btnWidthI, btnHeightI, btnType, Sheets(SH.Name))
        Else
            Set shpI = buttons_CreateMenu_addButton(shpI.Left + shpI.Width - 5, shpI.Top, shpI.Width, shpI.Height, btnType, Sheets(SH.Name))
        End If
    Next
    
    'TURNS TRUE
    buttons_CreateMenu = True
endWithError:
    'REMAINS FALSE
    Debug.Print Err.Number, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Function buttons_CreateMenu_addButton(leftI As Long, topI As Long, btnWidthI As Long, btnHeightI As Long, btnType As MsoAutoShapeType, hTarget As Worksheet) As Shape
    Dim BTTN As Variant
    
    Set BTTN = ActiveSheet.Shapes.AddShape(msoShapeChevron, Left:=leftI, Top:=topI, Width:=btnWidthI, Height:=btnHeightI)
    With BTTN.Parent.Shapes(BTTN.Name)
        With .TextFrame
            .Characters.Text = hTarget.Name
        End With
        
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        .ShapeStyle = msoShapeStylePreset77
    End With
    With BTTN.Parent
        .Hyperlinks.Add Anchor:=BTTN.Parent.Shapes(BTTN.Name), Address:="", SubAddress:="'" & hTarget.Name & "'!A1", ScreenTip:=hTarget.Name
    End With
    
    Set buttons_CreateMenu_addButton = BTTN
End Function
