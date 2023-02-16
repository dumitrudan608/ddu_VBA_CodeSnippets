
Sub shape_CreateMenu_MAIN()
    Debug.Print shape_CreateMenu(Range("C6").Left, Range("c6").Top, 120, 30, msoShapeChevron)
End Sub
Function shape_CreateMenu(leftI As Long, topI As Long, btnWidthI As Long, btnHeightI As Long, btnType As MsoAutoShapeType, Optional shTarget As Worksheet) As Boolean
    Dim dTarget As New Scripting.Dictionary, shpI As Shape, SH As Variant
    
    If shTarget Is Nothing Then: Set shTarget = ActiveSheet
    
    shape_CreateMenu = False
    'On Error GoTo endWithError
    
    dTarget.Add shHome, ""
    dTarget.Add shFiles, ""
    dTarget.Add shProjects, ""
    dTarget.Add shMapping, ""
    dTarget.Add shPreview, ""
    
    For Each SH In dTarget
        If shpI Is Nothing Then
            Set shpI = shape_CreateMenu_addButton(leftI, topI, btnWidthI, btnHeightI, btnType, Sheets(SH.Name), shTarget, True)
        Else
            Set shpI = shape_CreateMenu_addButton(shpI.Left + shpI.Width - 5, shpI.Top, shpI.Width, shpI.Height, btnType, Sheets(SH.Name), shTarget)
        End If
    Next
    
    'TURNS TRUE
    shape_CreateMenu = True
endWithError:
    'REMAINS FALSE
    Debug.Print Err.Number, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Function shape_CreateMenu_addButton(leftI As Long, topI As Long, btnWidthI As Long, btnHeightI As Long, btnType As MsoAutoShapeType, _
    shLink As Worksheet, Optional shTarget As Worksheet, Optional isSelected As Boolean = False) As Shape
    
    Dim BTTN As Variant, btnName As String
    
    If shTarget Is Nothing Then: Set shTarget = ActiveSheet
    
    btnName = "btn" & shTarget.Name & "|" & shLink.Name
    shape_deleteIfExists btnName
    
    Set BTTN = shTarget.Shapes.AddShape(msoShapeChevron, Left:=leftI, Top:=topI, Width:=btnWidthI, Height:=btnHeightI)
    BTTN.Name = btnName
    With BTTN.Parent.Shapes(BTTN.Name)
        
        If isSelected Then
            .ShapeStyle = msoShapeStylePreset41
            .Glow.Radius = 20
            .Glow.Color = RGB(128, 0, 255) 'blue
            .Glow.Transparency = 0.75
        Else
           .ShapeStyle = msoShapeStylePreset74
        End If
        With .TextFrame
            .Characters.Text = shLink.Name
            .Characters.Font.Color = vbWhite
            .Characters.Font.Size = 12
        End With
        
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
        End With
        
        
    End With
    With BTTN.Parent
        .Hyperlinks.Add Anchor:=BTTN.Parent.Shapes(BTTN.Name), Address:="", SubAddress:="'" & shLink.Name & "'!A1", ScreenTip:=shLink.Name
    End With
    
    Set shape_CreateMenu_addButton = BTTN
End Function

Function shape_deleteIfExists(btnName As String, Optional shParent As Worksheet)
    Dim BTN As Variant, isSuccess As Boolean
    
    isSuccess = True    'NOT VALIDATING ANITHING AT THE MOMENT
    
    shape_deleteIfExists = False
    If shParent Is Nothing Then: Set shParent = ActiveSheet
    On Error Resume Next
    Set BTN = shParent.Shapes(btnName)
    If Not BTN Is Nothing Then: BTN.Delete
    On Error GoTo 0
    
    shape_deleteIfExists = isSuccess
    
End Function
