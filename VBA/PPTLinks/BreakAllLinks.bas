Attribute VB_Name = "BreakAllLinks"
Sub BreakAllLinks()
Dim shp As shape
Dim sld As slide

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
        On Error Resume Next
            shp.LinkFormat.BreakLink
    Next shp
Next sld

End Sub
