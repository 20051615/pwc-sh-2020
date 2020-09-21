Attribute VB_Name = "ChangeAllLinks"

Function EndsWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function

Sub ChangeAllLinks()

Dim oldAddress As String
Dim t As String

Dim shp As Shape
Dim sld As Slide

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
        On Error GoTo Next_shp
            t = shp.LinkFormat.SourceFullName
            If EndsWith(t, ".xlsx") Then
                oldAddress = t
                GoTo Label
            End If
Next_shp:
Resume Next_shp_
Next_shp_:
    Next shp
Next sld

Label:

Dim newAddress As String
Dim fdgOpen As FileDialog
Set fdgOpen = Application.FileDialog(msoFileDialogOpen)
fdgOpen.Title = "Please select the new position"
fdgOpen.InitialFileName = "%USERPROFILE%\Desktop"
fdgOpen.Show
newAddress = fdgOpen.SelectedItems(1)

For Each sld In ActivePresentation.Slides
    For Each shp In sld.Shapes
        On Error GoTo Next_shp_2
            t = shp.LinkFormat.SourceFullName
            shp.LinkFormat.SourceFullName = Replace(t, oldAddress, newAddress)
            shp.LinkFormat.AutoUpdate = ppUpdateOptionAutomatic
Next_shp_2:
Resume Next_shp_2_
Next_shp_2_:
    Next shp
Next sld

End Sub
