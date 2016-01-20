Attribute VB_Name = "CorelMacros"
Option Explicit

'Call the dialog to create the color swatches
Public Sub CreateColorSwatch()
Attribute CreateColorSwatch.VB_Description = "Generates rectangles filled with colors from any color palette. Users can type in the date and printer name."
    frmColorSwatch.Show
End Sub

'Call the dialog to create the page numbering
Public Sub PageNumbering()
Attribute PageNumbering.VB_Description = "Adds artistic text labels with page numbers to the document. Users can define font styles and location. Works well with a BeforePrint event!"
    frmNumberPage.Show
End Sub

Sub LatexEdit()
    '
    ' Recorded 10.02.2006
    '
    ' Description:
    '
    '
    Dim frmEdit As New frmLatexEdit
    Dim s1 As Shape
    Set s1 = ActiveShape
    If s1 Is Nothing Then
        frmEdit.Show vbModal
    Else
        Load frmEdit
        frmEdit.TextBox1.Text = s1.ObjectData("Comments")
        frmEdit.Show vbModal
    End If
End Sub
