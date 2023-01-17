Imports Microsoft.Office.Tools.Ribbon

Public Class ManageTaskPaneRibbon
    Private Sub ShowTextEditPane_Click(sender As Object, e As RibbonControlEventArgs) Handles TextEditPaneShow.Click
        Globals.ThisAddIn.ChangeTaskPaneVisible(Globals.ThisAddIn.Application.ActiveWindow)
    End Sub
End Class
