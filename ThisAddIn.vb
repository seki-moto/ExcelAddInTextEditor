Imports System.Diagnostics
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Excel

Public Class ThisAddIn

    Private taskPanes As Dictionary(Of Integer, taskPane)

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        taskPanes = New Dictionary(Of Integer, taskPane)

        ' 1つのインスタンスにしようとすると、複数のWindowを開いた場合、次のエラーとなる。
        ' 「基になる RCW から分割された COM オブジェクトを使うことはできません。」
        ' taskPaneSingle = New TaskPaneControl()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
    End Sub

    Private Sub Application_WindowActivate(Wb As Workbook, Wn As Window) Handles Application.WindowActivate
        ChangeTaskPaneVisible(Wn)
    End Sub

    Public Sub ChangeTaskPaneVisible(Wn As Window)
        If taskPanes.ContainsKey(Wn.Hwnd) = False Then
            taskPanes(Wn.Hwnd) = New taskPane With {
                .Pane = Me.CustomTaskPanes.Add(New TaskPaneControl(), "テキスト編集", Wn)
            }
        End If
        taskPanes(Wn.Hwnd).Pane.Visible = Globals.Ribbons.ManageTaskPaneRibbon.TextEditPaneShow.Checked
    End Sub

    Private Class taskPane

        Private WithEvents taskPaneValue As Microsoft.Office.Tools.CustomTaskPane

        Private Sub taskPaneValue_VisibleChanged(sender As Object, e As EventArgs) Handles taskPaneValue.VisibleChanged
            If taskPaneValue.Visible = False Then
                Globals.Ribbons.ManageTaskPaneRibbon.TextEditPaneShow.Checked = False
            End If
        End Sub

        Public Property Pane As Microsoft.Office.Tools.CustomTaskPane
            Get
                Return taskPaneValue
            End Get
            Set(value As Microsoft.Office.Tools.CustomTaskPane)
                taskPaneValue = value
            End Set
        End Property

    End Class

End Class
