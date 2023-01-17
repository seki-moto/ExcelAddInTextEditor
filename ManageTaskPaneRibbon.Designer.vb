Partial Class ManageTaskPaneRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms クラス作成デザイナーのサポートに必要です
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()>
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'この呼び出しは、コンポーネント デザイナーで必要です。
        InitializeComponent()

    End Sub

    'Component は、コンポーネント一覧に後処理を実行するために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'コンポーネント デザイナーで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャはコンポーネント デザイナーで必要です。
    'コンポーネント デザイナーを使って変更できます。
    'コード エディターを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TabTextEditor4Excel = Me.Factory.CreateRibbonTab
        Me.GrpEditor = Me.Factory.CreateRibbonGroup
        Me.TextEditPaneShow = Me.Factory.CreateRibbonToggleButton
        Me.TabTextEditor4Excel.SuspendLayout()
        Me.GrpEditor.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabTextEditor4Excel
        '
        Me.TabTextEditor4Excel.Groups.Add(Me.GrpEditor)
        Me.TabTextEditor4Excel.KeyTip = "T"
        Me.TabTextEditor4Excel.Label = "テキスト編集"
        Me.TabTextEditor4Excel.Name = "TabTextEditor4Excel"
        '
        'GrpEditor
        '
        Me.GrpEditor.Items.Add(Me.TextEditPaneShow)
        Me.GrpEditor.Label = "操作"
        Me.GrpEditor.Name = "GrpEditor"
        '
        'TextEditPaneShow
        '
        Me.TextEditPaneShow.Description = "作業ウィンドウにテキストエディタを表示します。"
        Me.TextEditPaneShow.ImageName = "TextEdit の表示"
        Me.TextEditPaneShow.KeyTip = "S"
        Me.TextEditPaneShow.Label = "TextEdit の表示"
        Me.TextEditPaneShow.Name = "TextEditPaneShow"
        Me.TextEditPaneShow.OfficeImageId = "GroupSiteSummaryEdit"
        Me.TextEditPaneShow.ScreenTip = "TextEdit の表示"
        Me.TextEditPaneShow.ShowImage = True
        Me.TextEditPaneShow.SuperTip = "TextEditorの作業ウィンドウを表示します。"
        '
        'ManageTaskPaneRibbon
        '
        Me.Name = "ManageTaskPaneRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.TabTextEditor4Excel)
        Me.TabTextEditor4Excel.ResumeLayout(False)
        Me.TabTextEditor4Excel.PerformLayout()
        Me.GrpEditor.ResumeLayout(False)
        Me.GrpEditor.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TabTextEditor4Excel As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents GrpEditor As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents TextEditPaneShow As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ManageTaskPaneRibbon() As ManageTaskPaneRibbon
        Get
            Return Me.GetRibbon(Of ManageTaskPaneRibbon)()
        End Get
    End Property
End Class
