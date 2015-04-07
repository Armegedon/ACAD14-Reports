<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_PointsStaOffset
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_PointsStaOffset))
        Me.GroupBox_Points = New System.Windows.Forms.GroupBox()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Combo_Pgroup = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button_Deselect = New System.Windows.Forms.Button()
        Me.Button_Select = New System.Windows.Forms.Button()
        Me.ListView_Aligns = New System.Windows.Forms.ListView()
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox()
        Me.Label_PgroupNameValue = New System.Windows.Forms.Label()
        Me.Label_PgroupName = New System.Windows.Forms.Label()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.Combo_Align = New System.Windows.Forms.ComboBox()
        Me.Label_SelectAlign = New System.Windows.Forms.Label()
        Me.Button_SelectAlign = New System.Windows.Forms.Button()
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox()
        Me.Label_Save = New System.Windows.Forms.Label()
        Me.Button_Save = New System.Windows.Forms.Button()
        Me.Label_StaEquations = New System.Windows.Forms.Label()
        Me.Label_StaRange = New System.Windows.Forms.Label()
        Me.Label_StaEquationValue = New System.Windows.Forms.Label()
        Me.Label_StaRangeValue = New System.Windows.Forms.Label()
        Me.Label_AlignNameValue = New System.Windows.Forms.Label()
        Me.Label_AlignName = New System.Windows.Forms.Label()
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox()
        Me.Label_ReportDesc = New System.Windows.Forms.Label()
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox()
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar()
        Me.Button_CreateReport = New System.Windows.Forms.Button()
        Me.Button_Help = New System.Windows.Forms.Button()
        Me.Button_Done = New System.Windows.Forms.Button()
        Me.GroupBox_Points.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox_Points
        '
        resources.ApplyResources(Me.GroupBox_Points, "GroupBox_Points")
        Me.GroupBox_Points.Controls.Add(Me.ListBox1)
        Me.GroupBox_Points.Controls.Add(Me.Combo_Pgroup)
        Me.GroupBox_Points.Controls.Add(Me.Label1)
        Me.GroupBox_Points.Controls.Add(Me.Button_Deselect)
        Me.GroupBox_Points.Controls.Add(Me.Button_Select)
        Me.GroupBox_Points.Controls.Add(Me.ListView_Aligns)
        Me.GroupBox_Points.Name = "GroupBox_Points"
        Me.GroupBox_Points.TabStop = False
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        resources.ApplyResources(Me.ListBox1, "ListBox1")
        Me.ListBox1.Name = "ListBox1"
        '
        'Combo_Pgroup
        '
        resources.ApplyResources(Me.Combo_Pgroup, "Combo_Pgroup")
        Me.Combo_Pgroup.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Pgroup.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Pgroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Pgroup.FormattingEnabled = True
        Me.Combo_Pgroup.Name = "Combo_Pgroup"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Button_Deselect
        '
        resources.ApplyResources(Me.Button_Deselect, "Button_Deselect")
        Me.Button_Deselect.Name = "Button_Deselect"
        Me.Button_Deselect.UseVisualStyleBackColor = True
        '
        'Button_Select
        '
        resources.ApplyResources(Me.Button_Select, "Button_Select")
        Me.Button_Select.Name = "Button_Select"
        Me.Button_Select.UseVisualStyleBackColor = True
        '
        'ListView_Aligns
        '
        resources.ApplyResources(Me.ListView_Aligns, "ListView_Aligns")
        Me.ListView_Aligns.CheckBoxes = True
        Me.ListView_Aligns.FullRowSelect = True
        Me.ListView_Aligns.GridLines = True
        Me.ListView_Aligns.HideSelection = False
        Me.ListView_Aligns.MultiSelect = False
        Me.ListView_Aligns.Name = "ListView_Aligns"
        Me.ListView_Aligns.UseCompatibleStateImageBehavior = False
        '
        'GroupBox_Settings
        '
        resources.ApplyResources(Me.GroupBox_Settings, "GroupBox_Settings")
        Me.GroupBox_Settings.Controls.Add(Me.Label_PgroupNameValue)
        Me.GroupBox_Settings.Controls.Add(Me.Label_PgroupName)
        Me.GroupBox_Settings.Controls.Add(Me.SplitContainer1)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StaEquations)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StaRange)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StaEquationValue)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StaRangeValue)
        Me.GroupBox_Settings.Controls.Add(Me.Label_AlignNameValue)
        Me.GroupBox_Settings.Controls.Add(Me.Label_AlignName)
        Me.GroupBox_Settings.Name = "GroupBox_Settings"
        Me.GroupBox_Settings.TabStop = False
        '
        'Label_PgroupNameValue
        '
        resources.ApplyResources(Me.Label_PgroupNameValue, "Label_PgroupNameValue")
        Me.Label_PgroupNameValue.Name = "Label_PgroupNameValue"
        '
        'Label_PgroupName
        '
        resources.ApplyResources(Me.Label_PgroupName, "Label_PgroupName")
        Me.Label_PgroupName.Name = "Label_PgroupName"
        '
        'SplitContainer1
        '
        resources.ApplyResources(Me.SplitContainer1, "SplitContainer1")
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Combo_Align)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label_SelectAlign)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button_SelectAlign)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.TextBox_SaveReport)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label_Save)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button_Save)
        '
        'Combo_Align
        '
        resources.ApplyResources(Me.Combo_Align, "Combo_Align")
        Me.Combo_Align.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Align.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Align.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Align.FormattingEnabled = True
        Me.Combo_Align.Name = "Combo_Align"
        '
        'Label_SelectAlign
        '
        resources.ApplyResources(Me.Label_SelectAlign, "Label_SelectAlign")
        Me.Label_SelectAlign.Name = "Label_SelectAlign"
        '
        'Button_SelectAlign
        '
        resources.ApplyResources(Me.Button_SelectAlign, "Button_SelectAlign")
        Me.Button_SelectAlign.Name = "Button_SelectAlign"
        Me.Button_SelectAlign.UseVisualStyleBackColor = True
        '
        'TextBox_SaveReport
        '
        resources.ApplyResources(Me.TextBox_SaveReport, "TextBox_SaveReport")
        Me.TextBox_SaveReport.Name = "TextBox_SaveReport"
        '
        'Label_Save
        '
        resources.ApplyResources(Me.Label_Save, "Label_Save")
        Me.Label_Save.Name = "Label_Save"
        '
        'Button_Save
        '
        resources.ApplyResources(Me.Button_Save, "Button_Save")
        Me.Button_Save.Name = "Button_Save"
        Me.Button_Save.UseVisualStyleBackColor = True
        '
        'Label_StaEquations
        '
        resources.ApplyResources(Me.Label_StaEquations, "Label_StaEquations")
        Me.Label_StaEquations.Name = "Label_StaEquations"
        '
        'Label_StaRange
        '
        resources.ApplyResources(Me.Label_StaRange, "Label_StaRange")
        Me.Label_StaRange.Name = "Label_StaRange"
        '
        'Label_StaEquationValue
        '
        resources.ApplyResources(Me.Label_StaEquationValue, "Label_StaEquationValue")
        Me.Label_StaEquationValue.Name = "Label_StaEquationValue"
        '
        'Label_StaRangeValue
        '
        resources.ApplyResources(Me.Label_StaRangeValue, "Label_StaRangeValue")
        Me.Label_StaRangeValue.Name = "Label_StaRangeValue"
        '
        'Label_AlignNameValue
        '
        resources.ApplyResources(Me.Label_AlignNameValue, "Label_AlignNameValue")
        Me.Label_AlignNameValue.Name = "Label_AlignNameValue"
        '
        'Label_AlignName
        '
        resources.ApplyResources(Me.Label_AlignName, "Label_AlignName")
        Me.Label_AlignName.Name = "Label_AlignName"
        '
        'GroupBox_ReportDesc
        '
        resources.ApplyResources(Me.GroupBox_ReportDesc, "GroupBox_ReportDesc")
        Me.GroupBox_ReportDesc.Controls.Add(Me.Label_ReportDesc)
        Me.GroupBox_ReportDesc.Name = "GroupBox_ReportDesc"
        Me.GroupBox_ReportDesc.TabStop = False
        '
        'Label_ReportDesc
        '
        resources.ApplyResources(Me.Label_ReportDesc, "Label_ReportDesc")
        Me.Label_ReportDesc.Name = "Label_ReportDesc"
        '
        'GroupBox_Progress
        '
        resources.ApplyResources(Me.GroupBox_Progress, "GroupBox_Progress")
        Me.GroupBox_Progress.Controls.Add(Me.ProgressBar_Creating)
        Me.GroupBox_Progress.Name = "GroupBox_Progress"
        Me.GroupBox_Progress.TabStop = False
        '
        'ProgressBar_Creating
        '
        resources.ApplyResources(Me.ProgressBar_Creating, "ProgressBar_Creating")
        Me.ProgressBar_Creating.Name = "ProgressBar_Creating"
        Me.ProgressBar_Creating.Step = 1
        '
        'Button_CreateReport
        '
        resources.ApplyResources(Me.Button_CreateReport, "Button_CreateReport")
        Me.Button_CreateReport.Name = "Button_CreateReport"
        Me.Button_CreateReport.UseVisualStyleBackColor = True
        '
        'Button_Help
        '
        resources.ApplyResources(Me.Button_Help, "Button_Help")
        Me.Button_Help.Name = "Button_Help"
        Me.Button_Help.UseVisualStyleBackColor = True
        '
        'Button_Done
        '
        resources.ApplyResources(Me.Button_Done, "Button_Done")
        Me.Button_Done.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Done.Name = "Button_Done"
        Me.Button_Done.UseVisualStyleBackColor = True
        '
        'ReportForm_PointsStaOffset
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Points)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_PointsStaOffset"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_Points.ResumeLayout(False)
        Me.GroupBox_Points.PerformLayout()
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox_Points As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_Aligns As System.Windows.Forms.ListView
    Friend WithEvents Button_Deselect As System.Windows.Forms.Button
    Friend WithEvents Button_Select As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents Label_SelectAlign As System.Windows.Forms.Label
    Friend WithEvents Button_SelectAlign As System.Windows.Forms.Button
    Friend WithEvents Combo_Align As System.Windows.Forms.ComboBox
    Friend WithEvents Label_StaEquations As System.Windows.Forms.Label
    Friend WithEvents Label_StaRange As System.Windows.Forms.Label
    Friend WithEvents Label_AlignName As System.Windows.Forms.Label
    Friend WithEvents Label_StaEquationValue As System.Windows.Forms.Label
    Friend WithEvents Label_StaRangeValue As System.Windows.Forms.Label
    Friend WithEvents Label_AlignNameValue As System.Windows.Forms.Label
    Friend WithEvents GroupBox_ReportDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label_ReportDesc As System.Windows.Forms.Label
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents Label_Save As System.Windows.Forms.Label
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents TextBox_SaveReport As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Combo_Pgroup As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label_PgroupNameValue As System.Windows.Forms.Label
    Friend WithEvents Label_PgroupName As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
End Class
