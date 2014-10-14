<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_CorridorCrossSection
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_CorridorCrossSection))
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_Done = New System.Windows.Forms.Button
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.CheckBox_XLS = New System.Windows.Forms.CheckBox
        Me.CheckBox_HTML = New System.Windows.Forms.CheckBox
        Me.CheckBox_ASCII = New System.Windows.Forms.CheckBox
        Me.NumericUpDown_StationInc = New System.Windows.Forms.NumericUpDown
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.TextBox_EndStation = New System.Windows.Forms.TextBox
        Me.TextBox_StartStation = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.Label_EndStation = New System.Windows.Forms.Label
        Me.Label_StationInc = New System.Windows.Forms.Label
        Me.Label_StartStation = New System.Windows.Forms.Label
        Me.GroupBox_FeatureLines = New System.Windows.Forms.GroupBox
        Me.Button_UncheckAll = New System.Windows.Forms.Button
        Me.Button_CheckAll = New System.Windows.Forms.Button
        Me.ListView_FeatureLines = New System.Windows.Forms.ListView
        Me.GroupBox_Components = New System.Windows.Forms.GroupBox
        Me.Combo_SLG = New System.Windows.Forms.ComboBox
        Me.radioSurfaces = New System.Windows.Forms.RadioButton
        Me.radioLinks = New System.Windows.Forms.RadioButton
        Me.radioPoints = New System.Windows.Forms.RadioButton
        Me.Label_SelectSLG = New System.Windows.Forms.Label
        Me.Combo_Align = New System.Windows.Forms.ComboBox
        Me.Combo_Corridor = New System.Windows.Forms.ComboBox
        Me.Label_Alignment = New System.Windows.Forms.Label
        Me.Label_SelectCorridor = New System.Windows.Forms.Label
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        CType(Me.NumericUpDown_StationInc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_FeatureLines.SuspendLayout()
        Me.GroupBox_Components.SuspendLayout()
        Me.SuspendLayout()
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
        'GroupBox_Settings
        '
        resources.ApplyResources(Me.GroupBox_Settings, "GroupBox_Settings")
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_XLS)
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_HTML)
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_ASCII)
        Me.GroupBox_Settings.Controls.Add(Me.NumericUpDown_StationInc)
        Me.GroupBox_Settings.Controls.Add(Me.Button_Save)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_SaveReport)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_EndStation)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_StartStation)
        Me.GroupBox_Settings.Controls.Add(Me.Label_SaveTo)
        Me.GroupBox_Settings.Controls.Add(Me.Label_EndStation)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StationInc)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StartStation)
        Me.GroupBox_Settings.Name = "GroupBox_Settings"
        Me.GroupBox_Settings.TabStop = False
        '
        'CheckBox_XLS
        '
        resources.ApplyResources(Me.CheckBox_XLS, "CheckBox_XLS")
        Me.CheckBox_XLS.Name = "CheckBox_XLS"
        Me.CheckBox_XLS.UseVisualStyleBackColor = True
        '
        'CheckBox_HTML
        '
        resources.ApplyResources(Me.CheckBox_HTML, "CheckBox_HTML")
        Me.CheckBox_HTML.Checked = True
        Me.CheckBox_HTML.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_HTML.Name = "CheckBox_HTML"
        Me.CheckBox_HTML.UseVisualStyleBackColor = True
        '
        'CheckBox_ASCII
        '
        resources.ApplyResources(Me.CheckBox_ASCII, "CheckBox_ASCII")
        Me.CheckBox_ASCII.Name = "CheckBox_ASCII"
        Me.CheckBox_ASCII.UseVisualStyleBackColor = True
        '
        'NumericUpDown_StationInc
        '
        resources.ApplyResources(Me.NumericUpDown_StationInc, "NumericUpDown_StationInc")
        Me.NumericUpDown_StationInc.Maximum = New Decimal(New Integer() {1000000, 0, 0, 0})
        Me.NumericUpDown_StationInc.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.NumericUpDown_StationInc.Name = "NumericUpDown_StationInc"
        Me.NumericUpDown_StationInc.Value = New Decimal(New Integer() {25, 0, 0, 0})
        '
        'Button_Save
        '
        resources.ApplyResources(Me.Button_Save, "Button_Save")
        Me.Button_Save.Name = "Button_Save"
        Me.Button_Save.UseVisualStyleBackColor = True
        '
        'TextBox_SaveReport
        '
        resources.ApplyResources(Me.TextBox_SaveReport, "TextBox_SaveReport")
        Me.TextBox_SaveReport.Name = "TextBox_SaveReport"
        '
        'TextBox_EndStation
        '
        resources.ApplyResources(Me.TextBox_EndStation, "TextBox_EndStation")
        Me.TextBox_EndStation.Name = "TextBox_EndStation"
        '
        'TextBox_StartStation
        '
        resources.ApplyResources(Me.TextBox_StartStation, "TextBox_StartStation")
        Me.TextBox_StartStation.Name = "TextBox_StartStation"
        '
        'Label_SaveTo
        '
        resources.ApplyResources(Me.Label_SaveTo, "Label_SaveTo")
        Me.Label_SaveTo.Name = "Label_SaveTo"
        '
        'Label_EndStation
        '
        resources.ApplyResources(Me.Label_EndStation, "Label_EndStation")
        Me.Label_EndStation.Name = "Label_EndStation"
        '
        'Label_StationInc
        '
        resources.ApplyResources(Me.Label_StationInc, "Label_StationInc")
        Me.Label_StationInc.Name = "Label_StationInc"
        '
        'Label_StartStation
        '
        resources.ApplyResources(Me.Label_StartStation, "Label_StartStation")
        Me.Label_StartStation.Name = "Label_StartStation"
        '
        'GroupBox_FeatureLines
        '
        resources.ApplyResources(Me.GroupBox_FeatureLines, "GroupBox_FeatureLines")
        Me.GroupBox_FeatureLines.Controls.Add(Me.Button_UncheckAll)
        Me.GroupBox_FeatureLines.Controls.Add(Me.Button_CheckAll)
        Me.GroupBox_FeatureLines.Controls.Add(Me.ListView_FeatureLines)
        Me.GroupBox_FeatureLines.Name = "GroupBox_FeatureLines"
        Me.GroupBox_FeatureLines.TabStop = False
        '
        'Button_UncheckAll
        '
        resources.ApplyResources(Me.Button_UncheckAll, "Button_UncheckAll")
        Me.Button_UncheckAll.Name = "Button_UncheckAll"
        Me.Button_UncheckAll.UseVisualStyleBackColor = True
        '
        'Button_CheckAll
        '
        resources.ApplyResources(Me.Button_CheckAll, "Button_CheckAll")
        Me.Button_CheckAll.Name = "Button_CheckAll"
        Me.Button_CheckAll.UseVisualStyleBackColor = True
        '
        'ListView_FeatureLines
        '
        resources.ApplyResources(Me.ListView_FeatureLines, "ListView_FeatureLines")
        Me.ListView_FeatureLines.CheckBoxes = True
        Me.ListView_FeatureLines.FullRowSelect = True
        Me.ListView_FeatureLines.GridLines = True
        Me.ListView_FeatureLines.Name = "ListView_FeatureLines"
        Me.ListView_FeatureLines.UseCompatibleStateImageBehavior = False
        '
        'GroupBox_Components
        '
        resources.ApplyResources(Me.GroupBox_Components, "GroupBox_Components")
        Me.GroupBox_Components.Controls.Add(Me.Combo_SLG)
        Me.GroupBox_Components.Controls.Add(Me.radioSurfaces)
        Me.GroupBox_Components.Controls.Add(Me.radioLinks)
        Me.GroupBox_Components.Controls.Add(Me.radioPoints)
        Me.GroupBox_Components.Controls.Add(Me.Label_SelectSLG)
        Me.GroupBox_Components.Controls.Add(Me.Combo_Align)
        Me.GroupBox_Components.Controls.Add(Me.Combo_Corridor)
        Me.GroupBox_Components.Controls.Add(Me.Label_Alignment)
        Me.GroupBox_Components.Controls.Add(Me.Label_SelectCorridor)
        Me.GroupBox_Components.Name = "GroupBox_Components"
        Me.GroupBox_Components.TabStop = False
        '
        'Combo_SLG
        '
        Me.Combo_SLG.FormattingEnabled = True
        resources.ApplyResources(Me.Combo_SLG, "Combo_SLG")
        Me.Combo_SLG.Name = "Combo_SLG"
        '
        'radioSurfaces
        '
        resources.ApplyResources(Me.radioSurfaces, "radioSurfaces")
        Me.radioSurfaces.Name = "radioSurfaces"
        Me.radioSurfaces.TabStop = True
        Me.radioSurfaces.UseVisualStyleBackColor = True
        '
        'radioLinks
        '
        resources.ApplyResources(Me.radioLinks, "radioLinks")
        Me.radioLinks.Name = "radioLinks"
        Me.radioLinks.TabStop = True
        Me.radioLinks.UseVisualStyleBackColor = True
        '
        'radioPoints
        '
        resources.ApplyResources(Me.radioPoints, "radioPoints")
        Me.radioPoints.Name = "radioPoints"
        Me.radioPoints.TabStop = True
        Me.radioPoints.UseVisualStyleBackColor = True
        '
        'Label_SelectSLG
        '
        resources.ApplyResources(Me.Label_SelectSLG, "Label_SelectSLG")
        Me.Label_SelectSLG.Name = "Label_SelectSLG"
        '
        'Combo_Align
        '
        Me.Combo_Align.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Align.FormattingEnabled = True
        resources.ApplyResources(Me.Combo_Align, "Combo_Align")
        Me.Combo_Align.Name = "Combo_Align"
        '
        'Combo_Corridor
        '
        Me.Combo_Corridor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Corridor.FormattingEnabled = True
        resources.ApplyResources(Me.Combo_Corridor, "Combo_Corridor")
        Me.Combo_Corridor.Name = "Combo_Corridor"
        '
        'Label_Alignment
        '
        resources.ApplyResources(Me.Label_Alignment, "Label_Alignment")
        Me.Label_Alignment.Name = "Label_Alignment"
        '
        'Label_SelectCorridor
        '
        resources.ApplyResources(Me.Label_SelectCorridor, "Label_SelectCorridor")
        Me.Label_SelectCorridor.Name = "Label_SelectCorridor"
        '
        'ReportForm_CorridorCrossSection
        '
        Me.CancelButton = Me.Button_Done
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.GroupBox_Components)
        Me.Controls.Add(Me.GroupBox_FeatureLines)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_CorridorCrossSection"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        CType(Me.NumericUpDown_StationInc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_FeatureLines.ResumeLayout(False)
        Me.GroupBox_Components.ResumeLayout(False)
        Me.GroupBox_Components.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox_Components As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox_ReportDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label_ReportDesc As System.Windows.Forms.Label
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown_StationInc As System.Windows.Forms.NumericUpDown
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents TextBox_SaveReport As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_EndStation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_StartStation As System.Windows.Forms.TextBox
    Friend WithEvents Label_SaveTo As System.Windows.Forms.Label
    Friend WithEvents Label_EndStation As System.Windows.Forms.Label
    Friend WithEvents Label_StationInc As System.Windows.Forms.Label
    Friend WithEvents Label_StartStation As System.Windows.Forms.Label
    Friend WithEvents GroupBox_FeatureLines As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_FeatureLines As System.Windows.Forms.ListView
    Friend WithEvents Combo_Align As System.Windows.Forms.ComboBox
    Friend WithEvents Combo_Corridor As System.Windows.Forms.ComboBox
    Friend WithEvents Label_Alignment As System.Windows.Forms.Label
    Friend WithEvents Label_SelectCorridor As System.Windows.Forms.Label
    Friend WithEvents Button_UncheckAll As System.Windows.Forms.Button
    Friend WithEvents Button_CheckAll As System.Windows.Forms.Button
    Friend WithEvents CheckBox_XLS As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_HTML As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_ASCII As System.Windows.Forms.CheckBox
    Friend WithEvents Combo_SLG As System.Windows.Forms.ComboBox
    Friend WithEvents Label_SelectSLG As System.Windows.Forms.Label
    Friend WithEvents radioPoints As System.Windows.Forms.RadioButton
    Friend WithEvents radioSurfaces As System.Windows.Forms.RadioButton
    Friend WithEvents radioLinks As System.Windows.Forms.RadioButton
End Class
