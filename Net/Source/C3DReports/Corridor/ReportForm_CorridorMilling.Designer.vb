<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_CorridorMilling
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_CorridorMilling))
        Me.GroupBox_ReportCom = New System.Windows.Forms.GroupBox
        Me.Combo_Align = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Combo_Corridor = New System.Windows.Forms.ComboBox
        Me.Button_SelectCorridor = New System.Windows.Forms.Button
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_Done = New System.Windows.Forms.Button
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.TextBox_EndStation = New System.Windows.Forms.TextBox
        Me.TextBox_StartStation = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.Label_EndStation = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox_MillingTolerance = New System.Windows.Forms.GroupBox
        Me.TextBox_Thickness = New System.Windows.Forms.TextBox
        Me.Label_Thickness = New System.Windows.Forms.Label
        Me.TextBox_Width = New System.Windows.Forms.TextBox
        Me.Label_Width = New System.Windows.Forms.Label
        Me.GroupBox_ReportCom.SuspendLayout()
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_MillingTolerance.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox_ReportCom
        '
        resources.ApplyResources(Me.GroupBox_ReportCom, "GroupBox_ReportCom")
        Me.GroupBox_ReportCom.Controls.Add(Me.Combo_Align)
        Me.GroupBox_ReportCom.Controls.Add(Me.Label3)
        Me.GroupBox_ReportCom.Controls.Add(Me.Label1)
        Me.GroupBox_ReportCom.Controls.Add(Me.Combo_Corridor)
        Me.GroupBox_ReportCom.Controls.Add(Me.Button_SelectCorridor)
        Me.GroupBox_ReportCom.Name = "GroupBox_ReportCom"
        Me.GroupBox_ReportCom.TabStop = False
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
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Combo_Corridor
        '
        Me.Combo_Corridor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Corridor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Corridor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Corridor.FormattingEnabled = True
        resources.ApplyResources(Me.Combo_Corridor, "Combo_Corridor")
        Me.Combo_Corridor.Name = "Combo_Corridor"
        '
        'Button_SelectCorridor
        '
        resources.ApplyResources(Me.Button_SelectCorridor, "Button_SelectCorridor")
        Me.Button_SelectCorridor.Name = "Button_SelectCorridor"
        Me.Button_SelectCorridor.UseVisualStyleBackColor = True
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
        'GroupBox_Settings
        '
        resources.ApplyResources(Me.GroupBox_Settings, "GroupBox_Settings")
        Me.GroupBox_Settings.Controls.Add(Me.Button_Save)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_SaveReport)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_EndStation)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_StartStation)
        Me.GroupBox_Settings.Controls.Add(Me.Label_SaveTo)
        Me.GroupBox_Settings.Controls.Add(Me.Label_EndStation)
        Me.GroupBox_Settings.Controls.Add(Me.Label2)
        Me.GroupBox_Settings.Name = "GroupBox_Settings"
        Me.GroupBox_Settings.TabStop = False
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
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'GroupBox_MillingTolerance
        '
        resources.ApplyResources(Me.GroupBox_MillingTolerance, "GroupBox_MillingTolerance")
        Me.GroupBox_MillingTolerance.Controls.Add(Me.TextBox_Thickness)
        Me.GroupBox_MillingTolerance.Controls.Add(Me.Label_Thickness)
        Me.GroupBox_MillingTolerance.Controls.Add(Me.TextBox_Width)
        Me.GroupBox_MillingTolerance.Controls.Add(Me.Label_Width)
        Me.GroupBox_MillingTolerance.Name = "GroupBox_MillingTolerance"
        Me.GroupBox_MillingTolerance.TabStop = False
        '
        'TextBox_Thickness
        '
        resources.ApplyResources(Me.TextBox_Thickness, "TextBox_Thickness")
        Me.TextBox_Thickness.Name = "TextBox_Thickness"
        '
        'Label_Thickness
        '
        resources.ApplyResources(Me.Label_Thickness, "Label_Thickness")
        Me.Label_Thickness.Name = "Label_Thickness"
        '
        'TextBox_Width
        '
        resources.ApplyResources(Me.TextBox_Width, "TextBox_Width")
        Me.TextBox_Width.Name = "TextBox_Width"
        '
        'Label_Width
        '
        resources.ApplyResources(Me.Label_Width, "Label_Width")
        Me.Label_Width.Name = "Label_Width"
        '
        'ReportForm_CorridorMilling
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.GroupBox_MillingTolerance)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.Controls.Add(Me.GroupBox_ReportCom)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_CorridorMilling"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportCom.ResumeLayout(False)
        Me.GroupBox_ReportCom.PerformLayout()
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_MillingTolerance.ResumeLayout(False)
        Me.GroupBox_MillingTolerance.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox_ReportCom As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox_ReportDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label_ReportDesc As System.Windows.Forms.Label
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents TextBox_SaveReport As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_EndStation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_StartStation As System.Windows.Forms.TextBox
    Friend WithEvents Label_SaveTo As System.Windows.Forms.Label
    Friend WithEvents Label_EndStation As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Combo_Align As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Combo_Corridor As System.Windows.Forms.ComboBox
    Friend WithEvents Button_SelectCorridor As System.Windows.Forms.Button
    Friend WithEvents GroupBox_MillingTolerance As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox_Thickness As System.Windows.Forms.TextBox
    Friend WithEvents Label_Thickness As System.Windows.Forms.Label
    Friend WithEvents TextBox_Width As System.Windows.Forms.TextBox
    Friend WithEvents Label_Width As System.Windows.Forms.Label
End Class
