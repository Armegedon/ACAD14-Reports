<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_AlignSta
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_AlignSta))
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.GroupBox_Alignments = New System.Windows.Forms.GroupBox
        Me.ListView_Alignments = New System.Windows.Forms.ListView
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.TextBox_EndStation = New System.Windows.Forms.TextBox
        Me.TextBox_StartStation = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.Label_EndStation = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button_Done = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Alignments.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
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
        'GroupBox_Alignments
        '
        resources.ApplyResources(Me.GroupBox_Alignments, "GroupBox_Alignments")
        Me.GroupBox_Alignments.Controls.Add(Me.ListView_Alignments)
        Me.GroupBox_Alignments.Name = "GroupBox_Alignments"
        Me.GroupBox_Alignments.TabStop = False
        '
        'ListView_Alignments
        '
        resources.ApplyResources(Me.ListView_Alignments, "ListView_Alignments")
        Me.ListView_Alignments.Name = "ListView_Alignments"
        Me.ListView_Alignments.UseCompatibleStateImageBehavior = False
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
        'Button_Done
        '
        resources.ApplyResources(Me.Button_Done, "Button_Done")
        Me.Button_Done.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Done.Name = "Button_Done"
        Me.Button_Done.UseVisualStyleBackColor = True
        '
        'Button_Help
        '
        resources.ApplyResources(Me.Button_Help, "Button_Help")
        Me.Button_Help.Name = "Button_Help"
        Me.Button_Help.UseVisualStyleBackColor = True
        '
        'Button_CreateReport
        '
        resources.ApplyResources(Me.Button_CreateReport, "Button_CreateReport")
        Me.Button_CreateReport.Name = "Button_CreateReport"
        Me.Button_CreateReport.UseVisualStyleBackColor = True
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
        'ReportForm_AlignSta
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Alignments)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_AlignSta"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Alignments.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox_ReportDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label_ReportDesc As System.Windows.Forms.Label
    Friend WithEvents GroupBox_Alignments As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_Alignments As System.Windows.Forms.ListView
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents TextBox_SaveReport As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_EndStation As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_StartStation As System.Windows.Forms.TextBox
    Friend WithEvents Label_SaveTo As System.Windows.Forms.Label
    Friend WithEvents Label_EndStation As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
End Class
