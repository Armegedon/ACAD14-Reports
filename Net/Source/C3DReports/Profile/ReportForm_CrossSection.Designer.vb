<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_CrossSection
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_CrossSection))
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_Done = New System.Windows.Forms.Button
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.CheckBox_XLS = New System.Windows.Forms.CheckBox
        Me.CheckBox_HTML = New System.Windows.Forms.CheckBox
        Me.CheckBox_ASCII = New System.Windows.Forms.CheckBox
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.TextBox_EndStation = New System.Windows.Forms.TextBox
        Me.TextBox_StartStation = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.Label_EndStation = New System.Windows.Forms.Label
        Me.Label_StartStation = New System.Windows.Forms.Label
        Me.GroupBox_SLG = New System.Windows.Forms.GroupBox
        Me.ListView_SLG = New System.Windows.Forms.ListView
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_SLG.SuspendLayout()
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
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_XLS)
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_HTML)
        Me.GroupBox_Settings.Controls.Add(Me.CheckBox_ASCII)
        Me.GroupBox_Settings.Controls.Add(Me.Button_Save)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_SaveReport)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_EndStation)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_StartStation)
        Me.GroupBox_Settings.Controls.Add(Me.Label_SaveTo)
        Me.GroupBox_Settings.Controls.Add(Me.Label_EndStation)
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
        Me.CheckBox_HTML.Name = "CheckBox_HTML"
        Me.CheckBox_HTML.UseVisualStyleBackColor = True
        '
        'CheckBox_ASCII
        '
        resources.ApplyResources(Me.CheckBox_ASCII, "CheckBox_ASCII")
        Me.CheckBox_ASCII.Name = "CheckBox_ASCII"
        Me.CheckBox_ASCII.UseVisualStyleBackColor = True
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
        'Label_StartStation
        '
        resources.ApplyResources(Me.Label_StartStation, "Label_StartStation")
        Me.Label_StartStation.Name = "Label_StartStation"
        '
        'GroupBox_SLG
        '
        resources.ApplyResources(Me.GroupBox_SLG, "GroupBox_SLG")
        Me.GroupBox_SLG.Controls.Add(Me.ListView_SLG)
        Me.GroupBox_SLG.Name = "GroupBox_SLG"
        Me.GroupBox_SLG.TabStop = False
        '
        'ListView_SLG
        '
        resources.ApplyResources(Me.ListView_SLG, "ListView_SLG")
        Me.ListView_SLG.CheckBoxes = True
        Me.ListView_SLG.FullRowSelect = True
        Me.ListView_SLG.GridLines = True
        Me.ListView_SLG.HideSelection = False
        Me.ListView_SLG.Name = "ListView_SLG"
        Me.ListView_SLG.UseCompatibleStateImageBehavior = False
        '
        'ReportForm_CrossSection
        '
        Me.CancelButton = Me.Button_Done
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.GroupBox_SLG)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_CrossSection"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_SLG.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
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
    Friend WithEvents Label_StartStation As System.Windows.Forms.Label
    Friend WithEvents CheckBox_XLS As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_HTML As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_ASCII As System.Windows.Forms.CheckBox
    Friend WithEvents GroupBox_SLG As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_SLG As System.Windows.Forms.ListView
End Class
