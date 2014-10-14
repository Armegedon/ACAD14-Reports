<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_CorridorSlopeStake
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_CorridorSlopeStake))
        Me.GroupBox_ReportCom = New System.Windows.Forms.GroupBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Combo_Material_List = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Combo_Align = New System.Windows.Forms.ComboBox
        Me.Combo_Corridor = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button_SelectCorridor = New System.Windows.Forms.Button
        Me.Combo_Link = New System.Windows.Forms.ComboBox
        Me.Combo_LineGroup = New System.Windows.Forms.ComboBox
        Me.Button_AddLink = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
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
        Me.GroupBox_Corridors = New System.Windows.Forms.GroupBox
        Me.Button_Delete = New System.Windows.Forms.Button
        Me.ListView_Corridors = New System.Windows.Forms.ListView
        Me.GroupBox_ROW_Settings = New System.Windows.Forms.GroupBox
        Me.Label_ROW_PointCode = New System.Windows.Forms.Label
        Me.ComboBox_ROW_PointCode = New System.Windows.Forms.ComboBox
        Me.CheckBox_DisplayROW = New System.Windows.Forms.CheckBox
        Me.CheckBox_AllowGraphics = New System.Windows.Forms.CheckBox
        Me.CheckBox_DisplayCutFill = New System.Windows.Forms.CheckBox
        Me.GroupBox_ReportCom.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_Corridors.SuspendLayout()
        Me.GroupBox_ROW_Settings.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox_ReportCom
        '
        resources.ApplyResources(Me.GroupBox_ReportCom, "GroupBox_ReportCom")
        Me.GroupBox_ReportCom.Controls.Add(Me.SplitContainer1)
        Me.GroupBox_ReportCom.Name = "GroupBox_ReportCom"
        Me.GroupBox_ReportCom.TabStop = False
        '
        'SplitContainer1
        '
        resources.ApplyResources(Me.SplitContainer1, "SplitContainer1")
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Combo_Material_List)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label6)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Combo_Align)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Combo_Corridor)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer1.Panel1.Controls.Add(Me.Button_SelectCorridor)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Combo_Link)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Combo_LineGroup)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Button_AddLink)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label5)
        Me.SplitContainer1.Panel2.Controls.Add(Me.Label4)
        '
        'Combo_Material_List
        '
        resources.ApplyResources(Me.Combo_Material_List, "Combo_Material_List")
        Me.Combo_Material_List.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Material_List.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Material_List.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Material_List.FormattingEnabled = True
        Me.Combo_Material_List.Name = "Combo_Material_List"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
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
        'Combo_Corridor
        '
        resources.ApplyResources(Me.Combo_Corridor, "Combo_Corridor")
        Me.Combo_Corridor.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Corridor.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Corridor.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Corridor.FormattingEnabled = True
        Me.Combo_Corridor.Name = "Combo_Corridor"
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
        'Button_SelectCorridor
        '
        resources.ApplyResources(Me.Button_SelectCorridor, "Button_SelectCorridor")
        Me.Button_SelectCorridor.Name = "Button_SelectCorridor"
        Me.Button_SelectCorridor.UseVisualStyleBackColor = True
        '
        'Combo_Link
        '
        resources.ApplyResources(Me.Combo_Link, "Combo_Link")
        Me.Combo_Link.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Link.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Link.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Link.FormattingEnabled = True
        Me.Combo_Link.Name = "Combo_Link"
        '
        'Combo_LineGroup
        '
        resources.ApplyResources(Me.Combo_LineGroup, "Combo_LineGroup")
        Me.Combo_LineGroup.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_LineGroup.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_LineGroup.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_LineGroup.FormattingEnabled = True
        Me.Combo_LineGroup.Name = "Combo_LineGroup"
        '
        'Button_AddLink
        '
        resources.ApplyResources(Me.Button_AddLink, "Button_AddLink")
        Me.Button_AddLink.Name = "Button_AddLink"
        Me.Button_AddLink.UseVisualStyleBackColor = True
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
        '
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
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
        'GroupBox_Corridors
        '
        resources.ApplyResources(Me.GroupBox_Corridors, "GroupBox_Corridors")
        Me.GroupBox_Corridors.Controls.Add(Me.Button_Delete)
        Me.GroupBox_Corridors.Controls.Add(Me.ListView_Corridors)
        Me.GroupBox_Corridors.MinimumSize = New System.Drawing.Size(468, 99)
        Me.GroupBox_Corridors.Name = "GroupBox_Corridors"
        Me.GroupBox_Corridors.TabStop = False
        '
        'Button_Delete
        '
        resources.ApplyResources(Me.Button_Delete, "Button_Delete")
        Me.Button_Delete.Name = "Button_Delete"
        Me.Button_Delete.UseVisualStyleBackColor = True
        '
        'ListView_Corridors
        '
        resources.ApplyResources(Me.ListView_Corridors, "ListView_Corridors")
        Me.ListView_Corridors.MinimumSize = New System.Drawing.Size(426, 73)
        Me.ListView_Corridors.Name = "ListView_Corridors"
        Me.ListView_Corridors.UseCompatibleStateImageBehavior = False
        '
        'GroupBox_ROW_Settings
        '
        resources.ApplyResources(Me.GroupBox_ROW_Settings, "GroupBox_ROW_Settings")
        Me.GroupBox_ROW_Settings.Controls.Add(Me.Label_ROW_PointCode)
        Me.GroupBox_ROW_Settings.Controls.Add(Me.ComboBox_ROW_PointCode)
        Me.GroupBox_ROW_Settings.Controls.Add(Me.CheckBox_DisplayROW)
        Me.GroupBox_ROW_Settings.Name = "GroupBox_ROW_Settings"
        Me.GroupBox_ROW_Settings.TabStop = False
        '
        'Label_ROW_PointCode
        '
        resources.ApplyResources(Me.Label_ROW_PointCode, "Label_ROW_PointCode")
        Me.Label_ROW_PointCode.Name = "Label_ROW_PointCode"
        '
        'ComboBox_ROW_PointCode
        '
        resources.ApplyResources(Me.ComboBox_ROW_PointCode, "ComboBox_ROW_PointCode")
        Me.ComboBox_ROW_PointCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_ROW_PointCode.FormattingEnabled = True
        Me.ComboBox_ROW_PointCode.Name = "ComboBox_ROW_PointCode"
        '
        'CheckBox_DisplayROW
        '
        resources.ApplyResources(Me.CheckBox_DisplayROW, "CheckBox_DisplayROW")
        Me.CheckBox_DisplayROW.Name = "CheckBox_DisplayROW"
        Me.CheckBox_DisplayROW.UseVisualStyleBackColor = True
        '
        'CheckBox_AllowGraphics
        '
        resources.ApplyResources(Me.CheckBox_AllowGraphics, "CheckBox_AllowGraphics")
        Me.CheckBox_AllowGraphics.Checked = True
        Me.CheckBox_AllowGraphics.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_AllowGraphics.Name = "CheckBox_AllowGraphics"
        Me.CheckBox_AllowGraphics.UseVisualStyleBackColor = True
        '
        'CheckBox_DisplayCutFill
        '
        resources.ApplyResources(Me.CheckBox_DisplayCutFill, "CheckBox_DisplayCutFill")
        Me.CheckBox_DisplayCutFill.Checked = True
        Me.CheckBox_DisplayCutFill.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox_DisplayCutFill.Name = "CheckBox_DisplayCutFill"
        Me.CheckBox_DisplayCutFill.UseVisualStyleBackColor = True
        '
        'ReportForm_CorridorSlopeStake
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.CheckBox_DisplayCutFill)
        Me.Controls.Add(Me.CheckBox_AllowGraphics)
        Me.Controls.Add(Me.GroupBox_ROW_Settings)
        Me.Controls.Add(Me.GroupBox_Corridors)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.Controls.Add(Me.GroupBox_ReportCom)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_CorridorSlopeStake"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportCom.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.PerformLayout()
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.Panel2.PerformLayout()
        Me.SplitContainer1.ResumeLayout(False)
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_Corridors.ResumeLayout(False)
        Me.GroupBox_ROW_Settings.ResumeLayout(False)
        Me.GroupBox_ROW_Settings.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox_ReportCom As System.Windows.Forms.GroupBox
    Friend WithEvents Combo_Link As System.Windows.Forms.ComboBox
    Friend WithEvents Combo_LineGroup As System.Windows.Forms.ComboBox
    Friend WithEvents Combo_Align As System.Windows.Forms.ComboBox
    Friend WithEvents Combo_Corridor As System.Windows.Forms.ComboBox
    Friend WithEvents Button_AddLink As System.Windows.Forms.Button
    Friend WithEvents Button_SelectCorridor As System.Windows.Forms.Button
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
    Friend WithEvents GroupBox_Corridors As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_Corridors As System.Windows.Forms.ListView
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents Button_Delete As System.Windows.Forms.Button
    Friend WithEvents Combo_Material_List As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox_ROW_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox_AllowGraphics As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox_DisplayROW As System.Windows.Forms.CheckBox
    Friend WithEvents Label_ROW_PointCode As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ROW_PointCode As System.Windows.Forms.ComboBox
    Friend WithEvents CheckBox_DisplayCutFill As System.Windows.Forms.CheckBox
End Class
