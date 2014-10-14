<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_AlignStakeout
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_AlignStakeout))
        Me.GroupBox_StakeoutOption = New System.Windows.Forms.GroupBox
        Me.Edit_BacksightPoint = New System.Windows.Forms.TextBox
        Me.Edit_PointOccupied = New System.Windows.Forms.TextBox
        Me.Radio_Direction = New System.Windows.Forms.RadioButton
        Me.Button_Backsight = New System.Windows.Forms.Button
        Me.Radio_DeflectMinus = New System.Windows.Forms.RadioButton
        Me.Radio_TurnedPlus = New System.Windows.Forms.RadioButton
        Me.Radio_DeflectPlus = New System.Windows.Forms.RadioButton
        Me.Radio_TurnedMinus = New System.Windows.Forms.RadioButton
        Me.Label_PtOccupied = New System.Windows.Forms.Label
        Me.Label_AngleType = New System.Windows.Forms.Label
        Me.Button_Occupied = New System.Windows.Forms.Button
        Me.Label_BacksightPt = New System.Windows.Forms.Label
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.GroupBox_Alignments = New System.Windows.Forms.GroupBox
        Me.ListView_Alignments = New System.Windows.Forms.ListView
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.NumericUpDown_Offset = New System.Windows.Forms.NumericUpDown
        Me.Label1 = New System.Windows.Forms.Label
        Me.NumericUpDown_StationInc = New System.Windows.Forms.NumericUpDown
        Me.Label_StationInc = New System.Windows.Forms.Label
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.TextBox_EndStation = New System.Windows.Forms.TextBox
        Me.TextBox_StartStation = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.Label_EndStation = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_Done = New System.Windows.Forms.Button
        Me.GroupBox_StakeoutOption.SuspendLayout()
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Alignments.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        CType(Me.NumericUpDown_Offset, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown_StationInc, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox_Progress.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox_StakeoutOption
        '
        resources.ApplyResources(Me.GroupBox_StakeoutOption, "GroupBox_StakeoutOption")
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Edit_BacksightPoint)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Edit_PointOccupied)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Radio_Direction)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Button_Backsight)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Radio_DeflectMinus)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Radio_TurnedPlus)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Radio_DeflectPlus)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Radio_TurnedMinus)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Label_PtOccupied)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Label_AngleType)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Button_Occupied)
        Me.GroupBox_StakeoutOption.Controls.Add(Me.Label_BacksightPt)
        Me.GroupBox_StakeoutOption.Name = "GroupBox_StakeoutOption"
        Me.GroupBox_StakeoutOption.TabStop = False
        '
        'Edit_BacksightPoint
        '
        resources.ApplyResources(Me.Edit_BacksightPoint, "Edit_BacksightPoint")
        Me.Edit_BacksightPoint.Name = "Edit_BacksightPoint"
        '
        'Edit_PointOccupied
        '
        resources.ApplyResources(Me.Edit_PointOccupied, "Edit_PointOccupied")
        Me.Edit_PointOccupied.Name = "Edit_PointOccupied"
        '
        'Radio_Direction
        '
        resources.ApplyResources(Me.Radio_Direction, "Radio_Direction")
        Me.Radio_Direction.Name = "Radio_Direction"
        Me.Radio_Direction.TabStop = True
        Me.Radio_Direction.UseVisualStyleBackColor = True
        '
        'Button_Backsight
        '
        resources.ApplyResources(Me.Button_Backsight, "Button_Backsight")
        Me.Button_Backsight.Name = "Button_Backsight"
        Me.Button_Backsight.UseVisualStyleBackColor = True
        '
        'Radio_DeflectMinus
        '
        resources.ApplyResources(Me.Radio_DeflectMinus, "Radio_DeflectMinus")
        Me.Radio_DeflectMinus.Name = "Radio_DeflectMinus"
        Me.Radio_DeflectMinus.TabStop = True
        Me.Radio_DeflectMinus.UseVisualStyleBackColor = True
        '
        'Radio_TurnedPlus
        '
        resources.ApplyResources(Me.Radio_TurnedPlus, "Radio_TurnedPlus")
        Me.Radio_TurnedPlus.Name = "Radio_TurnedPlus"
        Me.Radio_TurnedPlus.TabStop = True
        Me.Radio_TurnedPlus.UseVisualStyleBackColor = True
        '
        'Radio_DeflectPlus
        '
        resources.ApplyResources(Me.Radio_DeflectPlus, "Radio_DeflectPlus")
        Me.Radio_DeflectPlus.Name = "Radio_DeflectPlus"
        Me.Radio_DeflectPlus.TabStop = True
        Me.Radio_DeflectPlus.UseVisualStyleBackColor = True
        '
        'Radio_TurnedMinus
        '
        resources.ApplyResources(Me.Radio_TurnedMinus, "Radio_TurnedMinus")
        Me.Radio_TurnedMinus.Name = "Radio_TurnedMinus"
        Me.Radio_TurnedMinus.TabStop = True
        Me.Radio_TurnedMinus.UseVisualStyleBackColor = True
        '
        'Label_PtOccupied
        '
        resources.ApplyResources(Me.Label_PtOccupied, "Label_PtOccupied")
        Me.Label_PtOccupied.Name = "Label_PtOccupied"
        '
        'Label_AngleType
        '
        resources.ApplyResources(Me.Label_AngleType, "Label_AngleType")
        Me.Label_AngleType.Name = "Label_AngleType"
        '
        'Button_Occupied
        '
        resources.ApplyResources(Me.Button_Occupied, "Button_Occupied")
        Me.Button_Occupied.Name = "Button_Occupied"
        Me.Button_Occupied.UseVisualStyleBackColor = True
        '
        'Label_BacksightPt
        '
        resources.ApplyResources(Me.Label_BacksightPt, "Label_BacksightPt")
        Me.Label_BacksightPt.Name = "Label_BacksightPt"
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
        Me.GroupBox_Settings.Controls.Add(Me.NumericUpDown_Offset)
        Me.GroupBox_Settings.Controls.Add(Me.Label1)
        Me.GroupBox_Settings.Controls.Add(Me.NumericUpDown_StationInc)
        Me.GroupBox_Settings.Controls.Add(Me.Label_StationInc)
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
        'NumericUpDown_Offset
        '
        Me.NumericUpDown_Offset.DecimalPlaces = 2
        resources.ApplyResources(Me.NumericUpDown_Offset, "NumericUpDown_Offset")
        Me.NumericUpDown_Offset.Name = "NumericUpDown_Offset"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'NumericUpDown_StationInc
        '
        Me.NumericUpDown_StationInc.DecimalPlaces = 2
        resources.ApplyResources(Me.NumericUpDown_StationInc, "NumericUpDown_StationInc")
        Me.NumericUpDown_StationInc.Name = "NumericUpDown_StationInc"
        '
        'Label_StationInc
        '
        resources.ApplyResources(Me.Label_StationInc, "Label_StationInc")
        Me.Label_StationInc.Name = "Label_StationInc"
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
        'ReportForm_AlignStakeout
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Alignments)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.Controls.Add(Me.GroupBox_StakeoutOption)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_AlignStakeout"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_StakeoutOption.ResumeLayout(False)
        Me.GroupBox_StakeoutOption.PerformLayout()
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Alignments.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        CType(Me.NumericUpDown_Offset, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown_StationInc, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox_StakeoutOption As System.Windows.Forms.GroupBox
    Friend WithEvents Edit_PointOccupied As System.Windows.Forms.TextBox
    Friend WithEvents Radio_TurnedPlus As System.Windows.Forms.RadioButton
    Friend WithEvents Label_AngleType As System.Windows.Forms.Label
    Friend WithEvents Label_BacksightPt As System.Windows.Forms.Label
    Friend WithEvents Label_PtOccupied As System.Windows.Forms.Label
    Friend WithEvents Button_Occupied As System.Windows.Forms.Button
    Friend WithEvents Edit_BacksightPoint As System.Windows.Forms.TextBox
    Friend WithEvents Button_Backsight As System.Windows.Forms.Button
    Friend WithEvents Radio_Direction As System.Windows.Forms.RadioButton
    Friend WithEvents Radio_DeflectMinus As System.Windows.Forms.RadioButton
    Friend WithEvents Radio_DeflectPlus As System.Windows.Forms.RadioButton
    Friend WithEvents Radio_TurnedMinus As System.Windows.Forms.RadioButton
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
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents NumericUpDown_StationInc As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label_StationInc As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_Offset As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
