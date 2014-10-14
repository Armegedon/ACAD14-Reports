<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_ParcelMapCheck
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_ParcelMapCheck))
        Me.Group_Objects = New System.Windows.Forms.GroupBox
        Me.ListView_Figures = New System.Windows.Forms.ListView
        Me.ListView_Parcels = New System.Windows.Forms.ListView
        Me.Button_Deselect = New System.Windows.Forms.Button
        Me.Button_Select = New System.Windows.Forms.Button
        Me.Radio_Figures = New System.Windows.Forms.RadioButton
        Me.Radio_Parcels = New System.Windows.Forms.RadioButton
        Me.Group_Analysis = New System.Windows.Forms.GroupBox
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.Text_PobXCoord = New System.Windows.Forms.TextBox
        Me.Label_X = New System.Windows.Forms.Label
        Me.Text_PobYCoord = New System.Windows.Forms.TextBox
        Me.Label_Y = New System.Windows.Forms.Label
        Me.Button_Pt = New System.Windows.Forms.Button
        Me.Check_AcrossChord = New System.Windows.Forms.CheckBox
        Me.Check_CounterClock = New System.Windows.Forms.CheckBox
        Me.Label_PtBegin = New System.Windows.Forms.Label
        Me.Group_Settings = New System.Windows.Forms.GroupBox
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.Label_Format = New System.Windows.Forms.Label
        Me.Label_DirecFormat = New System.Windows.Forms.Label
        Me.Label_Value = New System.Windows.Forms.Label
        Me.Label_PrecPrec = New System.Windows.Forms.Label
        Me.Label_Precision = New System.Windows.Forms.Label
        Me.Label_AreaPrec = New System.Windows.Forms.Label
        Me.Label_Direction = New System.Windows.Forms.Label
        Me.Label_PrecisionPerimeter = New System.Windows.Forms.Label
        Me.Label_DirePrec = New System.Windows.Forms.Label
        Me.Label_Area = New System.Windows.Forms.Label
        Me.Label_DistPrec = New System.Windows.Forms.Label
        Me.Label_CloPrec = New System.Windows.Forms.Label
        Me.Label_Distance = New System.Windows.Forms.Label
        Me.Label_Closure = New System.Windows.Forms.Label
        Me.GroupBox_ReportDesc = New System.Windows.Forms.GroupBox
        Me.Label_ReportDesc = New System.Windows.Forms.Label
        Me.GroupBox_Settings = New System.Windows.Forms.GroupBox
        Me.Button_Save = New System.Windows.Forms.Button
        Me.TextBox_SaveReport = New System.Windows.Forms.TextBox
        Me.Label_SaveTo = New System.Windows.Forms.Label
        Me.GroupBox_Progress = New System.Windows.Forms.GroupBox
        Me.ProgressBar_Creating = New System.Windows.Forms.ProgressBar
        Me.Button_CreateReport = New System.Windows.Forms.Button
        Me.Button_Help = New System.Windows.Forms.Button
        Me.Button_Done = New System.Windows.Forms.Button
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Group_Objects.SuspendLayout()
        Me.Group_Analysis.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.Group_Settings.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Group_Objects
        '
        resources.ApplyResources(Me.Group_Objects, "Group_Objects")
        Me.Group_Objects.Controls.Add(Me.ListView_Figures)
        Me.Group_Objects.Controls.Add(Me.ListView_Parcels)
        Me.Group_Objects.Controls.Add(Me.Button_Deselect)
        Me.Group_Objects.Controls.Add(Me.Button_Select)
        Me.Group_Objects.Controls.Add(Me.Radio_Figures)
        Me.Group_Objects.Controls.Add(Me.Radio_Parcels)
        Me.Group_Objects.Name = "Group_Objects"
        Me.Group_Objects.TabStop = False
        '
        'ListView_Figures
        '
        Me.ListView_Figures.CheckBoxes = True
        Me.ListView_Figures.GridLines = True
        resources.ApplyResources(Me.ListView_Figures, "ListView_Figures")
        Me.ListView_Figures.MultiSelect = False
        Me.ListView_Figures.Name = "ListView_Figures"
        Me.ListView_Figures.UseCompatibleStateImageBehavior = False
        '
        'ListView_Parcels
        '
        resources.ApplyResources(Me.ListView_Parcels, "ListView_Parcels")
        Me.ListView_Parcels.CheckBoxes = True
        Me.ListView_Parcels.FullRowSelect = True
        Me.ListView_Parcels.GridLines = True
        Me.ListView_Parcels.HideSelection = False
        Me.ListView_Parcels.MultiSelect = False
        Me.ListView_Parcels.Name = "ListView_Parcels"
        Me.ListView_Parcels.UseCompatibleStateImageBehavior = False
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
        'Radio_Figures
        '
        resources.ApplyResources(Me.Radio_Figures, "Radio_Figures")
        Me.Radio_Figures.Name = "Radio_Figures"
        Me.Radio_Figures.UseVisualStyleBackColor = True
        '
        'Radio_Parcels
        '
        Me.Radio_Parcels.Checked = True
        resources.ApplyResources(Me.Radio_Parcels, "Radio_Parcels")
        Me.Radio_Parcels.Name = "Radio_Parcels"
        Me.Radio_Parcels.TabStop = True
        Me.Radio_Parcels.UseVisualStyleBackColor = True
        '
        'Group_Analysis
        '
        resources.ApplyResources(Me.Group_Analysis, "Group_Analysis")
        Me.Group_Analysis.Controls.Add(Me.SplitContainer2)
        Me.Group_Analysis.Controls.Add(Me.Button_Pt)
        Me.Group_Analysis.Controls.Add(Me.Check_AcrossChord)
        Me.Group_Analysis.Controls.Add(Me.Check_CounterClock)
        Me.Group_Analysis.Controls.Add(Me.Label_PtBegin)
        Me.Group_Analysis.Name = "Group_Analysis"
        Me.Group_Analysis.TabStop = False
        '
        'SplitContainer2
        '
        resources.ApplyResources(Me.SplitContainer2, "SplitContainer2")
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.Text_PobXCoord)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label_X)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.Text_PobYCoord)
        Me.SplitContainer2.Panel2.Controls.Add(Me.Label_Y)
        Me.SplitContainer2.TabStop = False
        '
        'Text_PobXCoord
        '
        resources.ApplyResources(Me.Text_PobXCoord, "Text_PobXCoord")
        Me.Text_PobXCoord.Name = "Text_PobXCoord"
        Me.Text_PobXCoord.ReadOnly = True
        '
        'Label_X
        '
        resources.ApplyResources(Me.Label_X, "Label_X")
        Me.Label_X.Name = "Label_X"
        '
        'Text_PobYCoord
        '
        resources.ApplyResources(Me.Text_PobYCoord, "Text_PobYCoord")
        Me.Text_PobYCoord.Name = "Text_PobYCoord"
        Me.Text_PobYCoord.ReadOnly = True
        '
        'Label_Y
        '
        resources.ApplyResources(Me.Label_Y, "Label_Y")
        Me.Label_Y.Name = "Label_Y"
        '
        'Button_Pt
        '
        resources.ApplyResources(Me.Button_Pt, "Button_Pt")
        Me.Button_Pt.Name = "Button_Pt"
        Me.Button_Pt.UseVisualStyleBackColor = True
        '
        'Check_AcrossChord
        '
        resources.ApplyResources(Me.Check_AcrossChord, "Check_AcrossChord")
        Me.Check_AcrossChord.Name = "Check_AcrossChord"
        Me.Check_AcrossChord.UseVisualStyleBackColor = True
        '
        'Check_CounterClock
        '
        resources.ApplyResources(Me.Check_CounterClock, "Check_CounterClock")
        Me.Check_CounterClock.Name = "Check_CounterClock"
        Me.Check_CounterClock.UseVisualStyleBackColor = True
        '
        'Label_PtBegin
        '
        resources.ApplyResources(Me.Label_PtBegin, "Label_PtBegin")
        Me.Label_PtBegin.Name = "Label_PtBegin"
        '
        'Group_Settings
        '
        resources.ApplyResources(Me.Group_Settings, "Group_Settings")
        Me.Group_Settings.Controls.Add(Me.TableLayoutPanel1)
        Me.Group_Settings.Name = "Group_Settings"
        Me.Group_Settings.TabStop = False
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Format, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_DirecFormat, 2, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Value, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_PrecPrec, 1, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Precision, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_AreaPrec, 1, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Direction, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_PrecisionPerimeter, 0, 5)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_DirePrec, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Area, 0, 3)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_DistPrec, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_CloPrec, 1, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Distance, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Label_Closure, 0, 4)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'Label_Format
        '
        resources.ApplyResources(Me.Label_Format, "Label_Format")
        Me.Label_Format.Name = "Label_Format"
        '
        'Label_DirecFormat
        '
        resources.ApplyResources(Me.Label_DirecFormat, "Label_DirecFormat")
        Me.Label_DirecFormat.Name = "Label_DirecFormat"
        '
        'Label_Value
        '
        resources.ApplyResources(Me.Label_Value, "Label_Value")
        Me.Label_Value.Name = "Label_Value"
        '
        'Label_PrecPrec
        '
        resources.ApplyResources(Me.Label_PrecPrec, "Label_PrecPrec")
        Me.Label_PrecPrec.Name = "Label_PrecPrec"
        '
        'Label_Precision
        '
        resources.ApplyResources(Me.Label_Precision, "Label_Precision")
        Me.Label_Precision.Name = "Label_Precision"
        '
        'Label_AreaPrec
        '
        resources.ApplyResources(Me.Label_AreaPrec, "Label_AreaPrec")
        Me.Label_AreaPrec.Name = "Label_AreaPrec"
        '
        'Label_Direction
        '
        resources.ApplyResources(Me.Label_Direction, "Label_Direction")
        Me.Label_Direction.Name = "Label_Direction"
        '
        'Label_PrecisionPerimeter
        '
        resources.ApplyResources(Me.Label_PrecisionPerimeter, "Label_PrecisionPerimeter")
        Me.Label_PrecisionPerimeter.Name = "Label_PrecisionPerimeter"
        '
        'Label_DirePrec
        '
        resources.ApplyResources(Me.Label_DirePrec, "Label_DirePrec")
        Me.Label_DirePrec.Name = "Label_DirePrec"
        '
        'Label_Area
        '
        resources.ApplyResources(Me.Label_Area, "Label_Area")
        Me.Label_Area.Name = "Label_Area"
        '
        'Label_DistPrec
        '
        resources.ApplyResources(Me.Label_DistPrec, "Label_DistPrec")
        Me.Label_DistPrec.Name = "Label_DistPrec"
        '
        'Label_CloPrec
        '
        resources.ApplyResources(Me.Label_CloPrec, "Label_CloPrec")
        Me.Label_CloPrec.Name = "Label_CloPrec"
        '
        'Label_Distance
        '
        resources.ApplyResources(Me.Label_Distance, "Label_Distance")
        Me.Label_Distance.Name = "Label_Distance"
        '
        'Label_Closure
        '
        resources.ApplyResources(Me.Label_Closure, "Label_Closure")
        Me.Label_Closure.Name = "Label_Closure"
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
        'GroupBox_Settings
        '
        resources.ApplyResources(Me.GroupBox_Settings, "GroupBox_Settings")
        Me.GroupBox_Settings.Controls.Add(Me.Button_Save)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_SaveReport)
        Me.GroupBox_Settings.Controls.Add(Me.Label_SaveTo)
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
        'Label_SaveTo
        '
        resources.ApplyResources(Me.Label_SaveTo, "Label_SaveTo")
        Me.Label_SaveTo.Name = "Label_SaveTo"
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
        'SplitContainer1
        '
        resources.ApplyResources(Me.SplitContainer1, "SplitContainer1")
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Group_Analysis)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.Group_Settings)
        '
        'ReportForm_ParcelMapCheck
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.Controls.Add(Me.Group_Objects)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_ParcelMapCheck"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.Group_Objects.ResumeLayout(False)
        Me.Group_Analysis.ResumeLayout(False)
        Me.Group_Analysis.PerformLayout()
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.Panel2.PerformLayout()
        Me.SplitContainer2.ResumeLayout(False)
        Me.Group_Settings.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Group_Objects As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Select As System.Windows.Forms.Button
    Friend WithEvents Radio_Figures As System.Windows.Forms.RadioButton
    Friend WithEvents Radio_Parcels As System.Windows.Forms.RadioButton
    Friend WithEvents ListView_Figures As System.Windows.Forms.ListView
    Friend WithEvents ListView_Parcels As System.Windows.Forms.ListView
    Friend WithEvents Button_Deselect As System.Windows.Forms.Button
    Friend WithEvents Group_Analysis As System.Windows.Forms.GroupBox
    Friend WithEvents Group_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents Text_PobYCoord As System.Windows.Forms.TextBox
    Friend WithEvents Text_PobXCoord As System.Windows.Forms.TextBox
    Friend WithEvents Label_Y As System.Windows.Forms.Label
    Friend WithEvents Label_X As System.Windows.Forms.Label
    Friend WithEvents Label_PtBegin As System.Windows.Forms.Label
    Friend WithEvents Check_AcrossChord As System.Windows.Forms.CheckBox
    Friend WithEvents Check_CounterClock As System.Windows.Forms.CheckBox
    Friend WithEvents Button_Pt As System.Windows.Forms.Button
    Friend WithEvents Label_Value As System.Windows.Forms.Label
    Friend WithEvents Label_Format As System.Windows.Forms.Label
    Friend WithEvents Label_Precision As System.Windows.Forms.Label
    Friend WithEvents Label_Closure As System.Windows.Forms.Label
    Friend WithEvents Label_Area As System.Windows.Forms.Label
    Friend WithEvents Label_Direction As System.Windows.Forms.Label
    Friend WithEvents Label_Distance As System.Windows.Forms.Label
    Friend WithEvents Label_PrecisionPerimeter As System.Windows.Forms.Label
    Friend WithEvents Label_PrecPrec As System.Windows.Forms.Label
    Friend WithEvents Label_CloPrec As System.Windows.Forms.Label
    Friend WithEvents Label_AreaPrec As System.Windows.Forms.Label
    Friend WithEvents Label_DirecFormat As System.Windows.Forms.Label
    Friend WithEvents Label_DirePrec As System.Windows.Forms.Label
    Friend WithEvents Label_DistPrec As System.Windows.Forms.Label
    Friend WithEvents GroupBox_ReportDesc As System.Windows.Forms.GroupBox
    Friend WithEvents Label_ReportDesc As System.Windows.Forms.Label
    Friend WithEvents GroupBox_Settings As System.Windows.Forms.GroupBox
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents TextBox_SaveReport As System.Windows.Forms.TextBox
    Friend WithEvents Label_SaveTo As System.Windows.Forms.Label
    Friend WithEvents GroupBox_Progress As System.Windows.Forms.GroupBox
    Friend WithEvents ProgressBar_Creating As System.Windows.Forms.ProgressBar
    Friend WithEvents Button_CreateReport As System.Windows.Forms.Button
    Friend WithEvents Button_Help As System.Windows.Forms.Button
    Friend WithEvents Button_Done As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
End Class
