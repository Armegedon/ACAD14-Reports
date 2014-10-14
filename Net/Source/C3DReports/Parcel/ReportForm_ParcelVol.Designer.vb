<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportForm_ParcelVol
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ReportForm_ParcelVol))
        Me.Combo_Surface = New System.Windows.Forms.ComboBox
        Me.Label_Surface = New System.Windows.Forms.Label
        Me.TextBox_Ele = New System.Windows.Forms.TextBox
        Me.TextBox_Fill = New System.Windows.Forms.TextBox
        Me.TextBox_Cut = New System.Windows.Forms.TextBox
        Me.Label_Elevation = New System.Windows.Forms.Label
        Me.Label_CutCorrection = New System.Windows.Forms.Label
        Me.Label_FillCorrection = New System.Windows.Forms.Label
        Me.Label_CompSurfaceValue = New System.Windows.Forms.Label
        Me.Label_BaseSurfaceValue = New System.Windows.Forms.Label
        Me.Label_TypeValue = New System.Windows.Forms.Label
        Me.Label_CompSurface = New System.Windows.Forms.Label
        Me.Label_BaseSurface = New System.Windows.Forms.Label
        Me.Label_Type = New System.Windows.Forms.Label
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
        Me.GroupBox_Parcel = New System.Windows.Forms.GroupBox
        Me.ListView_Parcels = New System.Windows.Forms.ListView
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox_ReportDesc.SuspendLayout()
        Me.GroupBox_Settings.SuspendLayout()
        Me.GroupBox_Progress.SuspendLayout()
        Me.GroupBox_Parcel.SuspendLayout()
        Me.SuspendLayout()
        '
        'Combo_Surface
        '
        resources.ApplyResources(Me.Combo_Surface, "Combo_Surface")
        Me.Combo_Surface.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.Combo_Surface.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.Combo_Surface.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Surface.FormattingEnabled = True
        Me.Combo_Surface.Name = "Combo_Surface"
        '
        'Label_Surface
        '
        resources.ApplyResources(Me.Label_Surface, "Label_Surface")
        Me.Label_Surface.Name = "Label_Surface"
        '
        'TextBox_Ele
        '
        resources.ApplyResources(Me.TextBox_Ele, "TextBox_Ele")
        Me.TextBox_Ele.Name = "TextBox_Ele"
        '
        'TextBox_Fill
        '
        resources.ApplyResources(Me.TextBox_Fill, "TextBox_Fill")
        Me.TextBox_Fill.Name = "TextBox_Fill"
        '
        'TextBox_Cut
        '
        resources.ApplyResources(Me.TextBox_Cut, "TextBox_Cut")
        Me.TextBox_Cut.Name = "TextBox_Cut"
        '
        'Label_Elevation
        '
        resources.ApplyResources(Me.Label_Elevation, "Label_Elevation")
        Me.Label_Elevation.Name = "Label_Elevation"
        '
        'Label_CutCorrection
        '
        resources.ApplyResources(Me.Label_CutCorrection, "Label_CutCorrection")
        Me.Label_CutCorrection.Name = "Label_CutCorrection"
        '
        'Label_FillCorrection
        '
        resources.ApplyResources(Me.Label_FillCorrection, "Label_FillCorrection")
        Me.Label_FillCorrection.Name = "Label_FillCorrection"
        '
        'Label_CompSurfaceValue
        '
        resources.ApplyResources(Me.Label_CompSurfaceValue, "Label_CompSurfaceValue")
        Me.Label_CompSurfaceValue.Name = "Label_CompSurfaceValue"
        '
        'Label_BaseSurfaceValue
        '
        resources.ApplyResources(Me.Label_BaseSurfaceValue, "Label_BaseSurfaceValue")
        Me.Label_BaseSurfaceValue.Name = "Label_BaseSurfaceValue"
        '
        'Label_TypeValue
        '
        resources.ApplyResources(Me.Label_TypeValue, "Label_TypeValue")
        Me.Label_TypeValue.Name = "Label_TypeValue"
        '
        'Label_CompSurface
        '
        resources.ApplyResources(Me.Label_CompSurface, "Label_CompSurface")
        Me.Label_CompSurface.Name = "Label_CompSurface"
        '
        'Label_BaseSurface
        '
        resources.ApplyResources(Me.Label_BaseSurface, "Label_BaseSurface")
        Me.Label_BaseSurface.Name = "Label_BaseSurface"
        '
        'Label_Type
        '
        resources.ApplyResources(Me.Label_Type, "Label_Type")
        Me.Label_Type.Name = "Label_Type"
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
        Me.GroupBox_Settings.Controls.Add(Me.Label2)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_Fill)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_Ele)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_Cut)
        Me.GroupBox_Settings.Controls.Add(Me.Label1)
        Me.GroupBox_Settings.Controls.Add(Me.Label_CompSurfaceValue)
        Me.GroupBox_Settings.Controls.Add(Me.Combo_Surface)
        Me.GroupBox_Settings.Controls.Add(Me.Label_FillCorrection)
        Me.GroupBox_Settings.Controls.Add(Me.Label_CutCorrection)
        Me.GroupBox_Settings.Controls.Add(Me.Label_Elevation)
        Me.GroupBox_Settings.Controls.Add(Me.Label_BaseSurfaceValue)
        Me.GroupBox_Settings.Controls.Add(Me.Button_Save)
        Me.GroupBox_Settings.Controls.Add(Me.Label_TypeValue)
        Me.GroupBox_Settings.Controls.Add(Me.Label_Surface)
        Me.GroupBox_Settings.Controls.Add(Me.Label_CompSurface)
        Me.GroupBox_Settings.Controls.Add(Me.TextBox_SaveReport)
        Me.GroupBox_Settings.Controls.Add(Me.Label_BaseSurface)
        Me.GroupBox_Settings.Controls.Add(Me.Label_SaveTo)
        Me.GroupBox_Settings.Controls.Add(Me.Label_Type)
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
        'GroupBox_Parcel
        '
        resources.ApplyResources(Me.GroupBox_Parcel, "GroupBox_Parcel")
        Me.GroupBox_Parcel.Controls.Add(Me.ListView_Parcels)
        Me.GroupBox_Parcel.Name = "GroupBox_Parcel"
        Me.GroupBox_Parcel.TabStop = False
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
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'ReportForm_ParcelVol
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.Button_Done
        Me.Controls.Add(Me.GroupBox_Parcel)
        Me.Controls.Add(Me.GroupBox_Progress)
        Me.Controls.Add(Me.Button_CreateReport)
        Me.Controls.Add(Me.Button_Help)
        Me.Controls.Add(Me.Button_Done)
        Me.Controls.Add(Me.GroupBox_Settings)
        Me.Controls.Add(Me.GroupBox_ReportDesc)
        Me.MinimizeBox = False
        Me.Name = "ReportForm_ParcelVol"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.GroupBox_ReportDesc.ResumeLayout(False)
        Me.GroupBox_Settings.ResumeLayout(False)
        Me.GroupBox_Settings.PerformLayout()
        Me.GroupBox_Progress.ResumeLayout(False)
        Me.GroupBox_Parcel.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Label_Surface As System.Windows.Forms.Label
    Friend WithEvents Combo_Surface As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox_Ele As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Cut As System.Windows.Forms.TextBox
    Friend WithEvents TextBox_Fill As System.Windows.Forms.TextBox
    Friend WithEvents Label_Elevation As System.Windows.Forms.Label
    Friend WithEvents Label_CutCorrection As System.Windows.Forms.Label
    Friend WithEvents Label_FillCorrection As System.Windows.Forms.Label
    Friend WithEvents Label_CompSurfaceValue As System.Windows.Forms.Label
    Friend WithEvents Label_BaseSurfaceValue As System.Windows.Forms.Label
    Friend WithEvents Label_TypeValue As System.Windows.Forms.Label
    Friend WithEvents Label_CompSurface As System.Windows.Forms.Label
    Friend WithEvents Label_BaseSurface As System.Windows.Forms.Label
    Friend WithEvents Label_Type As System.Windows.Forms.Label
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
    Friend WithEvents GroupBox_Parcel As System.Windows.Forms.GroupBox
    Friend WithEvents ListView_Parcels As System.Windows.Forms.ListView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
End Class
