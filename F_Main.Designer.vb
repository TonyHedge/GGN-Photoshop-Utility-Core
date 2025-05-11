<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class F_Main
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
		components = New ComponentModel.Container()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(F_Main))
		MenuStrip1 = New MenuStrip()
		StartToolStripMenuItem = New ToolStripMenuItem()
		CloseStripMenuItem = New ToolStripMenuItem()
		ContextMenuStrip1 = New ContextMenuStrip(components)
		SS_Main = New StatusStrip()
		SL_Main = New ToolStripStatusLabel()
		Rtb_Log = New RichTextBox()
		MenuStrip1.SuspendLayout()
		SS_Main.SuspendLayout()
		SuspendLayout()
		' 
		' MenuStrip1
		' 
		MenuStrip1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
		MenuStrip1.Dock = DockStyle.None
		MenuStrip1.Items.AddRange(New ToolStripItem() {StartToolStripMenuItem, CloseStripMenuItem})
		MenuStrip1.Location = New Point(5, 5)
		MenuStrip1.Name = "MenuStrip1"
		MenuStrip1.Padding = New Padding(7, 2, 0, 2)
		MenuStrip1.Size = New Size(100, 24)
		MenuStrip1.TabIndex = 0
		MenuStrip1.Text = "MenuStrip1"
		' 
		' StartToolStripMenuItem
		' 
		StartToolStripMenuItem.Name = "StartToolStripMenuItem"
		StartToolStripMenuItem.Size = New Size(43, 20)
		StartToolStripMenuItem.Text = "Start"
		' 
		' CloseStripMenuItem
		' 
		CloseStripMenuItem.Name = "CloseStripMenuItem"
		CloseStripMenuItem.Size = New Size(48, 20)
		CloseStripMenuItem.Text = "Close"
		' 
		' ContextMenuStrip1
		' 
		ContextMenuStrip1.Name = "ContextMenuStrip1"
		ContextMenuStrip1.Size = New Size(61, 4)
		' 
		' SS_Main
		' 
		SS_Main.Items.AddRange(New ToolStripItem() {SL_Main})
		SS_Main.Location = New Point(0, 506)
		SS_Main.Name = "SS_Main"
		SS_Main.Padding = New Padding(1, 0, 16, 0)
		SS_Main.Size = New Size(718, 22)
		SS_Main.TabIndex = 1
		SS_Main.Text = "StatusStrip1"
		' 
		' SL_Main
		' 
		SL_Main.Name = "SL_Main"
		SL_Main.Size = New Size(0, 17)
		' 
		' Rtb_Log
		' 
		Rtb_Log.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
		Rtb_Log.Location = New Point(0, 31)
		Rtb_Log.Margin = New Padding(4, 3, 4, 3)
		Rtb_Log.Name = "Rtb_Log"
		Rtb_Log.Size = New Size(717, 467)
		Rtb_Log.TabIndex = 2
		Rtb_Log.Text = ""
		' 
		' F_Main
		' 
		AutoScaleDimensions = New SizeF(7F, 15F)
		AutoScaleMode = AutoScaleMode.Font
		ClientSize = New Size(718, 528)
		Controls.Add(Rtb_Log)
		Controls.Add(SS_Main)
		Controls.Add(MenuStrip1)
		Icon = CType(resources.GetObject("$this.Icon"), Icon)
		MainMenuStrip = MenuStrip1
		Margin = New Padding(4, 3, 4, 3)
		Name = "F_Main"
		Text = "GGN Photoshop Utility - Core"
		MenuStrip1.ResumeLayout(False)
		MenuStrip1.PerformLayout()
		SS_Main.ResumeLayout(False)
		SS_Main.PerformLayout()
		ResumeLayout(False)
		PerformLayout()

	End Sub

	Friend WithEvents MenuStrip1 As MenuStrip
	Friend WithEvents StartToolStripMenuItem As ToolStripMenuItem
	Friend WithEvents ContextMenuStrip1 As ContextMenuStrip
	Friend WithEvents SS_Main As StatusStrip
	Friend WithEvents SL_Main As ToolStripStatusLabel
	Friend WithEvents Rtb_Log As RichTextBox
	Friend WithEvents CloseStripMenuItem As ToolStripMenuItem
End Class
