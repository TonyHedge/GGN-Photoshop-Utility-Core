'
'	GGN Photoshop Utility
'
'	Copyright @2025 C.A.(Tony) Hedge
'
'   Does the processing of Colour and B&W photo files for input into GGN Magazine
'   Converts .jpg or .psd file to a .tif file with LGW Compression and ICC/sRGB off
'   Resizes photo without resampling to 300 ppi
'   Resizes with resampling to 19cm for colour photos and 9 cm wide for B&W photos
'   Converts colour photos to CMYK, and B&W photos to Greyscale
'   Does above for a single file or all .jpg (and .tif) files in the folder at the user's discretion
'   Saves resulting image in folder of user's choice
'
'   Change History:-
'   v1.0        19.08.11    Original version
'   v1.1        20.08.11    Restructured code so as to save the Tiff image just once
'   v1.2        21.08.11    Restructured code as VBScript does not support the GoTo statement
'                           This makes error handling very complicated and obsure
'                           We have to use nested if ... then ... else ... statements to handle errors
'   v1.3        16.04.12    Resize Colour photos to 19cm wide and B&W photos to 9 cm wide
'                           Option to process all currently opened images or ...
'                           ... all images in the input folder (either/or option)
'                           Use either MsComDlg.CommonDialog or UserAccounts.CommonDialog to display the File Open/Save dialog boxes
'   v1.4        02.06.12    If neither of the above work, use the Shell.BrowseForFolder method
'	v1.5		28.06.13	Replace the Microsoft Common Dialog Box DLLs with CAHCommonDialog DLL, which will (hopefully) work with all
'							versions of the Windows operating system
'	v2.0		15.02.25	Converted to Windows Forms Application
'	v2.1		10.05.25	Simplified implementation and made the selection of Phots more user friendly
'	v2.2		12.05.25	Tbd
'
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml

#Disable Warning CA1416
'
'************************************************************************************************************
'************************************************************************************************************
'
'	F_Main (Class)
'
Public Class F_Main
'
' Declare WinAPI functions and their constants
'
	Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, _
														ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As UInteger) As Boolean
	Public Const SWP_NOSIZE As Short = 1
	Public Const SWP_NOMOVE As Short = 2
	Public Const SWP_NOACTIVATE As Short = 16

	Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
	Public Const SW_RESTORE As Integer = 9
'
'	Global Data
'
	Dim BlackWhite As Boolean
	Dim MyPhotoPath As String
	Dim MyPhotoNames As New List(Of String)()
	Dim Persistent As List(Of String)
	Dim PhotoShopApp
	Dim PhotoShopProcess
	Dim PhotoTypes As New List(Of String) From {"*.jpg", "*.tiff", "*.png", "*.bmp", "*.gif", "*.psd"}
	Dim OutputFileName As String
	Dim OutputFolderName As String
	Dim RegexMatches As MatchCollection

	Dim fontFamily As New FontFamily("Arial")
	Dim LargeFont As New Font(fontFamily, 12, FontStyle.Underline Or FontStyle.Bold)
	Dim StndFont As New Font(fontFamily, 10, FontStyle.Regular)
	Dim BoldFont As New Font(fontFamily, 10, FontStyle.Bold)

	Public Const VersionNbr = "2.2"
'
'************************************************************************************************************
'
'	F_Main_Load (Event Procedure)
'
'	Called when
'		The Application is started
'
	Private Sub F_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

		AppendToRTB("GGN Photoshop Utility (v" & VersionNbr & ")" & vbCrLf, Color.Blue, LargeFont)

		Me.Location = New Point(5, 5)

		#if DEBUG
			Me.Text = Me.Text + " - Debug"
		#end if

		Try
			Persistent = File.ReadAllLines($"{Application.StartupPath}\\Settings.xml").ToList()

			RegexMatches = Regex.Matches(Persistent(1), "=""(.*?)""")
			MyPhotoPath = RegexMatches(0).Groups(1).Value
			OutputFolderName = RegexMatches(1).Groups(1).Value

		Catch
			MyPhotoPath = "C:\"
			OutputFolderName = "C:\"
		End Try

	End Sub

'
'************************************************************************************************************
'
'	F_Main_Closing (Event Procedure)
'
'	Called when
'		The Application is closing down
'
	Private Sub F_Main_Closing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing

		If PhotoShopProcess IsNot Nothing Then
			Try
				PhotoShopProcess(0).Kill()                                          ' Kill the Photoshop process
			Catch                                                                   ' Ignore any errors
			End Try
			PhotoShopProcess = Nothing
		End If

		Persistent.Clear()
		Persistent.Add( "<?xml version=""1.0"" encoding=""UTF-8""?>" )
		Persistent.Add( $"<Data SourceFolder=""{MyPhotoPath}"" OutputFolder=""{OutputFolderName}\""></Data>" )
		File.WriteAllLines($"{Application.StartupPath}\\Settings.xml", Persistent)

	End Sub

'
'************************************************************************************************************
'
'	Close_Click (Event Procedure)
'
'	Called when
'		User clicks on the Close menu 
'
'	Closes the application
'
	Private Sub Close_Click(sender As Object, e As EventArgs) Handles CloseStripMenuItem.Click

		Close                                                    ' Close the application

	End Sub

'
'************************************************************************************************************
'
'	Start_Click (Event Procedure)
'
'	Called when
'		User clicks on the Start menu
'
'	Launches PhotoShop CS2, if not already running
'	Creates a list of the Photos/Images to be converted to GGN format by
'		(a) Including in the list just the currently active Photo/Image in PhotoShop
'		(b) Including in the list all the Photos/Images in the same folder as the Photo/Image currently active in PhotoShop
'		(c) Prompting the user to select the Photos/Images to be processed
'	Converts the selected Photos/Images to GGN format
'
	Private Sub Start_Click(sender As Object, e As EventArgs) Handles StartToolStripMenuItem.Click

		Dim InFolder As DirectoryInfo
		Dim InFileList As FileInfo()
		Dim InFileInfo As FileInfo
		Dim i As Integer
		Dim r
		Dim c As CentreDialog
'
'   Launch Photoshop, if not already running, and re-size it's main windows to occupy half the monitor
'
		On Error Resume Next                                                        ' Handle errors internally
		Err.Clear

		Me.SL_Main.Text = "Launching Photoshop"
		Me.SS_Main.Update
		PhotoShopApp = CreateObject("Photoshop.Application")
		If Err.Number <> 0 Then
			AppendToRTB("Cannot launch PhotoShop, application aborted" & vbCrLf, Color.Red, BoldFont)
			AppendToRTB("Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont)
			Exit Sub
		End If
		Me.SL_Main.Text = vbNullString
		Me.SS_Main.Update
		AppendToRTB("Photoshop launched" & vbCrLf, Color.Black, StndFont)

		PhotoShopProcess = Process.GetProcessesByName("Photoshop")                  ' Array containing the Photoshop Process

		Call ResizePhotoshopWindow(PhotoShopProcess(0).MainWindowHandle)            ' Resize the Photoshop window to occupy the right-hand side of the monitor

		PhotoShopApp.Preferences.RulerUnits = 3                                     'for PsUnits --> 1 (psCm)
		PhotoShopApp.DisplayDialogs = 3                                             'for PsDialogModes --> 3 (psDisplayNoDialogs)
'
'	If there are an open Photos in Photoshop, ask the user if they are to be processed
'
		If PhotoShopApp.Documents.Count > 0 Then
			Using New CentreDialog(Me)
				r = MsgBox("Process all photos in the same folder as the opened photos ?" + Environment.NewLine + Environment.NewLine +
							"Yes - Process all photos in the same folder" + Environment.NewLine +
							"No - Process just the opened photos", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "GGN PhotoShop Automation")
			End Using

			If r = vbCancel Then
				Exit Sub
'
'	Just the opened photos are to be processed
'
			ElseIf r = vbNo Then
				MyPhotoNames.Clear()
				For i = 1 To PhotoShopApp.Documents.Count
					MyPhotoNames.Add(PhotoShopApp.Documents(i).FullName)
				Next
'
'	Create a list of all the photos in the same folder as the opened photos
'
			ElseIf r = vbYes Then
				MyPhotoNames.Clear()
				MyPhotoPath = PhotoShopApp.Documents(1).FullName.Substring(0, InStrRev(PhotoShopApp.Documents(1).FullName, "\"))
				Dim FileNames As List(Of String)
				For Each p As String In PhotoTypes                                  ' Loop over all Images types (*.jpg etc)
					FileNames = Directory.GetFiles(MyPhotoPath, p).ToList           ' All files of a particular type in the source folder
					MyPhotoNames.AddRange(FileNames)
				Next
			End If
'
' There is no open image, so ask the user to select the images to be processed
'
		Else
			MyPhotoNames = SelectImages()                                           ' Get the names of the images the user has selected
			If Not MyPhotoNames Is Nothing Then
			Else
				' User cancelled or error encountered
				Exit Sub
			End If
		End If
'
'	Display the source folder name and the list of photos to be processed
'
		MyPhotoPath = MyPhotoNames(0).Substring(0, InStrRev(MyPhotoNames(0), "\"))  ' Get the path of the first selected image
		AppendToRTB("Source folder:" & vbTab & MyPhotoPath & vbCrLf, Color.Black, StndFont)

		For Each p As String In MyPhotoNames
			AppendToRTB(vbTab & p.Substring(InStrRev(p, "\")) & vbCrLf, Color.Black, StndFont)
		Next
'
'   Select folder in which reformatted photos are to be stored
'
		OutputFolderName = GetOutputPath
		If OutputFolderName = vbNullString Then
			AppendToRTB("Failed to get name of Output Folder, run aborted" & vbCrLf, Color.Red, BoldFont)
			Exit Sub
		End If

		AppendToRTB("Output folder:" & vbTab & OutputFolderName & vbCrLf, Color.Black, StndFont)
'
'	Ask whether photo are to be converted to Black & White
'
		Using New CentreDialog(Me)
			r = MsgBox("Convert Photographs to Black and White ?", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "GGN PhotoShop Automation")
		End Using

		If r = vbYes
			' It is to be a B&W photo so ...
			BlackWhite = True
		ElseIf r = vbNo
			BlackWhite = False
		Else
			Exit Sub
		End If
'
'	Convert all the selected photos to GGN format
'
	For Each p As String In MyPhotoNames
			Err.Clear
			PhotoShopApp.Open(p)
			If Err.Number <> 0 Then
				AppendToRTB(vbTab & "Unable to open image " & p & vbCrLf, Color.Red, BoldFont)
				AppendToRTB(vbTab & "Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont)
			Else
				AppendToRTB(vbTab & PhotoShopApp.ActiveDocument.Name & vbCrLf, Color.Black, StndFont)
				OutputFileName = ProcessPhoto(PhotoShopApp.ActiveDocument, OutputFolderName, PhotoShopApp.ActiveDocument.Name, BlackWhite)
			End If
		Next
'
'   Let the user know that the script has completed
'
		AppendToRTB("All images have been processed" & vbCrLf, Color.Blue, BoldFont)

	End Sub
'
'************************************************************************************************************
'
'	AppendToRTB (Subroutine)
'
'	Called by
'		F_Main.ProcessPhoto
'		F_Main.Start_Click
'		F_Main.ResizePhotoshopWindow
'
'	Parameters
'		inText		- Text to be appended
'		inColour	- Foreground Text Colour
'		inFont		- Font
'
	Public Sub AppendToRTB(inText As String,
							inColour As Drawing.Color,
							inFont As Font)

		Call Me.Rtb_Log.SuspendLayout()
		Me.Rtb_Log.SelectionStart = Me.Rtb_Log.TextLength
		Me.Rtb_Log.SelectionLength = 0
		Me.Rtb_Log.SelectionFont = inFont
		Me.Rtb_Log.SelectionColor = inColour
		Me.Rtb_Log.SelectedText = inText
		Call Me.Rtb_Log.ResumeLayout()

		Me.Rtb_Log.ScrollToCaret

	End Sub
'
'************************************************************************************************************
'
'	GetOutputPath (Function)
'
'	Called by
'		F_Main.Start_Click
'
'	Returns
'		Null String			- Failed to select a folder
'		Output Folder Path
'
	Private Function GetOutputPath() As String

		Dim folderBrowserDialog1 As FolderBrowserDialog

		On Error Resume Next                                                    ' Handle errors internally
		folderBrowserDialog1 = New FolderBrowserDialog                          ' New instance of the File Browser dialobox

		folderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop
		folderBrowserDialog1.SelectedPath = OutputFolderName                    ' Initial selected folder
		folderBrowserDialog1.Description = "Select output FOLDER"
		folderBrowserDialog1.UseDescriptionForTitle = True                      ' Use the description as the title of the File Browser
		folderBrowserDialog1.ShowNewFolderButton = True                         ' Allow the browser to create new folders

		If folderBrowserDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
			Return vbNullString                                             ' Failed to select an output folder
		Else
			Return folderBrowserDialog1.SelectedPath                            ' Output folder Path
		End If
	End Function
'
'************************************************************************************************************
'
'	ProcessPhoto (Function)
'
'	Called by
'		F_Main.Start_Click
'
'	Parameters:-
'		MyPhoto				- Photograph to be processed
'		OutputFolderName	- Name of the Output Folder
'		OutputFileName		- File Name of the Processed Photo
'		BlackWhite			- Whether Balck & White Photo or Coloured Photo
'
	Function ProcessPhoto(ByVal MyPhoto, ByVal OutputFolderName, ByVal OutputFileName, ByVal BlackWhite) As String

		Dim i, TiffOptions

		On Error Resume Next
'
'   Make sure file extension of output filename is ".tif"
'
		OutputFileName = OutputFileName.Substring(0, InStrRev(OutputFileName, ".")) & "tif"
'
'   Now configure the photo with the required parameters
'
		MyPhoto.Flatten                                     ' Flatten the layers in the image
		MyPhoto.ColorProfileType = 1                        ' 1 = psNo; Turn ICC/sRGB off
		MyPhoto.ResizeImage(, , 300, 1)                 ' 1 = psNoResampling; Resize image to 300 ppi without ReSampling
'
' It is to be a B&W photo so ...
'
		If BlackWhite Then
			If MyPhoto.Width > 9 Then
				MyPhoto.ResizeImage(9 * 25 / 6, , , 5)      ' (CS2) Resize images to 9cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
'				 MyPhoto.ResizeImage 9, , , 5				' (CS3) Resize images to 9cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
			End If

			MyPhoto.ChangeMode(4)                           ' 4 = psConvertToLab; convert colour profile to Lab mode

			MyPhoto.Channels("a").Delete                    ' Delete the "a" channel (this also deletes the "b"  channel)
			MyPhoto.Channels("Alpha 2").Delete              ' Delete the "Alpha 2" channel leaving only the "Alpha 1" channel

			MyPhoto.ChangeMode(1)                           ' 1 = psConvertToGrayscale; convert colour profile to Grayscale
'
'   Suffix the filename of a B&W image with "_BW"
'
			i = InStrRev(OutputFileName, ".")               ' Position (if any) of "." in the Output File Name
			OutputFileName = OutputFileName.Substring(OutputFileName.Length - i) & "_BW" & OutputFileName.Substring(i + 1, OutputFileName.Length - i)
'
'	It's a colour photo so ...
'	
		Else
			If MyPhoto.Width > 19 Then
				MyPhoto.ResizeImage(19 * 25 / 6, , , 5)    ' (CS2) Resize images to 16cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
				'	            MyPhoto.ResizeImage (19, , , 5)				' (CS3) Resize images to 16cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
			End If

			MyPhoto.ChangeMode(3)                          ' 3 = psConvertToCMYK; convert colour profile to CMYK
		End If
'
'   Save the newly configured Tiff image
'
		Err.Clear
		TiffOptions = CreateObject("Photoshop.TIFFSaveOptions")
		If Err.Number <> 0 Then
			AppendToRTB("Cannot save output file '" & OutputFolderName & "\" & OutputFileName & "' as a TIFF, attempt aborted" & vbCrLf, Color.Red, BoldFont)
			AppendToRTB("Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont)
		Else
			TiffOptions.ImageCompression = 2                ' 2 = psTiffLZW; LZW encoding (compression)

			Err.Clear
			MyPhoto.SaveAs(OutputFolderName & "\" & OutputFileName, TiffOptions, True, 2)
			If Err.Number <> 0 Then
				AppendToRTB("Cannot save output file '" & OutputFolderName & "\" & OutputFileName & "' as a TIFF, attempt aborted" & vbCrLf, Color.Red, BoldFont)
				AppendToRTB("Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont)
			End If
		End If
'
'   Now close the source image without modifying it
'
		MyPhoto.Close(2)                                    ' 2 = psDoNotSaveChanges; Close without saving changes to Jpeg image

		TiffOptions = Nothing

		Return OutputFileName

	End Function
'
'************************************************************************************************************
'
'	ResizePhotoshopWindow (Subroutine)
'
'	Called by
'		F_Main.Start_Click
'
'	Parameters:-
'		WindowHandle				- Handle of Photoshop's Main Window
'
'   This subroutine controls the size and position of the InDesign window. The InDesign window is always displayed underlapping
'	the GGN Utilty window, but may either be displayed as encountered when the GGN Utility is first run or in the bottom right
'	hand half of the screen
'
	Sub ResizePhotoshopWindow(ByVal WindowHandle As IntPtr)

		Dim b As Boolean
'
'	Make the Photoshop window occupy the right-hand side of the monitor
'
		b = SetWindowPos(WindowHandle, 1, _                                  ' Bottom of the Z-order
							Screen.PrimaryScreen.WorkingArea.Width / 2,         ' Left = Halfway across the Monitor
							1,                                                  ' Top = Top of Monitor
							Screen.PrimaryScreen.WorkingArea.Width / 2,         ' Width = Half the Monitor's width
							Screen.PrimaryScreen.WorkingArea.Height - 1,            ' Height = Full monitor height
							SWP_NOACTIVATE)                                 ' Do not Activate the Window
		If Not b Then
			AppendToRTB("Failed to re-size the Photoshop window" & vbCrLf, Color.Red, BoldFont)
			Exit Sub
		End If

		b = ShowWindow(WindowHandle, SW_RESTORE)                                ' Restore (Normalise) the window
		If Not b Then
			AppendToRTB("Failed to Restore the Photoshop window" & vbCrLf, Color.Red, BoldFont)
			Exit Sub
		End If

	End Sub
'
'************************************************************************************************************
'
'	Select Images (Subroutine)
'
'	Called by
'		F_Main.Start_Click
'
'	Returns:-
'		List of selected images' names
'		Nothing	- user cancelled or errror encountered
'
'   This subroutine selects the images to be processed
'
	Function SelectImages() As List(Of String)

		Dim openFileDialog1 As New OpenFileDialog()

		openFileDialog1.InitialDirectory = MyPhotoPath
		openFileDialog1.Filter = "Photo files |*.jpg;*.png;*.tif;*.bmp;*.gif;*.psd|All files (*.*)|*.*"
		openFileDialog1.FilterIndex = 1
		openFileDialog1.RestoreDirectory = False
		openFileDialog1.Title = "Select the Photos to be processed"
		openFileDialog1.Multiselect = True

		Try
			If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				' Get the path of selected files
				Return openFileDialog1.FileNames.ToList()
			Else
				Using New CentreDialog(Me)
					MsgBox("User cancelled", vbOK + vbInformation + vbMsgBoxSetForeground, "GGN PhotoShop Utility")
				End Using
				Return Nothing
			End If
		Catch ex As Exception
			Using New CentreDialog(Me)
				MsgBox("Unable to display the File Open dialogbox" + Environment.NewLine + Environment.NewLine + ex.Message,
							vbOK + vbCritical + vbMsgBoxSetForeground, "GGN PhotoShop Utility")
			End Using
			Return Nothing
		End Try

	End Function

End Class
