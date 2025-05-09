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
'
Imports System.IO
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
	Public Declare Function SetWindowPos Lib "user32" ( ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal x As Integer, _
														ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, byVal uFlags As UInteger ) As Boolean
	Public Const SWP_NOSIZE As Short = 1
	Public Const SWP_NOMOVE As Short = 2
	Public Const SWP_NOACTIVATE As Short = 16

	Public Declare Function ShowWindow Lib "user32" ( ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
	Public Const SW_RESTORE As Integer =9
'
'	Global Data
'
	Dim BlackWhite As Boolean
	Dim MyPhoto As Object
	Dim MyPhotoPath As String
	Dim MyPhotoName As String
	Dim MyPhotoNames As new List(Of String) ()
	Dim PhotoShopApp
	Dim PhotoShopProcess
	Dim OutputFileName As String
	Dim OutputFolderName As String
	Dim OpenOutputFile

	Dim fontFamily As New FontFamily("Arial")
	Dim LargeFont As New Font( fontFamily, 12, FontStyle.Underline Or FontStyle.Bold)
	Dim StndFont As New Font( fontFamily, 10, FontStyle.Regular)
	Dim	BoldFont As New Font( fontFamily, 10, FontStyle.Bold )

	Public Const VersionNbr = "2.1"
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

		Me.Location = New Point( 5, 5 )

	End Sub
'
'************************************************************************************************************
'
'	Start_Click (Event Procedure)
'
'	Called when
'		User clicks on the Start menu
'
	Private Sub Start_Click(sender As Object, e As EventArgs) Handles StartToolStripMenuItem.Click

	Dim InFolder As DirectoryInfo
	Dim InFileList As FileInfo()
	Dim InFileInfo As FileInfo
	Dim i As Integer
	Dim r
	Dim c As CentreDialog
'
'   Create Photoshop object and re-size it's main windows to occupy half the monitor
'
		On Error Resume Next														' Handle errors internally
		Err.Clear

		Me.SL_Main.Text = "Launching Photoshop"
		Me.SS_Main.Update 
		PhotoShopApp = CreateObject( "Photoshop.Application" )
		If Err.Number <> 0 Then
			AppendToRTB ( "Cannot launch PhotoShop, application aborted" & vbCrLf, Color.Red, BoldFont )
			AppendToRTB ( "Error Description: " & Err.Description &vbCrLf, Color.Red, BoldFont )
			Exit Sub
		End if
		Me.SL_Main.Text = vbNullString
		Me.SS_Main.Update
		AppendToRTB ( "Photoshop launched" &vbCrLf, Color.Black, StndFont )

		PhotoShopProcess = Process.GetProcessesByName("Photoshop")					' Array containing the Photoshop Process

		Call ResizePhotoshopWindow(PhotoShopProcess(0).MainWindowHandle)			' Resize the Photoshop window to occupy the right-hand side of the monitor

		PhotoShopApp.Preferences.RulerUnits = 3										'for PsUnits --> 1 (psCm)
		PhotoShopApp.DisplayDialogs = 3												'for PsDialogModes --> 3 (psDisplayNoDialogs)
'
'	If there are an open photos in Photoshop, ask the user if they are to be processed
'
		If PhotoShopApp.Documents.Count > 0 Then
			Using New CentreDialog(Me)
				r = MsgBox( "Process just curently opened photo(s) ?" + Environment.NewLine + Environment.NewLine +
							"Yes - Just process the opened photos" + Environment.NewLine +
							"No - Process all photos in the folder", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "GGN PhotoShop Automation" )
			End Using

			If r = vbCancel Then
				Exit Sub
'
'	Just the opened photos are to be processed
'
			ElseIf r = vbYes Then
				MyPhotoNames.Clear()
				For i = 1 To PhotoShopApp.Documents.Count
					MyPhotoNames.Add( PhotoShopApp.Documents(i).FullName )
				Next
'
'	All the photos in the same folder as the opened photos are to be processed
'
			ElseIf r = vbNo Then
				MyPhotoNames.Clear()
				Dim FirstName As String = PhotoShopApp.Documents(1).FullName
				Dim FileNames As List (Of String)
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.jpg").ToList ' All the *.jpg files in the folder
				MyPhotoNames.AddRange( FileNames )
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.tif").ToList ' All the *.tif files in the folder
				MyPhotoNames.AddRange( FileNames )
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.png").ToList ' All the *.png files in the folder
				MyPhotoNames.AddRange( FileNames )
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.bmp").ToList ' All the *.bmp files in the folder
				MyPhotoNames.AddRange( FileNames )
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.gif").ToList ' All the *.gif files in the folder
				MyPhotoNames.AddRange( FileNames )
				FileNames = Directory.GetFiles(FirstName.Substring(0, InStrRev(FirstName, "\")), "*.psd").ToList ' All the *.psd files in the folder
				MyPhotoNames.AddRange( FileNames )
			End If

			MyPhoto = PhotoShopApp.ActiveDocument
			MyPhotoPath = MyPhoto.Path
			MyPhotoName = MyPhoto.Name
			OutputFileName = MyPhoto.Name
'
' There is no open image, so ask the user to select the images to be processed
'
		Else
			MyPhoto = Nothing
			MyPhotoPath = vbNullString
			MyPhotoName = vbNullString
			OutputFileName = vbNullString
			MyPhotoNames = SelectImages()											' Get the names of the images to be processed
			If Not MyPhotoNames Is Nothing Then
			Else
				' User cancelled or error encountered
				Exit Sub
			End If
		End If

		MyPhotoPath = MyPhotoNames(0).Substring(0, InStrRev(MyPhotoNames(0), "\"))  ' Get the path of the first selected image
		AppendToRTB( "Source folder:" & vbTab & MyPhotoPath & vbCrLf, Color.Black, StndFont )
'
'   Select folder in which reformatted photos are to be stored
'
		Outputfoldername = GetOutputPath
		if OutputFolderName = vbNullString Then
			AppendToRTB ( "Failed to get name of Output Folder, application aborted" & vbCrLf, Color.Red, BoldFont )
			Exit Sub
		End If

		AppendToRTB( "Output folder:" & vbTab & Outputfoldername & vbCrLf, Color.Black, StndFont )
		
		Using  New CentreDialog(Me)
			r =MsgBox( "Convert Photographs to Black and White ?", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "GGN PhotoShop Automation" )
		End Using

		If r = vbYes
			' It is to be a B&W photo so ...
			BlackWhite = True
		Else If r = vbNo
			BlackWhite = False
		Else
			Exit Sub
		End If
'
'	Process the first photograph, converting it to a form suitable for publication
'	After the photograph has been processed, it is closed (in Photoshop)
'				
		OpenOutputFile = True
		r = vbNo
		If PhotoShopApp.Documents.Count > 1 Then
			Using New CentreDialog(Me)
				r =MsgBox("Process all the other currently opened images ?", vbQuestion + vbYesNoCancel + vbMsgBoxSetForeground, "GGN PhotoShop Utility")
			End Using
		End if

		If MyPhotoName <> vbNullString then
			AppendToRTB ( "Re-formatting images open in Photoshop:-" & vbCrLf, Color.Black, StndFont )
			AppendToRTB ( vbTab & PhotoShopApp.ActiveDocument.Name & vbCrLf, Color.Black, StndFont )
			OutputFileName = ProcessPhoto( PhotoShopApp.ActiveDocument, OutputFolderName, MyPhotoName, BlackWhite )
		End if
'
'   Ask if all the rest of the currently opened images are to be similarly processed
'
'   Note: Due to the way in which the Documents Collection is maintained by the PhotShop Application
'   the 'For Loop' below always has to process PhotoShop.Documents(1)
'
		If r = vbYes Then                                            ' Yes, so ...
			OpenOutputFile = False
				
			For i = 1 To PhotoShopApp.Documents.Count
				PhotoShopApp.ActiveDocument = PhotoShopApp.Documents(1)
				AppendToRTB ( vbTab & PhotoShopApp.ActiveDocument.Name & vbCrLf, Color.Black, StndFont )
				OutputFileName = ProcessPhoto(PhotoShopApp.ActiveDocument, OutputFolderName, PhotoShopApp.ActiveDocument.Name, BlackWhite)
			Next
		Else If r = vbCancel then
			Exit Sub
		End If
'
'   Ask if all images in the input folder are to be similarly processed
'
		Using new CentreDialog(Me)
			r = MsgBox("Process all images in the source folder ?", vbQuestion + vbYesNoCancel, "GGN PhotoShop Automation")
		End Using

		If r = vbCancel Then
			Exit Sub
		Else If r = vbYes Then                                            ' Yes, so ...
			OpenOutputFile = False
'
'   Create a FileSystem object in order to work with folders and files, and then link to the input folder
'
			InFolder = New DirectoryInfo(MyPhotoPath)
			InFileList = InFolder.GetFiles()
'
'   Loops over all files in the input folder and, provided the file is a .jpg, .tif, png, bmp, gif, or .psd  file,
'   process it as for the initial file
'
			AppendToRTB ( "Re-formatting images in source folder '" & MyPhotoPath & "'" & vbCrLf, Color.Black, StndFont )
			For each InFileInfo In InFileList
				Select Case LCase( InFileInfo.Name.Substring(InStrRev(InFileInfo.Name, ".")) )
					Case "jpg", "tif", "png", "bmp", "gif", "psd"
						If InFileInfo.Name <> MyPhotoName Then
							Err.Clear
							PhotoShopApp.Open( MyPhotoPath & InFileInfo.Name )
							If Err.Number <> 0 Then
								AppendToRTB ( vbTab & "Unable to open image " & InFileInfo.Name & vbCrLf, Color.Red, BoldFont )
								AppendToRTB ( vbTab & "Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont )
							Else
								AppendToRTB ( vbTab & PhotoShopApp.ActiveDocument.Name & vbCrLf, Color.Black, StndFont )
								OutputFileName = ProcessPhoto(PhotoShopApp.ActiveDocument, OutputFolderName, PhotoShopApp.ActiveDocument.Name, BlackWhite)
							End If
						End If
					Case Else
				End Select
			Next InFileInfo
		End If
'
'   Open the output file if only one image has been processed
'
		If OpenOutputFile _
		Then
			Err.Clear()
			PhotoShopApp.Open ( OutputFolderName & "\" & OutputFileName )                  ' Open the single Tiff image produced
			If Err.Number <> 0 _
			Then
				AppendToRTB ( "Unable to open output file '" & OutputFolderName & "\" &  OutputFileName & "' attempt aborted" & vbCrLf, Color.Red, BoldFont )
				AppendToRTB ( "Error Description: " & Err.Description & vbCrLf,  Color.Red, BoldFont )
			End If
		End If
'
'   Let the user know that the script has completed
'
		AppendToRTB ( "All images have been processed" & vbCrLf, Color.Blue, BoldFont )

		MyPhoto = Nothing
		PhotoShopApp = Nothing

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
	Public Sub AppendToRTB( inText As String,
							inColour As Drawing.Color,
							inFont as Font )

		Call Me.Rtb_Log.SuspendLayout()
		Me.Rtb_Log.SelectionStart = Me.Rtb_Log.TextLength
		Me.Rtb_Log.SelectionLength = 0
		Me.Rtb_Log.SelectionFont = inFont
		Me.Rtb_Log.SelectionColor = inColour
		Me.Rtb_Log.SelectedText = inText
		Call Me.Rtb_Log.ResumeLayout()

	End Sub
'
'************************************************************************************************************
'
'	GetInputPath (Function)
'
'	Called by
'		F_Main.Start_Click
'
'	Returns
'		Null String			- Failed to select a folder
'		Input Folder Path
'
	Private Function GetInputPath() As string

		Dim folderBrowserDialog1 As FolderBrowserDialog

		On Error Resume Next													' Handle errors internally
		folderBrowserDialog1 = New FolderBrowserDialog							' New instance of the File Browser dialobox

		folderBrowserDialog1.RootFolder = Environment.SpecialFolder.Desktop		' Root of the folder tree displayed by the File Browser
		If Not MyPhoto Is Nothing Then
			folderBrowserDialog1.SelectedPath = MyPhotoPath						' Initial selected folder (same as active photo path)
		Else 
			folderBrowserDialog1.SelectedPath = Environment.ExpandEnvironmentVariables("%SYSTEMDRIVE%")
		End If
		folderBrowserDialog1.Description = "Select SOURCE folder"
		folderBrowserDialog1.UseDescriptionForTitle = True						' Use the description as the title of the File Browser
		folderBrowserDialog1.ShowNewFolderButton = true							' Allow the browser to create new folders

		If folderBrowserDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
			Return  vbnullstring												' Failed to select an input folder
		else
			Return folderBrowserDialog1.SelectedPath							' Output folder Path
		End if
	End Function
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
	Private Function GetOutputPath() As string

		Dim folderBrowserDialog1 As FolderBrowserDialog

		On Error Resume Next													' Handle errors internally
		folderBrowserDialog1 = New FolderBrowserDialog							' New instance of the File Browser dialobox

		folderBrowserDialog1.RootFolder = Environment.SpecialFolder.MyComputer	' Root of the folder tree displayed by the File Browser
		folderBrowserDialog1.SelectedPath = MyPhotoPath                         ' Initial selected folder (same as source path)
		folderBrowserDialog1.Description = "Select OUTPUT folder"
		folderBrowserDialog1.UseDescriptionForTitle = True                      ' Use the description as the title of the File Browser
		folderBrowserDialog1.ShowNewFolderButton = true							' Allow the browser to create new folders

		If folderBrowserDialog1.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then
			Return  vbnullstring												' Failed to select an output folder
		else
			Return folderBrowserDialog1.SelectedPath							' Output folder Path
		End if
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
		OutputFileName = OutputFileName.Substring(0,InStrRev(OutputFileName,".")) & "tif"
'
'   Now configure the photo with the required parameters
'
		MyPhoto.Flatten										' Flatten the layers in the image
		MyPhoto.ColorProfileType = 1						' 1 = psNo; Turn ICC/sRGB off
		MyPhoto.ResizeImage( , , 300, 1)					' 1 = psNoResampling; Resize image to 300 ppi without ReSampling
'
' It is to be a B&W photo so ...
'
		If BlackWhite Then
			If MyPhoto.Width > 9 Then
				MyPhoto.ResizeImage (9 * 25 / 6, , , 5)		' (CS2) Resize images to 9cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
	'            MyPhoto.ResizeImage 9, , , 5				' (CS3) Resize images to 9cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
			End If
		
			MyPhoto.ChangeMode(4)							' 4 = psConvertToLab; convert colour profile to Lab mode
		
			MyPhoto.Channels("a").Delete					' Delete the "a" channel (this also deletes the "b"  channel)
			MyPhoto.Channels("Alpha 2").Delete				' Delete the "Alpha 2" channel leaving only the "Alpha 1" channel
		
			MyPhoto.ChangeMode (1)							' 1 = psConvertToGrayscale; convert colour profile to Grayscale
'
'   Suffix the filename of a B&W image with "_BW"
'
			i = InStrRev(OutputFileName, ".")				' Position (if any) of "." in the Output File Name
			OutputFileName = OutputFileName.Substring(OutputFileName.Length - i) & "_BW" & OutputFileName.Substring(i+1,OutputFileName.Length - i)
'
'	It's a colour photo so ...
'	
		Else   
			If MyPhoto.Width > 19 Then
				MyPhoto.ResizeImage (19 * 25 / 6, , , 5)    ' (CS2) Resize images to 16cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
'	            MyPhoto.ResizeImage (19, , , 5)				' (CS3) Resize images to 16cm wide using BiCubicSharper resampling (5 = psBicubicSharper)
			End If
		
			MyPhoto.ChangeMode (3)                          ' 3 = psConvertToCMYK; convert colour profile to CMYK
		End If
'
'   Save the newly configured Tiff image
'
		Err.Clear
		TiffOptions = CreateObject( "Photoshop.TIFFSaveOptions" )
		If Err.Number <> 0 Then
			AppendToRTB ( "Cannot save output file '" & OutputFolderName & "\" & OutputFileName & "' as a TIFF, attempt aborted" & vbCrLf, Color.Red, BoldFont )
			AppendToRTB ( "Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont )
		Else
			TiffOptions.ImageCompression = 2				' 2 = psTiffLZW; LZW encoding (compression)
		
			Err.Clear
			MyPhoto.SaveAs( OutputFolderName & "\" & OutputFileName, TiffOptions, True, 2 )
			If Err.Number <> 0 Then
			AppendToRTB ( "Cannot save output file '" & OutputFolderName & "\" & OutputFileName & "' as a TIFF, attempt aborted" & vbCrLf, Color.Red, BoldFont )
			AppendToRTB ( "Error Description: " & Err.Description & vbCrLf, Color.Red, BoldFont )
			End If
		End if
'
'   Now close the source image without modifying it
'
		MyPhoto.Close (2)									' 2 = psDoNotSaveChanges; Close without saving changes to Jpeg image
	
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
	Sub ResizePhotoshopWindow( ByVal WindowHandle As IntPtr  )

		Dim b As Boolean
'
'	Make the Photoshop window occupy the right-hand side of the monitor
'
		b = SetWindowPos ( WindowHandle, 1, _									' Bottom of the Z-order
							Screen.PrimaryScreen.WorkingArea.Width/2,			' Left = Halfway across the Monitor
							1,													' Top = Top of Monitor
							Screen.PrimaryScreen.WorkingArea.Width/2,			' Width = Half the Monitor's width
							Screen.PrimaryScreen.WorkingArea.Height-1,			' Height = Full monitor height
							SWP_NOACTIVATE )									' Do not Activate the Window
		If Not b Then
			AppendToRTB ( "Failed to re-size the Photoshop window" &vbCrLf, Color.Red, BoldFont )
			Exit sub
		End If

		b = ShowWindow ( WindowHandle, SW_RESTORE )								' Restore (Normalise) the window
		If Not b Then
			AppendToRTB ( "Failed to Restore the Photoshop window" &vbCrLf, Color.Red, BoldFont )
			Exit sub
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

		openFileDialog1.InitialDirectory = "c:\"
		openFileDialog1.Filter = "Images files |*.jpg;*.png;*.tif;*.bmp;|All files (*.*)|*.*"
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
					MsgBox( "User cancelled", vbOK + vbInformation + vbMsgBoxSetForeground, "GGN PhotoShop Utility" )
				End Using
				Return Nothing
			End If
		Catch ex As Exception
			Using New CentreDialog(Me)
				MsgBox( "Unable to display the File Open dialogbox" + Environment.NewLine + Environment.NewLine + ex.Message, 
							vbOK + vbCritical + vbMsgBoxSetForeground, "GGN PhotoShop Utility" )
			End Using
			Return Nothing
		End Try

	End Function 

End Class
