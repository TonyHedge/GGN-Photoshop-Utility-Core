'
'   Class CentreDialog
'
'   GGN Photoshop Utility
'
'   C# version written by Hans Passant (see https:'stackoverflow.com/questions/2576156/winforms-how-can-i-make-messagebox-appear-centered-on-mainform )
'   and converted from C# to VB
'
'  Called by
'	    Any code that wishes to display a DialogBox
'
'   Contains
'      checkWindow
'      findDialog
'
'   Change History:-
'   v1.0        20.02.25	Initial version
'
'   This class displays the DialogBox centred in the middle of the form that call it.
'   (Normally such DialogBoxes are displayed centred on the screen, or whereever .NET decides it wishes to display them)
'
Imports System
Imports System.Text

#Disable Warning CA1416

Public Class CentreDialog

	Implements IDisposable

	Dim mTries As Integer = 0
	Dim mOwner As Form

	public Sub New( ByVal owner As Form )

		mOwner = owner
		owner.BeginInvoke(new MethodInvoker(AddressOf findDialog))

	End Sub

	Public Sub dispose() Implements IDisposable.dispose
		mTries = -1
	End sub
'
'*********************************************************************************************************************
'
'  findDialog
'
	Public Sub findDialog() 

		Dim callBack as EnumThreadDelegate = new EnumThreadDelegate(AddressOf checkWindow)

		Trace.WriteLine( $"CentreDialog.findDialog - Enter: Number of tries {mtries}")

		' Enumerate windows to find the message box
		if mTries < 0 Then Return

		if (EnumThreadWindows(GetCurrentThreadId(), callBack, IntPtr.Zero)) Then
			if ++mTries < 10 Then
				mOwner.BeginInvoke(new MethodInvoker(AddressOf findDialog))
			End If
		End If

	End Sub
'
'*********************************************************************************************************************
'
'  checkWindow
'
	Private Function checkWindow( ByVal hWnd as IntPtr, ByVal lp As IntPtr ) As Boolean

		Trace.WriteLine( "CentreDialog.checkWindow - Enter")
		
		Dim frmRect As Rectangle = new Rectangle( mOwner.Location, mOwner.Size )
		Dim dlgRect As RECT
		Dim sb As New StringBuilder("", 260)

		' Checks if <hWnd> is a dialog
		GetClassName( hWnd, sb, 260 )

		if sb.ToString <> "#32770" Then return True					' Got it

		GetWindowRect( hWnd, dlgRect )

		MoveWindow( hWnd,
					frmRect.Left + (frmRect.Width - dlgRect.Right + dlgRect.Left) / 2,
					frmRect.Top + (frmRect.Height - dlgRect.Bottom + dlgRect.Top) / 2,
					dlgRect.Right - dlgRect.Left,
					dlgRect.Bottom - dlgRect.Top, true )

		return false

	End Function
'
'*********************************************************************************************************************
'
' P/Invoke declarations
'
		Delegate Function EnumThreadDelegate ( ByVal hWnd As IntPtr, ByVal lParam As IntPtr ) As Boolean

		private declare Function EnumThreadWindows Lib "user32" ( dwThreadId as Int32, lpfn as EnumThreadDelegate, lParam as IntPtr ) as Boolean 
		
		private declare Function GetCurrentThreadId Lib "kernel32" () As UInteger
		
		private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As System.Text.StringBuilder, ByVal nMaxCount As Long) As Long        
		
		private declare Function GetWindowRect Lib "user32" ( ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
		
		private Declare Function MoveWindow Lib "user32" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean

		Public Structure RECT
			Public left As Integer
			Public top As Integer
			Public right As Integer
			Public bottom As Integer
		End Structure

End Class
