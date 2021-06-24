Option Strict Off
Option Explicit On
Friend Class frmDemo
	Inherits System.Windows.Forms.Form
	
	
	
	Const frmDemo_decl_end As Boolean = True
	
	
	Sub frmDemo_GO()
		Me.ShowDialog()
	End Sub
	
	
	Private Sub cmdButton1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdButton1.Click
		Me.Close()
		Exit Sub
	End Sub
	Private Sub cmdExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdExit.Click
		End
	End Sub
	
	
	Private Sub frmDemo_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		Call CenterOnScreen(Me)
		lblDisclaimer.Text = "This is a DEMONSTRATION version of the AdDesignS program. " & "This demonstration version may only load and simulate the " & "LIQUID.DAT and GAS.DAT files in the examples subdirectory. " & "For the full version of this program, please contact " & "Dr. David W. Hand (dwhand@mtu.edu or 906-487-2777). " & "Additional information about this program is available at " & "our web site (http://www.cpas.mtu.edu/etdot/)."
	End Sub
End Class