Attribute VB_Name = "Main_"
Public Money As Integer
Public StartTime As Date

Private Sub Main()
Money = 100
StartTime = Time
frmMain.Show
frmAbout.Show vbModal, frmMain
End Sub

