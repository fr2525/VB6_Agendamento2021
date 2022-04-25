VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private WithEvents pCombo As ComboBoxPrivate
'IsDelOrBack As Boolean
'
'Public Sub Bind(Cmb As ComboBox)
'    Set pCombo = Cmb
'End Sub
'
'Private Sub Class_Terminate()
'    Set pCombo = Nothing
'End Sub
'Private Sub pCombo_Change()
'    On Error Resume Next
'    Dim OldLen As Integer
'    Dim i As Integer
'    If Not pCombo.Text = "" And Not IsDelOrBack Then
'       OldLen = Len(pCombo.Text)
'       For i = 0 To pCombo.ListCount - 1
'           If InStr(1, Mid$(UCase(pCombo.List(i)), 1, OldLen), UCase(pCombo.Text)) Then
'               pCombo.Text = pCombo.List(i)
'               If pCombo.SelText = "" Then
'                   pCombo.SelStart = OldLen
'                Else
'                    pCombo.SelStart = InStr(pCombo.Text, pCombo.SelText)
'                End If
'                pCombo.SelLength = Len(pCombo.Text)
'                Exit For
'            End If
'       Next
'    End If
'End Sub
'Private Sub pCombo_KeyDown(KeyCode As Integer, Shift As Integer)
'    IsDelOrBack = (KeyCode = vbKeyBack Or KeyCode = vbKeyDelete)
'End Sub
