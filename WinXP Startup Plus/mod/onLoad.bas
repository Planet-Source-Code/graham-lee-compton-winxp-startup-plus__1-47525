Attribute VB_Name = "onLoad"
Option Explicit
Dim sMe, sConfig As String
Public Z As Integer
Public sLogoDir As String
Sub Main()
 sMe = App.Path
 sConfig = sMe & "\XpS+.sys\"
 If Dir(sConfig) = Null Then
  ' Missing Config
  Else
  If Dir(sConfig & "Settings.sys") = Null Then
   ' Missing Settings
   Else
   Dim sFile
   Dim sTemp As String
   sFile = FreeFile
   Open sConfig & "Settings.sys" For Input As #sFile
    Do While EOF(sFile) = False
     Input #sFile, sTemp
     With frmMain
      If InStr(1, sTemp, "logo=") Then
       sLogoDir = GetValue("logo", sTemp)
       If InStr(1, sTemp, "$me") Then sTemp = Replace(sTemp, "$me", sMe, 1, -1)
       If Not Dir(GetValue("logo", sTemp)) = "" Then
        frmAbout.imgLogo.Picture = LoadPicture(GetValue("logo", sTemp))
       End If
      End If
      If InStr(1, sTemp, "name=") Then .lst.ColumnHeaders.Item(1).Width = GetValue("name", sTemp)
      If InStr(1, sTemp, "style=") Then .lst.ColumnHeaders.Item(2).Width = GetValue("style", sTemp)
      If InStr(1, sTemp, "method=") Then .lst.ColumnHeaders.Item(3).Width = GetValue("method", sTemp)
      If InStr(1, sTemp, "path=") Then .lst.ColumnHeaders.Item(4).Width = GetValue("path", sTemp)
      If InStr(1, sTemp, "height=") Then .Height = GetValue("height", sTemp)
      If InStr(1, sTemp, "width=") Then .Width = GetValue("width", sTemp)
      If InStr(1, sTemp, "left=") Then .Left = GetValue("left", sTemp)
      If InStr(1, sTemp, "top=") Then .Top = GetValue("top", sTemp)
      If InStr(1, sTemp, "settings=") Then
       Dim sSetup As String
       sSetup = GetValue("settings", sTemp)
       If Len(sSetup) >= 4 Then
        If Mid(sSetup, 1, 1) = 1 Then .mnuOnConn.Checked = True
        If Mid(sSetup, 2, 1) = 1 Then .mnuMinimise.Checked = True
        If Mid(sSetup, 3, 1) = 1 Then
         .mnuAdd.Checked = True
         .pAdd.Visible = False
         .Refresh
        End If
        If Mid(sSetup, 4, 1) = 1 Then .mnuAlpha.Checked = True
       End If
      End If
      .lst.Height = .ScaleHeight
'      .lst.Width = .ScaleWidth '?
     End With
    Loop
   Close #sFile
  End If
  If Dir(sConfig & "files.sys") = Null Then
   ' Missing Files
   Else
   sFile = FreeFile
   Open sConfig & "files.sys" For Input As #sFile
    Do While EOF(sFile) = False
     Input #sFile, sTemp
     If InStr(1, sTemp, "$me") Then sTemp = Replace(sTemp, "$me", sMe, 1, -1)
     With frmMain
     If Not sTemp = "" Then
      .lst.ListItems.Add , , Mid(sTemp, 5, InStr(1, sTemp, "=") - 5)
      .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Add , , ""
      .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Add , , ""
      .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Add , , Mid(sTemp, InStr(1, sTemp, "=") + 1, Len(sTemp) - InStr(1, sTemp, "=") + 1)
      ' Tick?
      Select Case Mid(sTemp, 1, 1)
      Case 0
       .lst.ListItems.Item(.lst.ListItems.Count).Checked = False
      Case 1
       .lst.ListItems.Item(.lst.ListItems.Count).Checked = True
      End Select
      ' Style?
      Select Case Mid(sTemp, 2, 1)
      Case 1
       .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(1).Text = "Normal"
      Case 2
       .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(1).Text = "Hidden"
      End Select
      ' Method
      Select Case Mid(sTemp, 3, 1)
      Case 1
       .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(2).Text = "onStart"
      '###########################################
      'Check for ticked and run it
      If .lst.ListItems.Item(.lst.ListItems.Count).Checked = True Then
       If Not Dir(.lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(3)) = "" Then
          If .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(1).Text = "Normal" Then
           Shell .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(3), vbNormalFocus
          End If
          If .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(1).Text = "Hidden" Then
           Shell .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(3), vbHide
          End If
        End If
      End If
      '###########################################
      Case 2
       .lst.ListItems.Item(.lst.ListItems.Count).ListSubItems.Item(2).Text = "onConn"
      End Select
      End If
     End With
    Loop
   Close #sFile
  End If
  If frmMain.mnuAlpha.Checked = True And frmMain.mnuMinimise.Checked = False Then
   Call MakeTransparent(frmMain.hwnd, "0")
    frmMain.Show
    For Z = 0 To 255 Step 5
     Call MakeTransparent(frmMain.hwnd, Z)
    Next Z
   Else
   If frmMain.mnuMinimise.Checked = True Then
     frmMain.Show
     frmMain.WindowState = vbMinimized
    Else
     frmMain.Show
   End If
  End If
 End If
End Sub

Public Function GetValue(sText As String, sData As String) As String
 If InStr(1, sData, sText & "=") Then
  GetValue = Mid(sData, InStr(1, sData, sData & "=") + Len(sText & "=") + 1, Len(sData) - InStr(1, sData, sText & "="))
 End If
End Function

Public Function SetBuffer(sBuffer As String, lSize As Long) As String
 SetBuffer = Space(lSize)
End Function

Public Function WriteSettings()
Dim sFile
Dim sT As String
Dim sS As String
With frmMain
 If .mnuOnConn.Checked = True Then sS = sS & 1 Else sS = sS & 0
 If .mnuMinimise.Checked = True Then sS = sS & 1 Else sS = sS & 0
 If .mnuAdd.Checked = True Then sS = sS & 1 Else sS = sS & 0
 If .mnuAlpha.Checked = True Then sS = sS & 1 Else sS = sS & 0
 sT = sT & "logo=" & sLogoDir & vbCrLf
 sT = sT & "name=" & .lst.ColumnHeaders.Item(1).Width & vbCrLf
 sT = sT & "style=" & .lst.ColumnHeaders.Item(2).Width & vbCrLf
 sT = sT & "method=" & .lst.ColumnHeaders.Item(3).Width & vbCrLf
 sT = sT & "path=" & .lst.ColumnHeaders.Item(4).Width & vbCrLf
 sT = sT & "height=" & .Height & vbCrLf
 sT = sT & "width=" & .Width & vbCrLf
 sT = sT & "left=" & .Left & vbCrLf
 sT = sT & "top=" & .Top & vbCrLf
 sT = sT & "settings=" & sS & vbCrLf
End With
sFile = FreeFile
Open sConfig & "Settings.sys" For Output As #sFile
 Print #sFile, sT
Close #1
With frmMain
 sT = ""
 For Z = 1 To .lst.ListItems.Count
  If .lst.ListItems.Item(Z).Checked = False Then sT = sT & "0"
  If .lst.ListItems.Item(Z).Checked = True Then sT = sT & "1"
  If .lst.ListItems.Item(Z).ListSubItems.Item(1) = "Normal" Then sT = sT & "1"
  If .lst.ListItems.Item(Z).ListSubItems.Item(1) = "Hidden" Then sT = sT & "2"
  If .lst.ListItems.Item(Z).ListSubItems.Item(2) = "onStart" Then sT = sT & "1"
  If .lst.ListItems.Item(Z).ListSubItems.Item(2) = "onConn" Then sT = sT & "2"
  sT = sT & "."
  sT = sT & .lst.ListItems.Item(Z).Text & "="
  sT = sT & .lst.ListItems.Item(Z).ListSubItems.Item(3).Text & vbCrLf
 Next Z
End With
 If Not sT = "" Then
  sFile = FreeFile
  Open sConfig & "files.sys" For Output As #sFile
   Print #sFile, sT
  Close #1
 End If
End Function
