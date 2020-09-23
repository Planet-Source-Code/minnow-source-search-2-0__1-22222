Attribute VB_Name = "Resize"
Option Explicit

'This function handles resizing the elements of the form.
'You can figure out what it does for yourself :)
Public Function ResizefrmMain()

  With frmMain
    If .WindowState = vbMinimized Then Exit Function
    
    If .Width < 5000 Then .Width = 5000
    If .Height < 5000 Then .Height = 5000
    
    .TabStrip1.Width = frmMain.Width - 240
    .TabStrip1.Left = 60
    .TabStrip1.Height = frmMain.Height - 900
    
'******************
      .fraSearchList.Width = .TabStrip1.Width - 240
      .fraSearchList.Left = .TabStrip1.Left + 120
      .fraSearchList.Height = .TabStrip1.Height - 480
      .fraSearchList.Top = 480
      
        .fraSLList.Width = .fraSearchList.Width - 240
        .fraSLList.Height = .fraSearchList.Height - (.fraSLSingle.Height + 240)
        
          .PL.Width = .fraSLList.Width - 240
          .PL.Height = .fraSLList.Height - (.cmdRemove.Height + 480)
          
          .cmdRemove.Left = 120 + .PL.Width - .cmdRemove.Width
          .cmdRemove.Top = .PL.Height + 360
          
          .cmdUpdate.Left = .cmdRemove.Left - (240 + .cmdUpdate.Width)
          .cmdUpdate.Top = .cmdRemove.Top
          
          .cmdAdd.Left = .cmdUpdate.Left - (240 + .cmdAdd.Width)
          .cmdAdd.Top = .cmdRemove.Top
        
        .fraSLSingle.Width = .fraSLList.Width
        .fraSLSingle.Top = .fraSLList.Height + 120
        
          .cmbSeverity.Left = .fraSLSingle.Width - (120 + .cmbSeverity.Width)
          .lblSeverity.Left = .cmbSeverity.Left
          .txtPhrase.Width = .cmbSeverity.Left - 240
          .txtDescription.Width = .fraSLSingle.Width - 240
          
'******************
      .fraProjectList.Width = .TabStrip1.Width - 240
      .fraProjectList.Left = .TabStrip1.Left + 120
      .fraProjectList.Height = .TabStrip1.Height - 480
      .fraProjectList.Top = 480
          
        If .imgCustomsize.Left < 1000 Then .imgCustomsize.Left = 1000
        If .imgCustomsize.Left > .fraProjectList.Width - (.lblSearchingStatus.Width + 500) Then .imgCustomsize.Left = .fraProjectList.Width - (.lblSearchingStatus.Width + 500)
        .imgCustomsize.Height = .fraProjectList.Height - 480
        .PT.Height = .imgCustomsize.Height
        .PT.Width = .imgCustomsize.Left - 160
        
        .SL.Width = .fraProjectList.Width - (.PT.Width + 420)
        .SL.Left = .PT.Width + 260
        .SL.Height = .imgCustomsize.Height - 600
        
        .picProgressContainer.Left = .SL.Left
        .picProgressContainer.Top = .SL.Top + .SL.Height + 348
        .picProgressContainer.Width = .SL.Width
        
        .lblSearchingStatus.Top = .picProgressContainer.Top - 280
        .lblSearchingStatus.Left = .picProgressContainer.Left + (.picProgressContainer.Width - .lblSearchingStatus.Width) / 2
        
        .picProgress.Width = Int(.picProgressContainer.Width * SearchPercent) + 99
        
    
  End With

End Function
