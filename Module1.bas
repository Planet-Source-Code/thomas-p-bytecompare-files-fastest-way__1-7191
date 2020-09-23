Attribute VB_Name = "Module1"
Public rvFileName As String ' // returned file name
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

' // Type declarations
Public Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
    
Public Function Open_File(hWnd As Long) As String
   '
   Dim OpenFileDialog As OPENFILENAME
   Dim rv As Long
   
   ' // init dialog
   With OpenFileDialog
     .lStructSize = Len(OpenFileDialog)
     .hwndOwner = hWnd&
     .hInstance = App.hInstance
     .lpstrFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
     .lpstrFile = Space$(254)
     .nMaxFile = 255
     .lpstrFileTitle = Space$(254)
     .nMaxFileTitle = 255
     .lpstrInitialDir = CurDir
     .lpstrTitle = "Open File..."
     .flags = 0
   End With
  
   ' // call API to show the dialog that was just initialized
   rv& = GetOpenFileName(OpenFileDialog)
   
   If (rv&) Then
      Open_File = Trim$(OpenFileDialog.lpstrFile)
   Else
      Open_File = ""
   End If
   
End Function


Public Function ByteCompare(OriginalFile As String, PatchedFile As String, OffsetBox As ListBox, OriginalByte As ListBox, ChangedByte As ListBox)
Dim of() As Byte 'This Array will contain the whole Original File
Dim pf() As Byte 'This Array will contain the whole Changed File
On Error Resume Next 'Better than a SEH :p
ofsize = FileLen(OriginalFile) 'File Length in Byte
pfsize = FileLen(PatchedFile) 'File Length in Byte
If ofsize <> pfsize Then Exit Function 'You can remove this line, will allow you to compare 2 files with different file size
ReDim of(1 To ofsize) 'Dimension array to the exact filesize
ReDim pf(1 To pfsize) 'Dimension array to the exact filesize
DoEvents
'Make the ListBoxes invisible (faster)
OffsetBox.Visible = False
OriginalByte.Visible = False
ChangedByte.Visible = False
'Clear the ListBoxes
OffsetBox.Clear
OriginalByte.Clear
ChangedByte.Clear
'Open the Files in Binary Mode
Open OriginalFile For Binary As #1
Open PatchedFile For Binary As #2
Get #1, 1, of() 'Read the whole Original File into our ByteArray
Get #2, 1, pf() 'Read the whole Changed File into our ByteArray
'Close em again
Close #1
Close #2
For lv = 1 To ofsize Step 3 'Without Step its hella slow (you can increase the steps but dont forget to increas the code below too (you can add another loop too..)
        If of(lv) <> pf(lv) Then
        OffsetBox.AddItem Hex(lv)
        OriginalByte.AddItem Hex(of(lv))
        ChangedByte.AddItem Hex(pf(lv))
        End If
        If of(lv + 1) <> pf(lv + 1) Then
        OffsetBox.AddItem Hex(lv + 1)
        OriginalByte.AddItem Hex(of(lv + 1))
        ChangedByte.AddItem Hex(pf(lv + 1))
        End If
        If of(lv + 2) <> pf(lv + 2) Then
        OffsetBox.AddItem Hex(lv + 2)
        OriginalByte.AddItem Hex(of(lv + 2))
        ChangedByte.AddItem Hex(pf(lv + 2))
        End If
Next lv
'Make the ListBoxes visible again
OffsetBox.Visible = True
OriginalByte.Visible = True
ChangedByte.Visible = True
End Function
