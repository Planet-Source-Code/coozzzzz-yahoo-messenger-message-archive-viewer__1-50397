VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_Main 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahoo! Messenger Archive Viewer"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   5400
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   5415
      Left            =   4320
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9551
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      OLEDropMode     =   1
      TextRTF         =   $"frm_Main.frx":000C
   End
   Begin MSComctlLib.TreeView Profiles 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10186
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      BorderStyle     =   1
      Appearance      =   1
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is basically just an unfinished project I started some time ago. It had more
'features and no comments and since I didn't plan on working on it any further, I
'decided to comment it the best as I could from what I remember (from months ago)
'and add it to PSC. All my comments may not be fully correct as it's 9am and I
'haven't slept yet, but they should be close enough. I've removed the support for
'stripping HTML/ANSI from the messages as well as searching through the RichTextBox
'as I see no direct need for such. Nothing else to really say on it except ignore
'the bad methods i've used for doing this. :)

'-Coozzzzz

Option Explicit

'Using variables instead of constants in case you wish to change them at run-time
Private lngColorSent As Long
Private lngColorRecv As Long
Private strFolder As String

Private bLoading As Boolean
Private Sub Form_Load()
    'Assign colors of messages
    lngColorSent = &HC00000
    lngColorRecv = &HC0&
    'Assign default profiles folder
    strFolder = "C:\Program Files\Yahoo!\Messenger\Profiles\"
    'Load profiles into TreeView
    Call LoadProfiles(strFolder)
End Sub
Private Sub LoadProfiles(ByVal strPath As String)
    'Recursively search through files/directories in "strPath"
    Dim cProf1 As New Collection, cProf2 As New Collection, cProf3 As New Collection
    Dim cItem1 As Variant, cItem2 As Variant, cItem3 As Variant
    Set cProf1 = DirSearch(strPath)
    For Each cItem1 In cProf1
        Call Profiles.Nodes.Add(, , cItem1, cItem1)
        Set cProf2 = DirSearch(strPath & cItem1 & "\Archive\Messages\")
        For Each cItem2 In cProf2
            Call Profiles.Nodes.Add(cItem1, tvwChild, cItem1 & "," & cItem2, cItem2)
            Set cProf3 = DirSearch(strPath & cItem1 & "\Archive\Messages\" & cItem2 & "\")
            For Each cItem3 In cProf3
                Call Profiles.Nodes.Add(cItem1 & "," & cItem2, tvwChild, cItem1 & "," & cItem2 & "," & cItem3, cItem3)
            Next cItem3
        Next cItem2
    Next cItem1
End Sub
Private Function DirSearch(ByVal strDir As String) As Collection
    'Add all files and directories to the function's returning collection
    Set DirSearch = New Collection
    strDir = Dir(strDir, vbDirectory + vbNormal)
    Do Until strDir = ""
        If strDir <> "." And strDir <> ".." Then Call DirSearch.Add(strDir)
        strDir = Dir()
    Loop
End Function
Private Sub LoadMessages(ByVal strUsername As String, ByVal strFile As String)
    Dim intFreeFile As Integer, strBuffer As String
    Dim strArray() As String, lngCount As Long
    Dim intLen As Integer, strMessage As String, boolSent As Boolean
    Dim intLoop As Integer, intCount As Integer
    'If the file exists...
    If Dir(strFile, vbNormal) <> "" Then
        'Set loading variable to true. This is used for when users attempt to load
        'another archive while one is already loading, it won't let them.
        bLoading = True
        'Load file contents into strBuffer for processing
        intFreeFile = FreeFile
        Open strFile For Binary As intFreeFile
            strBuffer = Space(LOF(intFreeFile))
            Get #intFreeFile, 1, strBuffer
        Close intFreeFile
        'Split file contents into array to process. Best method known to me at the
        'time was by splitting with delimeter of 3 null bytes
        strArray = Split(strBuffer, String(3, vbNullChar))
        'Update our progress bar to prepare for processing
        Progress.Min = 0
        Progress.Max = UBound(strArray)
        'Clear out RichTextBox for new messages
        Messages.Text = vbNullString
        'Loop through every index in our array
        Do Until lngCount = UBound(strArray) + 1
            'If the current array item has an item after it and the current item
            'contains data then..
            If lngCount + 1 <= UBound(strArray) And Len(strArray(lngCount)) > 0 Then
                'After testing it seems I can decide sent from received messages
                'by the length of the item before the actual XOR'd message so we'll
                'set our boolean variable appropriately
                If Len(strArray(lngCount)) = 1 Then
                    boolSent = False
                ElseIf Len(strArray(lngCount)) = 2 Then
                    boolSent = True
                End If
                'After testing it seems the length of the XOR'd message is within
                'the previous item (item that tells us sent/recv) as the ASCII
                'value so we'll strip the null bytes and set our intLen variable
                'appropriately.
                intLen = Asc(Replace(strArray(lngCount), vbNullChar, ""))
                'If the length of the string is greater than 1 then..
                If intLen > 1 Then
                    'If our intLen is the same as the length of the next item (which
                    'is our XOR'd message) then we "have a winner!" :)
                    If intLen = Len(strArray(lngCount + 1)) Then
                        'Set our intCount to 1 (the start of our username) so we can
                        'begin XOR'ing the message.
                        intCount = 1
                        'Start our loop from the start and end of our XOR'd message in
                        'our array
                        For intLoop = 1 To Len(strArray(lngCount + 1))
                            'Replace the current char with the XOR'd char.
                            Mid(strArray(lngCount + 1), intLoop, 1) = Chr(Asc(Mid(strArray(lngCount + 1), intLoop, 1)) Xor Asc(Mid(strUsername, intCount, 1)))
                            'Increase our count so it'll use the next char in the username
                            'to XOR the next char in the XOR'd message
                            intCount = intCount + 1
                            'If our count is beyond the length of our username, set the
                            'count at 1 so it will reuse the username to XOR
                            If intCount > Len(strUsername) Then intCount = 1
                        Next intLoop
                        'Our XOR'd message is now XOR'd back to the original message so we'll
                        'add it to the RichTextBox
                        With Messages
                            .SelStart = Len(.Text)
                            .SelColor = IIf(boolSent = True, lngColorSent, lngColorRecv)
                            .SelText = strArray(lngCount + 1) & vbCrLf & vbCrLf
                            .SelStart = Len(.Text)
                        End With
                    End If
                End If
            End If
            'Update ProgressBar to reflect just-processed items
            Progress.Value = lngCount
            'Increment our array index count
            lngCount = lngCount + 1
            'Free up processing of other window-messages
            DoEvents
        Loop
        'Set our boolean to false so the user can now load other archives
        bLoading = False
    End If
End Sub
Private Sub Messages_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'Decided to keep this, for drag&drop, no need to explain
    Dim strUsername As String
    If bLoading = False Then
        If LCase(Right(Data.Files.Item(1), 3)) = "dat" Then
            strUsername = InputBox("Enter the Yahoo! Username which this archive file belongs to (incorrect name will result in incorrect decoding):", "Drag-Drop Support")
            If Len(strUsername) > 0 Then Call LoadMessages(strUsername, Data.Files.Item(1))
        End If
    End If
End Sub
Private Sub Profiles_NodeClick(ByVal Node As MSComctlLib.Node)
    'For browsing the TreeView.. no need to explain
    Dim strArray() As String
    If bLoading = False Then
        strArray = Split(Node.Key, ",")
        If UBound(strArray) = 2 Then
            Call LoadMessages(strArray(0), strFolder & strArray(0) & "\Archive\Messages\" & strArray(1) & "\" & strArray(2))
        End If
    Else
        Call MsgBox("Please wait for current messages to finish loading.")
    End If
End Sub
