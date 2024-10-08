VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FolderSelectBox 
   Caption         =   "Select Folder for Filing"
   ClientHeight    =   5520
   ClientLeft      =   15
   ClientTop       =   -45
   ClientWidth     =   10740
   OleObjectBlob   =   "FolderSelectBox.frx":0000
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FolderSelectBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Show a list of folders taken from the locations of a selected set of emails and/or conversations.
' Also shows a FULL list of all folders with a dynamic filter box, type some chars to filter the full list
' If a folder is selected (preference given to the select folder list), file all of the selected emails to that folder.
'
' Facilitates really easy filing once you have put one email in a conversation to a particular folder,
' you can easily file all of the other related emails without having to hunt through all of your
' folders. Just make sure you use a view with "Conversation View" turned on.
'
' Also has a View button to switch the current view in outlook to the selected folder instead of filing.
'
' Double-click on on a folder in either list is the same as pressing the "File" Button.
'
' WARNINGS: Assumed English folder names for exclusion of Inbox, Sent, etc. - special characters such as "/" are now supported.
'
' TO DO:
'   1) Add ability to move to another mailbox (http://www.slipstick.com/developer/working-vba-nondefault-outlook-folders/)
'   2) Allow multiple filters for full folder list
'
' Current Author: Corey Blakeborough
' Original Author: Julian Knight (Totally Information)
' Version: v2.2 2020-09-18
' History:
'   v2.2 2020-09-21 - Removal of extraneous fields/functionality. Bugfixes for selection, nested folders, and individual messages
'   v2.0 2020-09-18 - Forked project. Conversation support. Fixed issue with special characters.
'   v1.4 2015-06-29 - Chg default location to Win default. Add auto-selects. Add recents list (not yet working)
'   v1.3 2015-06-12 - Add copy link to clipboard after moving
'   v1.2 2015-05-18 - Add double-click processing
'   v1.1 2015-05-12 - Various improvements - add filter to full folder list, add view button
'   v1.0 2015-05-08 - Initial Release

Option Explicit

' Define form global variables
Dim folderNames() As String
Dim maxNames As Long
Dim folderPaths() As String
Dim maxPaths As Long
Dim folderAllPaths() As Variant
Dim maxFAP As Long
Dim folderAllNames() As String
Dim maxFAN As Long
Dim mailbox As String
Dim changingFolders As Boolean
Dim parsedConversations() As String

Private Sub btnCancel_Click()
    ' Do nothing other than cancel everything
    Unload Me
End Sub

' Only change the current view to the selected folder
Private Sub btnView_Click()
    Dim fldr As Outlook.MAPIFolder
    
    Set fldr = fldrDest
    ' If anywhere selected, change the explorer view now
    If IsObject(fldr) Then
        Set Application.ActiveExplorer.CurrentFolder = fldr
    End If
    
    ' End
    Set fldr = Nothing
    Unload Me
End Sub

' Do the move
Private Sub btnFileToFolder_Click()
    Dim fldr As Outlook.MAPIFolder
    Dim objItem As Variant
    Dim objConvHeader As Outlook.ConversationHeader
    Dim X As Long
    
    Set fldr = fldrDest
    On Error GoTo err
    ' If anywhere to move to, move each email now
    If IsObject(fldr) Then
        ' Start with any conversations
        Dim activeSelection, convSelection As Selection
        Set activeSelection = ActiveExplorer.Selection
        Set convSelection = activeSelection.GetSelection(olConversationHeaders)
        
        If convSelection.Count > 0 Then
            For Each objConvHeader In convSelection
                ' Cache conversation ID
                If IsInArray(parsedConversations, objConvHeader.ConversationID) = False Then
                    AddToArray parsedConversations, objConvHeader.ConversationID
            
                    For Each objItem In objConvHeader.GetItems ' Items in the conversation.
                        If TypeName(objItem) = "MailItem" Or TypeName(objItem) = "AppointmentItem" Or TypeName(objItem) = "MeetingItem" Then
                            ' Only move items not already in the dest folder
                            If objItem.Parent.Name <> fldr.Name Then
                                objItem.Move fldr
                                X = X + 1
                            End If
                        End If
                    Next
                End If
            Next objConvHeader
        End If
        
        objItem = Empty
        
        ' Now add any selected items that weren't part of those conversations
        X = 0
        For Each objItem In activeSelection
            If TypeName(objItem) = "MailItem" Or TypeName(objItem) = "AppointmentItem" Or TypeName(objItem) = "MeetingItem" Then
                If IsInArray(parsedConversations, objItem.ConversationID) = False And objItem.Parent.Name <> fldr.Name Then
                    objItem.Move fldr
                    X = X + 1
                End If
            End If
        Next objItem
    End If
    
    GoTo endit
err:
    MsgBox "Error processing selection, something odd selected?", vbCritical, "Folder Move Error"
    ' End
endit:
    On Error GoTo 0
    Set fldr = Nothing
    Set objItem = Nothing
    Unload Me
End Sub

Private Function fldrDest() As Outlook.MAPIFolder
    Dim obj As Object
    Dim destIdx, arr, e, i As Integer
    Dim X
    Dim pos, val
    
    ' Index = -1 if nothing selected
    If lstFolders.ListIndex > -1 Then
        Set fldrDest = ReturnDestinationFolder(folderPaths(lstFolders.ListIndex), _
            Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders _
        )
        ' set i=1 as we can only ever select one entry on this side
        i = 1
        
    ElseIf lstAllFolders.ListIndex > -1 Then
        
        i = 0
        For e = LBound(folderAllNames) To UBound(folderAllNames)
            If lstAllFolders.List(lstAllFolders.ListIndex) = folderAllNames(e) Then
                i = i + 1
                destIdx = e
            End If
        Next e
        
        ' If there is more than one matching folder, error, else move
        If i = 1 Then
            Set fldrDest = ReturnDestinationFolder(folderAllNames(destIdx), _
                Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders _
            )
        Else
            MsgBox "Zero or more than 1 folder was returned. Giving up", vbCritical, "Move to Folder Error"
        End If
    End If
    
    If i > 1 Or i = 0 Then
        fldrDest = Nothing
    End If
    
    Set obj = Nothing

End Function

Private Function ReturnDestinationFolder(findStr As Variant, fldrs As Outlook.Folders _
    ) As Outlook.MAPIFolder
    
    Dim fldr As Outlook.MAPIFolder
    Dim findArr As Variant
    Dim idx As Long
    Dim subI As Long
    
    ' Split the path into an array
    findStr = Replace(findStr, "\\", "")
    findArr = Split(findStr, "\")
    
    ' We are going to ignore the mailbox ID
    idx = LBound(findArr)
    If InStr(findArr(idx), "@") Or Len(findArr(idx)) = 0 Then idx = idx + 1
    
    For Each fldr In fldrs
        If fldr.Name = findArr(idx) Then
            ' Any more to find?
            If UBound(findArr) > idx Then
                ' Recurse if there are any sub folders
                If fldr.Folders.Count Then
                    Dim subStr As Variant
                    subStr = ""
                    
                    For subI = idx + 1 To UBound(findArr)
                        If (subStr <> "") Then
                            subStr = subStr & "\"
                        End If
                        subStr = subStr & findArr(subI)
                    Next subI

                
                    Set ReturnDestinationFolder = ReturnDestinationFolder( _
                        subStr, _
                        fldr.Folders _
                    )
                Else
                    ' No sub folders so we give up
                    Set ReturnDestinationFolder = Nothing
                End If
            Else
                ' No, so return the found folder
                Set ReturnDestinationFolder = fldr
            End If
            ' We either found it or failed to so no point in going further
            Exit For
        End If
    Next fldr
    
    Set fldr = Nothing

End Function

Private Sub lstAllFolders_Change()
    If changingFolders = False Then
        changingFolders = True
        'Deselect the previously selected folder
        lstFolders.Selected(0) = False
        changingFolders = False
    End If
End Sub

Private Sub lstAllFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If changingFolders = False Then
        changingFolders = True
        'Deselect the previously selected folder
        lstFolders.Selected(0) = False
        changingFolders = False
    End If
    
    Call btnFileToFolder_Click
    
End Sub

Private Sub lstFolders_Change()
    If changingFolders = False Then
        changingFolders = True
        'deselect the first from the all folders list
        lstAllFolders.Selected(0) = False
        changingFolders = False
    End If
End Sub

Private Sub lstFolders_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If changingFolders = False Then
        changingFolders = True
        'deselect the first from the all folders list
        lstAllFolders.Selected(0) = False
        changingFolders = False
    End If
    
    Call btnFileToFolder_Click
    
End Sub

' If the text box contents change, begin filtering
Private Sub tbFilterAllFolders_Change()
    
    With Me.tbFilterAllFolders
        If .Value = vbNullString Then
            Me.lstAllFolders.List = folderAllNames
        Else
            Me.lstAllFolders.List = Filter(SourceArray:=folderAllNames, Match:=.Value, Compare:=vbTextCompare)
        End If
    End With
    
    On Error Resume Next
    
    'When filtering, select the first from the all folders list
    lstAllFolders.Selected(0) = True
End Sub

Private Sub tbFilterAllFolders_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Dim X As Integer
    Dim listIdx As Integer
     If KeyCode = KeyCodeConstants.vbKeyDown And lstAllFolders.ListCount > 0 Then
        listIdx = 0
        
        For X = 0 To lstAllFolders.ListCount - 1
            If X = lstAllFolders.ListCount - 1 Then
                Exit Sub
            ElseIf lstAllFolders.Selected(X) = True Then
                lstAllFolders.Selected(X) = False
                listIdx = X + 1
                Exit For
            End If
        Next X
        
        lstAllFolders.Selected(listIdx) = True
     ElseIf KeyCode = KeyCodeConstants.vbKeyUp And lstAllFolders.ListCount > 0 Then
        listIdx = 0
        
        For X = 1 To lstAllFolders.ListCount - 1
            If lstAllFolders.Selected(X) = True Then
                lstAllFolders.Selected(X) = False
                listIdx = X - 1
                Exit For
            End If
        Next X
        
        lstAllFolders.Selected(listIdx) = True
     End If
End Sub

' Set up the form
Private Sub UserForm_Initialize()

    Dim objItem As Object
    Dim objConvHeader As Outlook.ConversationHeader
    Dim i As Long
    Dim numSelected As Long
    Dim numEmailsSelected As Long
    Dim mb() As String
    
    Dim X As Object
    Set X = Application
    
    ReDim Preserve folderNames(0)
    ReDim Preserve folderPaths(0)
    ReDim Preserve folderAllPaths(0)
    ReDim Preserve folderAllNames(0)
    ReDim Preserve parsedConversations(0)
    
    ' Center Window and adjust for scale
    screenScale = PointsPerPixel
    With Me
        .StartUpPosition = 3
        .left = (Application.ActiveWindow.left * screenScale) + (0.5 * (Application.ActiveWindow.Width * screenScale)) - (Me.Width / 2)
        .top = (Application.ActiveWindow.top * screenScale) + (0.5 * (Application.ActiveWindow.Height * screenScale)) - (Me.Height / 2)
    End With
    
    ' Walk through all selected emails and compile a list of folders
    ' that they are in.
    i = 0
    maxNames = 0
    maxPaths = 0
    numSelected = 0
    numEmailsSelected = 0
    
    'Start with any conversations
    Dim activeSelection, convSelection As Selection
    Set activeSelection = ActiveExplorer.Selection
    Set convSelection = activeSelection.GetSelection(olConversationHeaders)
    
    If convSelection.Count > 0 Then
        For Each objConvHeader In convSelection
            ' Cache conversation ID
            If IsInArray(parsedConversations, objConvHeader.ConversationID) = False Then
                AddToArray parsedConversations, objConvHeader.ConversationID
                
                For Each objItem In objConvHeader.GetItems ' Items in the conversation.
                    InitializeItem objItem, i, mb, numSelected, numEmailsSelected
                Next
            End If
        Next objConvHeader
    End If
    
    For Each objItem In activeSelection
        If TypeName(objItem) = "MailItem" Or TypeName(objItem) = "AppointmentItem" Or TypeName(objItem) = "MeetingItem" Then
            If IsInArray(parsedConversations, objItem.ConversationID) = True Then
                InitializeItem objItem, i, mb, numSelected, numEmailsSelected
            End If
        End If
    Next objItem
    
    
    ReDim parsedConversations(0)
    
    ' Show the list of folders where any of the selected items are already filed
    lstFolders.List = folderNames
    
    ' a selected email already filed so pre-select the first folder in that list
    If i > 0 Then
      lstFolders.Selected(0) = True
    End If
    
    'Create the AllFolders list
    maxFAP = 0
    maxFAN = 0
    ProcessFolder Application.Session.GetDefaultFolder(olFolderInbox).Parent
    
    'If no selected folder, select the first from the all folders list
    'Useful for filtering
    If lstFolders.Selected(0) = False Then
      lstAllFolders.Selected(0) = True
    End If
    
    Set objItem = Nothing
    
End Sub

Sub InitializeItem(ByRef objItem As Object, _
                    ByRef i As Long, _
                    ByRef mb() As String, _
                    ByRef numSelected As Long, _
                    ByRef numEmailsSelected As Long)
    numSelected = numSelected + 1
    ' How many items?
    numEmailsSelected = numEmailsSelected + 1
    ' Check that parent item really is a folder
    If objItem.Parent.Class = olFolder Then
        ' Edited to no longer exclude any folders they're in. Seems weird that any were excluded.
        If IsInArray(folderPaths, objItem.Parent.FolderPath) = False Then
                
            ' Save mailbox name
            If IsNull(mailbox) Or mailbox = "" Then
                mb = Split(objItem.Parent.FolderPath, "\")
                mailbox = mb(2)
            End If
            
            AddToArray folderPaths, objItem.Parent.FolderPath
            maxPaths = maxPaths + 1
            AddToArray folderNames, objItem.Parent.Name
            maxNames = maxNames + 1
            
            i = i + 1
            
        End If
    End If
End Sub

' Create the all-folder list
Sub ProcessFolder(objStartFolder As Outlook.MAPIFolder, _
                  Optional blnRecurseSubFolders As Boolean = True, _
                  Optional strFolderPath As String = "", _
                  Optional strFolderName As String = "")

    Dim objFolder As Outlook.MAPIFolder

    Dim i As Long, mb

     ' Loop through the items in the current folder
    For i = 1 To objStartFolder.Folders.Count

        Set objFolder = objStartFolder.Folders(i)

        ' Populate the listbox & save actual folder paths
        ' But only for NOT sent, drafts, etc
        ' Don't block the inbox in case it has sub-folders
        If objFolder.Name <> "Sent Items" And _
            objFolder.Name <> "Deleted Items" And _
            objFolder.Name <> "Outbox" And _
            objFolder.Name <> "Calendar" And _
            objFolder.Name <> "Contacts" And _
            objFolder.Name <> "Notes" And _
            objFolder.Name <> "Journal" And _
            objFolder.Name <> "Junk E-mail" And _
            objFolder.Name <> "News Feed" And _
            objFolder.Name <> "RSS Feeds" And _
            objFolder.Name <> "Conversation History" And _
            objFolder.Name <> "Conversation Action Settings" And _
            objFolder.Name <> "Quick Step Settings" And _
            objFolder.Name <> "LinkedIn" And _
            objFolder.Name <> "Suggested Contacts" And _
            objFolder.Name <> "Sync Issues" And _
            objFolder.Name <> "Tasks" And _
            objFolder.Name <> "My Site" And _
            objFolder.Name <> "Drafts" _
        Then
            ' Save mailbox name
            If maxFAP = 0 And mailbox = "" Then
                mb = Split(objFolder.FolderPath, "\")
                mailbox = mb(2)
            End If

            ' Add to the All Folder List
            lstAllFolders.AddItem strFolderName + "\" + objFolder.Name
            AddToArray folderAllPaths, objFolder.FolderPath
            maxFAP = maxFAP + 1
            AddToArray folderAllNames, strFolderName + "\" + objFolder.Name
            maxFAN = maxFAN + 1
            ' Recurse subfolders but not for subfolders of blocked folders
            If blnRecurseSubFolders Then
                ' Recurse through subfolders
                ProcessFolder objFolder, True, strFolderPath + "\" + objFolder.FolderPath, _
                    strFolderName + "\" + objFolder.Name
            End If
        End If

    Next
    
    Set objFolder = Nothing

End Sub


' ---- Functions from other people ----
Private Sub AddToArray(ByRef arr As Variant, val As Variant)
On Error GoTo err
    If arr(0) <> "" Or UBound(arr) > 0 Then
        ReDim Preserve arr(LBound(arr) To UBound(arr) + 1)
    End If
    
    On Error GoTo 0
    arr(UBound(arr)) = val
    Exit Sub
err:
    Debug.Print "Failed to add to array"
    Dim Msg As String
    Msg = "Error # " & Str(err.Number) & " was generated by " _
         & err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & err.Description
    MsgBox Msg, , "Error", err.HelpFile, err.HelpContext
End Sub

Function IsInArray(arr As Variant, valueToFind As Variant) As Boolean
    ' checks if valueToFind is found in arr, no loop!
    On Error GoTo err
    IsInArray = (UBound(Filter(arr, valueToFind)) > -1)
    Exit Function
err:
    Debug.Print "Failed to add to array"
End Function

Sub WaitFor(NumOfSeconds As Long)
    Dim SngSec As Long
    SngSec = Timer + NumOfSeconds
    
    Do While Timer < SngSec
        DoEvents
    Loop

End Sub
