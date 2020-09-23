VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'/ collection structs
Private Type cData
   BteP(0 To 2)     As Long
   Bid              As Long
   Count            As Long
   Res1             As Long
   FirstPtr         As Long
   LastPtr          As Long
   Res2             As Long
   Res3             As Long
   Res4             As Long
End Type

Private Type CElement
   Data             As Variant
   Key              As Long
   Prev             As Long
   Next             As Long
   Res1             As Long
   Res2             As Long
   Res3             As Long
End Type

Public Event eIStatus(bState As Boolean)

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function lstrlenW Lib "Kernel32" (ByVal lpString As Any) As Long


Public Sub Index_Dump(ByVal sIndPath As String)

'/*  dump collection keys and items to a binary file
'*~ The problem here was that collection keys are -write only-
'*~ this would have meant parsing key (file name) out of every
'*~ item, then formatting a string for a binary file dump..
'*~ This was a very expensive approach, as collection sizes
'*~ could top 500k items on large drives. I searched the internet
'*~ for hours, but all posts on the subject agreed that key could
'*~ not be read -until- some dude had little snippet of code that
'*~ illustrated collections structure, and said key could be hacked
'*~ from memory.. good idea! though his example crashed ide :o(
'*~ So, this was the result of that effort..
'*~ collection structure is enumerated and key/item extracted and
'*~ copied to arrays, arrays then use copymemory to build string for
'*~ the binary dump.. this proved to be hundreds of times faster
'*~ then original method..

Dim aKeys()     As String
Dim aItems()    As Variant
Dim lCt         As Long
Dim tCData      As cData
Dim tCElement   As CElement
Dim lIt         As Long
Dim lPos        As Long
Dim sResult     As String
Dim sTemp       As String
Dim lLen        As Long
Dim FF          As Integer
Dim cTemp       As New Collection

On Error GoTo Handler

    Set cTemp = New Collection
    Set cTemp = c_Storage
    
    If Len(sIndPath) = 0 Then Exit Sub
    If Not cTemp.Count > 0 Then Exit Sub
    FF = FreeFile
    
    '/* initialize arrays
    With cTemp
        ReDim aKeys(1 To .Count)
        ReDim aItems(1 To .Count)
    End With
    
    '/* copy to memory
    CopyMemory tCData, ByVal ObjPtr(cTemp), LenB(tCData)
    '/* first element
    tCElement.Next = tCData.FirstPtr

    For lCt = 1 To tCData.Count
        '/* fill data structure
        CopyMemory tCElement, ByVal tCElement.Next, LenB(tCElement)
        '/* copy key and item
        aKeys(lCt) = String_Pointer(tCElement.Key)
        aItems(lCt) = tCElement.Data
        '/* dimension tracking
        lLen = lLen + LenB(aKeys(lCt)) + LenB(aItems(lCt)) + 2
    Next
    
    '/* size string
    sTemp = Space$(lLen)
    '/* iterate through arrays
    '/* and copy to string
    For lIt = 1 To UBound(aKeys)
        sResult = aKeys(lIt) + Chr$(30) + aItems(lIt) + Chr$(29)
        CopyMemory ByVal StrPtr(sTemp) + lPos, ByVal StrPtr(sResult), LenB(sResult)
        lPos = lPos + LenB(sResult)
    Next lIt
    
    '/* dump to file
    If File_Exists(sIndPath) Then
        Kill sIndPath
    End If
    Open sIndPath For Binary As #FF
    Put #FF, , sTemp
    Close #FF

    '/* cleanup
    '/* do you need this here?
    '/* lets do it anyways..
    sResult = vbNullString
    sTemp = Space$(0)
    sTemp = vbNullString
    Erase aKeys
    Erase aItems
    
    Set cTemp = Nothing
    '/* weird bug caused me a lot of grief - don't ask
    c_Storage.Remove (c_Storage.Count)

Handler:
On Error GoTo 0

End Sub

Private Function String_Pointer(lAddress As Long) As String
'/* get the string address

   String_Pointer = String$(lstrlenW(lAddress), 0)
   CopyMemory ByVal StrPtr(String_Pointer), ByVal lAddress, lstrlenW(lAddress) * 2

End Function

Public Sub Index_Restore(ByVal sIndPath As String)

'/* open file, split and add to collection..
'*~ Tried copying entire collection
'*~ structure to file, then reloading it
'*~ at new collections entry point, but
'*~ just crashes ide.. :o(

Dim sResult As String
Dim l       As Long
Dim M       As Long
Dim c       As Long
Dim t       As String
Dim FF      As Integer
Dim cTemp   As New Collection

On Error Resume Next

    If Len(sIndPath) = 0 Then Exit Sub
    '~* create temporary container
    Set cTemp = New Collection
    
    FF = FreeFile
    '/* dump file
    Open sIndPath For Binary As #FF
    sResult = Space$(LOF(FF))
    Get #FF, , sResult
    Close #FF
    
    '/* 'loop until' is fastest..
    '/* concession added to events
    With cTemp
        Do
            M = l + 1
            l = InStr(M, sResult, Chr$(29))
            If Not l = 0 Then
                '~* not really needed.. ~*
                '/* at 2k iterations allow
                '/* system update
                c = c + 1
                If Not c < 2000 Then
                    c = 0
                    DoEvents
                End If
                '/* extract temp string, format
                '/* and add to collection
                t = Mid$(sResult, M, l - M)
                .Add Mid$(t, InStr(1, t, Chr$(30)) + 1), Left$(t, InStr(1, t, Chr$(30)) - 1)
            End If
        Loop Until l = 0
    End With
    
    Storage_Reset
    Set c_Storage = cTemp
    Set cTemp = Nothing

On Error GoTo 0

End Sub

Public Function Index_Status(ByVal sIndPath As String) As Boolean
'/* track status of index

Dim bState As Boolean

    If File_Exists(sIndPath) Then
        bState = True
        Index_Status = True
    End If
    
    '/* send notification
    RaiseEvent eIStatus(bState)

End Function

Public Function Storage_Reset()
'/* reset global storage object

    If Not c_Storage Is Nothing Then
        Set c_Storage = Nothing
    End If
    Set c_Storage = New Collection
    
End Function

Public Function File_Exists(ByVal sPath As String) As Boolean
'/* test file path

    If Len(Dir(sPath)) > 0 Then
        File_Exists = True
    End If
    
End Function

