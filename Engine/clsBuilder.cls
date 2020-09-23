VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'*~ lean and mean.. ~*
'*~
'*~ Some observations.. ~*
'*~ While making the file search routines, I did some experimenting
'*~ benching for the fastest possible recursion routines. What I found may
'*~ suprise some of you.. There are a number of extraneous factors
'*~ that occur at any given time that make benchmarking inexact..
'*~ drive/processor caches, thread spikes, etc.. but after several runs
'*~ I noticed some consistant behaviors.

'*~ 1) Constants slow performance down, hence I removed them where I could

'*~ 2) The boolean case select structure was faster then if/else, I think, (could be wrong), this is because
'*~ boolean only evaluates until statement is false, and if/else reads entire statement..

'*~ 3) open if statements (no closing end if) appeared to be, (marginally), faster
'*~ but I suspect code is the same when compiled

'*~ 4) avoid doevents like the plague!

'*~ 5) goto skip was faster then end if, (not sure why)

'*~ 6) shorter variable names are faster, (it all ends up in the bitstream)..
'*~ var names get compiled as numbers, yes, but it all gets serialized into binary
'*~ and even fixed length numeric variables are padded out with zeros, so I would
'*~ argue that a zero gets processed faster then a one (cache population, switch polarization, etc)
'*~ so this: 00000000 00010001 should be faster then this: 10101010 10101010.. because
'*~ any time you expend energy on a system, it costs you in time.. Thats my story, and I'm stickin to it..

'*~ 7) Raising status events are death! (in a recursion routine) If you must have a progress indicator
'*~ have it count through routines, or fake a timer based on HD size and run to 99% - addendum: added
'*~ a very inexpensive progress indicator that uses root folder enumeration for max count, and test for
'*~ path before recursion call - impact was nominal, but a user queue was needed

'*~ 8) The use of collections as done in this example yields the fastest combination of lookup/build results
'*~ The only thing faster then fetching an item by keyname from a collection (afaik), is a hashtable
'*~ and this takes time to build, and has a large (2x to 4x of array size) memory footprint.
'*~ Sorting/indexing a string array also takes time, (even with triquicksort), but a collection has none of this
'*~ additional overhead, builds are almost as fast as an array, needs no sorting, and lookups are very fast,
'*~ hence is ideal in this sort of situation. If you can provide an all-around faster (build + lookup) example
'*~ using arrays, send it along, (I would be very impressed)..

'*~ I was originally going for a system that forked the root folders
'*~ recursion calls out to seperate threads, that in turn registered with a central
'*~ processing hub that collected data from all streams. But this is both problematic
'*~ and not likely to be faster then the async method demonstrated here.
'*~ Spawning many worker threads means more overhead, and unless threads have
'*~ different activities with varying I/O wait times, dividing operations
'*~ would likely slow it all down

'*~ I went on some pretty wild tangents in attempts to improve performance with the recursion engine..
'*~ Highlites include: building assembler calls on the fly for findfile to call api directly, using copymemory
'*~ to join strings, using wndproc to map library addresses, spawning worker threads for each recursion call,
'*~ and much, much more.. but in each case, additional routines caused slowdowns, and nullified any gains.
'*~ If anyone can find some creative way of making this significantly faster (proofed), write a demo, send it along,
'*~ and I will build it in..

'*~ As it stands, in single mode the search function can return every match item on my full 4 gig partition (1ghz, 256MB)
'*~ in half a second. The recurse can dump returns from 20 item searches in about 3 seconds
'*~ the recursive search will slow as the number of files increase on a system, example sited was 25k +
'*~ but lag is one time only, during recursion and collection. while engine is engaged, balance of returns
'*~ will be very fast, as they are using keys to fetch results and return them to caller through
'*~ a bidirectional collection exchange.

'/* standard file type structures

'/* file time struct
Private Type FT
    lLD                             As Long
    lHD                             As Long
End Type

'/* file find struct
Private Type WFD
    lFA                             As Long
    fCT                             As FT
    fLA                             As FT
    fLWT                            As FT
    lFSH                            As Long
    lFSL                            As Long
    lR0                             As Long
    lR1                             As Long
    cFN                             As String * 260
    cAL                             As String * 14
End Type

Private lC                      As Long

'/* status event
Public Event eBCounter()
Public Event eBCountMax(lMax As Long)

'/* locals
Private c_Counter               As New Collection
Private W                       As String
Private N                       As String
Private X                       As String
Private lF                      As Long
Private W32                     As WFD

'/* file search api
Private Declare Function FindFirstFile Lib "Kernel32" Alias "FindFirstFileA" (ByVal lFN As String, _
                                                                              lFD As WFD) As Long

Private Declare Function FindNextFile Lib "Kernel32" Alias "FindNextFileA" (ByVal lHD As Long, _
                                                                            lFD As WFD) As Long

Private Declare Function FindClose Lib "Kernel32" (ByVal lHD As Long) As Long

Private Declare Function CloseHandle Lib "Kernel32" (ByVal lHD As Long) As Long


Public Sub Controller(ByVal P As String)
'/* pre-engage then recurse

    Preloader P
    Recurse P

End Sub

Private Sub Preloader(ByVal P As String)
'/* get root folder count for progress
'/* and pre-load drive cache

Dim E As Boolean
Dim H As Long

On Error Resume Next

    H = FindFirstFile(P + "*.*", W32)
    E = True
    Do While E
        With W32
            N = Left$(.cFN, InStr(1, .cFN, Chr$(0)) - 1)
        End With
        If Not LenB(N) > 0 Then GoTo Skip
        If AscW(N) = 46 Then GoTo Skip
        If W32.lFA And &H10 Then
            c_Counter.Add P + N
        End If
Skip:
        E = FindNextFile(H, W32)
    Loop
    FindClose H

    '/* send total back to caller for prg max
    RaiseEvent eBCountMax(c_Counter.Count)

On Error GoTo 0

End Sub

Private Sub Recurse(ByVal P As String)
'/* dump file paths into a collection
'/* and pass to caller for processing

Dim E As Boolean
Dim H As Long

On Error Resume Next

    H = FindFirstFile(P + "*.*", W32)
    E = True
    Do While E
        With W32
            N = Left$(.cFN, InStr(1, .cFN, Chr$(0)) - 1)
        End With
        If Not LenB(N) > 0 Then GoTo Skip
        If AscW(N) = 46 Then GoTo Skip
        Select Case True
        Case W32.lFA And &H10
            '/* lean progress counter
            '/* length check is relatively inexpensive
            If LenB(P + N) = LenB(c_Counter.Item(1)) Then
                '/* if they match
                If P + N = c_Counter.Item(1) Then
                    RaiseEvent eBCounter
                    c_Counter.Remove 1
                    '/* if you must have doevents
                    '/* this is the most inexpensive spot
                    DoEvents
                End If
            End If
            Recurse P + N + Chr$(92)
        Case Else
            '/* if key exists test
            W = N
            X = c_Storage.Item(W)
            If Not LenB(X) = 0 Then
                lC = 0
                '/* used to generate unique keys for collection
                '/* name key to next number with delimiter
                '/* test for key existence, move to next possible number
                '/* until no match is found
                '/* ex. desktop.ini 17
                Do Until LenB(X) = 0
                    lC = lC + 1
                    W = N + Chr$(31) & lC
                    X = vbNullString
                    X = c_Storage.Item(W)
                Loop
            End If
            '/* add file path/name and unique key
            c_Storage.Add P + N, W
        End Select
Skip:
        E = FindNextFile(H, W32)
    Loop
    FindClose H

On Error GoTo 0

End Sub

Private Sub Class_Initialize()
    Set c_Counter = New Collection
End Sub

Private Sub Class_Terminate()
    Set c_Counter = Nothing
End Sub


