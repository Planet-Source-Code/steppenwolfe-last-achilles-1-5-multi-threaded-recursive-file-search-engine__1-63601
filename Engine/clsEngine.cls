VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEngine"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***************************************************************************************
'*  Achilles!    Multi Threaded Recursive Search Engine ver. 1.5                       *
'*                                                                                     *
'*  Created:     December 21, 2005                                                     *
'*  Purpose:     Ultra-Fast File Search and Enumeration Engine                         *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.5.17                                                                *
'*  Compile:     Native                                                                *
'*  Referenced:  mControl, clsSearch                                                   *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*                                                                                     *
'***************************************************************************************


'*~ Exposed Functions ~*

'*~ Start                   - Start Asynchronous Processing Engine - Interface to Runnable_Start
'*~ Engine_Reload           - Reset all classes and collection objects
'*~ Engine_Reset            - Terminate all storage objects on class exit
'*~ Storage_Reset           - Reset central storage object
'*~ Thread_Accelerate       - Raises thread priority and process class


'*~ Exposed Properties ~*

'*~ p_EngineTask            - Assigns an engine task [Build | Save | Restore | Exact | Pattern]
'*~ p_CMultiTask            - Allows for serial processing of multiple commands [collection]
'*~ p_BuildPath             - Specified search path [string]
'*~ p_IndexId               - Index ID number [corresponds to saved file name i.e. index1.dat]
'*~ p_IndexPath             - Index path [fetch]
'*~ p_CReturn               - Search results return [collection]
'*~ p_CForward              - Search phrases passed into engine [collection]


'*~ List of Events ~*

'*~ eEComplete              - Engine task has completed
'*~ eEEngaged               - Engine is processing
'*~ eEReload                - Search collections have been erased
'*~ eEReset                 - All collections have been erased
'*~ eEReset                 - Search Engine has been Reset
'*~ eECount                 - Collapsed Search Counter
'*~ eECountMax              - Search Counter Max Param
'*~ eEBuild                 - Storage has been built
'*~ eEStoreReset            - Storage has been erased
'*~ eEProcessComplete       - Exact match scan has completed
'*~ eEPatternComplete       - Pattern scan has completed
'*~ eEDump                  - Index has been saved
'*~ eERestore               - Index has been restored


'*~ December 21, 2005
'*~ You can use/abuse this code in any way you like.. provided that you give credit to the author
'*~ and keep these header files intact. Place a credit line in appropriate place: about/help display,
'*~ reading: ex. engine based on the Achilles 1.5 Scan Engine by John Underhill, www.nspowertools.com

'/* runnable type library
Implements Runnable


'~* class linkage ~*

'/* search engine
Private cSearch                 As clsSearch
Attribute cSearch.VB_VarHelpID = -1
'/* index engine
Private WithEvents cBuild       As clsBuilder
Attribute cBuild.VB_VarHelpID = -1
'/* index engine
Private WithEvents cIndex       As clsIndex
Attribute cIndex.VB_VarHelpID = -1


'~* event notifications ~*

'*~ clsEngine ~*
'/* engine task completed
Public Event eEComplete()
'/* engine is busy
Public Event eEEngaged()
'/* search collections have been erased
Public Event eEReload()
'/* all collections have been erased
Public Event eEReset()
'/* storage reset
Public Event eEStoreReset()
'/* multitask event completed
Public Event eEMultiTask()

'*~ clsBuilder ~*
'/* indece returned notifier
Public Event eEBuild()
'/* recursion progress ECounter
Public Event eECount()
'/* recursion progress max
Public Event eECountMax(lMax As Long)

'*~ clsSearch ~*
'/* match search completed
Public Event eEProcessComplete()
'/* pattern search completed
Public Event eEPatternComplete()

'*~ clsIndex ~*
'/* index dump complete
Public Event eEDump()
'/* index restore complete
Public Event eERestore()
'/* index status notification
Public Event eEIndStatus(bState As Boolean)


'~* global enums ~*

'/* thread acceleration levels
Public Enum ePriority
    Low = 1
    Normal = 2
    High = 3
    Priority = 4
    High_Priority = 5
    Realtime = 6
End Enum

'/* engine task assignment
Public Enum eTask
    Index_Build = 1
    Index_Save = 2
    Index_Restore = 3
    Index_Status = 4
    Search_Exact = 5
    Search_Pattern = 6
    Engine_MultiTask = 7
End Enum


'~* local variables ~*

Private m_EngineTask            As Integer
Private m_IndexId               As Integer
Private m_bRunning              As Boolean
Private m_sBuildPath            As String
Private m_sIndexPath            As String


'/* thread priority api
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long

Private Declare Function SetThreadPriority Lib "Kernel32" (ByVal hThread As Long, _
                                                           ByVal nPriority As Long) As Long

Private Declare Function SetPriorityClass Lib "Kernel32" (ByVal hProcess As Long, _
                                                          ByVal dwPriorityClass As Long) As Long

Private Declare Function GetThreadPriority Lib "Kernel32" (ByVal hThread As Long) As Long

Private Declare Function GetPriorityClass Lib "Kernel32" (ByVal hProcess As Long) As Long

Private Declare Function GetCurrentThread Lib "Kernel32" () As Long


'***~ Properties ~***

'/* index builder switch
Public Property Get p_EngineTask() As eTask
    p_EngineTask = m_EngineTask
End Property

Public Property Let p_EngineTask(ByVal PropVal As eTask)
    m_EngineTask = PropVal
End Property

'/* search path (builder)
Public Property Get p_BuildPath() As String
    p_BuildPath = m_sBuildPath
End Property

Public Property Let p_BuildPath(ByVal PropVal As String)
    m_sBuildPath = PropVal
End Property

'/* index id (index)
Public Property Get p_IndexId() As Integer
    p_IndexId = m_IndexId
End Property

Public Property Let p_IndexId(ByVal PropVal As Integer)
    m_IndexId = PropVal
End Property

'/* index path (index)
Public Property Get p_IndexPath() As String
    p_IndexPath = m_sIndexPath
End Property

Public Property Let p_IndexPath(ByVal PropVal As String)
    m_sIndexPath = PropVal
End Property

'/* return results collection to caller
Public Property Get p_CReturn() As Collection
    Set p_CReturn = c_CReturn
End Property

Public Property Set p_CReturn(ByVal NewCollection As Collection)
    Set c_CReturn = NewCollection
End Property

'/* return index size
Public Property Get p_StorageCt() As Long
    p_StorageCt = c_Storage.Count
End Property


'/* pass search collection into class
Public Property Get p_CForward() As Collection
    Set p_CForward = c_CForward
End Property

Public Property Set p_CForward(ByVal NewCollection As Collection)
    Set c_CForward = NewCollection
End Property

'/* serialized task collection
Public Property Get p_CMultiTask() As Collection
    Set p_CMultiTask = c_CMultiTask
End Property

Public Property Set p_CMultiTask(ByVal NewCollection As Collection)
    Set c_CMultiTask = NewCollection
End Property


'***~ Events ~***

'/* return counter tick
Private Sub cBuild_eBCounter()
    RaiseEvent eECount
End Sub

'/* return counter max
Private Sub cBuild_eBCountMax(lMax As Long)
    RaiseEvent eECountMax(lMax)
End Sub

'/* return index status
Private Sub cIndex_eIStatus(bState As Boolean)
    RaiseEvent eEIndStatus(bState)
End Sub


'***~ Processors ~***

Private Sub Class_Initialize()
'/* instantiate search class

    Set cSearch = New clsSearch
    Set cBuild = New clsBuilder
    Set cIndex = New clsIndex
    Set c_Storage = New Collection
    Set c_CReturn = New Collection
    Set c_CForward = New Collection
    Set c_CMultiTask = New Collection

End Sub

Public Sub Start()
'/* start async processing

    If Not m_bRunning Then
        m_bRunning = True
        mControl.Start Me
    End If

End Sub

Private Sub Runnable_Start()
'/* search/build/index class calls
'/* forks out to index/builder/search modes
'/* set with p_Build/p_MultiTask property setting

On Error GoTo Handler

    '/* flag engaged
    RaiseEvent eEEngaged
    
    Select Case p_EngineTask
    Case 1
        '/* build an index
        cBuild.Controller p_BuildPath
        RaiseEvent eEBuild
    Case 2
        '/* save index
        cIndex.Index_Dump p_IndexPath
        RaiseEvent eEDump
    Case 3
        '/* restore index
        cIndex.Index_Restore p_IndexPath
        RaiseEvent eERestore
    Case 4
        '/* get index status
        cIndex.Index_Status p_IndexPath
    Case 5
        '/* exact search
        cSearch.Results_Process
        RaiseEvent eEProcessComplete
    Case 6
        '/* pattern search
        cSearch.Results_Pattern
        RaiseEvent eEPatternComplete
    Case 7
        '/* perform serialized tasks
        Dim vItem As Variant
        While p_CMultiTask.Count > 0
            Select Case CInt(p_CMultiTask.Item(1))
            Case 1
                '/* build an index
                cBuild.Controller p_BuildPath
                RaiseEvent eEBuild
            Case 2
                '/* save index
                cIndex.Index_Dump p_IndexPath
                RaiseEvent eEDump
            Case 3
                '/* restore index
            cIndex.Index_Restore p_IndexPath
            RaiseEvent eERestore
            End Select
            p_CMultiTask.Remove 1
        Wend
        '/* loop has been processed
        RaiseEvent eEMultiTask
    End Select

Handler:

    '/* we have completed
    RaiseEvent eEComplete
    
    '/* end async processing
    m_bRunning = False


End Sub

Public Sub Thread_Accelerate(ByVal eLevel As ePriority)
'/* process accelerator
'/* alter thread priority for current thread
'/* alter process class for process
'*~ not implemented at this time **
'*~ requires testing on all platforms **

Dim lProcess  As Long
Dim lThread   As Long
Dim lPriority As Long
Dim lClass    As Long

On Error Resume Next

    Select Case eLevel
    Case 1
        '/* not 100% sure on vals here..
        lPriority = &HFFFFFFFF
        lClass = &H20
    Case 2
        lPriority = &H0
        lClass = &H20
    Case 3
        lPriority = &H2
        lClass = &H80
    Case 4
        lPriority = &H6
        lClass = &H80
    Case 5
        lPriority = &HA
        lClass = &H80
    Case 6
        lPriority = &HF
        lClass = &H100
    Case Else
        Exit Sub
    End Select
    
    lThread = GetCurrentThread
    If lThread = 0 Then Exit Sub
    
    lProcess = GetCurrentProcess
    SetThreadPriority lThread, lPriority
    If lProcess = 0 Then Exit Sub
    SetPriorityClass lProcess, lClass

On Error GoTo 0

End Sub

Public Sub Engine_Reload()
'/* reset classes/search collections

On Error Resume Next

    Set p_CForward = Nothing
    Set p_CReturn = Nothing
    Set p_CMultiTask = Nothing
    Set p_CForward = New Collection
    Set p_CReturn = New Collection
    Set p_CMultiTask = New Collection
    Set cBuild = Nothing
    Set cSearch = Nothing
    Set cIndex = Nothing
    Set cBuild = New clsBuilder
    Set cSearch = New clsSearch
    Set cIndex = New clsIndex
    
    '/* locked and loaded
    RaiseEvent eEReload

On Error GoTo 0

End Sub

Public Sub Engine_Reset()
'/* reset search collections

On Error Resume Next

    Set c_CForward = Nothing
    Set c_CReturn = Nothing
    Set c_CMultiTask = Nothing
    Set c_Storage = Nothing
    Set cBuild = Nothing
    Set cSearch = Nothing
    Set cIndex = Nothing
    
    '/* we are done
    RaiseEvent eEReset

On Error GoTo 0

End Sub

Public Sub Storage_Reset()
'/* reset storage collection

On Error Resume Next

    Set c_Storage = Nothing
    Set c_Storage = New Collection
    
    '/* reset complete
    RaiseEvent eEStoreReset

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
'/* reset search class

    Engine_Reset

End Sub
