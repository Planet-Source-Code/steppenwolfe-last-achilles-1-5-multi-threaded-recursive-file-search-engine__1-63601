Attribute VB_Name = "mControl"
Option Explicit
'*~ timer and async thread launcher ~*
'*~ Based on an example written by Steve McMahon.. (my hero!) ~*

Private m_lTid                                  As Long
Private c_Run                                   As New Collection
Public c_Counter                                As New Collection
Public c_Storage                                As New Collection
Public c_CForward                               As New Collection
Public c_CReturn                                As New Collection
Public c_CMultiTask                             As New Collection

'/* version structure
Private Type OSVersion
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type

Private Declare Function CoLockObjectExternal Lib "ole32" (ByVal pUnk As IUnknown, _
                                                           ByVal fLock As Long, _
                                                           ByVal fLastUnlockReleases As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Private Sub Timer_Proc(ByVal lHwnd As Long, _
                       ByVal lMsg As Long, _
                       ByVal lTimerID As Long, _
                       ByVal lTime As Long)
'/* instantiate instance of runnable.tlb

Dim this As Runnable


    With c_Run
        Do While .Count > 0
            Set this = .Item(1)
            .Remove 1
            this.Start
            CoLockObjectExternal this, 0, 1
        Loop
    End With

    '/* release timer
    KillTimer 0, lTimerID
    m_lTid = 0
    Set c_Run = Nothing

End Sub

Public Sub Start(this As Runnable)
'/* start the timer

    CoLockObjectExternal this, 1, 1
    If c_Run Is Nothing Then
        Set c_Run = New Collection
    End If

    c_Run.Add this
    If Not m_lTid Then
        m_lTid = SetTimer(0, 0, 1, AddressOf Timer_Proc)
    End If

End Sub
