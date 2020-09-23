VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBenchMark 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Achilles 1.5 - Bench Mark"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   6825
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgIndex 
      Height          =   135
      Left            =   180
      TabIndex        =   9
      Top             =   1680
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdDir 
      Caption         =   "Dir Search >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5070
      TabIndex        =   7
      Top             =   2130
      Width           =   1485
   End
   Begin VB.CommandButton cmdBench 
      Caption         =   "Bench Mark >"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5010
      TabIndex        =   5
      Top             =   720
      Width           =   1485
   End
   Begin VB.Frame frmTime 
      Caption         =   "Current Status"
      Height          =   1335
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   4515
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "Found: "
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   540
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "Index Size: "
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   3
         Top             =   750
         Width           =   825
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "Time: "
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblStats 
         AutoSize        =   -1  'True
         Caption         =   "Search Count:"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Label lblProgress 
      AutoSize        =   -1  'True
      Caption         =   "Index Progress"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1500
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   $"frmBenchMark.frx":0000
      Height          =   645
      Left            =   210
      TabIndex        =   8
      Top             =   2190
      Width           =   4665
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "Total time:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1860
      Width           =   735
   End
End
Attribute VB_Name = "frmBenchMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'/~ Achilles 1.5 Benchmarking routines
'/~ Author John Underhill (Steppenwolfe)
'/~ if you need examples of other functions, just look at Hyperion..

Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private WithEvents cEngine          As clsEngine
Attribute cEngine.VB_VarHelpID = -1
Private c_Local                     As New Collection
Private c_Result                    As New Collection
Private m_bPreprocess               As Boolean
Private m_dTime                     As Double
Private m_dTotal                    As Double

Private Sub cEngine_eEBuild()
Debug.Print "build complete. Index size: " & cEngine.p_StorageCt & " items"

    m_dTotal = Format$(Timer - m_dTime, "#0.0000")
    Me.Caption = "Index loaded in: " & m_dTotal & " Seconds"

End Sub

Private Sub cEngine_eEComplete()
Debug.Print "task completed"

    cmdBench.Enabled = True
    
End Sub

Private Sub cEngine_eECount()
Debug.Print "progress tick"

On Error Resume Next

    prgIndex.Value = prgIndex.Value + 1
    lblProgress.Caption = Format((prgIndex.Value / prgIndex.Max) * 100, "###") & "%"

On Error GoTo 0

End Sub

Private Sub cEngine_eECountMax(lMax As Long)
Debug.Print "progress max count: " & lMax

    prgIndex.Max = lMax

End Sub

Private Sub cEngine_eEDump()
Debug.Print "index has been dumped"
End Sub

Private Sub cEngine_eEEngaged()
Debug.Print "the engine is engaged"
    cmdBench.Enabled = False
End Sub

Private Sub cEngine_eEIndStatus(bState As Boolean)
Debug.Print "Index status is: " & bState
End Sub

Private Sub cEngine_eEMultiTask()
Debug.Print "multitasking has completed"
End Sub

Private Sub cEngine_eEPatternComplete()
Debug.Print "pattern search has completed"
End Sub

Private Sub cEngine_eEReload()
Debug.Print "the engine has reloaded"
End Sub

Private Sub cEngine_eEReset()
Debug.Print "the engine has been reset"
End Sub

Private Sub cEngine_eERestore()
Debug.Print "the index has been restored"
End Sub

Private Sub cEngine_eEStoreReset()
Debug.Print "primary storage has been reset"
End Sub

Private Sub cEngine_eEProcessComplete()
Debug.Print "match search has completed"

Dim vItem   As Variant
Dim l       As Long
Dim dTemp   As Double


    lblStats(0).Caption = "Search Count " & c_Local.Count & " search items"
    dTemp = Format$(Timer - m_dTime, "#0.0000")
    m_dTotal = m_dTotal + dTemp
    lblTotal.Caption = "Scanned for 10,000 file names in total (index + scan) of: " & m_dTotal & " seconds"
    lblStats(1).Caption = "Time: " & dTemp & " Seconds"
    lblStats(2).Caption = "Index Size: " & cEngine.p_StorageCt & " items"
    lblStats(3).Caption = "Found: " & cEngine.p_CReturn.Count
    
    '/* transfer to local list
    Set c_Result = cEngine.p_CReturn
    '/* show results/counter
    For Each vItem In c_Result
        Debug.Print vItem
        l = l + 1
        If l > 100 Then
            l = 0
            DoEvents
        End If
    Next vItem

    'Me.Height = 3510
    
End Sub

Private Sub Engine_Scan()

    '/* reset counters
    m_dTime = 0
    m_dTime = Timer
    
    '/* initiate search
    With cEngine
        '/* scan path
        .p_BuildPath = "c:\"
        '/* add search items
        Set .p_CForward = c_Local
        '/* search type
        .p_EngineTask = Search_Exact
        '/* start processing
        .Start
    End With

End Sub

Private Sub Engine_Init()

    '/* initiate search
    With cEngine
        '/* scan path
        .p_BuildPath = "c:\"
        '/* search type
        .p_EngineTask = Index_Build
        '/* start processing
        .Start
    End With

End Sub

Private Sub cmdBench_Click()

    m_bPreprocess = True
    '/* build a search list
    Search_List
    '/* index launch
    Engine_Scan

End Sub

Private Sub cmdDir_Click()
Dir_Test
End Sub

Private Sub Form_Load()

    '/* dimension objects/set options
    Set cEngine = New clsEngine
    Set c_Local = New Collection
    Set c_Result = New Collection
    cmdBench.Enabled = False
    Me.Caption = "Loading Index.."
    '/* reset counters
    m_dTime = 0
    m_dTime = Timer
    Engine_Init

End Sub

Private Sub Dir_Test()
'/* parting shot..

    Open App.Path & "\dirtest.bat" For Append As #1
        Print #1, "@echo on"
        Print #1, "set _number=0"
        Print #1, "set _max=999"
        Print #1, ":Start"
        Print #1, "if %_number%==%_max% goto end"
        Print #1, "dir /b /s c:\msmin.exe"
        Print #1, "dir /b /s c:\pinball.exe"
        Print #1, "dir /b /s c:\wmplayer.exe"
        Print #1, "dir /b /s c:\slip.scp"
        Print #1, "dir /b /s c:\actmovie.exe"
        Print #1, "dir /b /s c:\msmom.dll"
        Print #1, "dir /b /s c:\nac.dll"
        Print #1, "dir /b /s c:\mobsync.exe"
        Print #1, "dir /b /s c:\freecell.exe"
        Print #1, "dir /b /s c:\diskmgmt.msc"
        Print #1, "set /a _number=_number + 1"
        Print #1, "goto start"
        Print #1, ":end"
        Print #1, "set _number="
        Print #1, "set _max="
    Close #1
    
    ShellExecute Me.hwnd, "open", App.Path & "\dirtest.bat", "", "", 1

End Sub

Private Sub Search_List()

Dim i As Integer

    Set c_Local = Nothing
    Set c_Result = Nothing
    Set c_Local = New Collection
    Set c_Result = New Collection
    
    '/* build a list of 10,000 file names
    '/* makes absolutely no difference if
    '/* names repeat, lookup process is
    '/* exactly the same..
    For i = 1 To 1000
        With c_Local
            .Add "msmin.exe"
            .Add "pinball.exe"
            .Add "wmplayer.exe"
            .Add "slip.scp"
            .Add "actmovie.exe"
            .Add "msmom.dll"
            .Add "nac.dll"
            .Add "mobsync.exe"
            .Add "freecell.exe"
            .Add "diskmgmt.msc"
        End With
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '/* release resource
    Set cEngine = Nothing
End Sub


'/* This is my last post on vb/psc.
'/* I am living in the world of C# now..
'/* .net is really not so bad, (you might want to think about switching..)
'/* see ya, and good luck
'/* John
