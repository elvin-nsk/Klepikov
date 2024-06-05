VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'===============================================================================
' Config v2023.11.19
' хранение, загрузка и сохранение настроек
'===============================================================================

Option Explicit

#Const ReadOnly = False
#Const DefaultContent = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public TopOffset As Double
Public LeftOffset As Double
Public RightOffset As Double
Public BottomOffset As Double
Public SpreadDistance As Double
Public MaxSheetWidth As Double
Public PlacesCount As Long
Public DeleteImages As Boolean

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.New_
    
        iniFilePrepare .Self, Action
        
        .Section = "SpreadForCutter"
        .KeyVar "TopOffset", TopOffset, 5
        .KeyVar "LeftOffset", LeftOffset, 5
        .KeyVar "RightOffset", RightOffset, 5
        .KeyVar "BottomOffset", BottomOffset, 5
        .KeyVar "SpreadDistance", SpreadDistance, 0
        .KeyVar "MaxSheetWidth", MaxSheetWidth, 339
        .KeyVar "PlacesCount", PlacesCount, 0
        .KeyVar "DeleteImages", DeleteImages, False
    
    End With
    
End Sub

'===============================================================================
'           * * * Контент, не предназначенный для изменения * * *
'===============================================================================
' # Constructors

Private Sub Class_Initialize()
    ForceLoad
End Sub

#If Not ReadOnly Then
Private Sub Class_Terminate()
    ForceSave
End Sub
#End If

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

#If Not ReadOnly Then
Public Sub ForceSave()
    Assign iniWrite
End Sub
#End If

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            #If DefaultContent Then
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
            #Else
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
            #End If
        End Select
    End With

End Sub

Private Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
