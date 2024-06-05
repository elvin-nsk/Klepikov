VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpreadForCutterView 
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3630
   OleObjectBlob   =   "SpreadForCutterView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpreadForCutterView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
' # State

Private Const MIN_VALUE As Double = 0

Public Parts As Parts

Public TopOffset As TextBoxHandler
Public LeftOffset As TextBoxHandler
Public RightOffset As TextBoxHandler
Public BottomOffset As TextBoxHandler
Public SpreadDistance As TextBoxHandler
Public MaxSheetWidth As TextBoxHandler
Public PlacesCount As TextBoxHandler

Public IsOk As Boolean
Public IsCancel As Boolean

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = "–‡ÒÍÎ‡‰ ‰Îˇ ÂÁ‡" & " (v" & APP_VERSION & ")"
    btnOk.Default = True
    Set TopOffset = _
        TextBoxHandler.New_(tbTopOffset, TextBoxTypeDouble, MIN_VALUE)
    Set LeftOffset = _
        TextBoxHandler.New_(tbLeftOffset, TextBoxTypeDouble, MIN_VALUE)
    Set RightOffset = _
        TextBoxHandler.New_(tbRightOffset, TextBoxTypeDouble, MIN_VALUE)
    Set BottomOffset = _
        TextBoxHandler.New_(tbBottomOffset, TextBoxTypeDouble, MIN_VALUE)
    Set SpreadDistance = _
        TextBoxHandler.New_(tbSpreadDistance, TextBoxTypeDouble)
    Set MaxSheetWidth = _
        TextBoxHandler.New_(tbMaxSheetWidth, TextBoxTypeDouble, 0)
    Set PlacesCount = _
        TextBoxHandler.New_(tbPlacesCount, TextBoxTypeLong, 0)
End Sub

'===============================================================================
' # Handlers

Private Sub UserForm_Activate()
    RefreshPlaces
End Sub

Private Sub tbBottomOffset_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbLeftOffset_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbRightOffset_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbSpreadDistance_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbTopOffset_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbMaxSheetWidth_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub tbPlacesCount_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    RefreshPlaces
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormŒ ()
    Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Hide
    IsCancel = True
End Sub

'===============================================================================
' # Helpers

Private Sub RefreshPlaces()
    Dim Places As CalculatedPlaces: Places = _
        CalcPlaces( _
            PlacesCount, Parts, LeftOffset, RightOffset, _
            SpreadDistance, MaxSheetWidth _
        )
    Dim Text As String: Text = Places.OnSheet
    If Places.Additional > 0 Then _
        Text = Text & " + " & Places.Additional & " ‰ÓÔ."
    lbCalcCount.Caption = Text
End Sub

'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
