VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpreadForCutterView 
   ClientHeight    =   4650
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

Public IsOk As Boolean
Public IsCancel As Boolean

'===============================================================================
' # Constructor

Private Sub UserForm_Initialize()
    Caption = "Ðàñêëàä äëÿ ðåçà" & " (v" & APP_VERSION & ")"
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
        TextBoxHandler.New_(tbSpreadDistance, TextBoxTypeDouble, MIN_VALUE)
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

Private Sub btnOk_Click()
    FormÎÊ
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================
' # Logic

Private Sub FormÎÊ()
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
    Dim Count As Long: Count = _
        CalcPlaces(Parts, LeftOffset, RightOffset, SpreadDistance)
    If Count < 0 Then Count = 0
    lbCalcCount.Caption = _
        "Êîëè÷åñòâî ìåñò: " & Count
    If Count = 0 Then
        btnOk.Enabled = False
    Else
        btnOk.Enabled = True
    End If
End Sub

'===============================================================================
' # Boilerplate

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Ñancel = True
        FormCancel
    End If
End Sub
