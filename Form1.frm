VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hyperlink Sample"
   ClientHeight    =   1668
   ClientLeft      =   120
   ClientTop       =   360
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1668
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.planet-source-code.com"
      ForeColor       =   &H00FF0000&
      Height          =   192
      Left            =   672
      TabIndex        =   0
      Top             =   672
      Width           =   2496
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuNew 
         Caption         =   "Open in New Window"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Target As..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Target..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopyShortcut 
         Caption         =   "Copy Shortcut"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFavorites 
         Caption         =   "Add to Favorites..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const SM_CXDRAG = 68
Private Const SM_CYDRAG = 69
Private Const IDC_ARROW = 32512&
Private Const IDC_HAND = 32649&

Private Const STR_LINK          As String = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=31572&optCodeRatingValue=5&intUserRatingTotal=0&intNumOfUserRatings=0"
Private Const CR_LINK_NORMAL    As Long = vbBlue
Private Const CR_LINK_OVER      As Long = vbBlue
Private Const CR_LINK_PRESSED   As Long = vbRed

Private m_bOver             As Boolean
Private m_bPressed          As Boolean
Private m_bDrag             As Boolean
Private m_bEatMouseEvent    As Boolean
Private m_sDownX            As Single
Private m_sDownY            As Single


Private Sub FireClick()
    If MsgBox(vbCrLf & "Do you want to vote for this submission?" & vbCrLf & vbCrLf & vbCrLf & _
            "Note:" & vbTab & "Please, do this only if you feel that this submission is worth it! You will be navigated" & vbCrLf & _
            vbTab & "to the PSC page of this entry where you can choose your vote." & vbCrLf & vbCrLf & _
            vbTab & "Thank you in advance!" & vbCrLf & vbCrLf & _
            vbTab & "</wqw>", vbQuestion Or vbYesNo) = vbYes Then
        ShellExecute 0, "open", STR_LINK, "", "", 5
    End If
End Sub

Private Sub FirePopup()
    PopupMenu mnuPopup
End Sub

Private Sub ChangeLinkState( _
            ByVal bOver As Boolean, _
            ByVal bPressed As Boolean, _
            ByVal bDrag As Boolean)
    '--- save state
    m_bOver = bOver
    m_bPressed = bPressed
    m_bDrag = bDrag
    '--- fix link's underline and color
    Label1.Font.Underline = bOver
    Label1.ForeColor = IIf(m_bOver, _
                            IIf(m_bPressed, CR_LINK_PRESSED, CR_LINK_OVER), _
                            CR_LINK_NORMAL)
End Sub

Private Sub mnuCopyShortcut_Click()
    Clipboard.SetText STR_LINK, vbCFText
End Sub

Private Sub mnuOpen_Click()
    FireClick
End Sub

Private Sub Label1_DblClick()
    '--- mousedown not fired on dbl clicking
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bEatMouseEvent Then
        m_bEatMouseEvent = False
        Exit Sub
    End If
    m_sDownX = X
    m_sDownY = Y
    '--- press the link
    ChangeLinkState m_bOver, True, False
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 0 And Y >= 0 And X < Label1.Width And Y < Label1.Height Then
        SetCursor LoadCursor(0, IDC_HAND)
        If Not m_bOver Then
            '--- "mouseover" the link
            ChangeLinkState True, m_bPressed, m_bDrag
            '--- set capture (if mouse not already pressed)
            If Not m_bPressed Then
                mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
                m_bEatMouseEvent = True
            End If
            Exit Sub
        End If
    Else
        SetCursor LoadCursor(0, IDC_ARROW)
        If m_bOver Then
            '--- "mouseout" the link
            ChangeLinkState False, m_bPressed, m_bDrag
            '--- release capture (if mouse not pressed)
            If Not m_bPressed Then
                mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                m_bEatMouseEvent = True
            End If
            Exit Sub
        End If
    End If
    '--- cancel mode support
    If Not m_bDrag And (m_bOver Or m_bPressed) And (GetCapture() = 0) Then
        Debug.Print "CANCELMODE "; Timer
        '--- reset link state
        ChangeLinkState False, False, False
    End If
    '--- drag & drop support
    If m_bPressed And Not m_bDrag And Button = vbLeftButton Then
        '--- check if enough movement occured to start dragging
        If Abs(X - m_sDownX) > ScaleX(GetSystemMetrics(SM_CXDRAG), vbPixels) _
                Or Abs(Y - m_sDownY) > ScaleY(GetSystemMetrics(SM_CYDRAG), vbPixels) Then
            Label1.OLEDrag
            Exit Sub
        End If
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '--- eat mouse event if signalled
    If m_bEatMouseEvent Then
        m_bEatMouseEvent = False
        Exit Sub
    End If
    '--- fix of problem with VB implementation ot setcapture
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    m_bEatMouseEvent = True
    '--- will signal if released over the link
    If m_bOver Then
        '--- fire click/popup upon left/right button released over the link
        If Button = vbLeftButton Then
            FireClick
        ElseIf Button = vbRightButton Then
            FirePopup
        End If
    End If
    '--- reset link state
    ChangeLinkState False, False, False
End Sub

Private Sub Label1_OLECompleteDrag(Effect As Long)
    '--- reset link state
    ChangeLinkState False, False, False
End Sub

Private Sub Label1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    '--- undocumented effect vbDropEffectLink :-)))
    AllowedEffects = 4
    '--- pass link URL as plain text
    Data.SetData STR_LINK, vbCFText
    '--- set pressed and drag to on
    ChangeLinkState m_bOver, True, True
End Sub

