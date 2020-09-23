VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13140
   LinkTopic       =   "Form1"
   ScaleHeight     =   540
   ScaleWidth      =   13140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6000
      Top             =   4740
   End
   Begin VB.PictureBox picTicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   0
      Width           =   8175
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         Caption         =   "TWgjp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   300
         MouseIcon       =   "Form1.frx":0000
         TabIndex        =   1
         Top             =   60
         Width           =   555
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   5760
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2, SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1, HWND_NOTOPMOST = -2

Dim NewsItem() As udtNewsItem

Private Sub Form_Click()
    Unload Me
End Sub
Sub SetTopmostWindow(ByVal hwnd As Long, Optional topmost As Boolean = True)
    Const HWND_NOTOPMOST = -2
    Const HWND_TOPMOST = -1
    Const SWP_NOMOVE = &H2
    Const SWP_NOSIZE = &H1
    SetWindowPos hwnd, IIf(topmost, HWND_TOPMOST, HWND_NOTOPMOST), 0, 0, 0, 0, _
        SWP_NOMOVE + SWP_NOSIZE
End Sub
Private Sub Form_Load()
    Me.Width = picTicker.Width
    Me.Left = Screen.Width - Me.Width - 920
    Me.BackColor = picTicker.BackColor
    Me.Top = 0
    
    SetTopmostWindow Me.hwnd
    
    LoadPage
    Timer1.Enabled = True
    
End Sub
Sub LoadPage()
Dim sPage As String
    sPage = Inet1.OpenURL("http://tickers.bbc.co.uk/tickerdata/story3.dat")
    
    ExtractNews sPage

    SetupTicker

End Sub
Sub SetupTicker()
Dim X As Integer
    picTicker.Font = lblItem(0).Font
    picTicker.FontSize = lblItem(0).FontSize
    lblItem(0).Height = picTicker.TextHeight("WTQjpg")
    picTicker.Height = lblItem(0).Height + 60
    Me.Height = picTicker.Height
    If lblItem.Count > 1 Then
        For X = 1 To lblItem.Count - 1
            Unload lblItem(X)
        Next
    End If
    For X = 0 To UBound(NewsItem)
        If X > 0 Then
            Load lblItem(X)
            lblItem(X).Visible = True
            lblItem(X).Left = lblItem(X - 1).Left + lblItem(X - 1).Width + 220
        Else
            lblItem(X).Left = 0
            lblItem(X).Top = 10
        End If
        With lblItem(X)
            If Len(NewsItem(X).URL) > 0 Then
                .MousePointer = vbCustom
                .ForeColor = RGB(51, 51, 152)
            Else
                .ForeColor = vbBlack
            End If
            .Caption = NewsItem(X).Headline
            picTicker.FontBold = lblItem(0).FontBold
            .Width = picTicker.TextWidth(.Caption)

        End With
    Next
    picTicker.Width = lblItem(X - 1).Left + lblItem(X - 1).Width + 240
End Sub
Sub ExtractNews(sIn As String)
Dim X As Long
Dim Y As Long
Dim SrchStr As String
    ReDim NewsItem(0)
    X = InStr(sIn, "STORY ")
    Y = X
    Do
        ReDim Preserve NewsItem(UBound(NewsItem) + 1)
        With NewsItem(UBound(NewsItem) - 1)
            SrchStr = "HEADLINE "
            X = InStr(Y, sIn, SrchStr)
            If X > 0 Then
                Y = InStr(X, sIn, vbLf)
                .Headline = Mid(sIn, X + Len(SrchStr), Y - X - Len(SrchStr))
                
                SrchStr = "URL "
                X = InStr(Y, sIn, SrchStr)
                Y = InStr(X, sIn, vbLf)
                If Y - X = Len(SrchStr) Then 'no url
                    .URL = ""
                Else
                    .URL = Mid(sIn, X + Len(SrchStr), Y - X - Len(SrchStr))
                End If
            End If
        End With
    Loop Until X = 0
    If UBound(NewsItem) > 0 Then ReDim Preserve NewsItem(UBound(NewsItem) - 1)
End Sub

Private Sub lblItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Unload Me
    End
Else
    If Len(NewsItem(Index).URL) > 0 Then
        ShellExecute 0, vbNullString, NewsItem(Index).URL, vbNullString, vbNullString, vbNormalFocus
    End If
End If
End Sub

Private Sub picTicker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    Unload Me
    End
End If

End Sub

Private Sub Timer1_Timer()
    picTicker.Left = picTicker.Left - 10
    If picTicker.Left + picTicker.Width < 0 Then
        Timer1.Enabled = False
        LoadPage
        picTicker.Left = Me.Width + 30
        Timer1.Enabled = True
    End If
End Sub
