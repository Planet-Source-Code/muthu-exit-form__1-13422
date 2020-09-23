VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Exit Form"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Click me to Exit"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        
Do: DoEvents
    
    If Me.WindowState = 1 Or Me.WindowState = 2 Then
        Me.WindowState = 0
        Me.Top = (Screen.Height - Me.Height) / 2
        Me.Left = (Screen.Width - Me.Width) / 2
    End If
        
        Me.Left = Me.Left + 1&
        Loop Until Me.Left > Screen.Width
        
    End

End Sub

Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
   
End Sub

Private Sub Form_Resize()

    If Me.Height <> 3600 Or Me.Width <> 4800 Or Me.WindowState = 1 Or 2 Then
        
        Me.WindowState = 0
        Me.Height = 3600
        Me.Width = 4800
        Me.Top = (Screen.Height - Me.Height) / 4 * Rnd
        Me.Left = (Screen.Width - Me.Width) / 4 * Rnd
    End If


    If Me.WindowState = 1 Or Me.WindowState = 2 Then
        Me.WindowState = 0
        Me.Top = (Screen.Height - Me.Height) / 4 * Rnd
        Me.Left = (Screen.Width - Me.Width) / 4 * Rnd
    End If

        Command1.Top = (Me.Height - Command1.Height) / 2
        Command1.Left = (Me.Width - Command1.Width) / 2

End Sub
