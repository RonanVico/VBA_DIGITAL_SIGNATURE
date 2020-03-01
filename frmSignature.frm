VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSignature 
   Caption         =   "ASSINE AQUI"
   ClientHeight    =   2685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10200
   OleObjectBlob   =   "frmSignature.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSignature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
 
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
 
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, _
ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long
 
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private OPENED As Boolean
 
Private Property Get circleSize() As Integer
    circleSize = 5
End Property


Private Sub btnENVIAR_Click()

    If OPENED = False Then
        Linha.Visible = True
        btnENVIAR.Caption = "ENVIAR"
        OPENED = True
        Exit Sub
    End If
    
    With New clClipBoardToImage
        'Printa a tela ativa
        .PrintTheActivedScreenAltPrint
        'Fecha o formulario
        Unload Me
        'Salva a imagem do print
        .SaveAsImage
        'Joguei no email
        .ToEmail
        'Com o meu email ja com a imagem
        With .outlookEmail
            'Mudei o remetente
            .sentOnBehalfOfName = "ronanvico"
            'Mostrei o email
            .display
        End With
    End With
End Sub


Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Position As POINTAPI
    'Verifica se o bot�o do mouse esquerdo est� ativado
    If Button = 1 Then
        'Pegar a posi��o do mouse
        GetCursorPos Position
        'Desenhar uma ellipse utilizando as posi��es do mouse
        Ellipse GetWindowDC(0), Position.X - circleSize(), Position.Y - circleSize(), _
            Position.X + circleSize(), Position.Y + circleSize()
    End If
End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                