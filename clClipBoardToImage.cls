VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clClipBoardToImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'
'
'
'--      ___           ___           ___           ___           ___                                             ___           ___
'--     /  /\         /  /\         /__/\         /  /\         /__/\                   ___        ___          /  /\         /  /\
'--    /  /::\       /  /::\        \  \:\       /  /::\        \  \:\                 /__/\      /  /\        /  /:/        /  /::\
'--   /  /:/\:\     /  /:/\:\        \  \:\     /  /:/\:\        \  \:\                \  \:\    /  /:/       /  /:/        /  /:/\:\
'--  /  /:/~/:/    /  /:/  \:\   _____\__\:\   /  /:/~/::\   _____\__\:\                \  \:\  /__/::\      /  /:/  ___   /  /:/  \:\
'-- /__/:/ /:/___ /__/:/ \__\:\ /__/::::::::\ /__/:/ /:/\:\ /__/::::::::\           ___  \__\:\ \__\/\:\__  /__/:/  /  /\ /__/:/ \__\:\
'-- \  \:\/:::::/ \  \:\ /  /:/ \  \:\~~\~~\/ \  \:\/:/__\/ \  \:\~~\~~\/          /__/\ |  |:|    \  \:\/\ \  \:\ /  /:/ \  \:\ /  /:/
'--  \  \::/~~~~   \  \:\  /:/   \  \:\  ~~~   \  \::/       \  \:\  ~~~           \  \:\|  |:|     \__\::/  \  \:\  /:/   \  \:\  /:/
'--   \  \:\        \  \:\/:/     \  \:\        \  \:\        \  \:\                \  \:\__|:|     /__/:/    \  \:\/:/     \  \:\/:/
'--    \  \:\        \  \::/       \  \:\        \  \:\        \  \:\                \__\::::/      \__\/      \  \::/       \  \::/
'--     \__\/         \__\/         \__\/         \__\/         \__\/                    ~~~~                   \__\/         \__\/
'
'
'Criado por Ronan Vico , na TECNUN
'dia 24 01 2020
'VBA PALADIN
'-----------------------------------------------
Private m_Folder                                As String
Private m_FileName                              As String
Private oXL                                     As Object
Private oPrintScreen                            As clPrintScreen
Public outlookEmail                             As Object
'-----------------------------------------------
Private width As Long
Private height As Long
'-----------------------------------------------

Public Property Get Caminho() As Variant
    Caminho = Pasta & "\" & Arquivo
End Property

Public Property Get Arquivo() As Variant
    Arquivo = m_FileName
End Property

Public Property Get Pasta() As Variant
    Pasta = m_Folder
End Property
'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com - Rotina: Public Sub PrintTheScreen()
' Data......: 24/01/2020
' Descrição.: Realiza Print Screen da tela completa
'---------------------------------------------------------------------------------------
Public Function PrintTheScreen()
    oPrintScreen.PrintTheScreen
End Function

'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com - Rotina: Public Sub PrintTheActivedScreenAltPrint()
' Data......: 24/01/2020
' Descrição.: Realiza PrintScreen apenas da tela que estiver ativa ALT + PRINT
'---------------------------------------------------------------------------------------
Public Function PrintTheActivedScreenAltPrint()
    oPrintScreen.PrintTheActivedScreenAltPrint
End Function


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com - Rotina: Public Function SaveAsImage()
' Data......: 27/01/2020
' Descrição.: Salva como imagem , o que estiver no clipboard , seja printscreen ou range
'---------------------------------------------------------------------------------------
Public Function SaveAsImage() As clClipBoardToImage
    Set SaveAsImage = Me
    On Error GoTo TError
    oXL.EnableEvents = False
    oXL.DisplayAlerts = False
    'Stop
Dim ws                                 As Object 'Worksheet
Dim obChart                            As Object 'Excel.Chart
Dim i                                  As Long

    'adciona uma planilha nova temporaria
    Set ws = oXL.ActiveWorkbook.Worksheets.Add
    With ws
        .Visible = 0  'xlHidden
        '        'Cola no range A1
        .Range("A1").PasteSpecial
        'Caso não seja Imagem (print screen)  , é um range normal
        If .Shapes.Count = 0 Then
            'Range normal devemos colar o column widths
            'Copiar como imagem , e colar como imagem para ser igual a uma foto / imagem / PrintScreen
            .Range("A1").PasteSpecial 8 'xlPasteColumnWidths
            .UsedRange.CopyPicture 1, -4147 ' xlScreen, xlPicture
            .Range("A1").PasteSpecial
        End If
        'Adcionar um bojeto do tipo ChartObject para podemos exportar posteriormente como imagem
        Call .ChartObjects.Add(50, 50, .Shapes(1).width + 5, .Shapes(1).height + 5)
        Set obChart = .ChartObjects(.ChartObjects.Count).Chart
    End With

    With obChart
        'Pegamos o objeto adcionado e colamos dentro dele o que queremos exportar como imagem
        ws.Shapes(1).CopyPicture 1, -4147 'xlScreen,xlPicture
        'ws.Visible = 0
        'Mas Ronan , Você colocando .select ? Meu deus , porém
        'o Excel na build posterior a 16.14 , o objeto Chart não funciona com planilhas escondidas.
        'tudo há um porque.
        .Parent.Select
        .Paste
        .Parent.width = .Shapes(1).width + 10
        .Parent.height = .Shapes(1).height + 10
        .Shapes(1).Top = 5
        .Shapes(1).Left = 5
        width = .Shapes(1).width
        height = .Shapes(1).height
        'salva o arquivo em GIF no caminho especificado
        m_FileName = VBA.Format(VBA.Now(), "yyyymmddhhmmss" & oXL.hwnd) & ".gif"
        .Export Filename:=Caminho, Filtername:="gif"
        For i = 1 To 1000: VBA.DoEvents: Next i
    End With

    If VBA.Len(VBA.Dir(Caminho)) = 0 Then
        Err.Raise 667, "SaveAsImage()", "File Not Created!"
    End If

Fim:
    On Error Resume Next
    ws.Delete
    oXL.EnableEvents = True
    oXL.DisplayAlerts = True
    Exit Function
    Resume
TError:
    'Call MOSTRAR_ERRO(Err.Number, Err.Description, "ToEmail()")
    Stop
    GoTo Fim
End Function


'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com - Rotina: Public Function ToEmail() As Object
' Data......: 24/01/2020
' Descrição.: Insere Imagem no corpo do email
'---------------------------------------------------------------------------------------
Public Function ToEmail(Optional mailToPutPicture As Object) As Object
    On Error GoTo TError
Dim ol                                 As Object
Dim em                                 As Object
Dim HTMLDoc                            As Object
Dim NEW_CODE_HTML                      As String

    'Verifica se foi passado parametro de email ja pronto
    If mailToPutPicture Is Nothing Then
        If outlookEmail Is Nothing Then
            On Error Resume Next
            Set ol = GetObject(, "outlook.application")
            If ol Is Nothing Then Set ol = CreateObject("outlook.application")
            On Error GoTo TError
            Set em = ol.CreateItem(0)
            em.htmlbody = ""
            Set outlookEmail = em
        Else
            Set em = outlookEmail
        End If
    Else
        Set em = mailToPutPicture
    End If


    '    NEW_CODE_HTML = "<br><img  src='cid:" & Me.Arquivo & "'> <br>"
    NEW_CODE_HTML = "<br><img width=" & width & " height=" & height & " src='cid:" & Me.Arquivo & "'> <br>"
    With em
        .Attachments.Add Me.Caminho
        Set HTMLDoc = CreateObject("htmlFIle")
        HTMLDoc.Body.innerHTML = .htmlbody
        HTMLDoc.Body.innerHTML = HTMLDoc.Body.innerHTML & "<br>" & NEW_CODE_HTML
        .htmlbody = HTMLDoc.Body.parentElement.outerhtml
    End With
    Set ToEmail = em
Fim:
    Set HTMLDoc = Nothing
    Exit Function
    Resume
TError:
    Call MOSTRAR_ERRO(Err.Number, Err.Description, "ToEmail()")
    GoTo Fim
End Function



Public Function RangeToClipboard(ByRef rgToCopy As Object)
On Error GoTo TError
    rgToCopy.Copy
Fim:
     Exit Function
TError:
     Call MOSTRAR_ERRO(Err.Number, Err.Description, "RangeToClipboard()")
     GoTo Fim
End Function

'---------------------------------------------------------------------------------------
' Autor.....: RONAN VICO
' Contato...: ronanvico@hotmail.com - Rotina: Private Sub Class_Initialize()
' Data......: 27/01/2020
' Descrição.:
'---------------------------------------------------------------------------------------
Private Sub Class_Initialize()
    On Error GoTo TError

    If Application.Name = "Microsoft Excel" Then
        Set oXL = Application
    Else
        Set oXL = CreateObject("Excel.Application")
        oXL.EnableEvents = False
        oXL.Visible = False
        oXL.DisplayAlerts = False
        oXL.Workbooks.Add
    End If

    Set oPrintScreen = New clPrintScreen
    m_Folder = VBA.Environ("TEMP") & "\tr2a872vic10"
    If VBA.Len(VBA.FileSystem.Dir(m_Folder, vbDirectory)) = 0 Then
        Call VBA.MkDir(m_Folder)
    End If
Fim:
    Exit Sub
    Resume
TError:
    Call MOSTRAR_ERRO(Err.Number, Err.Description, "Class_Initialize()")
    GoTo Fim
End Sub


Private Sub Class_Terminate()
On Error Resume Next
    Set oXL = Nothing
    Set oPrintScreen = Nothing
    Set outlookEmail = Nothing
    Kill Me.Caminho
End Sub



'--RONAN RAPHAEL VICO VBA PALADINO 05 02 2020
'
'--_________________________¶¶¶¶¶
'--_______________________¶¶¶¶¶11¶¶
'--_____________________¶¶¶¶111111¶¶
'--___________________¶¶¶111111111¶¶¶¶
'--__________________¶¶¶11111111111¶¶¶¶¶¶
'--_________________¶¶111111111111111111¶¶¶
'--________________¶¶1111111111111111111111¶
'--_______________¶¶1111111111111111111111¶¶
'--_____________¶¶¶111111111111111111111¶¶¶
'--__________¶¶¶¶¶1111111111111111111111¶¶¶
'--_________¶¶¶111111111111111111111111111¶¶¶
'--_______¶¶¶1111111111111111111111111111111¶¶¶
'--______¶¶11111111111111111111111111111111111¶¶
'--_____¶¶1111111111111111111111111111111111111¶¶
'--____¶¶111111111111111111111111111111111111111¶¶
'--___¶¶11111111111111111111111111111111111111111¶¶
'--__¶¶1111111111111111111111111¶11111111111111111¶¶
'--__¶¶¶111111111111111111111¶¶¶¶11111111111111111¶¶
'--_¶¶¶¶¶¶¶¶1111111111111¶¶¶¶¶¶¶¶¶11111111111111111¶¶
'--_¶¶¶¶¶¶¶¶¶¶¶¶111111¶¶¶¶¶¶¶¶¶¶¶¶¶1111111111111111¶¶
'--¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶111111111111111111111¶
'--¶¶1111¶________¶¶¶_________¶111111111111111111111¶
'--¶¶111¶_____¶¶¶_¶¶____¶¶_____¶11111111111111111111¶
'--¶¶111¶_____¶¶¶_¶¶___¶¶¶¶____¶11111111111111111111¶
'--¶¶1111¶______¶¶¶¶____¶¶____¶111111111111111111111¶
'--_¶1111¶¶___¶¶¶___¶¶_______¶¶11111111111111111111¶¶
'--_¶¶11111¶¶¶¶_______¶¶¶¶¶¶¶1111111111111111111111¶¶
'--__¶1111¶¶¶___________¶¶¶11111111111111111111111¶¶
'--__¶¶1¶¶¶_______________¶¶¶¶¶¶¶¶¶¶¶1111111111111¶¶
'--___¶¶¶¶___$$$$$$$$$$$$$¶¶_____¶¶¶¶¶¶¶111111111¶¶
'--____¶¶1¶¶$$$$$$$$_____¶¶____________¶¶¶111111¶¶
'--_____¶¶11¶¶_________¶¶¶_______________¶¶¶111¶¶
'--______¶¶11¶¶¶______¶¶___________________¶11¶¶
'--_______¶¶1¶__¶¶¶__¶¶____________________¶¶¶¶
'--_________¶¶¶____¶¶_____________________¶¶¶
'--___________¶¶¶_______________________¶¶¶
'--____________¶¶¶¶_________________¶¶¶¶
'--________________¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶
'
                                                                              