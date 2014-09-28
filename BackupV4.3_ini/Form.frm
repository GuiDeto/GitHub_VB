VERSION 5.00
Begin VB.Form Backup 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistema de Backup"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6945
   FillStyle       =   6  'Cross
   ForeColor       =   &H00000000&
   Icon            =   "Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form.frx":0CCA
   ScaleHeight     =   5100
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_DiaSem 
      Height          =   285
      Left            =   240
      TabIndex        =   27
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3000
      Top             =   6960
   End
   Begin VB.TextBox txt_command_shut 
      Height          =   285
      Left            =   240
      TabIndex        =   24
      Top             =   6840
      Width           =   2055
   End
   Begin VB.TextBox NomeBkp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   1920
      TabIndex        =   20
      Text            =   "Backup"
      Top             =   2040
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Efetuar backup &agora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox txt_Pasta_Winrar 
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   7560
      Width           =   6255
   End
   Begin VB.ComboBox txt_Metodo 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Form.frx":7EFD
      Left            =   1920
      List            =   "Form.frx":7F0A
      TabIndex        =   16
      Text            =   "SELECIONE"
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txt_Parametros 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   1920
      TabIndex        =   14
      Top             =   3495
      Width           =   4935
   End
   Begin VB.TextBox txt_Arquivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "config.ini"
      Top             =   3855
      Width           =   1215
   End
   Begin VB.TextBox txt_Log_Dir 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   3015
   End
   Begin VB.TextBox txt_Log_File 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "teste"
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2400
      Top             =   6960
   End
   Begin VB.TextBox txt_Hora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      TabIndex        =   6
      Text            =   "10:10:10"
      Top             =   4200
      Width           =   2415
   End
   Begin VB.TextBox txt_Hora_Backup 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3870
      Width           =   1125
   End
   Begin VB.TextBox txt_Dest 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1920
      TabIndex        =   1
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox txt_Ori 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   2400
      Width           =   4935
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Metodo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   28
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label txt_DiaSemana 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5040
      TabIndex        =   26
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "do dia da Semana:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Desligar as:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label txt_hora_desligar 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      TabIndex        =   22
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome backup:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vers„o: 4.3"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   6000
      TabIndex        =   19
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1065
      Left            =   1560
      Picture         =   "Form.frx":7F2A
      Top             =   120
      Width           =   3705
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Parametros:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   555
      TabIndex        =   15
      Top             =   3525
      Width           =   1575
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivo de Config:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3615
      TabIndex        =   13
      Top             =   3915
      Width           =   2280
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Desenvolvido por: Guilherme Cunha Milanez"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   420
      Left            =   120
      TabIndex        =   10
      Top             =   4725
      Width           =   4815
   End
   Begin VB.Label txt_status 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   9
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Log:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   8
      Top             =   3165
      Width           =   615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Hora do Backup:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3915
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1035
      TabIndex        =   4
      Top             =   2805
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BackStyle       =   0  'Transparent
      Caption         =   "Origem:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   2445
      Width           =   1095
   End
End
Attribute VB_Name = "Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal Secao As String, ByVal Parametro As Any, ByVal padrao As String, ByVal variavel As String, ByVal tam As Long, ByVal Arquivo As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Secao As String, ByVal Parametro As Any, ByVal valor As Any, ByVal Arquivo As String) As Long
Public ConfigFile As String
Public Origem As String
Public Destino As String
Public LogFile As String
Public HoraBkp As String
Public Parametro As String
Public Metodo As String
Public Winrar As String
Public NomeArquivo As String
Public HoraDesligar As String
Public ComandoDesligar As String
Public DiaSemanaDesligar As String
Public DataHoje As String
Private Sub Form_Load()
Dim retlen As Long
    
    'VERIFICA A EXISTEMCIA DO ARQUIVO DE CONFIGURA«√O
    If Dir$(App.Path + "\config.ini") = "" Then
        MsgBox "CadÍ a PORRA do arquivo de configuraÁ„o!!!", vbCritical
        Exit Sub
    Else
        ConfigFile = App.Path + "\config.ini"
    End If

Origem = String(256, 0)
Destino = String(256, 0)
LogFile = String(256, 0)
HoraBkp = String(256, 0)
Parametro = String(256, 0)
Metodo = String(256, 0)
Winrar = String(256, 0)
NomeArquivo = String(256, 0)
HoraDesligar = String(256, 0)
ComandoDesligar = String(256, 0)
DiaSemanaDesligar = String(256, 0)


retlen = GetPrivateProfileString("CONFIGURA«‡O", "Origem", "", Origem, 256, ConfigFile)
Origem = Left(Origem, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "Destino", "", Destino, 256, ConfigFile)
Destino = Left(Destino, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "LogFile", "", LogFile, 256, ConfigFile)
LogFile = Left(LogFile, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "HoraBkp", "", HoraBkp, 256, ConfigFile)
HoraBkp = Left(HoraBkp, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "Parametro", "", Parametro, 256, ConfigFile)
Parametro = Left(Parametro, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "Metodo", "", Metodo, 256, ConfigFile)
Metodo = Left(Metodo, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "Winrar", "", Winrar, 256, ConfigFile)
Winrar = Left(Winrar, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "NomeArquivo", "", NomeArquivo, 256, ConfigFile)
NomeArquivo = Left(NomeArquivo, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "HoraDesligar", "", HoraDesligar, 256, ConfigFile)
HoraDesligar = Left(HoraDesligar, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "ComandoDesligar", "", ComandoDesligar, 256, ConfigFile)
ComandoDesligar = Left(ComandoDesligar, retlen)

retlen = GetPrivateProfileString("CONFIGURA«‡O", "DiaSemanaDesligar", "", DiaSemanaDesligar, 256, ConfigFile)
DiaSemanaDesligar = Left(DiaSemanaDesligar, retlen)

'Verifica se tem algum campo em branco
If Origem = "" Then
    MsgBox "O par‚metro Origem n„o foi encontrado!", vbCritical
End If

If Destino = "" Then
    MsgBox "O par‚metro ( Destino ) n„o foi encontrado!", vbCritical
End If

If LogFile = "" Then
    MsgBox "O par‚metro ( LogFile ) n„o foi encontrado!", vbCritical
End If

If HoraBkp = "" Then
    MsgBox "O par‚metro ( HoraBkp ) n„o foi encontrado!", vbCritical
End If

If Parametro = "" Then
    MsgBox "O par‚metro ( Parametro ) n„o foi encontrado!", vbCritical
End If

If Metodo = "" Then
    MsgBox "O par‚metro ( Metodo ) n„o foi encontrado!", vbCritical
End If

If Winrar = "" Then
    MsgBox "O par‚metro ( Winrar ) n„o foi encontrado!", vbCritical
End If

If NomeArquivo = "" Then
    MsgBox "O par‚metro ( NomeArquivo ) n„o foi encontrado!", vbCritical
End If

If HoraDesligar = "" Then
    MsgBox "O par‚metro ( HoraDesligar ) n„o foi encontrado!", vbCritical
End If

If ComandoDesligar = "" Then
    MsgBox "O par‚metro ( ComandoDesligar ) n„o foi encontrado!", vbCritical
End If

If DiaSemanaDesligar = "" Then
    MsgBox "O par‚metro ( DiaSemanaDesligar ) n„o foi encontrado!", vbCritical
End If

'Atribui as configuraÁoes nos campos de texto
txt_Ori.Text = Origem
txt_Dest.Text = Destino
txt_Log_Dir.Text = LogFile
txt_Hora_Backup.Text = HoraBkp
txt_Parametros.Text = Parametro
txt_Metodo.Text = Metodo
txt_Pasta_Winrar.Text = Winrar
NomeBkp.Text = NomeArquivo
txt_hora_desligar.Caption = HoraDesligar
txt_command_shut.Text = ComandoDesligar
txt_DiaSem.Text = DiaSemanaDesligar
    
    'Verifica se as pastas existem!
    If Dir(txt_Ori.Text, vbDirectory) = "" Then
        If (MsgBox("A pasta de ORIGEM ( " & txt_Ori.Text & " ) n„o existe!!" & vbCrLf & "Posso Criala?", vbYesNo + vbQuestion, "Alert!") = vbYes) Then
        MkDir txt_Ori.Text
        End If
    End If
    
    If Dir(txt_Dest.Text, vbDirectory) = "" Then
        If (MsgBox("A pasta de DESTINO ( " & txt_Dest.Text & " ) n„o existe!" & vbCrLf & "Posso Criala?", vbYesNo + vbQuestion, "Alert!") = vbYes) Then
        MkDir txt_Dest.Text
        End If
    End If
    
    If Dir(txt_Log_Dir.Text, vbDirectory) = "" Then
        If (MsgBox("A pasta de LOG ( " & txt_Log_Dir.Text & " ) n„o existe!" & vbCrLf & "Posso Criala?", vbYesNo + vbQuestion, "Alert!") = vbYes) Then
        MkDir txt_Log_Dir.Text
        End If
    End If
    
    'Verifica se o Winrar est· instalado no sistema
    If Dir$(txt_Pasta_Winrar.Text) = "" Then
        MsgBox "O programa ( WINRAR ) usado para compactar os arquivos," & vbCrLf & "N„o foi encontrado em seu sistema!", vbInformation
    End If
    
End Sub

Private Sub Command1_Click()
On Error Resume Next
'Escolhe a condiÁ„o para ele copiar o arquivo
If txt_Metodo.Text = "xcopy" Then
    Shell "xcopy " & txt_Ori & " " & txt_Dest & " " & txt_Parametros.Text, vbNormalFocus
ElseIf txt_Metodo.Text = "Compactar" Then
    Shell txt_Pasta_Winrar & " " & txt_Parametros.Text & " " & txt_Dest.Text & NomeBkp.Text & "-" & DataHoje & ".rar " & txt_Ori.Text, vbNormalFocus
Else
    Shell "robocopy " & txt_Ori & " " & txt_Dest & " " & txt_Parametros.Text & txt_Log_Dir.Text & txt_Log_File.Text, vbNormalFocus
End If
'Fecha as condiÁıes
End Sub

Private Sub Image1_Click()
    Config.Show
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    
    DataHoje = Format(Date, "ddmmyyyy") 'PEGA A DATA DO SISTEMA PARA COLOCAR NO LOG
    txt_Log_File.Text = "Log_Dia-" & DataHoje & ".log" 'ESCREVE A DATA NO TEXT LOG
    txt_Hora.Text = Format(Time, "hh:mm:ss")
    
If txt_Hora.Text = txt_Hora_Backup.Text Then
    txt_status.Caption = "Efetuando BACKUP programado: " & txt_Hora_Backup

'Escolhe a condiÁ„o para ele copiar o arquivo
If txt_Metodo.Text = "xcopy" Then
    Shell "xcopy " & txt_Ori & " " & txt_Dest & " " & txt_Parametros.Text, vbNormalFocus
ElseIf txt_Metodo.Text = "Compactar" Then
    Shell Chr$(34) & txt_Pasta_Winrar.Text & Chr$(34) & " " & txt_Parametros.Text & " " & txt_Dest.Text & NomeBkp.Text & "-" & DataHoje & ".rar " & txt_Ori.Text, vbNormalFocus
Else
    Shell "robocopy " & txt_Ori & " " & txt_Dest & " " & txt_Parametros.Text & txt_Log_Dir.Text & txt_Log_File.Text, vbNormalFocus
End If
'Fecha as condiÁıes
Else
'Mostra No titulo do FORM o metodo e a hora do backup
    Backup.Caption = txt_Metodo & " as : " & txt_Hora_Backup & " e desligar as: " & txt_hora_desligar.Caption
    txt_status.Caption = ""
  
    'Alterar parametros de acordo com o metodo
    If txt_Metodo.Text = "Compactar" Then
        txt_Parametros.Text = "a"
  
    ElseIf txt_Metodo.Text = "xcopy" Then
        txt_Parametros.Text = "/E /V /C /R /M /I"
        txt_Metodo.BackColor = &HFF&
    
    ElseIf txt_Metodo.Text = "robocopy" Then
        txt_Parametros.Text = "/E /COPY:DATSOU /R:2 /W:2 /V /ETA /NP /LOG:"
        txt_Metodo.BackColor = &HFF&
    End If
    
End If
End Sub
'HORA DE DESLIGAR
Private Sub Timer2_Timer()
Dim DiaSem As String
DiaSem = Weekday(Date)

Select Case txt_DiaSem.Text
Case 0
    txt_DiaSemana.Caption = "Todos"
Case 1
    txt_DiaSemana.Caption = "Domingo"
Case 2
    txt_DiaSemana.Caption = "Segunda"
Case 3
    txt_DiaSemana.Caption = "TerÁa"
Case 4
    txt_DiaSemana.Caption = "Quarta"
Case 5
    txt_DiaSemana.Caption = "Quinta"
Case 6
    txt_DiaSemana.Caption = "Sexta"
Case 7
    txt_DiaSemana.Caption = "Sabado"
End Select

If txt_Hora.Text = txt_hora_desligar.Caption And (DiaSem = txt_DiaSem.Text Or txt_DiaSem.Text = "0") Then
    Shell txt_command_shut.Text, vbHide
End If
End Sub
'ABRIR PASTAS E ARQUIVO COM DOIS CLICKS
Private Sub txt_Ori_DblClick()
On Error Resume Next
Dim OpenOrig
OpenOrig = Shell("explorer " & txt_Ori.Text, vbNormalFocus)
End Sub
'DOIS CLICKS PARA ABRIR A PASTA DE DESTINO DOS BACKUPS
Private Sub txt_Dest_DblClick()
On Error Resume Next
Dim OpenOrig
OpenOrig = Shell("explorer " & txt_Dest.Text, vbNormalFocus)
End Sub
'DOIS CLICKS PARA ABRIR O ARQUIVO DE CONFIGURA«√O DO SISTEMA
Private Sub txt_Arquivo_DblClick()
Dim AbrirConfig
AbrirConfig = Shell("notepad.exe " & App.Path & "\" & txt_Arquivo.Text, vbNormalFocus)
End Sub
'DOIS CLICKS PARA ABRIR O DIRETORIO DE LOG
Private Sub txt_Log_Dir_DblClick()
On Error Resume Next
Dim OpenLog
OpenLog = Shell("explorer " & txt_Log_Dir.Text, vbNormalFocus)
End Sub
