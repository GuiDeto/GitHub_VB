VERSION 5.00
Begin VB.Form Config 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_DiaSemanaDesligar 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   10
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txt_ComandoDesligar 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txt_HoraDesligar 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txt_NomeArquivo 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txt_Winrar 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   3000
      Width           =   4095
   End
   Begin VB.TextBox txt_Metodo 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txt_Parametro 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txt_HoraBkp 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txt_LogFile 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.TextBox txt_Destino 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox txt_Origem 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIGURA钦ES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   22
      Top             =   0
      Width           =   4335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Dia da Semana:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   21
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comando para Desligar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   4080
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de desligar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do arquivo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Winrar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   17
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   16
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LogFile:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Menu Config 
      Caption         =   "Configurar"
      Begin VB.Menu Save 
         Caption         =   "Salvar Dados"
      End
   End
End
Attribute VB_Name = "Config"
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
Private Sub Form_Load()
On Error Resume Next

Dim retlen As Long
ConfigFile = App.Path + "\config.ini"

'CARREGAR ARQUIVOS DE CONFIGURA敲O
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

retlen = GetPrivateProfileString("CONFIGURA青O", "Origem", "", Origem, 256, ConfigFile)
Origem = Left(Origem, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "Destino", "", Destino, 256, ConfigFile)
Destino = Left(Destino, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "LogFile", "", LogFile, 256, ConfigFile)
LogFile = Left(LogFile, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "HoraBkp", "", HoraBkp, 256, ConfigFile)
HoraBkp = Left(HoraBkp, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "Parametro", "", Parametro, 256, ConfigFile)
Parametro = Left(Parametro, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "Metodo", "", Metodo, 256, ConfigFile)
Metodo = Left(Metodo, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "Winrar", "", Winrar, 256, ConfigFile)
Winrar = Left(Winrar, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "NomeArquivo", "", NomeArquivo, 256, ConfigFile)
NomeArquivo = Left(NomeArquivo, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "HoraDesligar", "", HoraDesligar, 256, ConfigFile)
HoraDesligar = Left(HoraDesligar, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "ComandoDesligar", "", ComandoDesligar, 256, ConfigFile)
ComandoDesligar = Left(ComandoDesligar, retlen)

retlen = GetPrivateProfileString("CONFIGURA青O", "DiaSemanaDesligar", "", DiaSemanaDesligar, 256, ConfigFile)
DiaSemanaDesligar = Left(DiaSemanaDesligar, retlen)

'Verifica se tem algum campo em branco
If Origem = "" Then
    MsgBox "O par鈓etro Origem n鉶 foi encontrado!", vbCritical
End If

If Destino = "" Then
    MsgBox "O par鈓etro ( Destino ) n鉶 foi encontrado!", vbCritical
End If

If LogFile = "" Then
    MsgBox "O par鈓etro ( LogFile ) n鉶 foi encontrado!", vbCritical
End If

If HoraBkp = "" Then
    MsgBox "O par鈓etro ( HoraBkp ) n鉶 foi encontrado!", vbCritical
End If

If Parametro = "" Then
    MsgBox "O par鈓etro ( Parametro ) n鉶 foi encontrado!", vbCritical
End If

If Metodo = "" Then
    MsgBox "O par鈓etro ( Metodo ) n鉶 foi encontrado!", vbCritical
End If

If Winrar = "" Then
    MsgBox "O par鈓etro ( Winrar ) n鉶 foi encontrado!", vbCritical
End If

If NomeArquivo = "" Then
    MsgBox "O par鈓etro ( NomeArquivo ) n鉶 foi encontrado!", vbCritical
End If

If HoraDesligar = "" Then
    MsgBox "O par鈓etro ( HoraDesligar ) n鉶 foi encontrado!", vbCritical
End If

If ComandoDesligar = "" Then
    MsgBox "O par鈓etro ( ComandoDesligar ) n鉶 foi encontrado!", vbCritical
End If

If DiaSemanaDesligar = "" Then
    MsgBox "O par鈓etro ( DiaSemanaDesligar ) n鉶 foi encontrado!", vbCritical
End If

'Atribui as configura鏾es nos campos de texto
txt_Origem.Text = Origem
txt_Destino.Text = Destino
txt_LogFile.Text = LogFile
txt_HoraBkp.Text = HoraBkp
txt_Parametro.Text = Parametro
txt_Metodo.Text = Metodo
txt_Winrar.Text = Winrar
txt_NomeArquivo.Text = NomeArquivo
txt_HoraDesligar.Text = HoraDesligar
txt_ComandoDesligar.Text = ComandoDesligar
txt_DiaSemanaDesligar.Text = DiaSemanaDesligar

End Sub

Private Sub Save_Click()
Call WritePrivateProfileString("CONFIGURA青O", "Origem", txt_Origem.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "Destino", txt_Destino.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "LogFile", txt_LogFile.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "HoraBkp", txt_HoraBkp.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "Parametro", txt_Parametro.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "Metodo", txt_Metodo.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "Winrar", txt_Winrar.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "NomeArquivo", txt_NomeArquivo.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "HoraDesligar", txt_HoraDesligar.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "ComandoDesligar", txt_ComandoDesligar.Text, ConfigFile)
Call WritePrivateProfileString("CONFIGURA青O", "DiaSemanaDesligar", txt_DiaSemanaDesligar.Text, ConfigFile)
MsgBox "Configura珲es atualizadas com sucesso!", vbInformation
End Sub
