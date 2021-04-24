VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const INTERNET_DEFAULT_FTP_PORT = 21               ' default for FTP servers
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000            ' used for FTP connections
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0                    ' use registry configuration
Private Const INTERNET_OPEN_TYPE_DIRECT = 1                        ' direct to net
Private Const INTERNET_OPEN_TYPE_PROXY = 3                         ' via named proxy
Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4   ' prevent using java/script/INS
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByRef hInet As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Const PassiveConnection As Boolean = True

'segura número de conexao
Private hConnection As Long
Private hOpen As Long

'Conecta a FTP
Public Function ConectarFTP(ByVal FTPServer As String, ByVal User As String, _
ByVal PassWord As String)
    
    'Abre conexão
    hOpen = InternetOpen("BrunoPimentelMachado", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    'Conecta com servidor FTP
    hConnection = InternetConnect(hOpen, FTPServer, INTERNET_DEFAULT_FTP_PORT, User, PassWord, INTERNET_SERVICE_FTP, IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)

    ConectarFTP = hConnection <> 0

End Function

'baixa arquivo de diretorio aberto
Public Sub DownloadFile(ByVal ArquivoFTP As String, ByVal ArquivoLocal As String)
    FtpGetFile hConnection, ArquivoFTP, ArquivoLocal, False, 0, FTP_TRANSFER_TYPE_UNKNOWN, 0
End Sub

'Envia Arquivo para diretorio aberto
Public Sub UpLoadFile(ByVal ArquivoLocal As String, ByVal ArquivoFTP As String)
    FtpPutFile hConnection, ArquivoLocal, ArquivoFTP, FTP_TRANSFER_TYPE_UNKNOWN, 0
End Sub

'Cria Diretório
Public Sub CriarDiretorio(ByVal Diretorio As String)
    FtpCreateDirectory hConnection, Diretorio
End Sub

'renomeia arquivo
Public Sub RenomeiaArquivo(ByVal ArquivoOriginal As String, ByVal NovoNome As String)
    FtpRenameFile hConnection, ArquivoOriginal, NovoNome
End Sub

'seta diretório atual
Public Sub AbrirDiretorio(ByVal Diretorio As String)
    FtpSetCurrentDirectory hConnection, Diretorio
End Sub

'Pegar diretorio atual
Public Function DiretorioAtual() As String
    DiretorioAtual = String(260, Chr$(0))
    FtpGetCurrentDirectory hConnection, DiretorioAtual, Len(DiretorioAtual)
End Function

'Apaga diretorio
Public Sub ApagaDiretorio(ByVal Diretorio As String)
    FtpRemoveDirectory hConnection, Diretorio
End Sub

'Apaga Arquivo
Public Sub ApagarArquivo(ByVal Arquivo As String)
    FtpDeleteFile hConnection, Arquivo
End Sub

'fecha conexao com ftp
Public Sub FechaConexao()
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
End Sub

'enumera arquivos de pasta corrente
Public Function Arquivos() As Variant
Dim pData As WIN32_FIND_DATA
Dim hFind As Long
Dim lRet As Long
Dim Arquivo() As String
Dim IndiceArquivo As Integer
    
    'prepara buffer para primeiro arquivo
    pData.cFileName = String(MAX_PATH, 0)
    'verifica se encontrou - 0 = nao
    hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
    If hFind = 0 Then Exit Function
    
    IndiceArquivo = 0
    ReDim Arquivo(0 To IndiceArquivo)
    Arquivo(IndiceArquivo) = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    
    Do
        'acerta buffer
        pData.cFileName = String(MAX_PATH, 0)
        'procura proximo arquivo - 0 = nao
        lRet = InternetFindNextFile(hFind, pData)
        If lRet = 0 Then Exit Do
        
        IndiceArquivo = IndiceArquivo + 1
        ReDim Preserve Arquivo(0 To IndiceArquivo)
        
        Arquivo(IndiceArquivo) = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
    Loop
    
    'fecha conexao
    InternetCloseHandle hFind
    
    Arquivos = Arquivo
    
End Function

'forca fechamento de conexao
Private Sub Class_Terminate()
    FechaConexao
End Sub

