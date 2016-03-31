Const COM_FSO           = "Scripting.FileSystemObject"
Const COM_SHELL         = "WScript.Shell"
Const COM_SHELLAPP      = "Shell.Application"
Const COM_HTML          = "htmlfile"
Const COM_HTTP          = "Msxml2.XMLHTTP"
Const COM_XMLHTTP       = "Msxml2.ServerXMLHTTP"
Const COM_WINHTTP       = "WinHttp.WinHttpRequest.5.1"
Const COM_ADOSTREAM     = "Adodb.Stream"
Const COM_XMLDOM        = "Microsoft.XMLDOM"
Const COM_WMI           = "winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2"
Const COM_WMP           = "WMPlayer.ocx"
Const COM_WORD          = "Word.Application"
Const COM_EXCEL         = "Excel.Application"
Const COM_ACCESS        = "Access.Application"
Const COM_PHOTOSHOP     = "PHOTOSHOP.APPLICATION"
Const COM_DICT          = "Scripting.Dictionary"
Const COM_ADO_CONN      = "ADODB.Connection"
Const COM_ADO_RECORDSET = "ADODB.Recordset"
Const COM_ADO_COMMAND   = "ADODB.Command"
Const COM_ADO_CATALOG   = "ADOX.Catalog"
Const COM_COMMONDIALOG  = "UserAccounts.CommonDialog"
Const COM_IE            = "InternetExplorer.Application"
Const COM_TYPELIB       = "Scriptlet.TypeLib"
Const COM_POCKET_HTTP   = "pocket.HTTP"
Const COM_CAPICOM_UTIL  = "CAPICOM.Utilities"
Const COM_CAPICOM_HASH  = "CAPICOM.HashedData"
Const COM_REGEXP        = "VBSCRIPT.REGEXP"
Const COM_CDO_MESSAGE   = "CDO.Message"
Const COM_CDO_CONFIG    = "CDO.Configuration"
Const COM_ITUNES        = "iTunes.Application"

'------------------------------------------------
' VB常数
' vbCrLf        Chr(13) + Chr(10)   回车/换行组合符
' vbCr          Chr(13)             回车符
' vbLf          Chr(10)             换行符
' vbNewLine     Chr(13) + Chr(10)   换行符
' vbNullChar    Chr(0)              值为 0 的字符
' vbNullString  值为 0 的字符串
' vbObjectError -2147221504         错误号。用户定义的错误号应当大于该值。例如：Err.Raise(Number) = vbObjectError + 1000
' vbTab         Chr(9)              Tab 字符
' vbBack        Chr(8)              退格字符

'------------------------------------------------
' FSO
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8
Const CreateIfNeeded = true



Const DESKTOP = &H10&
Const LOCAL_APPLICATION_DATA = &H1c&
Const TEMPORARY_INTERNET_FILES = &H20&
Const FOF_CREATEPROGRESSDLG = &H0&

'------------------------------------------------
' Registry
Const HKEY_CLASSES_ROOT     = &H80000000
Const HKEY_CURRENT_USER     = &H80000001
Const HKEY_LOCAL_MACHINE    = &H80000002
Const HKEY_USERS            = &H80000003
Const HKEY_CURRENT_CONFIG   = &H80000005
Const HKCR = &H80000000 'HKEY_CLASSES_ROOT
Const HKCU = &H80000001 'HKEY_CURRENT_USER
Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
Const HKU  = &H80000003 'HKEY_USERS
Const HKCC = &H80000005 'HKEY_CURRENT_CONFIG
Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7

'------------------------------------------------
' Valid Charset values for ADODB.Stream
Const CdoBIG5        = "big5"
Const CdoEUC_JP      = "euc-jp"
Const CdoEUC_KR      = "euc-kr"
Const CdoGB2312      = "gb2312"
Const CdoISO_2022_JP = "iso-2022-jp"
Const CdoISO_2022_KR = "iso-2022-kr"
Const CdoISO_8859_1  = "iso-8859-1"
Const CdoISO_8859_2  = "iso-8859-2"
Const CdoISO_8859_3  = "iso-8859-3"
Const CdoISO_8859_4  = "iso-8859-4"
Const CdoISO_8859_5  = "iso-8859-5"
Const CdoISO_8859_6  = "iso-8859-6"
Const CdoISO_8859_7  = "iso-8859-7"
Const CdoISO_8859_8  = "iso-8859-8"
Const CdoISO_8859_9  = "iso-8859-9"
Const cdoKOI8_R      = "koi8-r"
Const cdoShift_JIS   = "shift-jis"
Const CdoUS_ASCII    = "us-ascii"
Const CdoUTF_7       = "utf-7"
Const CdoUTF_8       = "utf-8"

'------------------------------------------------
' Constants used by MS ADO.DB 

'---- CursorTypeEnum Values ----
Const adOpenForwardOnly     = 0
Const adOpenKeyset          = 1
Const adOpenDynamic         = 2
Const adOpenStatic          = 3

'---- LockTypeEnum Values ----
Const adLockReadOnly        = 1
Const adLockPessimistic     = 2
Const adLockOptimistic      = 3
Const adLockBatchOptimistic = 4

'---- CursorLocationEnum Values ----
Const adUseServer           = 2
Const adUseClient           = 3

'---- SearchDirection Values ----
Const adSearchForward       = 1
Const adSearchBackward      = -1

'---- CommandTypeEnum Values ----
Const adCmdUnknown          = &H0008
Const adCmdText             = &H0001
Const adCmdTable            = &H0002
Const adCmdStoredProc       = &H0004



'------------------------------------------------
' ADODB.Stream file I/O constants
Const adTypeBinary          = 1
Const adTypeText            = 2
Const adSaveCreateNotExist  = 1
Const adSaveCreateOverWrite = 2
Const adModeUnknown         = 0
Const adModeRead            = 1
Const adModeWrite           = 2
Const adModeReadWrite       = 3


'------------------------------------------------
' CAPICOM
Const CAPICOM_HASH_ALGORITHM_SHA1   = 0
Const CAPICOM_HASH_ALGORITHM_MD2    = 1
Const CAPICOM_HASH_ALGORITHM_MD4    = 2
Const CAPICOM_HASH_ALGORITHM_MD5    = 3
Const CAPICOM_HASH_ALGORITHM_SHA256 = 4
Const CAPICOM_HASH_ALGORITHM_SHA384 = 5
Const CAPICOM_HASH_ALGORITHM_SHA512 = 6

'------------------------------------------------
' IE
Const OLECMDID_SAVE = 3
Const OLECMDEXECOPT_DODEFAULT = 0

'------------------------------------------------
' Base64
Const sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  