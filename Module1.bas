Attribute VB_Name = "Module1"
Public Declare Function PlaySound Lib "winmm.dll" _
  Alias "PlaySoundA" (ByVal lpszName As String, _
  ByVal hModule As Long, ByVal dwFlags As Long) _
  As Long
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public Const SND_LOOP = &H8
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_RESERVED = &HFF000000
Public Const SND_RESOURCE = &H40004
Public Const SND_SYNC = &H0
Public Const SND_TYPE_MASK = &H170007
Public Const SND_VALID = &H1F
Public Const SND_VALIDFLAGS = &H17201F




