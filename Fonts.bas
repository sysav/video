Attribute VB_Name = "Fonts"
Private Declare Function WinDir _
     Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
     ByVal nSize As Long) As Long
Private Declare Function CreateScalableFontResource Lib "gdi32" _
     Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, _
     ByVal lpszResourceFile As String, ByVal lpszFontFile As String, _
     ByVal lpszCurrentPath As String) As Long
Private Declare Function AddFontResource Lib "gdi32" _
     Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function SendMessage Lib "user32" _
     Alias "SendMessageA" (ByVal hwnd As Long, _
     ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D
Public Function AgregaFont(NFont$, Optional Ruta$) As Boolean
     'Lucida sans = L_10646.TTF"
     'Lucida consule= Lucon.ttf
     Dim WinPath$
     WinPath = String(255, 0)
     Res = WinDir(WinPath, 255)
     WinPath = QuitaNulos(WinPath) + "\"
     FileCopy Ruta + NFont, WinPath + NFont
     TTF_Font$ = Ruta + Mid(NFont, 1, InStr(NFont, ".")) + "FOT"
     respath$ = WinPath
     Result& = CreateScalableFontResource(0, TTF_Font$, NFont$, respath$)
     If Result& Then
          Result& = AddFontResource(WinPath + NFont$)
          If Result& Then
               Result& = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
               If Result& Then
                    FileCopy Ruta + NFont, WinPath + "fonts\" + NFont
                    AgregaFont = True
               End If
          End If
      End If
End Function
Private Function QuitaNulos(Hilera As String) As String
     Dim Largo%
     Dim Char As String * 1
     Largo = Len(Hilera)
     For I = 1 To Largo
          Char = Mid(Hilera, I, 1)
          If Char <> Chr(0) Then
               QuitaNulos = QuitaNulos + Char
          End If
     Next
End Function

