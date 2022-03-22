Attribute VB_Name = "Modulo"
Public Const SWP_NOSIZE = &H1
Public Const LB_FINDSTRING = &H18F

Public Declare Function SendMessage Lib "user32" _
     Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Any) As Long

Public Declare Function SetWindowPos Lib "user32" _
     (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
     ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
     ByVal cy As Long, ByVal wFlags As Long) As Long
Global CiA$
Global CaT$
Public Pos As Integer
'Parametros
Global PreCBoD%    'Precios por bodega o lista de precios
Global FacTDoL%    'Si factura en dolares o colones
Global DeCiMaleS%  'Cuantos decimales para el precio
Global CosTeO%     'Si usa el modulo de costeo de importaciones
Global MoNLoC$     'La moneda local
Global MoNExT$     'La moneda extranjera
Global TCambio@    'El tipo de cambio
Global UsaLotes%   'Si se usa la opcion de lotes o no
Global IV#
Global IV2#
Global EspaCio As Workspace
Global DatOS As Database
Global TipoVerS$ '00 = Una moneda,una compania 11=Multimoneda,multicompania
Global Actual As Form
Global RutaCXP$
Global RutaCXC$
Global RutaProdu$
Global bode$
Global ReaRMa%
Global TieneSeries%
Global DirTrA As String 'El subdirectorio de trabajo
Global NomBaSe As String 'El nombre de la base de datos
Global LoGiN As String 'El login del usuario que ingreso al sistema
Global Modo As Integer
Global SimBol As String * 1
Global NumLin%
Global NomSis$
Global CopYRigHT$
Global FreQ%
Global Indice As Integer
Global Parche As Integer
Global Letra As String
Public Function ConVierte(Hilera As Variant, Tipo As String) As String
     If IsNull(Hilera) Then
          ConVierte = ""
          Exit Function
     End If
     Dim Letra As String * 1
     Select Case Tipo
     Case "T"
          Largo = Len(Trim(Hilera))
          ConVierte = UCase(Mid(Hilera, 1, 1))
          J = 0
          For I = 2 To Largo
               Letra = Mid(Hilera, I, 1)
               If J = 1 Then
                    Letra = UCase(Letra)
                    J = 0
               Else
                    Letra = LCase(Letra)
               End If
               If Letra = " " Or Letra = "." Then J = 1
               ConVierte = ConVierte + Letra
           Next
     Case "N"
          If Hilera = "" Then Hilera = "0"
          ConVierte = Format(Hilera, "###,###,###,##0")
     Case "M"
          If Hilera = "" Then Hilera = "0"
          ConVierte = Format(Hilera, " ###,###,###,##0.00")
     Case "F"
          ConVierte = Format(Hilera, "dd/mm/yyyy")
     End Select
End Function
Public Sub Selecciona(ByVal Texto As TextBox)
     Texto.SelStart = 0
     Texto.SelLength = Len(Texto.Text)
     If Texto.Enabled Then Texto.SetFocus
End Sub
Public Sub Barra(Texto$, Optional Modo%, Optional Form As Form)
     With Inicio
     .StatusBar1.Panels(1) = Texto
     If Modo = 0 Then
          .Timer2.Enabled = False
          .StatusBar1.Panels(1).Picture = .StatusBar1.Panels(2).Picture
          If Not (Form Is Nothing) Then Form.MousePointer = 0
     ElseIf Modo = 1 Or Modo = 4 Then
          .Timer1.Enabled = True
          .Timer2.Enabled = False
          If Modo = 4 Then
               .StatusBar1.Panels(1).Picture = .ImageList2.ListImages(1).Picture
          Else
               .StatusBar1.Panels(1).Picture = .StatusBar1.Panels(2).Picture
          End If
          If Not (Form Is Nothing) Then Form.MousePointer = 0
     ElseIf Modo = 2 Then
          If .Timer2.Enabled Then
               .Timer2.Enabled = False
               .StatusBar1.Panels(1).Picture = .StatusBar1.Panels(2).Picture
               If Not (Form Is Nothing) Then Form.MousePointer = 0
          Else
               If Texto <> "" Then
                    .StatusBar1.Panels(1).Picture = .ImageList2.ListImages(1).Picture
                    .Timer2.Enabled = True
                    If Not (Form Is Nothing) Then Form.MousePointer = 11
               End If
          End If
     End If
     End With
End Sub
Public Function Nulo(Campo As Object) As Variant
On Error GoTo Errores
     If IsNull(Campo) Then
          Select Case Campo.Properties(3)
          Case dbText
               Nulo = ""
          Case dbLong, dbInteger, dbSingle, dbCurrency, dbDouble
               Nulo = 0
          Case dbDate
               Nulo = CDate("01/01/1900")
          Case Else
               Nulo = ""
          End Select
     Else
          Nulo = Campo
     End If
Errores:
    If Err.Number > 0 Then
        Exit Function
    End If
End Function
Public Sub Posiciona(Ventana As Form, Modo As Integer)
     On Error Resume Next
     '1=se descarga 0=se carga
     If Modo = 0 Then
          Ventana.Top = CSng(ReadKey(Ventana.Name + "Top", "xy"))
          If Not Ventana.MDIChild And Ventana.Top < 500 Then Ventana.Top = 500
          Ventana.Left = CSng(ReadKey(Ventana.Name + "Left", "xy"))
          If Not Ventana.MDIChild And Ventana.Left < 500 Then Ventana.Left = 500
          If Ventana.BorderStyle = 2 Then
               Ventana.Height = CSng(ReadKey(Ventana.Name + "Height", "xy"))
               If Ventana.Height < 3000 Then Ventana.Height = 3000
               Ventana.Width = CSng(ReadKey(Ventana.Name + "Width", "xy"))
          End If
     Else
          Call SaveKey(Ventana.Name + "Top", Ventana.Top, "xy")
          Call SaveKey(Ventana.Name + "Left", Ventana.Left, "xy")
          Call SaveKey(Ventana.Name + "Height", Ventana.Height, "xy")
          Call SaveKey(Ventana.Name + "Width", Ventana.Width, "xy")
     End If
     On Error GoTo 0
End Sub
Public Function Doble(Texto As Variant) As Currency
     On Error Resume Next
     If IsNull(Texto) Then
        Doble = 0
        Exit Function
     End If
     If Texto <> "" Then Doble = CDbl(Texto)
     On Error GoTo 0
End Function
Public Function Centra(Ventana As Form)
      Ventana.Top = (((Inicio.Height - Inicio.Toolbar1.Height - Inicio.StatusBar1.Height)) - Ventana.Height) / 2
      Ventana.Left = (Inicio.Width - Ventana.Width) / 2
End Function
Public Function Centra1(Ventana As Form)
      Ventana.Top = ((((Inicio.Height - Inicio.Toolbar1.Height - Inicio.StatusBar1.Height)) - Ventana.Height) / 2) + 1000
      Ventana.Left = (Inicio.Width - Ventana.Width + 500) / 2
End Function
Public Function CentraF(Ventana As Form)
      Ventana.Top = ((((Inicio.Height - Inicio.Toolbar1.Height - Inicio.StatusBar1.Height)) - Ventana.Height) / 2) + 3500
      Ventana.Left = (Inicio.Width - Ventana.Width) / 2
End Function
Public Function Centra2(Ventana As Form)
      Ventana.Top = ((((Inicio.Height - Inicio.Toolbar1.Height - Inicio.StatusBar1.Height)) - Ventana.Height) / 2) + 500
      Ventana.Left = (Inicio.Width - Ventana.Width) / 2
End Function
Public Function GeneraNumero2(Tipo As String, compa As String, Optional xxbodega As String) As String
     On Error GoTo Errores
     Dim SQL$
     Dim Tabla As Recordset
     Select Case Tipo
     Case "ETOCia" 'Entrada Traslado bodega principal (Tommy)
          SQL = "select max(n_entrada) from entradas where cia='" + compa + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          SQL = SQL + " and mid(n_entrada,1,2)='99'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "EOcia" 'Entrada Traslado bodega principal (Tommy)
          SQL = "select max(n_entrada) from entradas where cia='" + compa + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          SQL = SQL + " and mid(n_entrada,1,2)='99'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     End Select
     If Tabla.EOF Or IsNull(Tabla(0)) Then
          GeneraNumero2 = "1"
     Else
          GeneraNumero2 = Trim(Val(Tabla(0)) + 1)
     End If
     GeneraNumero2 = Format(GeneraNumero2, "00000000")
     Set Tabla = Nothing
Errores:
     If Err.Number = 3021 Then
          GeneraNumero2 = "1"
          Resume Next
     ElseIf Err.Number > 0 Then
          MsgBox Err.Description + Str(Err.Number), 16, "GeneraNumero"
     End If
     On Error GoTo 0
End Function
Public Function GeneraNumero(Tipo As String, Optional xxbodega As String) As String
     On Error GoTo Errores
     Dim SQL$
     Dim Tabla As Recordset
     Select Case Tipo
     Case "E" 'Entrada
          SQL = "select max(n_entrada) from entradas where cia='" + CiA + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          SQL = SQL + " and mid(n_entrada,1,2)<>'99' AND MID(N_ENTRADA,1,2)<>'OT'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "ET" 'Entrada Traslado bodega principal (Tommy)
          SQL = "select max(n_entrada) from entradas where cia='" + CiA + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          SQL = SQL + " and mid(n_entrada,1,2)='99'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "F" 'Factura
          SQL = "select consec from parametros where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
          GeneraNumero = Format(Tabla!Consec, "00000000")
          Exit Function
     Case "S" 'Salida
          SQL = "select max(n_salida) from salidas where cia='" + CiA + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "PE" 'Pedidos
          SQL = "select max(n_factura) from pedidos where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "PR" 'Factura proforma
          SQL = "select max(n_factura) from proformas where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "P" 'Precio
          SQL = "select max(codigo) from precios where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "T" 'Traslado
          SQL = "select max(n_traslado) from traslados where cia='" + CiA + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "NC" 'Nota de credito
          SQL = "select max(n_factura) from notascred where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "ND" 'Nota de debito
          SQL = "select max(n_factura) from notasdeb where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "DV" 'Devolucion
          SQL = "select max(n_devol) from devolucion where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "OR" 'orden de compra
          SQL = "select max(numero) from ordenes where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "BK"
          SQL = "select max(numero) from backorder where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "TF"
          SQL = "select max(numero) from tomafisica where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "PEX"    ''' Pedidos Exterior
          SQL = "select max(n_factura) from PediExte where cia='" + CiA + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     Case "PA"    ''' Pago de Apartados (Recibos)
          SQL = "select max(numero) from Pagos where cia='" + CiA + "'"
          If Not IsNull(xxbodega) Then SQL = SQL + " and c_bodega='" + xxbodega + "'"
          Set Tabla = DatOS.OpenRecordset(SQL)
     End Select
     If Tabla.EOF Or IsNull(Tabla(0)) Then
          GeneraNumero = "1"
     Else
          GeneraNumero = Trim(Val(Tabla(0)) + 1)
     End If
     GeneraNumero = Format(GeneraNumero, "00000000")
     Set Tabla = Nothing
Errores:
     If Err.Number = 3021 Then
          GeneraNumero = "1"
          Resume Next
     ElseIf Err.Number > 0 Then
          MsgBox Err.Description + Str(Err.Number), 16, "GeneraNumero"
     End If
     On Error GoTo 0
End Function
Public Function Tipos(Tipo As Integer) As String
     Select Case Tipo
     Case 0
          Tipos = "N.CR Esp"
     Case 1
          Tipos = "Factura"
     Case 2
          Tipos = "Nota de Crédito"
     Case 3
          Tipos = "Nota de Débito"
     Case 4
          Tipos = "Recibo"
     End Select
End Function
Public Function Formato(Campo As Field) As Variant
     Dim Letra As String * 1
     Select Case Campo.Properties(3)
     Case dbText
          If IsNull(Campo.Value) Then
               Formato = ""
               Exit Function
          End If
          Largo = Len(Trim(Campo.Value))
          Formato = UCase(Mid(Campo.Value, 1, 1))
          J = 0
          For I = 2 To Largo
               Letra = Mid(Campo.Value, I, 1)
               If J = 1 Then
                    Letra = UCase(Letra)
                    J = 0
               Else
                    Letra = UCase(Letra)
               End If
               If Letra = " " Or Letra = "." Then J = 1
               Formato = Formato + Letra
               Formato = UCase(Formato)
           Next
     Case dbLong, dbInteger, dbSingle
          If IsNull(Campo.Value) Then
               Formato = 0
               Exit Function
          End If
          Formato = FormatNumber(Campo.Value)
     Case dbCurrency, dbDouble
          If IsNull(Campo.Value) Then
               Formato = 0
               Exit Function
          End If
          Formato = FormatNumber(Campo.Value)
     Case dbDate
          If IsNull(Campo.Value) Then
               Formato = ""
               Exit Function
          End If
          Formato = Format(Campo.Value, "dd/mm/yyyy")
     Case Else
          Formato = CStr(Campo.Value)
     End Select
End Function
Public Function Meses(MeS As Integer) As String
     Select Case MeS
     Case 1
          Meses = "Enero"
     Case 2
          Meses = "Febrero"
     Case 3
          Meses = "Marzo"
     Case 4
          Meses = "Abril"
     Case 5
          Meses = "Mayo"
     Case 6
          Meses = "Junio"
     Case 7
          Meses = "Julio"
     Case 8
          Meses = "Agosto"
     Case 9
          Meses = "Setiembre"
     Case 10
          Meses = "Octubre"
     Case 11
          Meses = "Noviembre"
     Case 12
          Meses = "Diciembre"
     Case 13
          Meses = "Mes Trece"
     End Select
End Function
Public Function LinkTable(Tabla$, Tipo$) As Boolean
     On Error GoTo Errores
     Dim CoNeCTa$
     Select Case LCase(Tipo)
     Case "produccion"
          CoNeCTa = RutaProdu
     Case "cxc"
          CoNeCTa = RutaCXC
     Case "cxp"
          CoNeCTa = GetRegistryKey("software\Hovisys\cxp", "ruta", 3)
          CoNeCTa = IIf(Right(CoNeCTa, 1) <> "\", CoNeCTa + "\", CoNeCTa)
          CoNeCTa = CoNeCTa + "datos\" + GetRegistryKey("software\Hovisys\cxp", "nombase", 3)
          'CoNeCTa = RutaCXP
     Case "conta"
          CoNeCTa = GetRegistryKey("software\Hovisys\conta", "ruta", 3)
          CoNeCTa = IIf(Right(CoNeCTa, 1) <> "\", CoNeCTa + "\", CoNeCTa)
          CoNeCTa = CoNeCTa + "datos\" + GetRegistryKey("software\Hovisys\conta", "nombase", 3)
     Case "inventario"
          CoNeCTa = GetRegistryKey("software\Hovisys\inventario", "ruta", 3)
          CoNeCTa = IIf(Right(CoNeCTa, 1) <> "\", CoNeCTa + "\", CoNeCTa)
          CoNeCTa = CoNeCTa + "datos\" + GetRegistryKey("software\Hovisys\inventario", "nombase", 3)
     End Select
     On Error GoTo 0
     LinkTable = BuscaTabla(Tabla, CoNeCTa)
Errores:
     If Err.Number = 3044 Or Err.Number = 3024 Then
          S = "No se encontró la base de datos '" + DatOS.Name + "'"
          S = S + Chr(13) + "Contacte a su administrador."
          MsgBox S, 16, "Error"
     ElseIf Err.Number > 0 Then
          MsgBox Err.Description + Str(Err.Number), 16, "LinkTable"
     End If
     On Error GoTo 0
End Function
Private Function BuscaTabla(Tabla$, ConectSTR$) As Boolean
     On Error GoTo Errores
     Dim Tupla As TableDef
     If LCase(Mid(DatOS.TableDefs(Tabla).Connect, 11)) <> LCase(ConectSTR) Then
          DatOS.TableDefs(Tabla).Connect = ";DATABASE=" + ConectSTR + ";PWD=TORTAS"
     End If
     DatOS.TableDefs(Tabla).RefreshLink
     BuscaTabla = True
Errores:
     If Err.Number = 3012 Then
          S = "drop table " + Tabla
          DatOS.Execute S, 128
          Resume
     ElseIf Err.Number = 3031 Then
          S = "drop table " + Tabla
          DatOS.Execute S, 128
          Resume Next
     ElseIf Err.Number = 3265 Then
          Set Tupla = DatOS.CreateTableDef(ConVierte(Tabla, "T"))
          Tupla.SourceTableName = Tabla
          Tupla.Connect = ";DATABASE=" + ConectSTR + ";PWD=TORTAS"
          DatOS.TableDefs.Append Tupla
          BuscaTabla = True
     ElseIf Err.Number = 9999 Then
          Resume Next
     ElseIf Err.Number > 0 Then
          MsgBox Err.Description + Str(Err.Number), 16, "Buscar Tabla"
     End If
     On Error GoTo 0
End Function
Public Sub Menus(Activa As Boolean, Form As Form)

End Sub
Public Function Seguridad(Codigo$) As Boolean
     On Error Resume Next
     Dim Accesos As Recordset
     S = "select activo from accesos where login='" + LoGiN
     S = S + "' and codigo='" + Codigo + "'"
     Set Accesos = DatOS.OpenRecordset(S)
     If Not Accesos.EOF Then
          Seguridad = IIf(Accesos!Activo = 1, True, False)
     Else
          Seguridad = True
     End If
     Accesos.Close
     On Error GoTo 0
End Function
Public Function GBitacora(Tipo%, ByVal DesC As String)
     On Error GoTo BitErr
     Dim SQLStr As String
     Dim Hora$
     Hora = Format(Time, "hh:mm:ss AM/PM")
     If Tipo > 0 Then DesC = "(Compañía: " + CiA + ") " + DesC
     SQLStr = "insert into bitacora(login,fecha,hora,tipo,descripcion,cia,sistema)"
     SQLStr = SQLStr + " values('" & LoGiN + "',#" & Format(Date, "m/d/yyyy") + "#,'"
     SQLStr = SQLStr + Hora + "'," & Tipo & ",'" + DesC + "','"
     SQLStr = SQLStr + CiA + "','" + NomSis + "')"
     DatOS.Execute SQLStr, 128
BitErr:
     If Err.Number <> 0 Then
          MsgBox Err.Description + Str(Err.Number), 16, "Bitacora"
          'Resume Next
     End If
     On Error GoTo 0
End Function
Public Function Encripta(SerialNo$) As String
     Dim Char%
     Dim num%
     Dim Digitos$
     Dim Caracter%
     Dim temp As String * 1
     Randomize
     num = Int((99 * Rnd) + 14)
     Encripta = num 'Parametro con el que se calculo el XOR
     Digitos = Len(SerialNo)
     Digitos = IIf(Len(Digitos) = 1, "0" + Digitos, Digitos)
     Encripta = Encripta + Digitos 'Largo del password
     For I = 1 To Len(SerialNo) 'Encripta el password
          Char = Val((Asc(Mid(SerialNo, I, 1))))
          Caracter = (num Xor Char)
          temp = Chr(Caracter)
          If temp = "'" Or temp = """" Then
               Call Encripta(SerialNo)
          Else
               Encripta = Encripta + temp
          End If
     Next
     Dim Largo$
     Largo = Len(Encripta)
     Do Until Largo >= 15 'Caracteres de mas hasta completar 15
          Randomize
          num = Int((99 * Rnd) + 1)
          Encripta = Encripta + Chr(num)
          Largo = Len(Encripta)
     Loop
End Function
Public Function Decripta(Texto As Variant) As String
     'Decripta el string que encripto la funcion 'Encripta
     Dim Passw$
     Dim Char%
     Dim num%
     Dim Largo%
     If IsNull(Texto) Then Exit Function
     num = Val(Mid(Texto, 1, 2))
     Largo = Val(Mid(Texto, 3, 2))
     Texto = Mid(Texto, 5, Largo)
     For I = 1 To Len(Texto)
          Char = Asc(Mid(Texto, I, 1))
          Decripta = Decripta + Chr(num Xor Char)
     Next
End Function
Public Sub Imprime(Modo%, NomRep$, Form As Form, Optional Formula$, _
     Optional Param1$, Optional Param2$, Optional Param3$, _
     Optional Param4$, Optional Param5$, Optional Param6$, _
     Optional DefPrint As Boolean)
     On Error GoTo PrinTErr
     Dim Control As CrystalReport
     Set Control = Form.Report1
     'Printer.ForeColor = vbBlue
     Dim Puerto$
     
     Puerto = Printer.Port
     'Control.PrinterPort = Mid(Puerto, 1, Len(Puerto) - 1)
     Control.PrinterDriver = Printer.DriverName
     Control.PrinterName = Printer.DeviceName
     'Printer.Orientation = 2   para que si cada reporte lo define ??? 24/5/00
     Control.ReportFileName = DirTrA + "Reportes\" + NomRep
     Control.DataFiles(0) = DatOS.Name
     Control.SelectionFormula = Formula
     Control.WindowShowRefreshBtn = True
     Control.WindowShowPrintSetupBtn = True
     Debug.Print Formula
     Control.Formulas(2) = ""
     Control.Formulas(3) = ""
     Control.Formulas(4) = ""
     Control.Formulas(5) = ""
     Control.Formulas(6) = ""
     If Param1 <> "" Then Control.Formulas(0) = "comodin = '" + Param1 + "'"
     If Param2 <> "" Then Control.Formulas(1) = "comodin2 = '" + Param2 + "'"
     If Param3 <> "" Then Control.Formulas(2) = "comodin3 = '" + Param3 + "'"
     If Param4 <> "" Then Control.Formulas(3) = "comodin4 = '" + Param4 + "'"
     If Param5 <> "" Then Control.Formulas(4) = "comodin5 = '" + Param5 + "'"
     If Param6 <> "" Then Control.Formulas(5) = "comodin6 = '" + Param6 + "'"
     If Control.WindowTitle = "" Then Control.WindowTitle = Form.Caption
     Control.Destination = Modo
     If Modo = 2 Then
          Form.Dialog1.DialogTitle = "Guardar como"
          Form.Dialog1.Filter = "Archivo de Texto |*.txt|Todos los archivos |*.*"
          Form.Dialog1.InitDir = App.Path
          Form.Dialog1.ShowSave
          If Form.Dialog1.FileName <> "" Then
               Control.PrintFileType = 2
               Control.PrintFileName = Form.Dialog1.FileName
          Else
               Exit Sub
          End If
     End If
     Control.Action = 1
PrinTErr:
     If Err.Number = 20504 Then
          MsgBox "Reporte no encontrado !" + Chr(13) + Control.ReportFileName, 16, NomRep
     ElseIf Err.Number = 20515 Then
          MsgBox Control.SelectionFormula, 16, "Formula Incorrecta"
     ElseIf Err.Number = 20526 Then
          MsgBox "Al menos una impresora debe existir en el sistema !", 16, "Impresión de reporte"
     ElseIf Err.Number = 20518 Then
          Resume Next
     ElseIf Err.Number = 9997 Then
          MsgBox "La impresora de reportes no existe en el sistema !", 16, "Error"
     ElseIf Err.Number > 0 Then
          MsgBox Err.Description + Str(Err.Number) + Chr(13) + Control.ReportFileName, 16, "Imprime"
     End If
     On Error GoTo 0
End Sub
Public Function Dias(Dia As Variant) As String
     If IsNull(Dia) Then Exit Function
     Select Case Val(Dia)
     Case 1
          Dias = "Domingo"
     Case 2
          Dias = "Lunes"
     Case 3
          Dias = "Martes"
     Case 4
          Dias = "Miercoles"
     Case 5
          Dias = "Jueves"
     Case 6
          Dias = "Viernes"
     Case 7
          Dias = "Sabado"
     Case 8
          Dias = "Todos"
     End Select
End Function
Public Sub MenuBotones(Activa As Boolean, Form As Form)

End Sub
Public Function Licencia(ListView As Control, COnta%) As Boolean
     Dim Tipo$
     Tipo = ReadKey("IYFDV")
     If Tipo = "11" Then
          If ListView.ListItems.Count >= COnta Then
               MsgBox "El máximo de registros ya se alcanzó en esta versión", 16, App.Title + " (Versión Demo)"
          Else
               Licencia = True
          End If
     Else
          Licencia = True
     End If
End Function
Public Function Arregla(Largo%, Campo As Variant, Optional Letra$) As String
     Dim Char As String * 1
     Dim Leng%
     Char = IIf(Letra <> "", Letra, Chr(32))
     If IsNull(Campo) Then Exit Function
     Arregla = FormatNumber(Doble(Campo))
     Leng = Len(Arregla)
     Do Until Leng >= Largo
          Arregla = Char + Arregla
          Leng = Len(Arregla)
     Loop
End Function
Public Function Arregla0(Largo%, Campo As Variant, Optional Letra$) As String
     Dim Char As String * 1
     Dim Leng%
     Char = IIf(Letra <> "", Letra, Chr(32))
     If IsNull(Campo) Then Exit Function
     Arregla0 = Format(Doble(Campo), "###,###,###")
     Leng = Len(Arregla0)
     Do Until Leng >= Largo
          Arregla0 = Char + Arregla0
          Leng = Len(Arregla0)
     Loop
End Function

Private Function Tri(ByVal Texto As String, Miles As Boolean) As String
     Dim Modo$
     Modo = 0
     Texto = Trim(Texto)
     Do Until Len(Texto) = 3
          Texto = "0" + Texto
          Modo = Modo + 1
     Loop
     For Pos = 1 To 3
          Caracter = Mid(Texto, Pos, 1)
          Select Case Caracter
          Case "0"
               If Pos = 3 Then
                    Tri = Trim(Tri)
                    If Right(Tri, 1) = "y" Then
                         Tri = Mid(Tri, 1, Len(Tri) - 2)
                    End If
               End If
          Case "1"
               If Pos = 1 Then
                    If Right(Texto, 2) = "00" Then
                         Tri = Tri + "cien"
                    Else
                         Tri = Tri + Tri + "ciento "
                    End If
               ElseIf Pos = 2 Then
                    Select Case Right(Texto, 1)
                    Case "0"
                         Tri = Tri + "diez"
                    Case "1"
                         Tri = Tri + "once"
                    Case "2"
                         Tri = Tri + "doce"
                    Case "3"
                         Tri = Tri + "trece"
                    Case "4"
                         Tri = Tri + "catorce"
                    Case "5"
                         Tri = Tri + "quince"
                    Case "6"
                         Tri = Tri + "dieciseis"
                    Case "7"
                         Tri = Tri + "diecisiete"
                    Case "8"
                         Tri = Tri + "dieciocho"
                    Case "9"
                         Tri = Tri + "diecinueve"
                    End Select
                    Exit For
               ElseIf Pos = 3 Then
                    If Not Miles Then
                         Tri = Tri + "uno"
                    Else
                         Tri = Tri + "un"
                    End If
               End If
          Case "2"
               If Pos = 1 Then
                    Tri = Tri + "doscientos "
               ElseIf Pos = 2 Then
                    If Right(Texto, 1) = "0" Then Tri = Tri + "veinte"
                    If Right(Texto, 1) <> "0" Then Tri = Tri + "veinti"
               ElseIf Pos = 3 Then
                    Tri = Tri + "dos"
               End If
          Case "3"
               If Pos = 1 Then
                    Tri = Tri + "trescientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "treinta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "tres"
               End If
          Case "4"
               If Pos = 1 Then
                    Tri = Tri + "cuatrocientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "cuarenta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "cuatro"
               End If
          Case "5"
               If Pos = 1 Then
                    Tri = Tri + "quinientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "cincuenta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "cinco"
               End If
          Case "6"
               If Pos = 1 Then
                    Tri = Tri + "seiscientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "sesenta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "seis"
               End If
          Case "7"
               If Pos = 1 Then
                    Tri = Tri + "setecientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "setenta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "siete"
               End If
          Case "8"
               If Pos = 1 Then
                    Tri = Tri + "ochocientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "ochenta y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "ocho"
               End If
          Case "9"
               If Pos = 1 Then
                    Tri = Tri + "novecientos "
               ElseIf Pos = 2 Then
                    Tri = Tri + "noventa y "
               ElseIf Pos = 3 Then
                    Tri = Tri + "nueve"
               End If
          End Select
     Next
End Function
Function Arma(Texto As Double) As String
     Dim Enteros$
     Dim DeCiMaleS$
     Dim Largo%
     resul = Format(Texto, "#########0.00")
     Donde = InStr(resul, ".")
     Enteros = Mid(resul, 1, Donde - 1)
     DeCiMaleS = Mid(resul, Donde + 1)
     Largo = Len(Enteros)
     Dim UCD$
     Dim Mill$
     Dim Miles$
     Select Case Largo
     Case 1 To 3 'Centenas
          Arma = Tri(Enteros, False)
     Case 4 To 6 'Miles
          UCD = Right(Enteros, 3)
          Miles = Mid(Enteros, 1, Largo - 3)
          Arma = Tri(Miles, True) + " mil " + Tri(UCD, True)
     Case 7 To 9 'Millones
          UCD = Right(Enteros, 3)
          Mill = Mid(Enteros, 1, Largo - 6)
          Miles = Mid(Enteros, (Largo - 6) + 1, 3)
          Arma = Tri(Mill, True) 'El monto de los millones
          Arma = Arma + IIf(Arma = "un", " millon ", " millones ")
          Arma = Arma + Tri(Miles, True) 'El monto de los miles
          Arma = Arma + IIf(Tri(Miles, True) <> "", " mil ", "")
          Arma = Arma + Tri(UCD, True) 'El monto de las centenas
     Case 10 To 12 'Miles de millones
          Dim MilMill$
          UCD = Right(Enteros, 3)
          Miles = Mid(Right(Enteros, 6), 1, 3)
          Mill = Mid(Right(Enteros, 9), 1, 3) '1,454,556,565
          MilMill = Mid(Enteros, 1, Largo - 9)
          temp = Tri(MilMill, False)
          Arma = IIf(temp = "uno", "mil ", temp + " mil ") 'miles de millones
          Arma = Arma + Tri(Mill, True) 'El monto de los millones
          Arma = Arma + IIf(Arma = "un", " millon ", " millones ")
          Arma = Arma + Tri(Miles, True) 'El monto de los miles
          Arma = Arma + IIf(Tri(Miles, True) <> "", " mil ", "")
          Arma = Arma + Tri(UCD, True) 'El monto de las centenas
     End Select
     If Arma <> "" Then
          Arma = Arma + " con " + DeCiMaleS + "/100"
          Arma = UCase(Mid(Arma, 1, 1)) + Mid(Arma, 2) + "*******"
     Else
          Arma = "********************"
     End If
End Function
Public Sub Actualizaciones()
     On Error GoTo Errores
     Call Barra("Actualizando base de datos ...")
     'Existencias
     S = "alter table existencias add Costo double"
     DatOS.Execute S
     S = "alter table existencias add CostoPro double"
     DatOS.Execute S
     'Traslados
     S = "alter table traslados add temp varchar(15)"
     DatOS.Execute S
     S = "update traslados set temp=c_encargad"
     DatOS.Execute S
     S = "alter table traslados drop c_encargad"
     DatOS.Execute S
     S = "alter table traslados add C_ENCARGAD varchar(15) not null"
     DatOS.Execute S
     S = "update traslados set c_encargad=temp"
     DatOS.Execute S
     S = "alter table traslados drop temp"
     DatOS.Execute S
     S = "alter table traslados add monto CURRENCY"
     DatOS.Execute S

     S = "alter table desgcred add LOTE VARCHAR(15)"
     DatOS.Execute S
     S = "alter table DESGDEB add LOTE VARCHAR(15)"
     DatOS.Execute S
     S = "alter table desgdeb drop constraint PK"
     DatOS.Execute S
     S = "alter table desgcred drop constraint PK"
     DatOS.Execute S
     S = "alter table desgent add LOTE varchar(15)"
     DatOS.Execute S
     S = "select * into DiarioTemp from Diario"
     DatOS.Execute S
     S = "delete from DiarioTemp"
     DatOS.Execute S
     'TipoFact
     S = "alter table tipofact add Numero int"
     DatOS.Execute S
     S = "alter table tipofact add Autoriza int"
     DatOS.Execute S
     'Facturas
     S = "alter table facturas add Comision Currency"
     DatOS.Execute S
     DatOS.TableDefs("facturas").Fields("Comision").DefaultValue = 0
     S = "alter table facturas add RENTA double"
     DatOS.Execute S
     DatOS.TableDefs("facturas").Fields("renta").DefaultValue = 0
     S = "alter table facturas add MONTORENTA Currency"
     DatOS.Execute S
     DatOS.TableDefs("facturas").Fields("montorenta").DefaultValue = 0
     S = "alter table facturas add DOCUMENTO VARCHAR(50)"
     DatOS.Execute S
     S = "alter table facturas add AUTORIZA VARCHAR(50)"
     DatOS.Execute S
     S = "create index PK2 on Facturas (C_bodega,N_factura)"
     DatOS.Execute S
     Dim Ei%
     For I = 0 To DatOS.Relations.Count - 1
          If DatOS.Relations(I).Name = "Facturas>Desgfact" Then
               Call DatOS.Relations.Delete(DatOS.Relations(I).Name)
               Ei = 1
               Exit For
          End If
     Next
     If Ei = 1 Then
          S = "alter table FACTURAS drop constraint PK"
          DatOS.Execute S
          S = "alter table FACTURAS drop constraint PK2"
          DatOS.Execute S
          S = "alter table FACTURAS add temp varchar(15)"
          DatOS.Execute S
          S = "update FACTURAS set temp=N_FACTURA"
          DatOS.Execute S
          S = "alter table facturas drop n_factura"
          DatOS.Execute S
          S = "alter table facturas add N_FACTURA varchar(15) not null"
          DatOS.Execute S
          S = "update FACTURAS set N_FACTURA=temp"
          DatOS.Execute S
          S = "alter table FACTURAS drop temp"
          DatOS.Execute S
          S = "alter table FACTURAS ADD CONSTRAINT PK Primary Key(n_factura)"
          DatOS.Execute S
          S = "create index PK2 on FACTURAS (C_BODEGA,N_FACTURA)"
          DatOS.Execute S
          Dim Rela As Relation
          Set Rela = DatOS.CreateRelation("FacturasDesgFact", "Facturas", "desgfact", 4352)
          Rela.Fields.Append Rela.CreateField("N_Factura")
          Rela.Fields!n_factura.ForeignName = "N_Factura"
          DatOS.Relations.Append Rela
     End If

     'Series
     S = "Create table Series(Codigo varchar(15) not null,Serie varchar(50) not null,"
     S = S + "Entrada varchar(10) ,Salida varchar(10) null,Estado int)"
     DatOS.Execute S
     S = "alter table Series add constraint PK Primary Key(codigo,serie)"
     DatOS.Execute S

     'Desgfact
     S = "alter table desgfact add LOTE varchar(15)"
     DatOS.Execute S
     S = "create index PK2 on DesgFact (c_bodega,n_factura)"
     DatOS.Execute S
     S = "alter table DESGFACT add serie varchar(50)"
     DatOS.Execute S
     S = "alter table DESGFACT add constraint PK Primary Key"
     S = S + "(c_bodega,n_factura,lote,serie)"
     DatOS.Execute S
     'Estadisticas
     S = "alter table Estadisticas drop constraint PK"
     DatOS.Execute S
     S = "alter table estadisticas add moneda int"
     DatOS.Execute S
     'Parametros
     S = "alter table Parametros add Series int"
     DatOS.Execute S
     S = "alter table Parametros add AvisaMinimo int"
     DatOS.Execute S
     S = "alter table Parametros add AvisaExis int"
     DatOS.Execute S
     S = "alter table Parametros add AvisaTipoPrec int"
     DatOS.Execute S
     S = "alter table Parametros add simbolo char(1)"
     DatOS.Execute S
     S = "create table Plazos(Bodega varchar(2),"
     S = S + "Factura varchar(10),Monto currency,"
     S = S + "Plazo int)"
     DatOS.Execute S
Errores:
     Select Case Err.Number
     Case 3377, 3380, 3010, 3283, 3262, 3375
          Resume Next
     End Select
     Call Barra("")
End Sub
Public Sub Recalcula()
     On Error GoTo Errores
     S = "delete from estadisticas"
     DatOS.Execute S
     S = "insert into estadisticas(codart,codbod,fecha,unidades,monto,cliente,moneda) "
     S = S + "select c_Articulo,desgfact.c_bodega,format(f_factura,'yyyymm'),"
     S = S + "sum(cantidad*unidades),"
     S = S + "sum(total_neto*iif(facturas.dolares=1,facturas.tipocambio,1))"
     S = S + ",c_cliente,dolares "
     S = S + "from facturas,desgfact "
     S = S + "where facturas.c_bodega=desgfact.c_bodega "
     S = S + "and facturas.n_factura = desgfact.n_factura "
     S = S + "group by c_articulo,desgfact.c_bodega,format(f_factura,'yyyymm')"
     S = S + ",c_cliente,dolares"
     DatOS.Execute S
Errores:
     Select Case Err.Number
     Case Is <> 0
          MsgBox Err.Description + Str(Err.Number), 16, "Recalcula"
          Resume Next
     End Select
     On Error GoTo 0
End Sub
Public Function redo_05(xx_monto As Double) As Double
     redo_05 = Round(xx_monto * 2, 1) / 2
End Function
Public Function redo_5(xx_monto As Double) As Double
'Redondea a los 5 Colones Siguientes
     If (xx_monto / 5) - Int(xx_monto / 5) > 0 Then
        redo_5 = (Int(xx_monto / 5) * 5)
     Else
        redo_5 = xx_monto
     End If
End Function
Public Function Consecutivo(eltipo$) As Long
Dim Conse As Recordset, Sx$, Sale As Boolean, HayError As Boolean
'On Error GoTo Errores
Sx = "Select Consecutivo From Consecutivos where cia='" + CiA + "'"
Sx = Sx + " and tipoconse='" + eltipo + "'"
Set Conse = DatOS.OpenRecordset(Sx)
If Conse.EOF Then
   Sx = "Insert into Consecutivos(cia,tipoconse,consecutivo) Values('"
   Sx = Sx + CiA + "','" + eltipo + "',0)"
   DatOS.Execute Sx, 128
End If
Conse.Close
Sale = True
Do While Sale
     HayError = False
     Sx = "Select Consecutivo From Consecutivos where cia='" + CiA + "'"
     Sx = Sx + " and tipoconse='" + eltipo + "'"
     Set Conse = DatOS.OpenRecordset(Sx)
     Sale = HayError
Loop
Consecutivo = Conse!Consecutivo + 1
Conse.Close
''''''''''''''''''''''''''''
Errores:
     Select Case Err.Number
     Case 3008, 3262
          Call Barra("Bloquendo:" + eltipo)
          HayError = True
          Resume Next
     Case Is <> 0
          MsgBox Err.Description + Str(Err.Number), 16, "Consecutivo Tipo:" + eltipo
     End Select
     On Error GoTo 0
End Function
 

