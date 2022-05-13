Attribute VB_Name = "a_Menu_Teclas"
' ------------------------------------------------------------ '
' ---                 Macro creada por                     --- '
' ---         MILAGROS HUERTA GÓMEZ DE MERODIO             --- '
' ------------------------------------------------------------ '
' ---                For Automatic POP UP                  --- '
' ------------------------------------------------------------ '
' ---    Puedes usarla libremente en tus aplicaciones,     --- '
' ---    pero no asignarte la autoría.                     --- '
' ---    Sirve para facilitar la creación de menús         --- '
' ---    desplegables en tus aplicaciones de Excel         --- '
' ------------------------------------------------------------ '
Option Explicit
Public Const minombre = "For Automatic POP UP - Milagros Huerta"     ' Constante con el NOMBRE de la aplicación
Public Cancela_Boton_Excel As Integer ' Boolean
Public Muestra_Mensaje, CAMBIA_CONTROLES As String
Public CAMBIAR_ARCHIVO, CAMBIAR_HOJA As String
Sub Crear_Menu()
' -------------------------------------------------------------------- '
' --- Esta macro crea el menú desplegable, variando para cada hoja --- '
' -------------------------------------------------------------------- '
Dim N_Fila  As Integer
Dim i, j As Integer
Dim N_SubN As Integer
Dim N_Nivel, N_SubNivel As Integer
Dim N_Action, N_Id, N_Group As Integer
Dim N_NA, N_Iclude, N_Teclas As Integer
Dim N_Menu As Object
Dim MG_Item As Object               ' Nuevo Item para el Menu General
Dim Mi_Desplegable As Object        ' Menu Desplegable
Dim General 'As Object               ' Menu General NO ADMITE VARIABLE COMO OBJETO

Set N_Menu = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Menu")
    N_Menu.ListColumns("N/A").Range.Calculate           ' Calcula la fórmula, porque cambiará en función de la hoja en la que estemos
    N_Fila = 1
' --------------------------------------------------------------------- '
' --- Definimos el NUMERO de las COLUMNAS en la que estan los datos --- '
' --------------------------------------------------------------------- '
    N_SubN = 1                          ' Nº Columna del número de subniveles
    N_Nivel = 2                         ' Nº Columna del nombre principal
    N_SubNivel = 3                      ' Nº Columna del nombre secundario
    N_Action = 4                        ' Nº Columna del nombre de la macro
    N_Id = 5                            ' Nº Columna del numero del ICONO
    N_Group = 6                         ' Si quieres que aparezca la linea de separacion de GRUPO
    N_NA = 9                            ' Si Aplica o NO Aplica (N/A), en funcion de las hojas en las que tenga que aparecer
    N_Iclude = 10                       ' Nº de los elementos de un submenu no hay que incluir
    N_Teclas = 12                       ' Atajo de teclado entre corchetes [  ]
' --------------------------------------------------------------------- '
    On Error Resume Next
    Application.CommandBars("Menu_Desplegable").Delete
    Set Mi_Desplegable = CommandBars.Add(Name:="Menu_Desplegable", Position:=msoBarPopup, Temporary:=True)
    With Mi_Desplegable
        For j = 1 To N_Menu.ListRows.Count    ' Num de filas en de la Tabla
            With General
' --- Si hay SubMenús debe generar un  - msoControlPopup -
                If N_Menu.DataBodyRange(N_Fila, N_SubN).Value <> 1 Then
                    If N_Menu.DataBodyRange(N_Fila, N_Iclude).Value <> N_Menu.DataBodyRange(N_Fila, N_SubN).Value Then
                        Set General = Mi_Desplegable.Controls.Add(Type:=msoControlPopup)
                        .Caption = N_Menu.DataBodyRange(N_Fila, N_Nivel).Value
                        .FaceId = N_Menu.DataBodyRange(N_Fila, N_Id).Value
                        .BeginGroup = N_Menu.DataBodyRange(N_Fila, N_Group).Value
                    End If
                End If
' --- Añade los Niveles, tantos como subniveles tenga, si solo tiene 1, lo añade al DESPLEGABLE
                For i = 1 To N_Menu.DataBodyRange(N_Fila, N_SubN).Value              ' Num de SubNiveles que tiene ese Nivel
                    If N_Menu.DataBodyRange(N_Fila, N_NA).Value <> "N/A" Then
                        If N_Menu.DataBodyRange(N_Fila, N_SubN).Value = 1 Then      ' Pueden ser valores 0 o mayores que 1
                            Set MG_Item = Mi_Desplegable.Controls.Add(Type:=msoControlButton)
                        Else
                            Set MG_Item = General.Controls.Add(Type:=msoControlButton)
                        End If
                        
                        With MG_Item
                            .Caption = N_Menu.DataBodyRange(N_Fila, N_SubNivel).Value & " " & N_Menu.DataBodyRange(N_Fila, N_Teclas).Value
                            .OnAction = N_Menu.DataBodyRange(N_Fila, N_Action).Value
                            .FaceId = N_Menu.DataBodyRange(N_Fila, N_Id).Value
                            .BeginGroup = N_Menu.DataBodyRange(N_Fila, N_Group).Value
                        End With
                    End If
                    N_Fila = N_Fila + 1         ' Pasa a la siguiente fila, dentro de un mismo grupo
                Next i
            End With
            j = N_Fila - 1      ' El valor de j se es el de la fila en la que está, hay que restar 1, el FOR lo suma automáticamente
        Next j
    End With
End Sub
Sub Activa_Menu_EXCEL()
Attribute Activa_Menu_EXCEL.VB_ProcData.VB_Invoke_Func = " \n14"
' Si pincha en la opcion del Menu Personalizado Desplegable EXCEL, desactiva el menu para que no se repita
    Cancela_Boton_Excel = 0
    If Muestra_Mensaje = "SI" Then MsgBox "Pulsa la tecla de Menu del teclado para ver las funciones de Excel", _
                                          vbInformation, minombre
    Muestra_Mensaje = "NO"
End Sub
Sub Activa_Funciones()
' ------------------------------------------------------------------- '
' ------------- CODIGOS DE TECLAS PARA COMBINAR --------------------- '
' --- SHIFT     +           CTRL      ^             ' ALT       % --- '
' ------------------------------------------------------------------- '
Dim N_T_Menu As Object
Dim N_T_Atajos As Object
Dim N_T_Comandos As Object
Dim C_Menu_Action As Integer
Dim C_Menu_Teclas As Integer
Dim C_Atajo_Action As Integer
Dim C_Atajo_Teclas As Integer
Dim C_Comandos_Activa As Integer
Dim C_Comandos_Tipo As Integer
Dim i As Integer

Application.ScreenUpdating = False
Set N_T_Menu = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Menu")
Set N_T_Atajos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Atajos")
Set N_T_Comandos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Comandos")

C_Menu_Action = 4
C_Menu_Teclas = 7
C_Atajo_Action = 2
C_Atajo_Teclas = 3
C_Comandos_Tipo = 1
C_Comandos_Activa = 3

    ExecuteExcel4Macro ("show.toolbar(""ribbon"",1)")   ' Activa la cinta de opciones 1 , 0 desactiva
    With Application
' --- MENU DESPLEGABLE ----------------------------------------------------------------------
        For i = 1 To N_T_Menu.ListRows.Count
            If N_T_Menu.DataBodyRange(i, C_Menu_Teclas).Value <> "" Then
              .OnKey N_T_Menu.DataBodyRange(i, C_Menu_Teclas).Value, N_T_Menu.DataBodyRange(i, C_Menu_Action).Value
            End If
        Next i
'--- ATAJOS DE TECLADO ----------------------------------------------------------------------
        For i = 1 To N_T_Atajos.ListRows.Count
            If N_T_Atajos.DataBodyRange(i, C_Atajo_Action).Value <> ThisWorkbook.Worksheets(hj_Tablas.Name).Range("No_Desact").Value And _
               N_T_Atajos.DataBodyRange(i, C_Atajo_Teclas).Value <> "" Then
              .OnKey N_T_Atajos.DataBodyRange(i, C_Atajo_Teclas).Value, N_T_Atajos.DataBodyRange(i, C_Atajo_Action).Value
            End If
        Next i
'--- COMANDOS ----------------------------------------------------------------------
        For i = 1 To N_T_Comandos.ListRows.Count
            .CommandBars(N_T_Comandos.DataBodyRange(i, C_Comandos_Tipo).Value).Enabled = _
                         N_T_Comandos.DataBodyRange(i, C_Comandos_Activa).Value
        Next i
            .StatusBar = minombre                       ' Pone el nombre en la parte inferior izquierda de Excel
            .DisplayFullScreen = True                   ' True si Microsoft Excel está en el modo de pantalla completa
            .DisplayFormulaBar = True
    End With
End Sub
Sub Desactiva_Funciones()
' ------------------------------------------------------------------- '
' ------------- CODIGOS DE TECLAS PARA COMBINAR --------------------- '
' --- SHIFT     +           CTRL      ^             ' ALT       % --- '
' ------------------------------------------------------------------- '
Dim N_T_Menu As Object
Dim N_T_Atajos As Object
Dim N_T_Comandos As Object
Dim C_Menu_Teclas As Integer
Dim C_Atajo_Action As Integer
Dim C_Atajo_Teclas As Integer
Dim C_Comandos_Tipo As Integer
Dim i As Integer

Application.ScreenUpdating = False
Set N_T_Menu = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Menu")
Set N_T_Atajos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Atajos")
Set N_T_Comandos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Comandos")

C_Menu_Teclas = 7
C_Atajo_Action = 2
C_Atajo_Teclas = 3
C_Comandos_Tipo = 1

    ExecuteExcel4Macro ("show.toolbar(""ribbon"",1)")   ' Activa la cinta de opciones 1 , 0 desactiva
    With Application
' --- MENU DESPLEGABLE ----------------------------------------------------------------------
        For i = 1 To N_T_Menu.ListRows.Count
            If N_T_Menu.DataBodyRange(i, C_Menu_Teclas).Value <> "" Then
              .OnKey N_T_Menu.DataBodyRange(i, C_Menu_Teclas).Value
            End If
        Next i
'--- ATAJOS DE TECLADO ----------------------------------------------------------------------
        For i = 1 To N_T_Atajos.ListRows.Count
            If N_T_Atajos.DataBodyRange(i, C_Atajo_Action).Value <> ThisWorkbook.Worksheets(hj_Tablas.Name).Range("No_Desact").Value And _
               N_T_Atajos.DataBodyRange(i, C_Atajo_Teclas).Value <> "" Then
              .OnKey N_T_Atajos.DataBodyRange(i, C_Atajo_Teclas).Value
            End If
        Next i
'--- COMANDOS ----------------------------------------------------------------------
        For i = 1 To N_T_Comandos.ListRows.Count
            .CommandBars(N_T_Comandos.DataBodyRange(i, C_Comandos_Tipo).Value).Enabled = True
        Next i
            .StatusBar = False                          ' Pone el nombre en la parte inferior izquierda de Excel
            .DisplayFullScreen = False                  ' True si Microsoft Excel está en el modo de pantalla completa
            .DisplayFormulaBar = True
    End With
End Sub
Sub Atajos_Teclado()
' -------------------------------------------------------------------- '
' --- Información de otros atajos de teclado, los toma del a TABLA --- '
' -------------------------------------------------------------------- '
Dim i As Integer
Dim N_Inicial, N_Final As Integer
Dim Mensaje_A As String
Dim Mensaje_Atajos As String
Dim N_Menu As Object
Dim N_Atajos As Object

Set N_Menu = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Menu")
Set N_Atajos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Atajos")
    N_Inicial = 1
    N_Final = N_Atajos.ListRows.Count
' Si hubiera un gran número en la tabla de Atajos de Teclado, habría que cambiar este bucle, pues no admite más de 20
    For i = N_Inicial To N_Final
        If N_Atajos.DataBodyRange(i, 2).Value <> ThisWorkbook.Worksheets(hj_Tablas.Name).Range("No_Desact").Value Then
            Mensaje_A = "-  " & N_Atajos.DataBodyRange(i, 1).Value & "     " & N_Atajos.DataBodyRange(i, 4).Value & vbNewLine & vbNewLine
            Mensaje_Atajos = Mensaje_Atajos & Mensaje_A
        End If
    Next i
    MsgBox "OTROS ATAJOS DE TECLADO:" & vbNewLine & vbNewLine & vbNewLine & Mensaje_Atajos, vbInformation, minombre
End Sub
Sub Pantalla_Mostar()
'   Activa o desactiva mostrar en pantalla la barra de fómulas, barra vertical/horizontal...
Dim a As Byte
Dim i As Integer
Dim N_Hoja As Integer
Dim N_Ribbon As Integer
Dim N_BFormula As Integer
Dim N_BD_Vertical As Integer
Dim N_BD_Horizontal As Integer
Dim N_Encabezados As Integer
Dim N_Filas_Pantalla As Integer
Dim N_T_Pantalla As Object
    Application.ScreenUpdating = False
    Set N_T_Pantalla = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Pantalla")

    N_Hoja = 1
    N_Ribbon = 2
    N_BFormula = 3
    N_BD_Vertical = 4
    N_BD_Horizontal = 5
    N_Encabezados = 6
    N_Filas_Pantalla = N_T_Pantalla.ListRows.Count

    For i = 1 To N_Filas_Pantalla
        If ActiveSheet.Name = N_T_Pantalla.DataBodyRange(i, N_Hoja).Value Then
            ' ----------------------------------------------------------------------------- '
            ' --- Para mostar o no la cinta de opciones: 1 - ACTIVA , 0 - DESACTIVA     --- '
            ' --- El valor no se puede poner como variable, está en una cadena de texto --- '
            ' ----------------------------------------------------------------------------- '
            a = N_T_Pantalla.DataBodyRange(i, N_Ribbon).Value
            If a = 1 Then
              ExecuteExcel4Macro ("show.toolbar(""ribbon"",1)")
            Else
              ExecuteExcel4Macro ("show.toolbar(""ribbon"",0)")
            End If
' -------------------------------------------------------------------------------------------- '
' --- Para mostar o no en PANTALLA la Barra de FORMULAS - DESPLAZAMIENTO H/V - ENCABEZADOS ---
' -------------------------------------------------------------------------------------------- '
            Application.DisplayFormulaBar = N_T_Pantalla.DataBodyRange(i, N_BFormula).Value
            ActiveWindow.DisplayVerticalScrollBar = N_T_Pantalla.DataBodyRange(i, N_BD_Vertical).Value
            ActiveWindow.DisplayHorizontalScrollBar = N_T_Pantalla.DataBodyRange(i, N_BD_Horizontal).Value
            ActiveWindow.DisplayHeadings = N_T_Pantalla.DataBodyRange(i, N_Encabezados).Value
            Exit For
        End If
    Next i
End Sub
Sub NO_DESACTIVAR()
' Esta macro no hace nada, es para que los atajos de teclado no hagan nada
End Sub
