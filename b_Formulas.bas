Attribute VB_Name = "b_Formulas"
Sub F_T_Menu()
Attribute F_T_Menu.VB_ProcData.VB_Invoke_Func = " \n14"
' ------------------------------------------------------- '
' --- Formulas y Encabezados de las diferentes TABLAS --- '
' ------------------------------------------------------- '
Dim i As Integer
Dim N_Menu As Object
Dim N_Atajos As Object
Dim N_Comandos As Object
Dim N_Pantalla As Object

Set N_Menu = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Menu")
Set N_Atajos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Atajos")
Set N_Comandos = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Comandos")
Set N_Pantalla = ThisWorkbook.Worksheets(hj_Tablas.Name).ListObjects("T_Pantalla")
    
    hj_Tablas.Select
    
' ---------------------------------------------- '
' --- Encabezados y fórmula de la tabla MENU --- '
' ---------------------------------------------- '
    For i = 1 To N_Menu.ListRows.Count                      ' Num de filas en de la Tabla
        N_Menu.HeaderRowRange(1, 1).Value = "N Sub"
        N_Menu.HeaderRowRange(1, 2).Value = "Nivel"
        N_Menu.HeaderRowRange(1, 3).Value = "Sub Nivel"
        N_Menu.HeaderRowRange(1, 4).Value = "On Action"
        N_Menu.HeaderRowRange(1, 5).Value = "Face Id"
        N_Menu.HeaderRowRange(1, 6).Value = "Begin Group"
        N_Menu.HeaderRowRange(1, 7).Value = "Teclas"
        N_Menu.HeaderRowRange(1, 8).Value = "HOJA NO MUESTRA"
        N_Menu.HeaderRowRange(1, 9).Value = "N/A"
        N_Menu.HeaderRowRange(1, 10).Value = "Num N/A"
        N_Menu.HeaderRowRange(1, 11).Value = "&"
        N_Menu.HeaderRowRange(1, 12).Value = "[Teclas]"
        
        N_Menu.DataBodyRange(1, 1).FormulaR1C1 = "=IF([@Nivel]=""C""&ROW()-1,0,COUNTIF([Nivel],[@Nivel]))"
        N_Menu.DataBodyRange(1, 9).FormulaR1C1 = "=IF([@[HOJA NO MUESTRA]]=Ninguna,""N/A""," & Chr(10) & "IF(IFERROR(SEARCH(MID(CELL(""nombrearchivo""),FIND(""]"",CELL(""nombrearchivo""),1)+1,100),[@[HOJA NO MUESTRA]],1),0)>0,""N/A"",""""))"
        N_Menu.DataBodyRange(1, 10).FormulaR1C1 = "=IF([@[N/A]]="""",,COUNTIF([&],[@[&]]))"
        N_Menu.DataBodyRange(1, 11).FormulaR1C1 = "=[@Nivel]&[@[N/A]]"
        N_Menu.DataBodyRange(1, 12).FormulaR1C1 = "=IFERROR(IF([@Teclas]="""","""",""[""&IFERROR(REPLACE(SEARCH(""%"",[@Teclas],1),1,1,""Alt+""),"""")&IFERROR(REPLACE(SEARCH(""^"",[@Teclas],1),1,1,""Ctrl+""),"""")&IFERROR(REPLACE(SEARCH(""+"",[@Teclas],1),1,1,""May+""),"""")&IF(IFERROR(SEARCH(""+"",[@Teclas],1),"""")<>"""",UPPER(MID([@Teclas],FIND(""{"",[@Teclas],1)+1,LEN([@Teclas])-FIND(""{"",[@Teclas],1)-1)),MID([" & _
                                                   "@Teclas],FIND(""{"",[@Teclas],1)+1,LEN([@Teclas])-FIND(""{"",[@Teclas],1)-1))&""]""),"""")"
    Next i
' --- Encabezados y fórmula de la tabla ATAJOS--- '
    For i = 1 To N_Atajos.ListRows.Count                      ' Num de filas en de la Tabla
        N_Atajos.HeaderRowRange(1, 1).Value = "Nombre"
        N_Atajos.HeaderRowRange(1, 2).Value = "On Action"
        N_Atajos.HeaderRowRange(1, 3).Value = "Teclas"
        N_Atajos.HeaderRowRange(1, 4).Value = "[Teclas]"
        
        N_Atajos.DataBodyRange(1, 4).FormulaR1C1 = "=IFERROR(IF([@Teclas]="""","""",""[""&IFERROR(REPLACE(SEARCH(""%"",[@Teclas],1),1,1,""Alt+""),"""")&        IFERROR(REPLACE(SEARCH(""^"",[@Teclas],1),1,1,""Ctrl+""),"""")&" & Chr(10) & "                                                                      IFERROR(REPLACE(SEARCH(""+"",[@Teclas],1),1,1,""May+""),"""")&        IF(IFERROR(SEARCH(""+"",[@Teclas],1),"""")<>""""," & Chr(10) & "UPPER(" & _
            "MID([@Teclas],FIND(""{"",[@Teclas],1)+1,LEN([@Teclas])-FIND(""{"",[@Teclas],1)-1))," & Chr(10) & "                   MID([@Teclas],FIND(""{"",[@Teclas],1)+1,LEN([@Teclas])-FIND(""{"",[@Teclas],1)-1))&""]""),"""")"
    Next i
' --- Encabezados de la tabla COMANDOS --- '
    N_Comandos.HeaderRowRange(1, 1).Value = "CommandBars"
    N_Comandos.HeaderRowRange(1, 2).Value = "Descripción"
    N_Comandos.HeaderRowRange(1, 3).Value = "Activa / Desactiva"
' --- Encabezados de la tabla PANTALLA --- '
    N_Pantalla.HeaderRowRange(1, 1).Value = "NOMBRE HOJA"
    N_Pantalla.HeaderRowRange(1, 2).Value = "Ribbon"
    N_Pantalla.HeaderRowRange(1, 3).Value = "Barra Formulas"
    N_Pantalla.HeaderRowRange(1, 4).Value = "B.D. Vertical"
    N_Pantalla.HeaderRowRange(1, 5).Value = "B.D. Horizontal"
    N_Pantalla.HeaderRowRange(1, 6).Value = "Encabezados"
End Sub
