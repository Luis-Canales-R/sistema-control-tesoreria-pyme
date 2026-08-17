Attribute VB_Name = "Modulo_ReportesFinancieros"
Sub Actualizar_ReportesFinancieros()
    Dim ws As Worksheet
    Dim wsPlan As Worksheet
    Dim totalFilas As Long
    
    Set ws = Worksheets("Reportes_Financieros")
    Set wsPlan = Worksheets("Plan_de_Cuentas")
    
    ws.Cells.Clear
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    
    totalFilas = wsPlan.Range("A" & wsPlan.Rows.Count).End(xlUp).Row
    If totalFilas < 2 Then Exit Sub
    
    ' --- SECCIÓN A: ESTADO DE RESULTADOS (DEVENGADO) ---
    ws.Range("A1").Value = "ESTADO DE RESULTADOS (DEVENGADO)"
    ws.Range("A1:D1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 12
    ws.Range("A1").Interior.Color = RGB(41, 128, 185)
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    
    ws.Range("A2:D2").Value = Array("Codigo", "Categoria", "Tipo", "Monto_Devengado")
    ws.Range("A2:D2").Font.Bold = True
    ws.Range("A2:D2").Interior.Color = RGB(235, 245, 251)
    
    ws.Range("A3:A" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!A2"
    ws.Range("B3:B" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!C2"
    ws.Range("C3:C" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!B2"
    
    ws.Range("D3:D" & (1 + totalFilas)).FormulaLocal = _
        "=LET(ing; SUMAR.SI.CONJUNTO(T_Compromisos[Monto_Neto]; T_Compromisos[Categoria]; B3; T_Compromisos[Tipo]; ""COBRAR""); " & _
        "egr; SUMAR.SI.CONJUNTO(T_Compromisos[Monto_Neto]; T_Compromisos[Categoria]; B3; T_Compromisos[Tipo]; ""PAGAR""); " & _
        "SI(C3=""Ingreso Operacional""; ing; -egr))"
    
    ws.Range("A" & (2 + totalFilas)).Value = "TOTAL GENERAL"
    ws.Range("A" & (2 + totalFilas) & ":C" & (2 + totalFilas)).Merge
    ws.Range("A" & (2 + totalFilas)).Font.Bold = True
    ws.Range("D" & (2 + totalFilas)).FormulaLocal = "=SUMA(D3:D" & (1 + totalFilas) & ")"
    ws.Range("D" & (2 + totalFilas)).Font.Bold = True
    ws.Range("A" & (2 + totalFilas) & ":D" & (2 + totalFilas)).Interior.Color = RGB(214, 234, 248)
    
    ' --- SECCIÓN B: FLUJO DE CAJA (EFECTIVO) ---
    ws.Range("F1").Value = "FLUJO DE CAJA (EFECTIVO)"
    ws.Range("F1:I1").Merge
    ws.Range("F1").Font.Bold = True
    ws.Range("F1").Font.Size = 12
    ws.Range("F1").Interior.Color = RGB(39, 174, 96)
    ws.Range("F1").Font.Color = RGB(255, 255, 255)
    
    ws.Range("F2:I2").Value = Array("Codigo", "Categoria", "Tipo_Mov", "Monto_Efectivo")
    ws.Range("F2:I2").Font.Bold = True
    ws.Range("F2:I2").Interior.Color = RGB(234, 250, 241)
    
    ws.Range("F3:F" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!A2"
    ws.Range("G3:G" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!C2"
    ws.Range("H3:H" & (1 + totalFilas)).FormulaLocal = "=Plan_de_Cuentas!B2"
    
    ws.Range("I3:I" & (1 + totalFilas)).FormulaLocal = _
        "=LET(ing; SUMAR.SI.CONJUNTO(T_MovimientosBancarios[Monto]; T_MovimientosBancarios[Categoria]; G3; T_MovimientosBancarios[Tipo_Movimiento]; ""INGRESO""); " & _
        "egr; SUMAR.SI.CONJUNTO(T_MovimientosBancarios[Monto]; T_MovimientosBancarios[Categoria]; G3; T_MovimientosBancarios[Tipo_Movimiento]; ""EGRESO""); " & _
        "ing - egr)"
    
    ws.Range("F" & (2 + totalFilas)).Value = "TOTAL GENERAL"
    ws.Range("F" & (2 + totalFilas) & ":H" & (2 + totalFilas)).Merge
    ws.Range("F" & (2 + totalFilas)).Font.Bold = True
    ws.Range("I" & (2 + totalFilas)).FormulaLocal = "=SUMA(I3:I" & (1 + totalFilas) & ")"
    ws.Range("I" & (2 + totalFilas)).Font.Bold = True
    ws.Range("F" & (2 + totalFilas) & ":I" & (2 + totalFilas)).Interior.Color = RGB(212, 239, 223)
    
    ws.Range("D3:D" & (2 + totalFilas) & ", I3:I" & (2 + totalFilas)).NumberFormat = "$#,##0;($#,##0);$0"
    ws.Columns("A:I").AutoFit
    
    MsgBox "Reportes Financieros actualizados correctamente con formato de región local.", vbInformation, "Tesorería PYME"
End Sub
