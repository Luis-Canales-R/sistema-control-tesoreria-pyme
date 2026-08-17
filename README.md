Markdown
# Sistema de Control de Tesorería PYME

Sistema de gestión financiera en Excel que integra el control del **Devengado** (compromisos) y el **Efectivo** (flujo de caja real).

## 📐 Estructura de Tablas y Reglas

### 1. `T_PlanCuentas` (`Plan_de_Cuentas`)
* **`Codigo`**: Identificador contable numérico.
* **`Tipo_Cuenta`**: Clasificación funcional (`Ingreso Operacional`, `Costo Directo`, `Gasto Operacional`, `Gasto Financiero`, `Transferencia Interna`).
* **`Categoria`**: Nombre de la cuenta (rango nombrado `=Lista_Categorias`).
* **`Afecta_Reporte`**: Criterio de despliegue (`Ambos`, `Estado de Resultados`, `Flujo de Caja`).
* **`Estado`**: Estado operativo (`ACTIVO` / `INACTIVO`).

### 2. `T_Compromisos` (`Compromisos`)
* **`ID_Compromiso`**: Clave primaria (`CMP-001`, `CMP-002`).
* **`Tipo`**: Lista desplegable (`COBRAR` / `PAGAR`).
* **`Categoria`**: Validación con `=Lista_Categorias`.
* **`Monto_Neto`**: Subtotal sin impuestos.
* **`Monto_IVA`**: Impuesto aplicado.
* **`Monto_Total`**: Campo calculado: `=[@[Monto_Neto]]+[@[Monto_IVA]]`.
* **`Estado_Pago`**: Lista desplegable (`PENDIENTE`, `PAGADO`, `PARCIAL`, `ANULADO`).

### 3. `T_MovimientosBancarios` (`Movimientos_Bancarios`)
* **`ID_Movimiento`**: Clave primaria (`MOV-001`, `MOV-002`).
* **`Tipo_Movimiento`**: Lista desplegable (`INGRESO` / `EGRESO`).
* **`Monto`**: Valor bruto abonado o cargado.
* **`ID_Compromiso`**: Enlace opcional a la factura o compromiso liquidado.
* **`Categoria`**: Validación con `=Lista_Categorias`.
* **`Estado_Conciliacion`**: Estado del registro (`CONCILIADO` / `PENDIENTE`).

## 🧮 Lógica de Consolidación Financiera (`Reportes_Financieros`)

* **Estado de Resultados (Devengado):**
  ```excel
  =LET(ing; SUMAR.SI.CONJUNTO(T_Compromisos[Monto_Neto]; T_Compromisos[Categoria]; B3; T_Compromisos[Tipo]; "COBRAR"); egr; SUMAR.SI.CONJUNTO(T_Compromisos[Monto_Neto]; T_Compromisos[Categoria]; B3; T_Compromisos[Tipo]; "PAGAR"); SI(C3="Ingreso Operacional"; ing; -egr))
Flujo de Caja (Efectivo):

Excel
=LET(ing; SUMAR.SI.CONJUNTO(T_MovimientosBancarios[Monto]; T_MovimientosBancarios[Categoria]; G3; T_MovimientosBancarios[Tipo_Movimiento]; "INGRESO"); egr; SUMAR.SI.CONJUNTO(T_MovimientosBancarios[Monto]; T_MovimientosBancarios[Categoria]; G3; T_MovimientosBancarios[Tipo_Movimiento]; "EGRESO"); ing - egr)
⚙️ Módulo de Automatización
La carpeta /src almacena Modulo_ReportesFinancieros.bas, encargado de reestructurar dinámicamente la hoja de reportes, inyectar fórmulas locales en español con separador ; y dar formato de moneda $#,##0;($#,##0);$0.


3. Haz clic en **"Commit changes..."** > **"Commit changes"**.

¡Avísame cuando hagas los 3 pasos en la web para sincronizar la copia local en VS Code!
