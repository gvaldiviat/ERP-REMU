# Reglas de Negocio y Desarrollo para el Motor de Remuneraciones

Este documento establece las directrices y reglas críticas para el cálculo de liquidaciones y auditorías de Membrantec. Cualquier agente o desarrollador que modifique esta base de código debe respetar estrictamente estas directrices para evitar diferencias de cuadratura frente a sistemas externos (Rex+).

---

## 1. Tratamiento de Parámetros Previsionales (Previred)
* **Actualización Mensual**: Los parámetros de UF, UTM, Sueldo Mínimo (IMM), tasa SIS y mutuales cambian periódicamente. Siempre se debe registrar el IMM exacto vigente para el mes de proceso para que la gratificación legal tope (Art. 50, tope de 4.75 IMM / 12) se calcule con precisión centavera.
* **Tasa SIS**: A partir de Julio 2026, la tasa del Seguro de Invalidez y Sobrevivencia (SIS) aumentó al **2.00%** (anteriormente era 1.62%).
* **Políticas de Fallback**: Si no existe archivo físico del mes de Previred, se debe buscar en reversa cronológica el parámetro del mes inmediatamente anterior más cercano, no usar valores predeterminados fijos de meses antiguos.

## 2. Escudo Fiscal de APV (Régimen B)
* **Ajuste de Alcance Líquido**: Bajo el Régimen B de APV (Ahorro Previsional Voluntario), la rebaja tributaria reduce el Impuesto Único.
* Contablemente en los reportes (para cuadrar con Rex+), el ahorro del impuesto generado por el APV debe sumarse contablemente a los descuentos legales y rebajarse de los descuentos voluntarios. Esto mantiene el `Sueldo Líquido` real inalterado, pero cuadra la columna de `Alcance Líquido` al centavo:
  $$\text{Alcance Líquido} = \text{Total Haberes} - (\text{Descuentos Legales} + \text{Escudo Fiscal APVI})$$

## 3. Proporcionalidad de Topes Imponibles por Licencias Médicas
* En Chile, ante la existencia de licencias médicas, los topes imponibles de AFP, Salud y AFC **deben proporcionalizarse** usando los días cotizables calculados como:
  $$\text{Días Cotizables} = \max(0, 30 - \text{Días de Licencia})$$
* La fórmula de tope imponible proporcional que se debe aplicar es:
  $$\text{Tope Proporcional} = \text{Tope UF} \times \text{Valor UF} \times \left( \frac{\text{Días Cotizables}}{30.0} \right)$$
* Nunca se debe aplicar el tope completo de 30 días cuando el colaborador registre días de licencia médica en el periodo.

## 4. Carga y Procesamiento de Sueldo Base
* **No Escalar Sueldo Base Existente**: El sueldo base guardado en la base de datos de empleados representa el sueldo base contractual/pactado. **No se debe multiplicar** por $30 / \text{días}$ cuando el trabajador tiene proporcionalidad de días, a menos que el dato provenga directamente de una liquidación histórica de Rex+ donde el sueldo base original no esté registrado.

## 5. Visualización del Frontend
* **Días de Licencia**: Si el colaborador no registra días de licencia en el mes, la interfaz web debe mostrar un valor de **`0`** limpio en vez de `undefined`.
