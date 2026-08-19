# Análisis de Reposición / Trade Marketing

Este proyecto genera un Excel de análisis (`SC_Analisis_TradeMarketing_<fecha_hora>.xlsx`) con 3 hojas: **ReposicionxRuta**, **ReposicionxCliente** y **ClientesSinReposicion**.

## Cómo ejecutarlo

```
python scripts\build_analysis.py
```

Cada vez que se corre, se genera un **archivo nuevo** (no se sobreescribe el anterior), con el nombre:

```
SC_Analisis_TradeMarketing_AAAAMMDD_HHMMSS.xlsx
```

## Archivos de entrada requeridos

Deben colocarse dentro de la carpeta `Data/` con exactamente estos nombres:

| Archivo | Contenido | Columnas clave que usa el análisis |
|---|---|---|
| `clientes.xlsx` | Maestro de clientes y a qué ruta pertenece cada uno | `Codigo Cliente`, `Nombre cliente`, `Ruta asignada` |
| `sales.xlsx` | Ventas (hoja "Ventas Combo Armado") | `Fecha`, `Cliente` (código), `Articulo`, `Venta Neta` |
| `clientes_reccorridos.xlsx` | Registro de visitas del reponedor (recorridos de julio) | `Fecha`, `Cliente` (nombre), `Ruta`, `Tiempo Atendido`, `Estado` |

El periodo de ventas analizado es **enero a julio 2026** (se filtra automáticamente por fecha).

## Filtro de rutas

Solo se analizan las rutas cuyo nombre contiene "WHS" (bodegas/mercados) más la ruta `SC-LICORES RAMADA`. Se excluye explícitamente `SC-WHS-MERC-LA-RAMADA-2`.

Las rutas que son la misma zona pero divididas en sub-rutas (sufijo `-1`, `-2`, `-3`, `-4`) se **unifican en un solo "Mercado"** usando un diccionario fijo dentro del script (ej. `SC-WHS-MERC-LA-RAMADA-1/3/4` → `MERCADO LA RAMADA`). `SC-LICORES RAMADA` es un canal distinto (licorerías) y se deja tal cual, sin unificar. El diccionario completo está en `scripts/build_analysis.py`, variable `RUTA_A_MERCADO`.

## Cómo se calcula cada medida

### Hoja 1 — ReposicionxRuta

- **Total ventas Ruta**: suma de "Venta Neta" de todas las ventas (enero–julio) de los clientes que pertenecen a ese mercado.
- **Clasificación ABC (de la ruta)**: Pareto 80/15/5 sobre el total de ventas de todos los mercados. Se ordenan los mercados de mayor a menor venta; los que en conjunto acumulan hasta el 80% de la venta total son **A**, el siguiente tramo hasta 95% es **B**, y el resto es **C**.
- **Cantidad de clientes de esa ruta**: cantidad de clientes distintos asignados a ese mercado según el maestro de clientes.
- **Cantidad de clientes atendidos**: cantidad de clientes distintos de ese mercado que tuvieron al menos una visita con Estado "VISITADO" o "INCOMPLETO" durante julio.
- **Clientes sin Atención**: Cantidad de clientes − Cantidad de clientes atendidos.
- **% Atención**: Cantidad de clientes atendidos ÷ Cantidad de clientes de esa ruta.
- **Tiempo Atendido (en base a la moda)**: tiempo típico que dura una visita en ese mercado. Se calcula así:
  1. Se descartan las visitas con 0 segundos de atención (no reflejan trabajo real).
  2. Para cada día del mes se calcula el promedio del tiempo de atención de ese mercado.
  3. De todos esos promedios diarios, se toma el valor que **más se repite** (la moda), redondeado a minutos (sin segundos).
- **Media clientes atendidos por Día**: rango de cuántos clientes puede cubrir un reponedor en un día en ese mercado, expresado como `mínimo-máximo`:
  - **Mínimo** = promedio de clientes atendidos por día a lo largo del mes.
  - **Máximo** = percentil 75 de clientes atendidos por día (el valor que deja al 75% de los días por debajo; representa "un buen día", sin ser un extremo aislado).

### Hoja 2 — ReposicionxCliente (solo clientes atendidos al menos una vez)

- **ABC_Ruta**: la clasificación ABC del mercado al que pertenece (igual que en Hoja 1).
- **Total vendido BS.**: suma de "Venta Neta" del cliente (enero–julio).
- **ABC_Ventas**: Pareto 80/15/5, pero calculado sobre el total de venta de **todos los clientes** (atendidos y no atendidos juntos), para que la clasificación sea comparable entre las hojas 2 y 3.
- **Cantidad de SKUs**: cantidad de artículos distintos que compró ese cliente en el periodo.
- **ABC_SKUs**: Pareto 80/15/5 sobre la cantidad de SKUs comprados, calculado sobre todos los clientes.
- **Moda de tiempo**: igual método que "Tiempo Atendido" de la Hoja 1 (promedio diario → moda de esos promedios, en minutos), pero calculado solo con las visitas de ese cliente específico.
- **Cantidad de Visitas**: número de días de julio en que ese cliente fue atendido (Estado "VISITADO" o "INCOMPLETO").

### Hoja 3 — ClientesSinReposicion (clientes que nunca fueron atendidos)

Mismas columnas que la Hoja 2 excepto "Moda de tiempo" y "Cantidad de Visitas" (no aplican porque no hubo visita). Un cliente puede aparecer aquí con $0 en ventas si no compró nada en el periodo — igual se incluye.

## Notas de cruce de datos

- El cruce entre `clientes_reccorridos.xlsx` y `clientes.xlsx` se hace principalmente por **nombre de cliente** (normalizado: mayúsculas, sin tildes, espacios limpios), porque el archivo de recorridos no tiene el código de cliente. La ruta del archivo de recorridos se usa solo como desempate cuando el mismo nombre existe en más de un mercado.
- Cuando el maestro de clientes tiene dos códigos distintos con el mismo nombre y el mismo mercado (probable duplicado de carga), se fusionan en una sola fila sumando sus ventas y SKUs, y la visita se les asigna conjuntamente (no se puede saber con certeza a cuál de los dos códigos corresponde).
- Un pequeño número de nombres del archivo de recorridos (visible en la consola al ejecutar el script) no logra cruzarse con el maestro de clientes (nombre no encontrado, o nombre ambiguo sin ruta que lo desempate) y queda excluido del análisis.
