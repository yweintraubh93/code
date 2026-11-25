# ☕ Cuentas de la Cafetería

Toma una foto (o un PDF escaneado) de la hoja donde los niños anotan lo que consumen fiado.
La IA lee la letra y empareja cada renglón con la lista de alumnos. Tú revisas, corriges lo que haga falta y guardas.
Al guardar:

- se actualizan las columnas **DEBE** y **A FAVOR** de la pestaña **Alumnos**,
- los consumos de los Barzilay se suman en su pestaña **BARZILAY** (renglón del día, columna de cada hijo),
- y todo queda registrado en una pestaña nueva, **Movimientos** (historial con saldo antes y después).

Todo vive dentro de tu misma hoja de Google: no hay servidor ni nada que pagar aparte de la IA.

## Cómo funcionan los saldos

| Caso | Qué pasa |
|---|---|
| Consumo | Primero se descuenta de **A FAVOR**; lo que falte se suma a **DEBE**. |
| Depósito | Primero se descuenta de **DEBE**; lo que sobre se suma a **A FAVOR**. |
| `60 Sophie 25-25` | Depósito de $60 y dos consumos de $25. |
| `25-25` o `25+25+20` | Varios consumos, uno por cada número. |
| Delgadillo, Guindi, Nagar, Nissan | Siempre van a su cuenta familiar, aunque el niño ponga su nombre. |
| Barzilay | Se suma en la pestaña BARZILAY. Los depósitos de ellos se siguen anotando a mano. |
| Renglón tachado | Aparece desmarcado. |
| Hoja que ya se había guardado | Avisa y deja todo desmarcado para no cobrar dos veces. |

Los renglones dudosos (letra difícil, hermanos con el mismo apellido, sin monto) salen en **amarillo**. No se puede guardar hasta que todos tengan cuenta y monto.

## Instalación (una sola vez, ~15 minutos)

### 1. Pasar el Excel a Hoja de Google
1. En Google Drive abre `Lista_Alumnos_Alfabetica.xlsx`.
2. **Archivo → Guardar como Hoja de cálculo de Google**.
3. A partir de ahora tú y tu mamá usan **esa** hoja de Google, no el Excel. Si alguien sigue editando el Excel, los cambios no se juntan.

### 2. Pegar el código
1. En la hoja de Google: **Extensiones → Apps Script**.
2. Borra lo que haya en `Código.gs` y pega el contenido de [`apps-script/Code.gs`](apps-script/Code.gs).
3. Haz clic en **+ → HTML**, nómbralo exactamente `Index` y pega el contenido de [`apps-script/Index.html`](apps-script/Index.html).
4. En ⚙️ **Configuración del proyecto**, activa *"Mostrar el archivo de manifiesto appsscript.json"* y reemplaza su contenido con [`apps-script/appsscript.json`](apps-script/appsscript.json).
5. Guarda (💾).

### 3. Clave de la IA (Anthropic)
1. Entra a <https://console.anthropic.com>, crea tu cuenta y en **Billing** carga crédito ($5 USD alcanzan para muchas hojas).
2. En **API Keys** crea una clave y cópiala (empieza con `sk-ant-`).
3. Regresa a tu hoja de Google y recárgala. Aparece el menú **☕ Cafetería → Configurar clave de API**. Pega la clave ahí.
   La primera vez Google te pide permisos: acepta. Es tu propio script.

### 4. Publicar la app para el celular
1. En Apps Script: **Implementar → Nueva implementación → Tipo: App web**.
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Solo yo**
2. **Implementar** y copia el link (también lo ves en **☕ Cafetería → Abrir app para fotos**).
3. Abre el link en tu celular (con tu cuenta de Google) y agrégalo a la pantalla de inicio.

> Si después cambias el código: **Implementar → Administrar implementaciones → ✏️ → Versión: Nueva** para que el link use la versión nueva.

## Uso diario
1. Abre la app → **Tomar foto o subir hoja** (puedes elegir varias fotos o un PDF de varias páginas).
2. Espera unos 30 segundos por hoja.
3. Revisa lo amarillo: elige la cuenta (toca una sugerencia o escribe para buscar), corrige montos y fecha.
4. **Guardar**. Si te equivocaste: **Deshacer este lote**, o desde la hoja **☕ Cafetería → Deshacer el último lote guardado**.

## Costo aproximado
Cada página cuesta unos centavos de dólar (aprox. 1 a 2 pesos), según qué tan llena esté.

## Para cambiar reglas
Al principio de `Code.gs`, en `CONFIG`:
- `FAMILIAS_CUENTA_UNICA`: familias que siempre van a su cuenta familiar.
- `HOJA_ALUMNOS`, `HOJA_BARZILAY`: nombres de las pestañas.
