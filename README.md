# DMS Contract Normalizer

Herramienta de normalización de contratos DMS para Andreani.

## 🚀 Características

- 📂 **Drag & drop** para archivos `.xlsx`
- ✅ **Validación avanzada** de campos en hoja `Formulario` (filas 3-24 y 27-30)
- 🔄 **Conversión de fórmulas** a valores estáticos
- 🗑️ **Eliminación automática** de hoja `Formulario`
- 📑 **Movimiento** de hoja `Referencias` al final
- 🏢 **Limpieza automática** de direcciones vacías en hoja `Direcciones`
- 🔢 **Normalización inteligente** de `CodigoSolicitanteERP` (6 dígitos / 10 dígitos con 4 ceros)
- 📝 **Nombre de archivo personalizado** (`Alta de contrato DMS [ERP] [Solicitante] [Cliente].xlsx`)
- 💬 **Mensaje automático** para copiar al portapapeles con lista de archivos procesados
- 🎨 **Diseño corporativo Andreani**
- 📥 **Descarga automática** de archivos normalizados

## 🛠️ Tecnologías

- HTML5
- CSS3 (Sistema de diseño Andreani)
- JavaScript (Vanilla)
- SheetJS (xlsx) - v0.18.5

## 📦 Uso

1. Abrir `index.html` en cualquier navegador moderno
2. Arrastrar uno o varios archivos `.xlsx`
3. Click en **"Procesar y Normalizar"**
4. Los archivos normalizados se descargan automáticamente
5. El mensaje con los nombres de los archivos se muestra en la card derecha para copiar

## 📋 Proceso de Normalización

### 1. Validación de `Formulario`
- Verifica que los campos requeridos en `Formulario` tengan valores válidos según la hoja `Referencias`
- Valida que `CodigoContratoERP` tenga 9 dígitos
- Valida que `IDSociedadSAP` sea `1100`
- Valida que `Habilitado` sea `1` o `0`
- Valida campos booleanos (true/false, 1/0, si/no)
- **Normaliza automáticamente** `HasDeliveryNote`:
  - `true`, `1`, `si` → `CON_REMITO`
  - `false`, `0`, `no` → `SIN_REMITO`
  - `CON_REMITO`, `SIN_REMITO`, `DIGITALIZADO` → se mantienen

### 2. Normalización de `CodigoSolicitanteERP`
- **6 dígitos** → válido (ej: `123456`)
- **10 dígitos comenzando con 4 ceros** → normaliza eliminando los 4 ceros iniciales (ej: `0000120761` → `120761`)
- **10 dígitos sin 4 ceros iniciales** → ❌ Error
- **Menos de 6 o más de 10 dígitos** → ❌ Error
- **No numérico** → ❌ Error

### 3. Limpieza de Direcciones
- Elimina automáticamente las filas de la hoja `Direcciones` que tengan **todos los campos vacíos o en 0**
- Solo conserva las direcciones que tienen al menos un campo con información real

### 4. Nombre de archivo personalizado
- Formato: `Alta de contrato DMS [CodigoContratoERP] [CodigoSolicitanteERP] [ClientName].xlsx`
- `CodigoContratoERP` (C3) → **Obligatorio**
- `CodigoSolicitanteERP` (C5) → **Obligatorio** (con normalización automática)
- `ClientName` (C14) → **Opcional** (si está vacío, se omite)

### 5. Conversión y Limpieza final
- **Conversión**: Todas las fórmulas se convierten a valores estáticos
- **Limpieza**: Se elimina la hoja `Formulario`
- **Organización**: Se mueve la hoja `Referencias` al final

### 6. Mensaje automático
- Se genera automáticamente con el formato:

Hola @Djirikian Damian Uriel, te paso [cantidad] archivos para alta de contratos DMS

[nombre del archivo 1]
[nombre del archivo 2]
[nombre del archivo 3]

- Botón **"📋 Copiar mensaje"** para copiar al portapapeles

## 🔧 Requisitos

- Navegador moderno (Chrome, Firefox, Edge, Safari)
- Conexión a internet (para cargar SheetJS desde CDN)

## 📄 Licencia

MIT
