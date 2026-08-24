# DMS Contract Normalizer

Herramienta de normalización de contratos DMS para Andreani.

## 🚀 Características

- 📂 Drag & drop para archivos `.xlsx`
- ✅ Validación de campos en hoja `Formulario` (filas 3-24 y 27-30)
- 🔄 Conversión de fórmulas a valores estáticos
- 🗑️ Eliminación automática de hoja `Formulario`
- 📑 Movimiento de hoja `Referencias` al final
- 🎨 Diseño corporativo Andreani
- 📥 Descarga automática de archivos normalizados

## 🛠️ Tecnologías

- HTML5
- CSS3 (Sistema de diseño Andreani)
- JavaScript (Vanilla)
- SheetJS (xlsx) - v0.18.5

## 📦 Uso

1. Abrir `index.html` en cualquier navegador moderno
2. Arrastrar uno o varios archivos `.xlsx`
3. Click en "Procesar y Normalizar"
4. Los archivos normalizados se descargan automáticamente

## 📋 Proceso de Normalización

1. **Validación**: Verifica que los campos requeridos en `Formulario` tengan valores válidos
2. **Conversión**: Todas las fórmulas se convierten a valores estáticos
3. **Limpieza**: Se elimina la hoja `Formulario`
4. **Organización**: Se mueve la hoja `Referencias` al final

## 🔧 Requisitos

- Navegador moderno (Chrome, Firefox, Edge, Safari)
- Conexión a internet (para cargar SheetJS desde CDN)

## 📄 Licencia

MIT
