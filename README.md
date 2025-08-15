# Email Migration Tool

Herramienta para procesar archivos Excel con pares de correos electrónicos y generar un JSON estructurado.

## Cómo usar

1. Coloca tu archivo Excel (`correos.xlsx`) en la carpeta `/input`
2. Ejecuta: `node processEmails.js`
3. El resultado se guardará en `/output/processed_emails.json`

## Requisitos

- Node.js v14+
- Archivo Excel con al menos 2 columnas:
  - Columna 1: Correo actual
  - Columna 2: Nuevo correo

## Instalación

```bash
npm install

## Configuración de columnas

El script buscará automáticamente columnas que contengan estos nombres (no es necesario que coincidan exactamente):
Puedes modificar estos valores editando el archivo `processEmails.js`:
```javascript
const COLUMN_NAMES = {
    CURRENT_EMAIL: 'current_email',
    NEW_EMAIL: 'new_email'
};