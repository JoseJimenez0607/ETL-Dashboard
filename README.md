# ☕ Café Aurora — Dashboard ETL

## Cómo ejecutar en VS Code (3 pasos)

### Requisitos previos
- Tener instalado **Node.js** → descárgalo en https://nodejs.org (versión LTS)
- Verificar que funciona: abre una terminal y escribe `node --version`

---

### Paso 1 — Abrir la carpeta en VS Code
1. Descomprime este proyecto donde quieras (ej: Escritorio)
2. Abre VS Code
3. Ve a **File → Open Folder** y selecciona la carpeta `cafeteria-dashboard`

---

### Paso 2 — Instalar dependencias
1. En VS Code abre la terminal: **Terminal → New Terminal** (o Ctrl + `)
2. Escribe este comando y presiona Enter:

```bash
npm install
```

Espera ~1 minuto mientras descarga React y las librerías.

---

### Paso 3 — Ejecutar el proyecto
En la misma terminal escribe:

```bash
npm start
```

¡Listo! Se abrirá automáticamente en tu navegador en http://localhost:3000

---

## Estructura del proyecto

```
cafeteria-dashboard/
├── public/
│   └── index.html        ← Página HTML base
├── src/
│   ├── index.js          ← Punto de entrada
│   └── App.jsx           ← Dashboard completo (aquí está toda la lógica)
└── package.json          ← Dependencias del proyecto
```

## Próximos pasos

- [ ] Conectar con archivos Excel reales (.xlsx)
- [ ] Agregar módulo de preguntas automáticas para columnas desconocidas
- [ ] Exportar reportes en PDF
- [ ] Conectar a base de datos
