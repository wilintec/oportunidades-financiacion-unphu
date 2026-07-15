# Portal de oportunidades de financiamiento

Este repositorio genera `index.html` automáticamente a partir del Excel maestro ubicado en:

```text
datos/oportunidades_financiacion.xlsx
```

## Estructura

```text
.
├── index.html
├── datos/
│   └── oportunidades_financiacion.xlsx
├── scripts/
│   └── generar_portal.py
├── .github/
│   └── workflows/
│       └── actualizar-portal.yml
├── requirements.txt
└── README.md
```

## Actualización normal

1. Abra `datos/oportunidades_financiacion.xlsx`.
2. Añada o modifique las oportunidades en la hoja **Oportunidades**.
3. Use `Sí` en la columna **Publicar** para mostrar una fila en el portal.
4. Sustituya el Excel del repositorio por la nueva versión, conservando exactamente el nombre:

   ```text
   oportunidades_financiacion.xlsx
   ```

5. Confirme el cambio en la rama `main`.

GitHub Actions ejecutará el generador, actualizará `index.html` y guardará automáticamente el resultado en el repositorio.

## Formas de ejecución

El flujo se ejecuta:

- al cambiar el Excel maestro;
- al cambiar el script generador;
- manualmente desde la pestaña **Actions**;
- automáticamente cada tres días a las 06:00, hora de República Dominicana.

## Ejecución manual en GitHub

1. Abra la pestaña **Actions** del repositorio.
2. Seleccione **Actualizar portal desde Excel**.
3. Pulse **Run workflow**.
4. Seleccione la rama `main` y confirme.

## Ejecución local opcional

Requiere Python 3.10 o superior.

```bash
python -m pip install -r requirements.txt
python scripts/generar_portal.py
```

## GitHub Pages

En **Settings → Pages**, configure:

- **Source:** Deploy from a branch
- **Branch:** `main`
- **Folder:** `/ (root)`

## Reglas importantes del Excel

- No cambie el nombre de la hoja `Oportunidades`.
- No cambie los encabezados de las columnas requeridas.
- Guarde las fechas como fechas reales de Excel.
- Use enlaces oficiales completos, comenzando por `https://`.
- Una oportunidad solo se publica cuando `Publicar` contiene `Sí` o `Si`.

## Seguridad

El workflow utiliza el permiso integrado `GITHUB_TOKEN`. No necesita almacenar una contraseña ni un token personal para hacer el commit automático dentro del mismo repositorio.
