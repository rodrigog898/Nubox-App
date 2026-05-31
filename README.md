# Miguel-App-Contractal

Panel local para integrar consultas de Nubox, visualizar resultados y descargar PDFs en lote (ZIP) mediante un proxy Python local.

## Vista previa

> Puedes subir 1 imagen de tu proyecto y dejarla en `assets/img/preview.png`.

![Vista previa del panel](assets/img/preview.png)

## Caracteristicas

- Consulta de documentos con rango de fechas.
- Configuracion local de cookies para autenticacion.
- Visualizacion tabular de resultados con paginacion.
- Apertura de PDF por registro.
- Descarga masiva de PDFs en archivo ZIP.
- Userscript para ejecutar consulta desde Nubox y enviar datos al panel local.

## Estructura del proyecto

- `app/`: paginas HTML.
- `assets/css/`: estilos.
- `assets/js/`: scripts frontend.
- `assets/img/`: imagenes del proyecto.
- `server/`: servidor y proxy local en Python.
- `scripts/`: userscript para Nubox.
- `docs/`: documentacion adicional.

## Requisitos

- Python 3.9+.
- Navegador moderno.
- Dependencias Python: `pypdf` y `openpyxl`.

Instalacion de dependencias:

```powershell
pip install pypdf openpyxl
```

## Ejecucion local

```powershell
python server/web_db_server.py
```

Luego abre:

- http://127.0.0.1:8765/app/index.html
- http://127.0.0.1:8765/app/login.html

## Userscript (opcional)

Archivo: `scripts/nubox_userscript.user.js`

El userscript abre el panel local y envia la respuesta de Nubox usando `postMessage`.

## Generacion Excel por ZIP

En la barra lateral existe la seccion **Generacion Excel por ZIP**:

- Carga un archivo ZIP que contenga PDFs de facturas.
- Pulsa **Analizar ZIP** para extraer RUT, folio, fecha e items.
- Se muestra una tabla en pantalla con una fila por item.
- Pulsa **Descargar Excel** para exportar el resultado.

## Notas

- El servidor local expone endpoints proxy para evitar problemas de CORS.
- Las cookies se almacenan localmente en el navegador.
