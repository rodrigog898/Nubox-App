# Miguel-App-Contractal

Panel local para integrar consultas de Nubox, visualizar resultados y descargar PDFs en lote (ZIP) mediante un proxy Python local.

## Vista previa

<img width="1917" height="901" alt="image" src="https://github.com/user-attachments/assets/7a47f2ab-b71a-466b-beaa-06c472df0465" />


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

## Notas

- El servidor local expone endpoints proxy para evitar problemas de CORS.
- Las cookies se almacenan localmente en el navegador.
