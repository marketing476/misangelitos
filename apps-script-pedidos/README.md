# Seguimiento de pedidos – Mis Angelitos Baby

Una web (Apps Script) con dos pantallas:
- **Panel privado** (`…/exec?p=admin`): con clave; crea pedidos y cambia estados.
- **Página pública** (`…/exec`): el cliente consulta su pedido con número + email o teléfono.

La hoja de Google Sheets solo es la base de datos (se crean solas las pestañas Pedidos, Historial y Config).

## Instalación
1. En la hoja: Extensiones > Apps Script. Deja un solo archivo `Código.gs` con el contenido de `TodoEnUno.gs`.
2. Guardar. Implementar > Nueva implementación > Aplicación web (ejecutar como: yo, acceso: cualquier persona).
3. Abre el enlace terminado en `/exec?p=admin`, crea tu clave y listo.

Si cambias el código, usa Implementar > Administrar implementaciones > Editar > Nueva versión.

`Codigo.gs`, `Seguimiento.html` y `Admin.html` son las fuentes; `TodoEnUno.gs` es el resultado de juntarlas.
