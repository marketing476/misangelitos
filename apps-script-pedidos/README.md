# Seguimiento de pedidos – Mis Angelitos Baby (Apps Script)

Hoja de Google Sheets como base de datos + página pública de seguimiento + correos automáticos.

## Instalación (5 minutos)
1. Crea una hoja de Google Sheets nueva.
2. **Extensiones > Apps Script**. Crea los archivos y pega el contenido de esta carpeta:
   `Codigo.gs`, `Seguimiento.html`, `NuevoPedido.html`.
   (En Configuración del proyecto activa "Mostrar appsscript.json" y pega también `appsscript.json`.)
3. Recarga la hoja → aparece el menú **📦 Pedidos > ⚙️ Configuración inicial** (acepta los permisos).
4. **Implementar > Nueva implementación > Aplicación web**: ejecutar como *yo*, acceso *cualquier persona*.
5. **📦 Pedidos > 🔗 Ver enlace de seguimiento** y comparte ese enlace con tus clientes.

## Uso diario
- **📦 Pedidos > ➕ Nuevo pedido** crea el pedido y envía el correo de confirmación.
- Cambia el **Estado** en la hoja *Pedidos* (lista desplegable): se guarda en *Historial* y se avisa al cliente.
- Rellena *Transportadora*, *Guía* y *Nota para el cliente* para que aparezcan en la página y en el correo.
- La hoja *Config* permite cambiar el nombre de la tienda, desactivar correos o poner un enlace de rastreo.

## Seguridad
El cliente solo ve su pedido si escribe el número **y** su email o teléfono; hay límite de 8 intentos por hora y pedido.
Si cambias el código, crea una **nueva versión** de la implementación para que se publique.
