# CUADRANTE AUTOMOCIÓN — Guía de instalación

Aplicación web de Google Apps Script, optimizada para móvil, ligada a una hoja de Google Sheets.

## 1. Crear la hoja de cálculo

1. Entra en Google Sheets y crea una hoja nueva (por ejemplo, "Cuadrante Automoción").
2. Menú **Extensiones > Apps Script**. Se abrirá el editor.

## 2. Pegar los archivos

En el editor de Apps Script:

1. En el archivo **Código.gs** que aparece por defecto, borra TODO el contenido y pega el de `Code.gs`.
2. Botón **+ > HTML**. Nómbralo exactamente **Index**, borra el contenido de ejemplo que crea Google y pega el de `Index.html`. Las imágenes (escudo y miniaturas) ya van incluidas dentro de este archivo.
3. Guarda (icono del disquete o Ctrl+S).

> Importante: al pegar cada archivo, comprueba que la ÚLTIMA línea del editor coincide con la última línea del archivo original. Así te aseguras de que el pegado no se ha cortado.

## 3. Crear la estructura de la hoja

1. En el desplegable de funciones (arriba, junto a "Depurar"), selecciona **configurarHoja**.
2. Pulsa **Ejecutar**. La primera vez pedirá permisos: acepta con tu cuenta.
3. Vuelve a la hoja de cálculo: verás las pestañas **Config, Departamentos, Trabajadores, Motivos, Vacaciones y Registros** con datos de ejemplo.

## 4. Personalizar la configuración

- **Config**: cambia la contraseña (fila PASSWORD). Por defecto es `1234`.
- **Departamentos**: escribe tus departamentos reales y su orden de aparición.
- **Trabajadores**: nombre, color (en formato #RRGGBB), departamento (debe coincidir con la pestaña Departamentos) y Activo (SI/NO).
- **Motivos**: motivos de no disponibilidad y su icono (un emoji). Los motivos nuevos que se creen con la opción "Otro" en la app se añadirán aquí automáticamente.

Las pestañas **Vacaciones** y **Registros** las rellena la aplicación sola; no hace falta tocarlas (aunque puedes corregir datos a mano si lo necesitas).

## 5. Publicar la aplicación web

1. En el editor de Apps Script: **Implementar > Nueva implementación**.
2. Tipo: **Aplicación web**.
3. Configuración:
   - **Ejecutar como**: Tú (tu cuenta).
   - **Quién tiene acceso**: *Cualquier usuario con el enlace* (la app ya tiene su propia contraseña de gestor) o *Solo yo* si prefieres exigir además sesión de Google.
4. Pulsa **Implementar** y copia la **URL de la aplicación web**.

## 6. Instalar en el móvil

1. Abre la URL en el navegador del móvil.
2. **Android (Chrome)**: menú ⋮ > *Añadir a pantalla de inicio*.
3. **iPhone (Safari)**: botón compartir > *Añadir a pantalla de inicio*.

Quedará como una app con el icono de Cuadrante Automoción.

## Notas de uso

- **Vacaciones**: toca días libres para seleccionarlos y pulsa *Asignar* para elegir al trabajador. Toca un día ya coloreado (sin selección activa) para ver quién está de vacaciones y poder quitarlo. Cuando coinciden varios trabajadores el día se muestra con rayas de sus colores.
- **Disponibilidad y fichaje**: toca uno o varios trabajadores (si van a hacer lo mismo) y pulsa *Salida* (elegirás el motivo) o *Entrada*. Los trabajadores de vacaciones aparecen en gris y no se pueden fichar.
- **Informes**: filtra por fechas, trabajador o departamento. El tiempo por motivo se calcula emparejando cada salida con su entrada del mismo día.
- **Actualizar datos**: botón ⟳ de la barra superior (útil si otro dispositivo ha hecho cambios).

## Si cambias el código más adelante

Tras cualquier modificación: **Implementar > Administrar implementaciones > editar (lápiz) > Versión: Nueva versión > Implementar**. La URL no cambia.


## Si ya lo habías instalado antes (actualización)

1. Sustituye el contenido de **Código.gs** por el nuevo `Code.gs`.
2. Sustituye el contenido de **Index** por el nuevo `Index.html`.
3. Si existe un archivo **Imagenes** (HTML o secuencia de comandos) de versiones anteriores, elimínalo (menú ⋮ del archivo > Eliminar). Ya no se usa.
4. **Implementar > Administrar implementaciones > icono del lápiz > Versión: "Nueva versión" > Implementar.**
   Este último paso es imprescindible: si no creas una nueva versión, la URL /exec seguirá sirviendo el código antiguo aunque hayas cambiado los archivos.


## Acceso recordado por dispositivo (desde v5)

- Al introducir el PIN con la casilla "Recordar este dispositivo" marcada, ese móvil ya no volverá a pedirlo: la app entra directa.
- Si cambias la contraseña en la pestaña Config, todos los dispositivos recordados caducan automáticamente y vuelven a pedir el PIN.
- Para dejar de recordar un móvil concreto: botón "Cerrar sesión en este dispositivo" al final del menú principal.
- Nota: el sensor de huella no se puede usar porque Google Apps Script bloquea esa función del navegador (WebAuthn) dentro de su entorno aislado. El acceso recordado cumple la misma función práctica.

## Guardar el código en GitHub (opcional)

GitHub no participa en la instalación de la app en los móviles (eso se hace abriendo la URL /exec y usando "Añadir a pantalla de inicio"), pero es muy útil para guardar versiones del código y no perderlo:

1. Crea una cuenta en github.com si no la tienes.
2. Arriba a la derecha: + > New repository. Nombre: cuadrante-automocion. Marca "Private" (recomendado: así solo tú ves el código). Pulsa Create repository.
3. En la página del repositorio: "uploading an existing file" (o botón Add file > Upload files).
4. Arrastra Code.gs, Index.html e INSTRUCCIONES.md y pulsa Commit changes.
5. Cada vez que cambies el código, repite la subida: GitHub guarda el historial completo y puedes recuperar cualquier versión anterior.

Para instalar en cada móvil de los gestores solo hace falta: abrir la URL /exec en Chrome > menú ⋮ > "Añadir a pantalla de inicio". Aparecerá con el escudo como una app más.
