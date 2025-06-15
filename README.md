# FormuleViewer - DESARROLLO DE UN ADD-IN VISUALIZACIONES QUE FACILITEN EL USO DE MS- EXCEL 

FormuleViewer es un complemento personalizado para Microsoft Excel que permite visualizar las fórmulas de la hoja activa desde un panel lateral. Está preparado para ejecutarse tanto **en local**, 
como **en Excel Online** utilizando archivos servidos a través de **GitHub Pages**.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

## Ejecución en Ms-Excel local (usada durante el desarrollo)

### Requisitos

Para ejecutar el complemento en local necesitas:

- [Node.js](https://nodejs.org) (versión 14 o superior)
- Microsoft Excel instalado en tu equipo (versión compatible con Office Add-ins)


### Instrucciones de uso local

1. Clona este repositorio
2. Ejecuta los siguientes comando en la consola de comandos desde el modo administrador:
    * cd tuRuta/FormuleViewer
    * `npm install` Para instalar las depencias
    * `npm start`  Esto lanzará un servidor local con el panel del complemento y abrirá automáticamente Excel con una hoja de cálculo en blanco. En la pestaña de complementos verás cargado FormuleViewer.
      Este proceso se conoce como sideloading, y permite probar el complemento sin publicarlo en AppSource.

El manifiesto utilizado en este entorno es manifest.xml, el cual contiene rutas apuntando a https://localhost, ejecutando el código pricial clonado, a excepción de logos e imágenes.

3. `npm stop` Para detener el servidor
   
  Nota:  En algunos entornos Windows es necesario habilitar el loopback para WebView2, ya que, por seguridad, puede estar deshabilitado y bloquear la carga del panel lateral del complemento

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
## Ejecución en Ms-Excel online (Usado pruebas online)
1. Abre Excel Online.

2. Crea o abre una hoja de cálculo.

3. Ve a Insertar > Complementos > Mis complementos > Agregar un complemento personalizado > Agregar desde archivo.

4. Selecciona el archivo manifest_online.xml ubicado en la carpeta /manifest_online de este repositorio.

Ese manifiesto apunta a archivos servidos desde GitHub Pages, por lo que todo el complemento funcionará online sin depender de localhost.

--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Durante el desarrollo se ha observado que el complemento funciona más eficientemente en entorno local, debido a que todos los recursos se cargan desde el mismo equipo, reduciendo la latencia.
En cambio, en el entorno remoto (como GitHub Pages), el rendimiento puede verse afectado por la conexión de red o la disponibilidad del servidor.
