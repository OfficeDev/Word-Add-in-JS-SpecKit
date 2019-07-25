---
page_type: sample
products:
- office-word
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/24/2016 12:45:01 PM
---
# <a name="word-add-in-javascript-speckit"></a>SpecKit de JavaScript del complemento de Word

Obtenga información sobre cómo crear un complemento que capture e inserte texto reutilizable y cómo puede implementar un proceso de validación de documento simple.

## <a name="table-of-contents"></a>Tabla de contenido
* [Historial de cambios](#change-history)
* [Requisitos previos](#prerequisites)
* [Configurar el proyecto](#configure-the-project)
* [Ejecutar el proyecto](#run-the-project)
* [Entender el código](#understand-the-code)
* [Preguntas y comentarios](#questions-and-comments)
* [Recursos adicionales](#additional-resources)

## <a name="change-history"></a>Historial de cambios

31 de marzo de 2016:
* Versión de ejemplo inicial.

## <a name="prerequisites"></a>Requisitos previos

* Word 2016 para Windows, compilación 16.0.6727.1000 o posterior.
* [Nodo y npm](https://nodejs.org/en/)
* [Git Bash](https://git-scm.com/downloads): debe usar una compilación posterior, ya que las compilaciones anteriores pueden mostrar un error al generar los certificados.

## <a name="configure-the-project"></a>Configurar el proyecto

Ejecute los siguientes comandos desde el shell de Bash en la raíz de este proyecto:

1. Clone este repositorio en el equipo local.
2. ```npm install``` para instalar todas las dependencias en package.json.
3. ```bash gen-cert.sh``` para crear los certificados necesarios para ejecutar este ejemplo. Después, en el repositorio en el equipo local, haga doble clic en ca.crt y seleccione **Instalar certificado**. Seleccione **Máquina local** y seleccione **Siguiente** para continuar. Seleccione **Colocar todos los certificados en el siguiente almacén** y, después, seleccione **Examinar**.  Seleccione **Entidades de certificación raíz de confianza** y después seleccione **Aceptar**. Seleccione **Siguiente** y después **Finalizar**. Ahora, se ha agregado el certificado de la autoridad de certificación al almacén de certificados.
4. ```npm start``` para iniciar el servicio.

Llegados a este punto, ya habrá implementado este complemento de ejemplo. Ahora debe indicarle a Microsoft Word dónde encontrar el complemento.

1. Cree un recurso compartido de red o [comparta una carpeta en la red](https://technet.microsoft.com/en-us/library/cc770880.aspx) y coloque allí el archivo de manifiesto [word-add-in-javascript-speckit-manifest.xml](word-add-in-javascript-speckit-manifest.xml).
3. Inicie Word y abra un documento.
4. Seleccione la pestaña **Archivo** y haga clic en **Opciones**.
5. Haga clic en **Centro de confianza** y seleccione el botón **Configuración del Centro de confianza**.
6. Seleccione **Catálogos de complementos de confianza**.
7. En el campo **Dirección URL del catálogo**, escriba la ruta de red al recurso compartido de carpeta que contiene word-add-in-javascript-speckit-manifest.xml y después elija **Agregar catálogo**.
8. Seleccione la casilla **Mostrar en menú** y, luego, elija **Aceptar**.
9. Aparecerá un mensaje para informarle de que la configuración se aplicará la próxima vez que inicie Microsoft Office. Cierre y vuelva a iniciar Word.

## <a name="run-the-project"></a>Ejecutar el proyecto

1. Abra un documento de Word.
2. En la pestaña **Insertar** de Word 2016, elija **Mis complementos**.
3. Seleccione la pestaña **CARPETA COMPARTIDA**.
4. Elija **el complemento SpecKit de Word** y, después, seleccione **Aceptar**.
5. Si su versión de Word admite los comandos de complemento, la interfaz de usuario le informará de que se ha cargado el complemento.

### <a name="ribbon-ui"></a>Interfaz de usuario de la cinta de opciones
En la cinta de opciones, puede:
* Seleccionar la pestaña del **complemento SpecKit** para iniciarlo en la interfaz de usuario.
* Seleccionar **Insert spec template (Insertar plantilla de especificación)** para iniciar el panel de tareas e insertar una plantilla de especificación en el documento.
* Usar los botones de validación de la cinta de opciones o hacer clic con el botón derecho en el menú contextual para validar el documento con una lista de palabras prohibidas.

 > Nota: El complemento se cargará en un panel de tareas si los comandos del complemento no son compatibles con su versión de Word.

### <a name="task-pane-ui"></a>Interfaz de usuario del panel de tareas
En el panel de tareas, puede:
* Guardar una oración al colocar el cursor sobre ella, asignarle un nombre en el campo que hay sobre **Add sentence to boilerplate (Agregar oración al texto reutilizable)* en el panel de tareas y seleccionar **Add sentence to boilerplate (Agregar oración al texto reutilizable)**. Puede hacer lo mismo para los párrafos.
* Guardar oraciones y párrafos también hará que el texto repetitivo esté disponible en el menú desplegable **Insert boilerplate** (Insertar texto reutilizable).
* Coloque el cursor en el documento. Seleccione un texto reutilizable del menú desplegable y el texto reutilizable se insertará en el documento.
* Cambie la propiedad *Autor* del documento. Para ello, cambie el nombre del autor y seleccione el botón **Update author name** (Actualizar nombre de autor). Esto actualizará tanto la propiedad del documento como el contenido de un control de contenido enlazado.

## <a name="understand-the-code"></a>Entender el código

En este ejemplo se usa el [conjunto de requisitos](http://dev.office.com/reference/add-ins/office-add-in-requirement-sets?product=word) 1.2 durante el período de vista previa, pero necesitará el conjunto de requisitos 1.3 una vez que el conjunto de requisitos esté disponible de forma general.

### <a name="task-pane"></a>Panel de tareas

La funcionalidad del panel de tareas está configurada en sample.js. Sample.js contiene la siguiente funcionalidad:

* Configurar la interfaz de usuario y los controladores de eventos.
* Obtener la plantilla de especificación de un servicio e insertarla en el documento.
* Cargar una lista negra que contiene palabras que se usan para validar el documento. Estas palabras se consideran palabras no autorizadas para este ejemplo.
* Cargar un texto reutilizable predeterminado de un servicio y almacenarlo en caché en el almacenamiento local.
* Código de esqueleto para probar el código de archivo de función. Quiere desarrollar el código de comando del complemento en el panel de tareas antes de moverlo a un archivo de función porque no puede adjuntar un depurador al archivo de función.
* Cargar el nombre de autor predeterminado de las propiedades del documento en el panel de tareas. Esto muestra cómo puede acceder y cambiar un elemento XML personalizado en un documento.
* Publicar el texto reutilizable en el servicio.

### <a name="document-validation-and-the-dialog-api"></a>Validación de documentos y la API de diálogo

Validation.js contiene el código para validar el documento con una lista de palabras prohibidas. El método validateContentAgainstBlacklist() usa el nuevo método splitTextRanges para dividir el documento en intervalos según delimitadores. Los delimitadores en esta función identifican palabras en el documento. Identificamos la intersección de palabras en el documento y la lista negra y pasamos esos resultados al almacenamiento local. Después, usamos el método displayDialogAsync para abrir un diálogo (dialog.html). El diálogo obtiene los resultados de la validación del almacenamiento local y muestra los resultados.

### <a name="boilerplate-text-functionality"></a>Funcionalidad de texto reutilizable

boilerplate.js contiene ejemplos de cómo puede guardar texto reutilizable en el almacenamiento local, actualizar una lista desplegable de Fabric con entradas que corresponden a texto reutilizable guardado e insertar texto reutilizable seleccionado de una lista desplegable. Este archivo contiene ejemplos de:
* splitTextRanges (novedad en el conjunto de requisitos 1.3 de WordApi): split() reemplazará a esta API en una versión futura.
* compareLocationWith (novedad en el conjunto de requisitos 1.3 de WordApi)
* Actualizar la lista desplegable de Fabric con las nuevas entradas.
* Insertar texto reutilizable en el documento.

### <a name="custom-xml-binding-to-core-document-properties"></a>Enlace XML personalizado a propiedades del documento principal

authorCustomXml.js contiene métodos para obtener y establecer las propiedades de documento predeterminadas.

* Cargar la propiedad de autor en el panel de tareas cuando se carga el panel de tareas. Observe que el documento también contiene el valor de la propiedad de autor. Esto se debe a que la plantilla contiene un control de contenido que se enlaza a la propiedad de este documento. Esto le permite establecer valores predeterminados en el documento según el contenido de un elemento XML personalizado.
* Actualizar la propiedad de autor del panel de tareas. Esto actualizará la propiedad del documento y el control de contenido enlazado en el documento.

### <a name="add-in-commands"></a>Comandos de complemento

Las declaraciones del comando del complemento se encuentran en word-add-in-javascript-speckit-manifest.xml. En este ejemplo se muestra cómo crear comandos de complemento en la cinta de opciones y en un menú contextual.

## <a name="questions-and-comments"></a>Preguntas y comentarios

Nos encantaría recibir sus comentarios sobre el ejemplo de SpecKit de Word. Puede enviarnos sus comentarios a través de la sección *Problemas* de este repositorio.

Las preguntas generales sobre desarrollo en Microsoft Office 365 deben publicarse en [Stack Overflow](http://stackoverflow.com/questions/tagged/office-js+API). Asegúrese de que sus preguntas se etiquetan con [office-js] y [API].

## <a name="additional-resources"></a>Recursos adicionales

* [Documentación de complementos de Office](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Centro de desarrollo de Office](http://dev.office.com/)
* [Proyectos de inicio y ejemplos de código de las API de Office 365](http://msdn.microsoft.com/en-us/office/office365/howto/starter-projects-and-code-samples)

## <a name="copyright"></a>Copyright
Copyright (c) 2016 Microsoft Corporation. Todos los derechos reservados.



Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, consulte las [preguntas más frecuentes sobre el Código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
