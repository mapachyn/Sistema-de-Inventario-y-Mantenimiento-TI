# Sistema de Inventario y Mantenimiento TI

Herramienta desarrollada en **PowerShell y Batch** para automatizar el inventario, diagnóstico y mantenimiento preventivo de equipos Windows en entornos corporativos.

El sistema recopila información del hardware y software, ejecuta tareas de mantenimiento, permite registrar información adicional del activo y genera un informe técnico en formato Word.

Además, incorpora integración con **Microsoft Power Automate** para enviar y registrar información del inventario.

## Funcionalidades principales

### 1. Inventario de equipos
Recopila automáticamente información como:

- Fabricante, modelo y número de serie.
- Procesador (CPU), memoria RAM y almacenamiento.
- Sistema operativo y características del equipo.
- Información de red y adaptadores.
- Estado de la batería.

También permite ingresar información manual:

- Placa de activo.
- Usuario y departamento.
- Fecha de finalización de garantía.
- Accesorios del equipo.
- Observaciones adicionales.

### 2. Mantenimiento automatizado
Permite ejecutar diferentes tareas de diagnóstico y mantenimiento:

- Verificación de actualizaciones de Windows.
- Comprobación del estado del antivirus.
- Validación de perfiles del firewall.
- Análisis de eventos del sistema.
- Limpieza de archivos temporales.
- Análisis de distribución del almacenamiento por carpetas.
- Verificación y corrección de permisos administrativos locales.

### 3. Modalidades de ejecución

**Modo automático:** ejecuta las tareas de mantenimiento de manera automatizada, muestra su progreso y permite continuar con el registro de información del activo.

**Modo manual:** permite seleccionar y ejecutar tareas individuales según las necesidades del técnico.

### 4. Generación de informes
Genera un informe técnico en formato **Microsoft Word (.docx)** que consolida la información del equipo y los resultados de las tareas de mantenimiento.

### 5. Integración con Power Automate
Permite enviar los datos del inventario mediante solicitudes HTTP POST en formato JSON a un flujo de Microsoft Power Automate, para su procesamiento y registro en una lista de inventario.

## Tecnologías utilizadas

- **PowerShell:** recopilación de información, diagnóstico y mantenimiento.
- **Batch (.bat):** inicio y ejecución del sistema.
- **Power Automate:** automatización del registro de información.
- **JSON:** intercambio y almacenamiento temporal de datos.
- **HTTP POST:** comunicación con el flujo.
- **Microsoft Word:** generación de informes técnicos.

## Requisitos

- Sistema operativo Windows.
- Windows PowerShell 5.1 o compatible con los módulos utilizados.
- Permisos administrativos para las tareas que los requieren.
- Conectividad de red para utilizar Power Automate.
- Archivos y dependencias del proyecto en la misma estructura de carpetas.

## Cómo utilizarlo

1. Descargar o copiar los archivos del proyecto.
2. Ejecutar el archivo `.bat` principal como administrador.
3. Seleccionar el modo de ejecución: **Automático** o **Manual**.
4. Seguir las instrucciones mostradas en pantalla.
5. Completar la información solicitada del activo.
6. Revisar los resultados de las tareas ejecutadas.
7. Generar y guardar el informe de mantenimiento en Word.
8. Enviar los datos a Power Automate cuando la integración esté habilitada.

## Consideraciones de seguridad

El sistema está diseñado para utilizarse por personal autorizado de TI.

Algunas funciones pueden modificar configuraciones del equipo, incluidos los permisos administrativos locales. Se recomienda probar la herramienta antes de implementarla en un entorno de producción.

**Importante:** no publicar contraseñas, credenciales, tokens, URL privadas de flujos de Power Automate ni información sensible de la infraestructura corporativa en el repositorio.

## Objetivo del proyecto

Centralizar y automatizar las actividades de inventario y mantenimiento preventivo de equipos corporativos, reduciendo el trabajo manual, estandarizando los procedimientos técnicos y facilitando la trazabilidad de los activos tecnológicos.

## Versión pública del proyecto

Este repositorio contiene una versión depurada del sistema de inventario y mantenimiento TI.

Por motivos de seguridad, se han excluido las credenciales, configuraciones privadas y otros datos potencialmente sensibles de la infraestructura corporativa.

Algunas funcionalidades, como la integración con Microsoft Power Automate, requieren configurar los parámetros correspondientes antes de su utilización.

El propósito de este repositorio es documentar el desarrollo de la herramienta y demostrar la implementación de procesos automatizados mediante PowerShell, Batch y tecnologías Microsoft.
