# Agentes de Apoyo a la Gestión – UPROGE

Programa de escritorio de la Unidad de Programación y Gestión Presupuestaria (UPROGE) del Ministerio de Agricultura de Chile. Reúne agentes que automatizan informes presupuestarios. Por ahora trae uno:

- **Ejecución Presupuestaria - Power BI**: arma el consolidado mensual MINAGRI que alimenta el Informe de Ejecución en Power BI, a partir de los informes de SIGFE/ERP de los servicios y de las planillas de Presupuesto Vigente.

## Qué hay en este repositorio

Solo el programa: código (`programa/`), íconos y logo (`recursos/`), los accesos `.bat` y un `config.ini` sin rutas. **No contiene datos institucionales**: el Catálogo Maestro, el manual de usuario, el historial de versiones y los informes se distribuyen solo dentro de UPROGE.

## Actualizaciones

Cada versión se publica como *Release* con la etiqueta `vX.Y` y un archivo `Agentes-UPROGE-vX.Y.zip`. Al abrirse, el programa consulta el último Release y, si es más nuevo que el instalado, ofrece instalarlo. La instalación conserva `config.ini` y el Catálogo Maestro, y respalda la versión anterior.

## Requisitos

Windows con Python 3.10 o superior. La primera vez se ejecuta `1_INSTALAR.bat`.
