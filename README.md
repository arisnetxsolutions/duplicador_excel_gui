# Duplicador de Grupos en Excel

Esta aplicación permite duplicar grupos de datos en un archivo Excel, según una columna específica. Los grupos se identifican por la presencia de valores en una columna y las filas vacías que siguen a esos valores. La aplicación puede duplicar los grupos de acuerdo con el número de filas vacías encontradas a continuación, proporcionando un archivo Excel duplicado con los grupos repetidos.

## Características

- **Carga de archivos Excel**: Permite seleccionar un archivo Excel `.xlsx` desde tu computadora.
- **Selección de hoja**: Puedes elegir la hoja del archivo de Excel que deseas procesar.
- **Selección de columna**: La columna donde se identifican los grupos a duplicar.
- **Soporte de encabezado**: Puedes especificar si el archivo Excel tiene un encabezado que debe ser ignorado al procesar los datos.
- **Proceso en segundo plano**: Utiliza hilos para procesar los datos sin que la aplicación se congele durante la ejecución.
- **Interfaz gráfica**: La interfaz gráfica está construida con `Tkinter` para facilitar su uso.

## Requisitos

Para ejecutar esta aplicación, necesitas tener instalados los siguientes paquetes:

- Python 3.x
- `openpyxl` - Para manipular archivos Excel.
- `tkinter` - Para la interfaz gráfica.
- `threading` - Para ejecutar procesos en segundo plano.
- `subprocess` - Para abrir archivos una vez procesados.

Puedes instalar las dependencias con el siguiente comando:

```bash
pip install openpyxl
```
## Crear el ejecutable
- * python -m PyInstaller --onefile --windowed --icon=fav.ico duplicador_excel_gui.py
