# Ex-to-Gsheet

Transfer data Excel to G-Spreadsheet

usamos dos funciones

update_data() - para actualizar los datos de excel

save_to_gsheets(df) - para cargar los datos a Gsheet

# Authentication
crear servicio "credenciales" para habilitar el programa a escribir en Google Sheets. 

flujo de activación
ir a Google Cloud Platform [GCP]('https://console.cloud.google.com/welcome?') y crear un proyecto nuevo o seleccionar uno existente. 

Si nunca ha usado GCP, los paso son:

En la barra de Buscar productos y recursos buscar Google Drive API y habilitarla.

En la barra de Buscar productos y recursos buscar Google Sheets API y habilitarla.

En la barra de Buscar productos y recursos buscar cuentas de servicio, en esa página:
+ Crear cuenta de servicio, completar el formulario; con el nombre de la cuenta de servicio es suficiente.
Una vez creada, seleccionarla para entrar en los detalles de la cuenta.

Claves > Agregar clave > Crear clave nueva > JSON > Crear. Aceptar la descarga de la cuenta de servicio. 

El resultado puede guardarse como: key.json.

Mover el archivo a la carpeta de trabajo. Debe estar en lugar seguro.
La cuenta de servicio servirá para todas las planillas de cálculo que necesitemos dentro de un mismo proyecto de GCP.

# Runtime settings
Python 3.10.7

# Install requirements
pip install -r requirements.txt

la funciones de cargas de datos estan anidadas en excel, ejecutadas una vez se abra el documento.

