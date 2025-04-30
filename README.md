# Optical Shop Management Software

Este proyecto es un software de gestión para ópticas, diseñado para manejar clientes y recetas de manera eficiente. El código ha sido refactorizado para mejorar la legibilidad, la organización y la compatibilidad con Visual Studio 2022.

## Características
- **Gestión de clientes**:
  - Agregar, limpiar y eliminar clientes.
  - Visualizar el historial de clientes en una tabla.
- **Gestión de recetas**:
  - Agregar nuevas recetas.
  - Visualizar el historial de recetas.

## Requisitos previos
Antes de comenzar, asegúrate de cumplir con los siguientes requisitos:
- **Sistema operativo**: Windows 10 o superior.
- **IDE**: Visual Studio 2022 o superior.
- **Base de datos**: Microsoft Access (`OpticalDb.mdb`).
- **Conexión a la base de datos**:
  - Configura la base de datos en la ruta `C:/OpticalShop/Database/OpticalDb.mdb`.

## Configuración del proyecto

### Paso 1: Clonar el repositorio
Clona este repositorio en tu máquina local:
```bash
git clone https://github.com/tu-usuario/tu-repositorio.git
cd tu-repositorio
```

### Paso 2: Abrir en Visual Studio
1. Abre Visual Studio 2022.
2. Haz clic en **Archivo > Abrir > Proyecto/Solución**.
3. Selecciona el archivo `.csproj` del proyecto.

### Paso 3: Configurar la base de datos
1. Asegúrate de que la base de datos `OpticalDb.mdb` esté en la ruta `C:/OpticalShop/Database/`.
2. Si la ruta es diferente, actualiza la cadena de conexión en el archivo `Form1.cs`:
   ```csharp
   private string databaseConnectionString = 
       "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=TU_RUTA/OpticalDb.mdb;Jet OLEDB:Engine Type=5";
   ```

### Paso 4: Restaurar dependencias
1. Si el proyecto utiliza paquetes NuGet, restaura las dependencias:
   ```bash
   dotnet restore
   ```

### Paso 5: Ejecutar el proyecto
1. Presiona `F5` o haz clic en el botón **Iniciar** en Visual Studio para ejecutar el proyecto.
2. La interfaz gráfica mostrará dos pestañas principales:
   - **Customer Details**: Para agregar y gestionar clientes.
   - **Prescriptions**: Para agregar y gestionar recetas.

## Probar las funcionalidades

### Gestión de clientes
1. Navega a la pestaña **Customer Details**.
2. Ingresa la información del cliente en los campos correspondientes:
   - Nombre.
   - Contacto.
   - Edad.
   - Dirección.
3. Haz clic en **Save Customer** para guardar los datos.
4. Haz clic en **Clear** para limpiar el formulario.

### Gestión de recetas
1. Navega a la pestaña **Prescriptions**.
2. Ingresa el tipo de lente y el precio en los campos correspondientes.
3. Haz clic en **Save Prescription** para guardar la receta.

### Validar historial
- Verifica los datos ingresados en las tablas de historial (clientes o recetas) dentro de las pestañas respectivas.

## Problemas comunes y soluciones

### Error de conexión a la base de datos
Si el proyecto no puede conectarse a la base de datos:
1. Verifica que el archivo `OpticalDb.mdb` exista en la ruta especificada.
2. Asegúrate de que el proveedor de conexión `Microsoft.Jet.OLEDB.4.0` esté instalado en tu máquina.

### Error de permisos en la base de datos
Si encuentras errores relacionados con permisos:
1. Asegúrate de que la carpeta `C:/OpticalShop/Database/` tenga permisos de lectura/escritura.
2. Ejecuta Visual Studio como administrador.

## Estructura del proyecto
```plaintext
OpticalShop/
├── Database/
│   └── OpticalDb.mdb
├── Forms/
│   └── Form1.cs
├── Program.cs
├── OpticalShop.csproj
└── README.md
```

## Contribuir
Si deseas contribuir a este proyecto:
1. Haz un fork del repositorio.
2. Crea una nueva rama:
   ```bash
   git checkout -b feature/nueva-funcionalidad
   ```
3. Realiza tus cambios y haz un commit:
   ```bash
   git commit -m "Agregada nueva funcionalidad"
   ```
4. Envía tus cambios:
   ```bash
   git push origin feature/nueva-funcionalidad
   ```
5. Abre un Pull Request.

## Licencia
Este proyecto está bajo la licencia MIT. Consulta el archivo `LICENSE` para más detalles.
