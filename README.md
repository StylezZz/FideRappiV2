# FideRAPPI v2.0

**Aplicación de Cargos y Abonos Automatizados**

![Python](https://img.shields.io/badge/Python-3.13-blue.svg)
![CustomTkinter](https://img.shields.io/badge/GUI-CustomTkinter-green.svg)
![Status](https://img.shields.io/badge/Status-Stable-success.svg)

---

## 📋 Descripción

FideRAPPI es una aplicación de escritorio diseñada para automatizar procesos de cargos y abonos. La aplicación permite procesar múltiples tipos de operaciones de forma eficiente y automatizada.

### ✨ Características principales

- 🔄 **Automatización completa** de procesos de cargos y abonos
- 💼 **Múltiples tipos de operaciones**: CCE, Ahorros, Cuentas Corrientes, LBTR, Cargo
- 📊 **Procesamiento de archivos Excel** con validación automática
- 🌐 **Integración con navegador web** mediante Selenium
- 📝 **Sistema de logs detallado** para trazabilidad
- 🎨 **Interfaz moderna** desarrollada con CustomTkinter
- ⚙️ **Configuración personalizable** por tipo de operación

---

## 🚀 Instalación

### Para usuarios finales (Ejecutable)

1. **Descargar** el archivo `FideRAPPI_v2.0.exe` desde la carpeta `releases/`
2. **Ejecutar** el archivo descargado
3. ¡**Listo para usar**! No requiere instalación adicional

### Para desarrolladores

1. **Clonar el repositorio**:
   ```bash
   git clone https://github.com/tu-usuario/fiderappi_v2.git
   cd fiderappi_v2
   ```

2. **Crear entorno virtual**:
   ```bash
   python -m venv venv
   venv\Scripts\activate  # Windows
   ```

3. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Ejecutar la aplicación**:
   ```bash
   python main.py
   ```

---

## 🖥️ Requisitos del Sistema

- **Sistema Operativo**: Windows 10 o superior
- **Memoria RAM**: 4 GB mínimo, 8 GB recomendado
- **Espacio en disco**: 500 MB libres
- **Microsoft Edge**: Requerido para automatización web
- **Permisos**: Administrador (en algunos casos)

---

## 📖 Uso de la Aplicación

### 1. Inicio de la aplicación
Al ejecutar FideRAPPI, se abrirá la ventana principal con las siguientes opciones:

- **CCE** - Cámara Compensadora Electrónica
- **AHORROS** - Cuentas de Ahorro
- **CTA_CTES** - Cuentas Corrientes
- **LBTR** - Liquidación Bruta en Tiempo Real
- **CARGO** - Operaciones de Cargo

### 2. Configuración por operación
1. **Seleccionar** el tipo de operación deseada
2. **Configurar** los parámetros específicos haciendo clic en ⚙️
3. **Cargar** el archivo Excel con los datos a procesar
4. **Iniciar** el procesamiento automático

### 3. Monitoreo del proceso
- La aplicación mostrará el progreso en tiempo real
- Los logs se guardan automáticamente en la carpeta `logs/`
- Se puede detener el proceso en cualquier momento

---

## 🔧 Configuración

### Archivos de configuración
La aplicación utiliza archivos JSON para configurar cada tipo de operación:

```
src/config/
├── cce_config.json
├── ahorros_config.json
├── cta_ctes_config.json
├── lbtr_config.json
└── cargo_config.json
```

### Personalización
Cada configuración incluye:
- **Selectores web** para automatización
- **Validaciones** específicas por operación
- **Rutas de archivos** por defecto
- **Parámetros de tiempo** de espera

---

## 📁 Estructura del Proyecto

```
fiderappi_v2/
├── src/                    # Código fuente
│   ├── interface/         # Interfaces de usuario
│   ├── core/             # Lógica de negocio
│   ├── utils/            # Utilidades
│   └── config/           # Configuraciones
├── assets/               # Recursos (iconos, imágenes)
├── logs/                 # Archivos de log
├── releases/             # Versiones publicadas
├── docs/                 # Documentación
├── main.py              # Punto de entrada
└── requirements.txt     # Dependencias
```

---

## 🐛 Solución de Problemas

### Problemas comunes

**❌ La aplicación no abre**
- Verificar que tienes permisos de administrador
- Ejecutar como administrador si es necesario
- Verificar que el antivirus no esté bloqueando el archivo

**❌ Error al procesar archivo Excel**
- Verificar formato del archivo (debe ser .xlsx)
- Revisar que las columnas tengan los nombres esperados
- Comprobar que no hay celdas vacías en datos críticos

**❌ Falla la automatización web**
- Verificar que Microsoft Edge esté instalado
- Comprobar conexión a internet
- Revisar logs para errores específicos

### Logs y diagnóstico
Los archivos de log se encuentran en la carpeta `logs/` con formato:
```
fiderapp_YYYYMMDD.log
```

---

## 🔄 Actualizaciones

### Historial de versiones
Ver [CHANGELOG.md](CHANGELOG.md) para el historial completo de cambios.

### Versión actual: 2.0.0
- ✅ Nueva interfaz con CustomTkinter
- ✅ Soporte para múltiples operaciones
- ✅ Mejoras en rendimiento
- ✅ Sistema de logs mejorado

---

## 🤝 Contribución

### Para desarrolladores
1. **Fork** el repositorio
2. **Crear** una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. **Commit** tus cambios (`git commit -am 'feat: agregar nueva funcionalidad'`)
4. **Push** a la rama (`git push origin feature/nueva-funcionalidad`)
5. **Crear** un Pull Request

### Reportar bugs
Usa la sección [Issues](https://github.com/tu-usuario/fiderappi_v2/issues) para reportar problemas.


## 📄 Licencia

Este proyecto es para usos propios de las personas que lo requieran.

---

## 🏗️ Construir desde el código fuente

### Generar ejecutable
```bash
# Usando cx_Freeze
python setup.py build

# Usando PyInstaller
pyinstaller --onefile --windowed --icon=assets/icons/logo-banco-nacion.ico main.py
```


---
Modificado por Christian Carrillo @2025
