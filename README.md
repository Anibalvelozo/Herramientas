## [v1.5.0] - 2026-02-09
### Agregado
- **Suite Office Completa:** Se añadieron opciones para reiniciar Microsoft Word (`winword.exe`) y Microsoft Excel (`excel.exe`).
- **Navegación:** Se expandió el menú principal a 11 opciones para acomodar las nuevas herramientas.
  
# Panel de Herramientas Pro - Historial de Versiones

Este documento detalla los cambios, mejoras y correcciones realizadas en el script de mantenimiento `ToolKit_Pro.bat`.

## [v1.4.0] - 2026-02-09
### Agregado
- **Gestión de Aplicaciones:** Se añadieron opciones específicas para reiniciar aplicaciones comunes de oficina:
  - **Microsoft Teams:** Detecta y cierra tanto la versión "New Teams" (trabajo) como la "Classic Teams".
  - **Google Chrome:** Fuerza el cierre de todos los procesos/pestañas y lo reabre.
  - **Adobe Acrobat:** Soporte para cerrar Acrobat Pro/DC y Adobe Reader.

### Cambios
- **Interfaz de Usuario:** Reorganización completa del Menú Principal.
  - **Grupo 1 (Ops. 1-4):** Herramientas de Aplicaciones (Prioridad alta para soporte rápido).
  - **Grupo 2 (Ops. 5-9):** Herramientas de Sistema y Diagnóstico.

## [v1.3.0] - 2026-02-09
### Agregado
- **Reparación de Audio:** Nueva opción (6) específica para HP EliteBook G10 / Drivers Realtek.
  - Detiene y reinicia los servicios `Audiosrv` y `AudioEndpointBuilder`.
  - Incluye validación estricta de permisos de Administrador.

### Mejoras
- **Validación de Admin:** Ahora el script impide ejecutar acciones críticas (como limpieza profunda o audio) si no se detectan permisos elevados, mostrando una advertencia en rojo.

---

## [v1.2.0] - 2026-02-08
### Interfaz (UI)
- **Códigos de Colores ANSI:** Se implementó un sistema visual para mejorar la lectura:
  - 🟢 **Verde:** Menús y operaciones exitosas (`[OK]`).
  - 🟡 **Amarillo:** Advertencias y procesos en curso.
  - 🔴 **Rojo:** Errores críticos y falta de permisos.

---

## [v1.1.0] - 2026-02-07
### Optimización de Código
- **Outlook:** Se cambió el método de `taskkill` a protocolo URI (`start outlook:`) para mayor compatibilidad con Office 365/2019.
- **Hardware:** Se reemplazaron comandos `wmic` lentos por llamadas a `PowerShell` para obtener datos de RAM y Batería con formato limpio.
- **Limpieza:** Se mejoró la lógica de borrado para incluir subcarpetas (`rd /s`) y no solo archivos sueltos.

---

## [v1.0.0] - 2026-02-01
### Lanzamiento Inicial
- Menú básico con opciones:
  1. Reiniciar Outlook.
  2. Estado de Batería.
  3. Info Hardware.
  4. Informe HTML de batería.
  5. Limpieza de Temporales básicos.
- Detección básica de Administrador.
