# Supervisión de Planta (React + Vite)

Sistema integral para supervisión de operaciones, catálogo de OP, inventario de piezas, control de máquinas y seguimiento de producción en planta.

## Características

- **Inventario**: Registro de piezas por empaque (cajas y paquetes), cálculo automático de unidades, autocompletado con catálogo de OP, selección de modelos, exportación agrupada a Excel y sistema de respaldo/restauración.
- **Consulta de Catálogo**: Búsqueda rápida por OP, descripción, cliente o SAP, copia de especificaciones y modelos para WhatsApp/portapapeles.
- **Máquinas**: Gestión de líneas de producción agrupadas por área (*Vasos y Envases*, *Papel y Cartón*), clasificación PP/PT, bloqueo seguro de edición contra cambios accidentales y sincronización con Google Sheets.
- **Producción por Turno**: Control de turnos Día/Noche, registro ágil de producción por unidades o cajas, pliegos para máquinas PP, control de mermas y motivos de parada.
- **Reportes y Activación**: Formateo de reportes listos para WhatsApp, generación de hojas de cálculo con estilos de celda personalizados en ExcelJS y exportación PDF/impresión.
- **Historial e Integración con Supabase**: Almacenamiento centralizado de turnos, control de versiones, edición en línea y exportación de migración para SAP.

## Tecnologías

- React 18 + TypeScript
- Vite
- Tailwind CSS
- Lucide React
- ExcelJS & XLSX (SheetJS)
