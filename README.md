# SUNAT RUC Pro

Consolida los documentos diarios (DNI y RUC) con sus datos: los **RUC** se consultan masivamente en SUNAT
(*e-consultaruc*, consulta por archivo, lotes de 100) y los **DNI** se completan con el archivo de respuesta del día.
Se opera desde una ventana Tkinter.

> Estado al 23/10/2026: la consulta a SUNAT funciona. El **modo automático por fechas aún falla** y hay una lista de
> correcciones pendientes (ver [Estado y pendientes](#estado-y-pendientes)).

## Estructura

```
23_10_2026/
├── README.md                      ← este archivo
├── LIBRERIAS.txt                  ← notas sueltas de instalación (pendiente de ordenar)
├── app/                           ← código
│   ├── V10.py                     ← aplicación principal (GUI: modo manual y automático)
│   ├── verificacion.py            ← comparador de dos Excel / dos listas
│   ├── diagnostico.py             ← inspecciona un archivo de entrada (solo lectura)
│   └── requirements.txt
├── DATOS SAT/                     ← datos de trabajo, una carpeta por día
│   └── AAAA-MM-DD/                ← 43 días: 2026-07-20 … 2026-08-31
│       ├── DNI ddmmaaaa.xlsx            lista maestra mixta (DNI 8 díg. + RUC 11 díg.)
│       ├── Placa ddmmaaaa.xlsx          placas, alineadas fila a fila con la lista maestra
│       ├── DNI ddmmaaaa_Respuesta.xlsx  datos personales de los DNI
│       └── RUC ddmmaaaa_Respuesta.xlsx  datos de los RUC
├── docs/
│   ├── DOCUMENTACION_PROYECTO.html  ← documentación completa (para pegar en OneNote)
│   ├── REGISTRO_DE_CAMBIOS.md       ← qué se movió, renombró o archivó el 23/10/2026
│   ├── flujo_version_anterior.mermaid
│   └── Manual_GATSO_ES.pdf
├── archivo/                       ← material histórico, no se usa
│   ├── version_anterior_sunat_ruc/  sunat_ruc.py (predecesor de V10), pruebas y respaldos
│   └── salidas_corrida_previa_2026-07-20/
└── _DUPLICADOS_PARA_ELIMINAR/     ← duplicados exactos verificados por hash; se pueden borrar
```

## Uso

```bash
python -m pip install -r app/requirements.txt
python app/V10.py                                   # abre la aplicación
python app/verificacion.py                          # comparador de Excel
python app/diagnostico.py "DATOS SAT/2026-07-20/DNI 20072026.xlsx"
```

En V10: pestaña **Modo Manual** para un archivo, pestaña **Automatización Masiva** para un rango de fechas sobre la
carpeta `DATOS SAT`.

## Los archivos de cada día

| Archivo | Contenido | Columnas |
|---|---|---|
| `DNI ddmmaaaa.xlsx` | **Entrada principal.** Todos los documentos del día, en orden, como texto | `Documento` |
| `Placa ddmmaaaa.xlsx` | Placas, misma cantidad de filas que la lista maestra | `Placa` |
| `DNI ddmmaaaa_Respuesta.xlsx` | Datos de **solo los DNI** | DNI, Apellido_Paterno, Apellido_Materno, Nombres, EstadoCivil, Sexo, Direccion, Ubigeo, Restriccion |
| `RUC ddmmaaaa_Respuesta.xlsx` | Datos de **solo los RUC** | RUC, Departamento, Provincia, Distrito, Ciudad, Estado, Fecha_Actual, Fecha_Alta, Nombre, Domicilio |

El archivo `DNI …_Respuesta` **no contiene RUC**; el archivo mixto es `DNI ddmmaaaa.xlsx`.

## Salida (consolidado)

`NumDocumento · TipoDocumento · Apellido_Paterno · Apellido_Materno · Nombres · Direccion · EstadoCivil · Ubigeo · Restriccion`

- **DNI:** los datos vienen del archivo `DNI …_Respuesta`.
- **RUC:** `Nombres` = razón social, `Ubigeo` = DEP/PROV/DIST, `Restriccion` = estado (ACTIVO, BAJA…) según SUNAT; los apellidos y el estado civil quedan en `N/A`.

## Cómo consulta a SUNAT (no modificar, funciona)

1. `GET` a `BASE_URL` para abrir sesión.
2. `POST` con `accion=consArchivo`, `tokenArch` aleatorio (52 caracteres) y un `.txt` con un `RUC|` por línea (máx. 100).
3. SUNAT devuelve un ZIP `RMaaaammdd….zip` con un `.txt` separado por `|` (latin-1). Si hay error responde HTML y se guarda `resultado_sunat_batch_N_ERROR.txt`.

## Estado y pendientes

Causa del error "No se encontraron filas con datos válidos": el selector toma el **primer** archivo de la carpeta. Del
20/07 al 23/07 eso era `DNI … - Respuesta.xlsx` (solo DNI), y el programa descarta todo lo que no tenga 11 dígitos.
Al unificar los nombres a `_Respuesta` el síntoma desaparece, pero el defecto de fondo sigue en el código:

- [ ] Elegir el archivo por su rol (`DNI ddmmaaaa.xlsx`), no por el orden en disco.
- [ ] Aceptar DNI (8 dígitos; 7 → `zfill(8)`) y RUC (11); no completar con ceros los de 10 dígitos.
- [ ] Armar el consolidado desde la lista maestra (hoy sale de `RUC.xlsx` y se pierden los DNI).
- [ ] Usar una subcarpeta limpia por corrida (hoy se leen ZIP de corridas anteriores).
- [ ] Pausa entre lotes, un reintento ante fallo de red y reporte de lotes con error.
- [ ] GUI: validar fechas y actualizar widgets solo desde el hilo principal (`root.after`).
- [ ] `verificacion.py`: normalizar valores y avisar de claves duplicadas; requiere `pandas`.
- [ ] Pruebas con pytest para `load_input_excel_data`.

Detalle completo, con severidades, en `docs/DOCUMENTACION_PROYECTO.html`.

## Datos a tener en cuenta

- Todos los días la lista maestra trae documentos repetidos (1 a 11).
- El 20/07 hay 34 DNI sin fila de respuesta; el 26/07, 28/07, 27/07, 07/08 y 27/08 faltan unos pocos más.
- Placa tiene una fila más que la lista de DNI el 28/07, 13/08, 15/08 y 19/08.
- Hasta el 23/07 el DNI de respuesta viene como número (pierde el cero inicial: 7 dígitos); desde el 24/07 viene como texto de 8 dígitos con espacios al final.
