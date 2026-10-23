# Registro de cambios — reorganización del 23/10/2026

Todo lo que se hizo es reversible: no se borró ningún archivo, solo se movió o renombró.

## Renombrados (8 archivos)
Se unificó el nombre al estándar `_Respuesta`:

| Carpetas | Antes | Ahora |
|---|---|---|
| 2026-07-20 … 2026-07-23 | `DNI ddmmaaaa - Respuesta.xlsx`, `RUC ddmmaaaa - Respuesta.xlsx` | `DNI ddmmaaaa_Respuesta.xlsx`, `RUC ddmmaaaa_Respuesta.xlsx` |

## Movidos a `_DUPLICADOS_PARA_ELIMINAR/` (duplicados exactos, comprobados por hash MD5)
| Origen | Motivo |
|---|---|
| `DATA_OLD/DATOS SAT/` | Copia idéntica, archivo por archivo, de `DATOS SAT/` |
| `DATOS SAT/2026-08-01/2026-08-02/` | Carpeta anidada idéntica a `DATOS SAT/2026-08-02/` |
| `DATOS SAT/2026-08-30/nova.xlsx` | Idéntico a `RUC 29082026_Respuesta.xlsx` |
| `DATOS SAT/2026-08-30/RUC 29082026_Respuesta.xlsx` | Archivo de otro día; idéntico al que ya está en `2026-08-29/` |
| `__pycache__` (raíz y sunat_ruc) | Caché de Python regenerable |

## Archivados en `archivo/`
| Qué | Dónde |
|---|---|
| `sunat_ruc.py` (versión anterior), `test_*.txt`, `requirements.txt`, `resultados_sunat/` de mayo 2026 | `archivo/version_anterior_sunat_ruc/` |
| `A.zip` | `archivo/version_anterior_sunat_ruc/A_respaldo.zip` |
| `codigo.py` (en realidad un RAR) | `archivo/version_anterior_sunat_ruc/codigo_respaldo.rar` |
| `RUC.xlsx` y `resultados_sunat/` generados el 21/10 en `2026-07-20` | `archivo/salidas_corrida_previa_2026-07-20/` |

## Reubicados
| Archivo | Nuevo lugar |
|---|---|
| `V10.py`, `verificacion.py`, `diagnostico.py` | `app/` |
| `DOCUMENTACION_PROYECTO.html` | `docs/` |
| `Manual de usuario GATSO ESPAÑOL (1).pdf` | `docs/Manual_GATSO_ES.pdf` |
| `flujo_sunat_ruc.mermaid` | `docs/flujo_version_anterior.mermaid` |

## Nuevos
`README.md`, `app/requirements.txt`, este registro.

## Resultado
`DATOS SAT/`: 43 carpetas × 4 archivos = 172 archivos. Cada día tiene exactamente: `DNI`, `Placa`, `DNI_Respuesta`, `RUC_Respuesta`.

## Pendiente de borrar manualmente
- Carpeta `_DUPLICADOS_PARA_ELIMINAR/` (cuando confirmes que no la necesitas).
- Carpeta vacía `DATA_OLD/` (el sistema bloqueó el borrado automático).
- Quedó sin tocar `C:\DATA SAT\21_10_2026\`, la versión anterior completa.
