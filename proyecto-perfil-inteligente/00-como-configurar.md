# Cómo configurar el proyecto "Perfil Inteligente" en Claude

## Archivos de esta carpeta

| Archivo | Dónde va |
| --- | --- |
| `01-instrucciones-del-proyecto.md` | Se pega en **Instrucciones** del proyecto |
| `02-manual-perfil-inteligente.md` | Se sube a **Conocimiento** del proyecto |
| `03-plantilla-nuevo-cliente.md` | Se usa para abrir cada cliente en una conversación nueva |
| `../marca/julio-canchano-logo.png` | Se sube a **Conocimiento** (logo para los entregables) |

## Pasos (una sola vez)

1. En claude.ai, abre **Proyectos** y crea uno llamado **Perfil Inteligente**, o abre el que ya creaste.
2. En **Instrucciones del proyecto**, pega el contenido de `01-instrucciones-del-proyecto.md` (todo lo que está debajo de la línea).
3. En **Conocimiento del proyecto**, sube `02-manual-perfil-inteligente.md` y el logo `julio-canchano-logo.png`.
4. No subas nada más al conocimiento. Lo que va ahí lo ven todas las conversaciones del proyecto.

## Por cada cliente

1. Dentro del proyecto, abre una **conversación nueva**. Esa conversación es el contenedor del cliente.
2. Pega `03-plantilla-nuevo-cliente.md`, complétala y adjunta los documentos del cliente: Word, PDF, fotos o texto.
3. Cambia el nombre de la conversación a `PLI-2026-001 · Nombre Apellido` (Claude te lo propone).
4. Trabaja con los comandos: `FICHA`, `ESTADO`, `FASE N`, `OK FASE`, `PENDIENTES`, `LISTA CLIENTE`, `VACANTES`, `ENTREGA`, `CERRAR`.
5. Para retomar un cliente otro día, entra a su conversación y escribe `ESTADO`.

## Por qué así no se mezclan los clientes

- Cada conversación solo ve sus propios mensajes y archivos, más el manual del proyecto.
- Los documentos de cada cliente se adjuntan dentro de su conversación, nunca en el conocimiento del proyecto.
- Las instrucciones le prohíben a Claude usar datos de otro cliente y le piden detenerse si aparecen datos de otra persona.
- Si tienes activada la memoria entre conversaciones en tu cuenta, desactívala para este proyecto o revisa que no traiga datos de otros clientes. Es la única vía por la que algo podría cruzarse.

## Alternativa en Claude Code

En el repositorio ya existe el agente `perfil-laboral-inteligente` con una carpeta por cliente (`clientes/<nombre>/`). Hace lo mismo, pero guarda los entregables como archivos. Úsalo cuando quieras automatizar más o trabajar con archivos en lote.
