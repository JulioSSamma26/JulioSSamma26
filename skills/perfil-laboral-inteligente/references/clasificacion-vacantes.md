# Fase 5 · Tablero de vacantes

El corazón del servicio: el cliente deja de revisar portales y solo mira las vacantes que valen la pena. Recibes alertas de los portales (texto de correos, enlaces o listados pegados) y devuelves cada vacante clasificada.

## Entrada

- Perfil del cliente: cargo objetivo, ciudad o modalidad, aspiración salarial, años de experiencia, palabras clave y requisitos excluyentes (por ejemplo, "no turnos nocturnos").
- Las vacantes: texto de los correos de alerta o descripciones pegadas.

Si falta la aspiración salarial, clasifica sin ese criterio y avísalo.

## Puntuación (0 a 100)

| Criterio | Puntos | Cómo calificar |
| --- | --- | --- |
| Cargo | 30 | 30 igual o equivalente; 15 relacionado; 0 distinto |
| Requisitos | 30 | Proporción de requisitos obligatorios que el cliente cumple |
| Salario | 20 | 20 dentro o sobre el rango; 10 si no lo publican; 0 por debajo |
| Ubicación y modalidad | 20 | 20 coincide; 10 aceptable; 0 no le sirve |

Un requisito excluyente que el cliente no cumple (título obligatorio, idioma, tarjeta profesional) deja la vacante en "Baja", sin importar el puntaje.

**Encaje:** Alta (75 a 100), Media (50 a 74), Baja (menos de 50).

## Duplicados y calidad

- Una misma vacante publicada en varios portales va una sola vez; anota los portales.
- Marca como **"Revisar"** las ofertas sospechosas: piden pagos, datos bancarios o documentos antes de entrevistar, salarios irreales para el cargo, o no identifican a la empresa ni la forma de contacto. Avísale al cliente: son señales comunes de estafa.

## Salida: filas para el tablero (Google Sheets)

Usa exactamente estas columnas, en este orden, para que se puedan pegar en la plantilla `assets/plantilla-tablero.csv`:

`Fecha | Cargo | Empresa | Ciudad/Modalidad | Salario | Portal | Enlace | Puntaje | Encaje | Por qué encaja | Qué le falta | Estado`

- **Por qué encaja / Qué le falta:** frases cortas y concretas ("Pide Power BI, que domina" / "Pide inglés B2; tiene B1").
- **Estado:** siempre "Nueva" al crear la fila; el cliente lo cambia a Postulada, Entrevista, Descartada u Oferta.
- Ordena por puntaje de mayor a menor.

## Informe semanal (plan Búsqueda Acompañada y mantenimiento)

Debajo de la tabla, un mensaje corto para WhatsApp:

```
Hola, [nombre]. Esta semana llegaron [N] vacantes y [N] encajan bien contigo. Las 3 mejores:
1. [Cargo] en [Empresa] · [puntaje] · [enlace]
2. ...
3. ...
Recomendación: [una acción concreta, ej.: ajustar la hoja de vida para la vacante 1 resaltando X].
```
