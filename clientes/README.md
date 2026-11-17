# Clientes

Una carpeta por cliente, creada a partir de `_plantilla/`:

```
clientes/
└── nombre-apellido/
    ├── ficha.md          # datos del cliente y plan contratado
    ├── insumos/          # hoja de vida actual, PDF de LinkedIn, certificados, alertas de vacantes
    └── entregables/      # lo que produce el agente, fase por fase, y pendientes.md
```

Para atender a un cliente en Claude Code:

1. Copia `_plantilla/` con el nombre del cliente y llena `ficha.md`.
2. Pon sus archivos en `insumos/`: los PDF originales (hoja de vida, LinkedIn, certificados) y, si copias texto de otro documento, guárdalo como Markdown limpio, sin restos de otras herramientas.
3. Pide: "Usa el agente perfil-laboral-inteligente para atender a nombre-apellido".

**Privacidad:** esta carpeta contiene datos personales. No subas carpetas de clientes reales a un repositorio público. Mantén el repositorio privado o agrega `clientes/*/` a `.gitignore`, y borra los insumos al terminar el servicio si el cliente lo pide (Ley 1581 de 2012).
