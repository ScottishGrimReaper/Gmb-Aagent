# PobreVermut — Propuesta de campañas 2026

Propuesta comercial para la gestión de campañas de PobreVermut, orientadas a
traccionar el ecommerce.

## Archivos

| Archivo | Qué es |
|---|---|
| `PobreVermut - Propuesta Campanas 2026.pptx` | El entregable. 5 slides, editable. |
| `generar_propuesta.js` | Script que genera el `.pptx`. Editar acá y regenerar. |
| `marca/` | Activos de marca extraídos del PDF `Barra_Pobrevermut`. El script los necesita. |

## Regenerar el deck

```bash
npm install pptxgenjs
node generar_propuesta.js "PobreVermut - Propuesta Campanas 2026.pptx"
```

## Los 5 slides

1. Portada
2. El plan — tres fases
3. Qué hago cada mes
4. Honorario — $350.000 líquidos al mes
5. Trabajemos juntos — cierre y contacto

## Definiciones comerciales

- **Honorario:** $350.000 líquidos mensuales. En el deck va solo el monto líquido,
  sin desglose tributario. Para emitir la boleta: bruto $412.979, retención de
  segunda categoría 15,25% ($62.979). El 15,25% es la tasa 2026 según Ley 21.133 y
  la retiene y declara el cliente.
- **Plazo:** un primer ciclo de tres meses. Al cierre se revisan los resultados y se
  define la continuidad. No hay variable sobre ventas comprometido en esta propuesta.
- **Alcance:** solo gestión de campañas — plan de medios, requerimiento de creativos,
  configuración, optimización y reportería.
- **Pauta:** no está incluida en el honorario. Va a nombre de PobreVermut, con su
  propio medio de pago. El rango del deck ($200.000 – $450.000 al mes) es referencial.

## Las tres fases

| Fase | Canal | Objetivo |
|---|---|---|
| 1 | Meta | Que nos conozcan — instalar la marca y dejar la medición andando |
| 2 | Meta | Que compren — venta directa al ecommerce |
| 3 | Google | Que nos encuentren — capturar a quien ya busca vermut |

El deck no menciona meses en las fases: la secuencia se entiende por el orden.

## Diseño

Sigue la identidad de PobreVermut, tomada del PDF `Barra_Pobrevermut`.

**Colores** — muestreados directamente del PDF:

| Rol | Hex |
|---|---|
| Fondo dominante (púrpura) | `3A1842` |
| Acento sobre púrpura (magenta) | `DE27A0` |
| Secundario sobre púrpura (lila) | `BFA3C8` |
| Fondo del bloque de valor (teal) | `284E58` |
| Acento sobre teal (verde) | `46BD7C` |

El deck usa bloques de color plano como la marca: cuatro slides en púrpura y el
del honorario en teal con verde, replicando cómo PobreVermut presenta el precio
en su propio material.

**Tipografía** — Arial Bold en caja alta para titulares, Courier New en caja alta
y espaciada para etiquetas y texto. La monoespaciada aproxima la de la marca;
Courier New se eligió por ser la única mono disponible tanto en PowerPoint como
en Google Slides. Si aparece el archivo de fuente real, se cambia en la constante
`MONO` del script.

**Activos** — el vaso de línea y la bajada manuscrita "Salud, vermút y buena vida"
se extrajeron del PDF de marca, se recortaron con canal alfa y se tiñeron al color
que corresponde en cada slide. Están en `marca/`.

## Antecedentes

- `Propuesta PobreVermut + Sherpas` (enero 2026, Drive) — propuesta Media Partner
  previa: fee $400.000 con descuento 30% → $280.000 + IVA.
- `Plan Ads – Pobrevermut` (febrero 2026, Drive) — fase base de testeo, $100.000–$200.000
  de pauta mensual, sin pago en esa etapa.
