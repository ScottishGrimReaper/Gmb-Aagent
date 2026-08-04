# PobreVermut — Propuesta de gestión de campañas 2026

Presentación comercial para la gestión de campañas de PobreVermut, orientadas a
traccionar el ecommerce.

## Archivos

| Archivo | Qué es |
|---|---|
| `PobreVermut - Propuesta Campanas 2026.pptx` | El entregable. 5 slides, editable. |
| `generar_propuesta.js` | Script que genera el `.pptx`. Editar acá y regenerar. |

## Los 5 slides

1. Portada
2. El plan — las tres fases
3. Qué hago cada mes — y qué no incluye
4. Honorario — $350.000 líquidos, pauta aparte, revisión al mes 3
5. Qué necesito para partir — y contacto

Hubo una versión previa de 11 slides (commit `494232c`), descartada por larga.

## Regenerar el deck

```bash
npm install pptxgenjs
node generar_propuesta.js "PobreVermut - Propuesta Campanas 2026.pptx"
```

## Definiciones comerciales que están en el deck

- **Honorario:** $350.000 líquidos mensuales vía boleta de honorarios.
  Bruto $412.979 · retención de segunda categoría 15,25% ($62.979) · líquido $350.000.
  El 15,25% es la tasa 2026 según Ley 21.133; la retiene y declara el cliente.
- **Alcance:** solo gestión de campañas — mediaplan, assetplan, configuración,
  optimización y reportería. No incluye producción de creativos ni community management.
- **Pauta:** no está incluida en el honorario. Va a nombre de PobreVermut, con su
  propio medio de pago. Los rangos del slide de inversión son referenciales.
- **Modelo de socio:** los primeros 3 meses van a honorario fijo. Al cierre del mes 3
  queda agendada la revisión para sumar un variable sobre ventas del ecommerce por
  sobre una meta base. No se cobra hoy.

## Estructura de las fases

| Fase | Período | Canal | Objetivo |
|---|---|---|---|
| 1 | Mes 1–2 | Meta (FB + IG) | Que nos conozcan · setup de medición y construcción de audiencias |
| 2 | Mes 3–4 | Meta (FB + IG) | Que compren · Advantage+ Shopping y remarketing dinámico |
| 3 | Mes 5+ | Google (Search + Shopping) | Que nos encuentren · capturar la demanda ya generada |

## Antecedentes

- `Propuesta PobreVermut + Sherpas` (enero 2026, Drive) — propuesta Media Partner
  previa: fee $400.000 con descuento 30% → $280.000 + IVA.
- `Plan Ads – Pobrevermut` (febrero 2026, Drive) — fase base de testeo, $100.000–$200.000
  de pauta mensual, sin pago en esa etapa.
