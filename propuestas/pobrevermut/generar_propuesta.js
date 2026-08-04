/**
 * PobreVermut — Propuesta de Gestión de Campañas Digitales 2026
 * Genera el .pptx de la propuesta comercial.
 *
 * Uso: node generar_propuesta.js [ruta_salida.pptx]
 */

const pptxgen = require("pptxgenjs");

// ---------------------------------------------------------------- paleta
const INK = "2B1B14"; // texto principal
const WINE_DARK = "4A1220"; // fondos oscuros
const WINE = "8C2F39"; // rojo vermut
const GOLD = "C8963E"; // acento
const CREAM = "F6F1EA"; // texto sobre oscuro
const WHITE = "FFFFFF";
const MUTED = "7D6A5E"; // texto secundario
const CARD = "F7F3EF"; // relleno de tarjetas
const LINE = "E3D9CF"; // bordes suaves

const SERIF = "Cambria";
const SANS = "Calibri";

const W = 13.333;
const H = 7.5;
const M = 0.75; // margen lateral
const CW = W - M * 2; // ancho de contenido

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nicolas - Sherpas Studio";
pres.company = "Sherpas Studio";
pres.title = "PobreVermut - Propuesta de Gestion de Campanas 2026";

// ---------------------------------------------------------------- helpers

/** Encabezado de slide claro: kicker dorado + título serif. */
function head(slide, kicker, title, sub) {
  slide.addText(kicker.toUpperCase(), {
    x: M,
    y: 0.5,
    w: CW,
    h: 0.28,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2,
    margin: 0,
  });
  slide.addText(title, {
    x: M,
    y: 0.82,
    w: CW,
    h: 0.62,
    fontFace: SERIF,
    fontSize: 34,
    bold: true,
    color: INK,
    margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: M,
      y: 1.46,
      w: CW,
      h: 0.34,
      fontFace: SANS,
      fontSize: 14,
      color: MUTED,
      margin: 0,
    });
  }
}

/** Círculo dorado con número o texto — el motivo visual del deck. */
function badge(slide, x, y, d, label, opts = {}) {
  slide.addShape(pres.ShapeType.ellipse, {
    x,
    y,
    w: d,
    h: d,
    fill: { color: opts.fill || GOLD },
    line: { color: opts.line || GOLD, width: 1.25 },
  });
  slide.addText(label, {
    x,
    y,
    w: d,
    h: d,
    align: "center",
    valign: "middle",
    fontFace: SERIF,
    fontSize: opts.fontSize || 18,
    bold: true,
    color: opts.color || WINE_DARK,
    margin: 0,
  });
}

/** Tarjeta de fondo suave. Las de color sólido van sin borde. */
function card(slide, x, y, w, h, fill) {
  const o = {
    x,
    y,
    w,
    h,
    rectRadius: 0.08,
    fill: { color: fill || CARD },
  };
  if (!fill) o.line = { color: LINE, width: 0.75 };
  slide.addShape(pres.ShapeType.roundRect, o);
}

/** Lista con viñetas, espaciada. */
function bullets(slide, items, o) {
  slide.addText(
    items.map((t, i) => ({
      text: t,
      options: { bullet: { indent: 16 }, breakLine: i < items.length - 1 },
    })),
    {
      x: o.x,
      y: o.y,
      w: o.w,
      h: o.h,
      fontFace: SANS,
      fontSize: o.fontSize || 13.5,
      color: o.color || INK,
      lineSpacingMultiple: 1.05,
      paraSpaceAfter: o.gap === undefined ? 10 : o.gap,
      margin: 0,
      valign: "top",
    }
  );
}

// ================================================================ 1 · PORTADA
{
  const s = pres.addSlide();
  s.background = { color: WINE_DARK };

  // motivo: sello circular dorado
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.55,
    y: 2.02,
    w: 3.0,
    h: 3.0,
    fill: { color: WINE_DARK },
    line: { color: GOLD, width: 1.5 },
  });
  s.addShape(pres.ShapeType.ellipse, {
    x: 9.78,
    y: 2.25,
    w: 2.54,
    h: 2.54,
    fill: { color: WINE_DARK },
    line: { color: GOLD, width: 0.75 },
  });
  s.addText("PV", {
    x: 9.55,
    y: 2.02,
    w: 3.0,
    h: 3.0,
    align: "center",
    valign: "middle",
    fontFace: SERIF,
    fontSize: 62,
    bold: true,
    color: GOLD,
    margin: 0,
  });

  s.addText("DOCUMENTO CONFIDENCIAL  ·  AGOSTO 2026", {
    x: M,
    y: 0.72,
    w: 7.6,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2,
    margin: 0,
  });

  s.addText("POBREVERMUT", {
    x: M,
    y: 2.15,
    w: 8.4,
    h: 0.45,
    fontFace: SANS,
    fontSize: 17,
    bold: true,
    color: CREAM,
    charSpacing: 5,
    margin: 0,
  });

  s.addText("Gestión de\nCampañas Digitales", {
    x: M,
    y: 2.62,
    w: 8.4,
    h: 1.65,
    fontFace: SERIF,
    fontSize: 46,
    bold: true,
    color: WHITE,
    lineSpacingMultiple: 1.0,
    margin: 0,
  });

  s.addText(
    "De que nos conozcan a que nos compren. Un plan por fases para traccionar el ecommerce, partiendo por Meta y sumando Google.",
    {
      x: M,
      y: 4.45,
      w: 7.9,
      h: 0.75,
      fontFace: SANS,
      fontSize: 15,
      color: "D9C7BC",
      lineSpacingMultiple: 1.15,
      margin: 0,
    }
  );

  s.addText("Preparado por Nicolás  ·  Sherpas Studio", {
    x: M,
    y: 6.28,
    w: 7.0,
    h: 0.3,
    fontFace: SANS,
    fontSize: 13,
    bold: true,
    color: CREAM,
    margin: 0,
  });
  s.addText("nicolas@sherpasstudio.cl  ·  www.sherpasstudio.cl", {
    x: M,
    y: 6.6,
    w: 7.0,
    h: 0.3,
    fontFace: SANS,
    fontSize: 12,
    color: "B79A8B",
    margin: 0,
  });

  s.addNotes(
    "Propuesta de gestión de campañas para PobreVermut. Fee de gestión únicamente; la pauta va a nombre del cliente."
  );
}

// ================================================================ 2 · PUNTO DE PARTIDA
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(
    s,
    "Contexto",
    "Dónde estamos hoy",
    "El punto de partida cambió, y con él cambia lo que hay que hacer."
  );

  const items = [
    {
      n: "01",
      t: "Hubo una fase base",
      d: "En febrero corrimos una fase de testeo sin costo: se levantó comunidad, aprendimos qué contenido funciona y se crearon las primeras audiencias.",
    },
    {
      n: "02",
      t: "Ahora hay ecommerce",
      d: "El objetivo dejó de ser solo que nos vean. Hoy hay un carrito que llenar, y eso exige medición, catálogo y campañas pensadas para vender.",
    },
    {
      n: "03",
      t: "Falta el método",
      d: "Publicar cuando se puede no escala. Se necesita un plan por fases, con presupuesto asignado y metas claras por etapa.",
    },
  ];

  const cw = 3.65;
  const gap = (CW - cw * 3) / 2;
  items.forEach((it, i) => {
    const x = M + i * (cw + gap);
    card(s, x, 2.15, cw, 3.3);
    badge(s, x + 0.34, 2.48, 0.62, it.n, { fontSize: 15 });
    s.addText(it.t, {
      x: x + 0.34,
      y: 3.3,
      w: cw - 0.68,
      h: 0.6,
      fontFace: SERIF,
      fontSize: 18,
      bold: true,
      color: WINE,
      valign: "top",
      margin: 0,
    });
    s.addText(it.d, {
      x: x + 0.34,
      y: 3.96,
      w: cw - 0.68,
      h: 1.35,
      fontFace: SANS,
      fontSize: 13,
      color: INK,
      lineSpacingMultiple: 1.15,
      valign: "top",
      margin: 0,
    });
  });

  s.addText(
    "PobreVermut ya tiene producto, marca y canal de venta. Lo que falta es el motor.",
    {
      x: M,
      y: 5.78,
      w: CW,
      h: 0.45,
      fontFace: SERIF,
      fontSize: 16,
      italic: true,
      color: MUTED,
      margin: 0,
    }
  );
}

// ================================================================ 3 · LA RUTA
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(
    s,
    "Estrategia",
    "La ruta en tres fases",
    "Cada fase construye el activo que la siguiente necesita para funcionar."
  );

  const fases = [
    {
      n: "1",
      t: "Que nos conozcan",
      p: "Mes 1 – 2",
      c: "Meta · Facebook + Instagram",
      d: "Instalamos la marca y dejamos la medición andando. Lo que se construye acá son las audiencias que la Fase 2 va a cosechar.",
    },
    {
      n: "2",
      t: "Que compren",
      p: "Mes 3 – 4",
      c: "Meta · Facebook + Instagram",
      d: "Campañas de venta directa al ecommerce y remarketing sobre todo lo que se levantó en la fase anterior. Se optimiza por ROAS.",
    },
    {
      n: "3",
      t: "Que nos encuentren",
      p: "Mes 5 en adelante",
      c: "Google · Search + Shopping",
      d: "Meta genera la demanda; Google la captura cuando la persona ya está buscando. Suma al plan, no reemplaza a Meta.",
    },
  ];

  const cw = 3.65;
  const gap = (CW - cw * 3) / 2;
  fases.forEach((f, i) => {
    const x = M + i * (cw + gap);
    const dark = i === 1;
    card(s, x, 2.1, cw, 3.55, dark ? WINE_DARK : CARD);
    badge(s, x + 0.34, 2.4, 0.64, f.n, {
      fontSize: 22,
      fill: dark ? GOLD : WINE,
      line: dark ? GOLD : WINE,
      color: dark ? WINE_DARK : CREAM,
    });
    s.addText(`FASE ${f.n}  ·  ${f.p}`, {
      x: x + 0.34,
      y: 3.18,
      w: cw - 0.68,
      h: 0.26,
      fontFace: SANS,
      fontSize: 10.5,
      bold: true,
      color: dark ? GOLD : MUTED,
      charSpacing: 1.5,
      margin: 0,
    });
    s.addText(f.t, {
      x: x + 0.34,
      y: 3.44,
      w: cw - 0.68,
      h: 0.72,
      fontFace: SERIF,
      fontSize: 20,
      bold: true,
      color: dark ? WHITE : INK,
      valign: "top",
      margin: 0,
    });
    s.addText(f.c, {
      x: x + 0.34,
      y: 4.22,
      w: cw - 0.68,
      h: 0.26,
      fontFace: SANS,
      fontSize: 12,
      bold: true,
      color: dark ? "E0C9A6" : WINE,
      margin: 0,
    });
    s.addText(f.d, {
      x: x + 0.34,
      y: 4.54,
      w: cw - 0.68,
      h: 1.0,
      fontFace: SANS,
      fontSize: 12.5,
      color: dark ? "E4D5CB" : INK,
      lineSpacingMultiple: 1.15,
      valign: "top",
      margin: 0,
    });
  });

  card(s, M, 5.95, CW, 0.72, CARD);
  s.addText(
    [
      { text: "Hito del mes 3:  ", options: { bold: true, color: WINE } },
      {
        text: "con los primeros resultados de venta en mano, revisamos el modelo comercial y evaluamos pasar a un esquema de socio con variable sobre ventas.",
        options: { color: INK },
      },
    ],
    {
      x: M + 0.35,
      y: 5.95,
      w: CW - 0.7,
      h: 0.72,
      fontFace: SANS,
      fontSize: 13,
      valign: "middle",
      margin: 0,
    }
  );
}

// ---------------------------------------------------------------- slide de fase
function slideFase(cfg) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, `Fase ${cfg.n}  ·  ${cfg.periodo}`, cfg.titulo, cfg.canal);

  const leftW = 7.15;
  bullets(s, cfg.acciones, {
    x: M,
    y: 2.2,
    w: leftW,
    h: 3.9,
    fontSize: 13.5,
    gap: 13,
  });

  // panel de KPI a la derecha
  const px = M + leftW + 0.55;
  const pw = CW - leftW - 0.55;
  card(s, px, 2.2, pw, 3.4, WINE_DARK);
  s.addText("QUÉ MEDIMOS EN ESTA FASE", {
    x: px + 0.35,
    y: 2.5,
    w: pw - 0.7,
    h: 0.3,
    fontFace: SANS,
    fontSize: 10.5,
    bold: true,
    color: GOLD,
    charSpacing: 1.5,
    margin: 0,
  });
  cfg.kpis.forEach((k, i) => {
    s.addText(k, {
      x: px + 0.35,
      y: 2.98 + i * 0.6,
      w: pw - 0.7,
      h: 0.48,
      fontFace: SERIF,
      fontSize: 15,
      bold: true,
      color: CREAM,
      valign: "middle",
      margin: 0,
    });
  });

  s.addText(cfg.nota, {
    x: M,
    y: 6.05,
    w: CW,
    h: 0.5,
    fontFace: SANS,
    fontSize: 12.5,
    italic: true,
    color: MUTED,
    margin: 0,
  });
  return s;
}

// ================================================================ 4 · FASE 1
slideFase({
  n: "1",
  periodo: "Mes 1 – 2",
  titulo: "Que nos conozcan",
  canal: "Meta · Facebook + Instagram",
  acciones: [
    "Setup técnico primero: píxel y API de Conversiones instalados en el ecommerce, catálogo de productos cargado y eventos de compra midiendo desde el día uno. Sin esto, la Fase 2 arranca a ciegas.",
    "Campañas de Reconocimiento y Reproducciones de Video con segmentación amplia en la Región Metropolitana, dejando que el algoritmo encuentre los intereses latentes.",
    "Contenido que explica lo básico: qué es el vermut, qué es PobreVermut, en qué se diferencia y cómo se toma. Partimos del supuesto de que quien ve el anuncio no nos conoce.",
    "Construcción del activo real de esta fase: audiencias de remarketing con quienes vieron el video, visitaron el sitio o interactuaron en Instagram.",
  ],
  kpis: [
    "Alcance y CPM",
    "Costo por visita al sitio",
    "Tamaño de audiencias",
    "Seguidores nuevos",
  ],
  nota: "En esta fase no se le exige venta a la pauta. Se le exige volumen de gente y data limpia.",
});

// ================================================================ 5 · FASE 2
slideFase({
  n: "2",
  periodo: "Mes 3 – 4",
  titulo: "Que compren",
  canal: "Meta · Facebook + Instagram",
  acciones: [
    "Advantage+ Shopping Campaigns apuntando directo al ecommerce, con la data de conversión que la Fase 1 ya dejó acumulada.",
    "Remarketing dinámico con catálogo sobre los públicos construidos: carrito abandonado, visitantes del sitio y audiencias de video de los meses anteriores.",
    "Testeo de oferta y ticket: pack de dos botellas, despacho gratis sobre cierto monto, promoción de primera compra. La oferta se testea igual que el creativo.",
    "Optimización semanal sobre resultado de negocio, no sobre interacción. Se corta lo que no convierte y se escala lo que sí.",
  ],
  kpis: [
    "ROAS",
    "Costo por adquisición",
    "Tasa de conversión del sitio",
    "Ticket promedio",
  ],
  nota: "Acá es donde el trabajo de los dos primeros meses se transforma en ventas medibles.",
});

// ================================================================ 6 · FASE 3
slideFase({
  n: "3",
  periodo: "Mes 5 en adelante",
  titulo: "Que nos encuentren",
  canal: "Google · Search + Shopping",
  acciones: [
    "Search de marca para defender el término «pobrevermut»: que nadie más compre nuestro nombre y que quien nos busca llegue al ecommerce y no a un tercero.",
    "Google Shopping conectado al feed del ecommerce, mostrando botella, precio y disponibilidad directo en los resultados.",
    "Search genérico acotado sobre intención de compra real: «comprar vermut», «vermut artesanal», «vermut chileno».",
    "Performance Max solo cuando haya volumen de conversión suficiente para que aprenda. Antes de eso, quema presupuesto.",
  ],
  kpis: [
    "ROAS de Google",
    "CAC combinado",
    "Búsquedas de marca",
    "Participación por canal",
  ],
  nota: "Meta genera la demanda; Google la captura. Entra como canal adicional, con Meta corriendo en paralelo.",
});

// ================================================================ 7 · ALCANCE
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(
    s,
    "Alcance",
    "Qué incluye el fee mensual",
    "Gestión completa de las campañas, de principio a fin, todos los meses."
  );

  const items = [
    [
      "Mediaplan mensual",
      "La hoja de ruta del mes: estructura de campañas, distribución del presupuesto por objetivo y metas comprometidas.",
    ],
    [
      "Assetplan",
      "El requerimiento técnico de creativos para el partner creativo: cuántas piezas, en qué formatos, duraciones y mensajes.",
    ],
    [
      "Setup y activación",
      "Armado de campañas, públicos, píxel, API de Conversiones, catálogo y eventos de compra. Todo queda a nombre de PobreVermut.",
    ],
    [
      "Optimización",
      "Revisión continua durante el mes con ajustes de público, presupuesto y creativo según lo que muestren los números.",
    ],
    [
      "Reportería mensual",
      "Informe con métricas de negocio — inversión, ventas, ROAS y costo de adquisición — no con métricas de vanidad.",
    ],
    [
      "Reunión de revisión",
      "Una al mes para leer juntos los resultados y dejar definido el plan del mes siguiente.",
    ],
  ];

  const cw = 3.65;
  const ch = 1.75;
  const gx = (CW - cw * 3) / 2;
  items.forEach(([t, d], i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = M + col * (cw + gx);
    const y = 2.2 + row * (ch + 0.3);
    card(s, x, y, cw, ch);
    s.addText(t, {
      x: x + 0.3,
      y: y + 0.2,
      w: cw - 0.6,
      h: 0.34,
      fontFace: SERIF,
      fontSize: 16,
      bold: true,
      color: WINE,
      valign: "top",
      margin: 0,
    });
    s.addText(d, {
      x: x + 0.3,
      y: y + 0.58,
      w: cw - 0.6,
      h: 1.05,
      fontFace: SANS,
      fontSize: 12,
      color: INK,
      lineSpacingMultiple: 1.12,
      valign: "top",
      margin: 0,
    });
  });

  s.addText(
    [
      { text: "No incluye:  ", options: { bold: true, color: WINE } },
      {
        text: "producción de creativos, community management ni la inversión publicitaria.",
        options: { color: MUTED },
      },
    ],
    {
      x: M,
      y: 6.35,
      w: CW,
      h: 0.4,
      fontFace: SANS,
      fontSize: 12.5,
      margin: 0,
    }
  );
}

// ================================================================ 8 · INVERSIÓN
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(
    s,
    "Inversión",
    "Honorarios e inversión publicitaria",
    "Dos partidas separadas: mi trabajo y la plata que se pone en las plataformas."
  );

  // --- Honorario (bloque destacado, oscuro)
  const aw = 5.9;
  card(s, M, 2.2, aw, 3.3, WINE_DARK);
  s.addText("HONORARIO MENSUAL DE GESTIÓN", {
    x: M + 0.4,
    y: 2.5,
    w: aw - 0.8,
    h: 0.3,
    fontFace: SANS,
    fontSize: 10.5,
    bold: true,
    color: GOLD,
    charSpacing: 1.5,
    margin: 0,
  });
  s.addText("$350.000", {
    x: M + 0.4,
    y: 2.85,
    w: aw - 0.8,
    h: 0.85,
    fontFace: SERIF,
    fontSize: 52,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  s.addText("líquidos mensuales", {
    x: M + 0.4,
    y: 3.68,
    w: aw - 0.8,
    h: 0.3,
    fontFace: SANS,
    fontSize: 14,
    color: "D9C7BC",
    margin: 0,
  });

  const filas = [
    ["Honorario bruto (boleta de honorarios)", "$412.979"],
    ["Retención de segunda categoría · 15,25%", "– $62.979"],
    ["Pago directo al prestador", "$350.000"],
  ];
  filas.forEach(([l, v], i) => {
    const y = 4.15 + i * 0.4;
    const last = i === filas.length - 1;
    s.addText(l, {
      x: M + 0.4,
      y,
      w: aw - 2.2,
      h: 0.34,
      fontFace: SANS,
      fontSize: 12,
      bold: last,
      color: last ? CREAM : "C4AEA1",
      valign: "middle",
      margin: 0,
    });
    s.addText(v, {
      x: M + aw - 1.9,
      y,
      w: 1.5,
      h: 0.34,
      align: "right",
      fontFace: SANS,
      fontSize: 12,
      bold: last,
      color: last ? GOLD : "C4AEA1",
      valign: "middle",
      margin: 0,
    });
  });

  // --- Pauta
  const bx = M + aw + 0.5;
  const bw = CW - aw - 0.5;
  card(s, bx, 2.2, bw, 3.3);
  s.addText("INVERSIÓN PUBLICITARIA · REFERENCIAL", {
    x: bx + 0.4,
    y: 2.5,
    w: bw - 0.8,
    h: 0.3,
    fontFace: SANS,
    fontSize: 10.5,
    bold: true,
    color: WINE,
    charSpacing: 1.5,
    margin: 0,
  });

  const pauta = [
    ["Fase 1 · Mes 1 – 2", "Meta", "$200.000 – $250.000"],
    ["Fase 2 · Mes 3 – 4", "Meta", "$300.000 – $400.000"],
    ["Fase 3 · Mes 5 +", "Meta + Google", "desde $450.000"],
  ];
  pauta.forEach(([f, c, v], i) => {
    const y = 3.05 + i * 0.8;
    s.addText(f, {
      x: bx + 0.4,
      y,
      w: 2.4,
      h: 0.3,
      fontFace: SANS,
      fontSize: 12,
      bold: true,
      color: INK,
      margin: 0,
    });
    s.addText(c, {
      x: bx + 0.4,
      y: y + 0.29,
      w: 2.4,
      h: 0.28,
      fontFace: SANS,
      fontSize: 11,
      color: MUTED,
      margin: 0,
    });
    s.addText(v, {
      x: bx + bw - 2.7,
      y: y + 0.02,
      w: 2.3,
      h: 0.46,
      align: "right",
      fontFace: SERIF,
      fontSize: 16,
      bold: true,
      color: WINE,
      valign: "middle",
      margin: 0,
    });
  });

  s.addText(
    "La retención de segunda categoría (15,25% en 2026, Ley 21.133) la efectúa y declara PobreVermut al SII. La inversión publicitaria no está incluida en el honorario: va a nombre de PobreVermut, con su propio medio de pago y bajo su control. Los rangos de pauta son referenciales y se ajustan según resultados.",
    {
      x: M,
      y: 5.78,
      w: CW,
      h: 0.7,
      fontFace: SANS,
      fontSize: 12,
      color: MUTED,
      lineSpacingMultiple: 1.15,
      margin: 0,
    }
  );
}

// ================================================================ 9 · MODELO DE SOCIO
{
  const s = pres.addSlide();
  s.background = { color: WINE_DARK };

  s.addText("PROYECCIÓN", {
    x: M,
    y: 0.62,
    w: CW,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2,
    margin: 0,
  });
  s.addText("Más allá del fee: el modelo de socio", {
    x: M,
    y: 0.95,
    w: CW,
    h: 0.62,
    fontFace: SERIF,
    fontSize: 34,
    bold: true,
    color: WHITE,
    margin: 0,
  });

  s.addText(
    "Los primeros tres meses van a honorario fijo. Es el período en que se instala la marca, se levanta la data y se prueban las ofertas — no hay todavía una línea base de ventas contra la cual medir un variable con justicia.",
    {
      x: M,
      y: 1.78,
      w: 11.2,
      h: 0.75,
      fontFace: SANS,
      fontSize: 14.5,
      color: "D9C7BC",
      lineSpacingMultiple: 1.2,
      margin: 0,
    }
  );

  card(s, M, 2.72, CW, 2.6, "5E1D2A");
  badge(s, M + 0.45, 3.06, 0.66, "M3", { fontSize: 15 });
  s.addText("Al cierre del mes 3 · punto de revisión", {
    x: M + 1.32,
    y: 3.06,
    w: 9.5,
    h: 0.66,
    fontFace: SERIF,
    fontSize: 22,
    bold: true,
    color: WHITE,
    valign: "middle",
    margin: 0,
  });

  const puntos = [
    "Con los resultados de las primeras campañas de venta en mano, revisamos el modelo comercial.",
    "La propuesta es mantener el honorario fijo y sumar un variable sobre las ventas del ecommerce que superen una meta base acordada entre los dos.",
    "En esa conversación se define lo concreto: meta base mensual, porcentaje del variable y vigencia del acuerdo.",
  ];
  bullets(s, puntos, {
    x: M + 0.5,
    y: 3.85,
    w: CW - 1.0,
    h: 1.4,
    fontSize: 13.5,
    color: "EADCD3",
    gap: 9,
  });

  s.addText(
    "Por qué importa: alinea mi ingreso con el crecimiento real de PobreVermut. Si el ecommerce vende más, ganamos los dos. Este punto no se cobra hoy — queda agendado.",
    {
      x: M,
      y: 5.62,
      w: 11.5,
      h: 0.6,
      fontFace: SERIF,
      fontSize: 15,
      italic: true,
      color: GOLD,
      lineSpacingMultiple: 1.15,
      margin: 0,
    }
  );
}

// ================================================================ 10 · ARRANQUE
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(
    s,
    "Siguiente paso",
    "Para arrancar necesito",
    "Con esto en mano, las primeras campañas quedan activas en menos de dos semanas."
  );

  const reqs = [
    [
      "Accesos de administrador",
      "Business Manager, cuenta publicitaria, Instagram y Facebook a nombre de PobreVermut.",
    ],
    [
      "Acceso al ecommerce",
      "Para instalar píxel, API de Conversiones y feed de catálogo — o el contacto de quien lo administra.",
    ],
    [
      "Presupuesto de pauta",
      "Monto mensual confirmado y medio de pago cargado a nombre de PobreVermut.",
    ],
    [
      "Creativos según Assetplan",
      "Entregados por el partner creativo con los formatos y duraciones del requerimiento.",
    ],
    [
      "Reunión de kickoff",
      "45 minutos en los primeros dos días hábiles para alinear objetivos y calendario.",
    ],
  ];

  reqs.forEach(([t, d], i) => {
    const y = 2.22 + i * 0.83;
    badge(s, M, y, 0.54, String(i + 1), { fontSize: 15 });
    s.addText(t, {
      x: M + 0.82,
      y: y - 0.02,
      w: 3.9,
      h: 0.35,
      fontFace: SERIF,
      fontSize: 16.5,
      bold: true,
      color: INK,
      valign: "middle",
      margin: 0,
    });
    s.addText(d, {
      x: M + 4.85,
      y: y - 0.02,
      w: CW - 4.85,
      h: 0.58,
      fontFace: SANS,
      fontSize: 13,
      color: MUTED,
      lineSpacingMultiple: 1.1,
      valign: "top",
      margin: 0,
    });
  });

  card(s, M, 6.42, CW, 0.6, CARD);
  s.addText(
    "Firma → accesos y setup técnico (días 1–5) → mediaplan y assetplan (días 5–10) → campañas activas (semana 2).",
    {
      x: M + 0.35,
      y: 6.42,
      w: CW - 0.7,
      h: 0.6,
      fontFace: SANS,
      fontSize: 12.5,
      bold: true,
      color: WINE,
      valign: "middle",
      margin: 0,
    }
  );
}

// ================================================================ 11 · CIERRE
{
  const s = pres.addSlide();
  s.background = { color: WINE_DARK };

  s.addShape(pres.ShapeType.ellipse, {
    x: 6.17,
    y: 1.15,
    w: 1.0,
    h: 1.0,
    fill: { color: WINE_DARK },
    line: { color: GOLD, width: 1.25 },
  });
  s.addText("PV", {
    x: 6.17,
    y: 1.15,
    w: 1.0,
    h: 1.0,
    align: "center",
    valign: "middle",
    fontFace: SERIF,
    fontSize: 22,
    bold: true,
    color: GOLD,
    margin: 0,
  });

  s.addText(
    "PobreVermut ya tiene producto,\nmarca y dónde vender.",
    {
      x: 1.6,
      y: 2.55,
      w: 10.13,
      h: 1.4,
      align: "center",
      fontFace: SERIF,
      fontSize: 36,
      bold: true,
      color: WHITE,
      lineSpacingMultiple: 1.05,
      margin: 0,
    }
  );
  s.addText("Lo que falta es el motor que lo mueva.", {
    x: 1.6,
    y: 4.02,
    w: 10.13,
    h: 0.5,
    align: "center",
    fontFace: SERIF,
    fontSize: 24,
    italic: true,
    color: GOLD,
    margin: 0,
  });

  s.addText("Nicolás  ·  Sherpas Studio", {
    x: 1.6,
    y: 5.55,
    w: 10.13,
    h: 0.35,
    align: "center",
    fontFace: SANS,
    fontSize: 15,
    bold: true,
    color: CREAM,
    margin: 0,
  });
  s.addText("nicolas@sherpasstudio.cl  ·  +56 9 3125 6539  ·  www.sherpasstudio.cl", {
    x: 1.6,
    y: 5.94,
    w: 10.13,
    h: 0.35,
    align: "center",
    fontFace: SANS,
    fontSize: 13,
    color: "B79A8B",
    margin: 0,
  });
}

// ---------------------------------------------------------------- salida
const out =
  process.argv[2] || "PobreVermut - Propuesta Gestion de Campanas 2026.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK ->", out));
