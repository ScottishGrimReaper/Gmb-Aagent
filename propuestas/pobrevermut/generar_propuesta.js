/**
 * PobreVermut — Propuesta de campañas 2026
 * Versión corta: 5 slides, una idea por slide.
 *
 * Uso: node generar_propuesta.js [ruta_salida.pptx]
 */

const pptxgen = require("pptxgenjs");

// ---------------------------------------------------------------- paleta
const INK = "2B1B14";
const WINE_DARK = "4A1220";
const WINE = "8C2F39";
const GOLD = "C8963E";
const CREAM = "F6F1EA";
const WHITE = "FFFFFF";
const MUTED = "7D6A5E";
const SOFT = "C4AEA1"; // secundario sobre oscuro

const SERIF = "Cambria";
const SANS = "Calibri";

const W = 13.333;
const M = 0.9; // margen lateral
const CW = W - M * 2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nicolas - Sherpas Studio";
pres.company = "Sherpas Studio";
pres.title = "PobreVermut - Propuesta de campanas 2026";

// ---------------------------------------------------------------- helpers

function head(slide, kicker, title, sub) {
  slide.addText(kicker.toUpperCase(), {
    x: M,
    y: 0.72,
    w: CW,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2.5,
    margin: 0,
  });
  slide.addText(title, {
    x: M,
    y: 1.08,
    w: CW,
    h: 0.72,
    fontFace: SERIF,
    fontSize: 40,
    bold: true,
    color: INK,
    margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: M,
      y: 1.86,
      w: CW,
      h: 0.36,
      fontFace: SANS,
      fontSize: 15,
      color: MUTED,
      margin: 0,
    });
  }
}

/** Círculo dorado — el único motivo gráfico del deck. */
function dot(slide, x, y, d, label, opts = {}) {
  slide.addShape(pres.ShapeType.ellipse, {
    x,
    y,
    w: d,
    h: d,
    fill: { color: opts.fill || GOLD },
  });
  if (label) {
    slide.addText(label, {
      x,
      y,
      w: d,
      h: d,
      align: "center",
      valign: "middle",
      fontFace: SERIF,
      fontSize: opts.fontSize || 20,
      bold: true,
      color: opts.color || WINE_DARK,
      margin: 0,
    });
  }
}

// ================================================================ 1 · PORTADA
{
  const s = pres.addSlide();
  s.background = { color: WINE_DARK };

  s.addShape(pres.ShapeType.ellipse, {
    x: 9.6,
    y: 2.3,
    w: 2.7,
    h: 2.7,
    fill: { color: WINE_DARK },
    line: { color: GOLD, width: 1.25 },
  });
  s.addText("PV", {
    x: 9.6,
    y: 2.3,
    w: 2.7,
    h: 2.7,
    align: "center",
    valign: "middle",
    fontFace: SERIF,
    fontSize: 56,
    bold: true,
    color: GOLD,
    margin: 0,
  });

  s.addText("PROPUESTA  ·  AGOSTO 2026", {
    x: M,
    y: 0.85,
    w: 7.5,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2.5,
    margin: 0,
  });

  s.addText("POBREVERMUT", {
    x: M,
    y: 2.5,
    w: 8.2,
    h: 0.42,
    fontFace: SANS,
    fontSize: 16,
    bold: true,
    color: CREAM,
    charSpacing: 5,
    margin: 0,
  });

  s.addText("Campañas para\nvender online", {
    x: M,
    y: 2.95,
    w: 8.2,
    h: 1.75,
    fontFace: SERIF,
    fontSize: 50,
    bold: true,
    color: WHITE,
    lineSpacingMultiple: 1.0,
    margin: 0,
  });

  s.addText("Ya hay marca, comunidad y ecommerce. Ahora, a vender.", {
    x: M,
    y: 4.85,
    w: 8.2,
    h: 0.4,
    fontFace: SANS,
    fontSize: 16,
    color: SOFT,
    margin: 0,
  });

  s.addText("Nicolás  ·  Sherpas Studio  ·  nicolas@sherpasstudio.cl", {
    x: M,
    y: 6.35,
    w: 8.5,
    h: 0.32,
    fontFace: SANS,
    fontSize: 13,
    color: SOFT,
    margin: 0,
  });
}

// ================================================================ 2 · EL PLAN
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "El plan", "Tres fases, un objetivo", "Cada fase construye lo que la siguiente necesita.");

  const fases = [
    {
      n: "1",
      meta: "MES 1 – 2  ·  META",
      t: "Que nos\nconozcan",
      d: "Instalamos la marca y dejamos la medición andando.",
    },
    {
      n: "2",
      meta: "MES 3 – 4  ·  META",
      t: "Que\ncompren",
      d: "Campañas de venta directa al ecommerce.",
    },
    {
      n: "3",
      meta: "MES 5 +  ·  GOOGLE",
      t: "Que nos\nencuentren",
      d: "Capturamos a quien ya está buscando vermut.",
    },
  ];

  const cw = 3.4;
  const gap = (CW - cw * 3) / 2;
  fases.forEach((f, i) => {
    const x = M + i * (cw + gap);
    dot(s, x, 2.75, 0.78, f.n, { fontSize: 26 });
    s.addText(f.meta, {
      x,
      y: 3.8,
      w: cw,
      h: 0.28,
      fontFace: SANS,
      fontSize: 11,
      bold: true,
      color: WINE,
      charSpacing: 1.5,
      margin: 0,
    });
    s.addText(f.t, {
      x,
      y: 4.12,
      w: cw,
      h: 1.1,
      fontFace: SERIF,
      fontSize: 26,
      bold: true,
      color: INK,
      lineSpacingMultiple: 1.0,
      valign: "top",
      margin: 0,
    });
    s.addText(f.d, {
      x,
      y: 5.3,
      w: cw,
      h: 0.7,
      fontFace: SANS,
      fontSize: 14,
      color: MUTED,
      lineSpacingMultiple: 1.15,
      valign: "top",
      margin: 0,
    });
  });
}

// ================================================================ 3 · QUÉ HAGO
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Mi trabajo", "Qué hago cada mes");

  const tareas = [
    "El plan de medios del mes",
    "El requerimiento de creativos al equipo creativo",
    "Configuración de campañas, píxel y catálogo",
    "Optimización durante todo el mes",
    "Reporte mensual con ventas, no con likes",
  ];

  tareas.forEach((t, i) => {
    const y = 2.55 + i * 0.78;
    dot(s, M, y + 0.06, 0.42, String(i + 1), { fontSize: 15 });
    s.addText(t, {
      x: M + 0.78,
      y,
      w: 6.4,
      h: 0.54,
      fontFace: SANS,
      fontSize: 18,
      color: INK,
      valign: "middle",
      margin: 0,
    });
  });

  // lo que queda fuera, a la derecha
  const nx = 8.15;
  const nw = W - nx - M;
  s.addText("NO INCLUYE", {
    x: nx,
    y: 2.58,
    w: nw,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: WINE,
    charSpacing: 2,
    margin: 0,
  });
  ["Producción de creativos", "Community management", "La inversión en pauta"].forEach(
    (t, i) => {
      const y = 3.15 + i * 0.62;
      dot(s, nx, y + 0.17, 0.16, null, { fill: "D8CCC3" });
      s.addText(t, {
        x: nx + 0.42,
        y,
        w: nw - 0.42,
        h: 0.5,
        fontFace: SANS,
        fontSize: 16,
        color: MUTED,
        valign: "middle",
        margin: 0,
      });
    }
  );
}

// ================================================================ 4 · CUÁNTO
{
  const s = pres.addSlide();
  s.background = { color: WINE_DARK };

  s.addText("HONORARIO", {
    x: M,
    y: 0.72,
    w: CW,
    h: 0.3,
    fontFace: SANS,
    fontSize: 11,
    bold: true,
    color: GOLD,
    charSpacing: 2.5,
    margin: 0,
  });

  s.addText("$350.000", {
    x: M,
    y: 1.5,
    w: 6.2,
    h: 1.5,
    fontFace: SERIF,
    fontSize: 76,
    bold: true,
    color: WHITE,
    margin: 0,
  });
  s.addText("líquidos al mes", {
    x: M,
    y: 3.02,
    w: 6.2,
    h: 0.42,
    fontFace: SANS,
    fontSize: 19,
    color: SOFT,
    margin: 0,
  });

  const filas = [
    ["Honorario bruto", "$412.979"],
    ["Retención 15,25%", "– $62.979"],
    ["Lo que recibo", "$350.000"],
  ];
  filas.forEach(([l, v], i) => {
    const y = 4.05 + i * 0.5;
    const last = i === filas.length - 1;
    s.addText(l, {
      x: M,
      y,
      w: 3.4,
      h: 0.42,
      fontFace: SANS,
      fontSize: 14,
      bold: last,
      color: last ? CREAM : SOFT,
      valign: "middle",
      margin: 0,
    });
    s.addText(v, {
      x: M + 3.4,
      y,
      w: 2.2,
      h: 0.42,
      align: "right",
      fontFace: SANS,
      fontSize: 14,
      bold: last,
      color: last ? GOLD : SOFT,
      valign: "middle",
      margin: 0,
    });
  });
  s.addText("Boleta de honorarios. La retención la declara PobreVermut.", {
    x: M,
    y: 5.66,
    w: 6.2,
    h: 0.34,
    fontFace: SANS,
    fontSize: 12,
    italic: true,
    color: "9C8377",
    margin: 0,
  });

  // dos notas a la derecha
  const nx = 7.55;
  const nw = W - nx - M;
  const notas = [
    [
      "LA PAUTA VA APARTE",
      "Entre $200.000 y $450.000 al mes según la fase, a nombre de PobreVermut.",
    ],
    [
      "LOS PRIMEROS 3 MESES, FIJO",
      "Al mes 3 revisamos el modelo y evaluamos sumar un variable sobre las ventas del ecommerce.",
    ],
  ];
  notas.forEach(([t, d], i) => {
    const y = 1.75 + i * 2.15;
    s.addText(t, {
      x: nx,
      y,
      w: nw,
      h: 0.3,
      fontFace: SANS,
      fontSize: 11,
      bold: true,
      color: GOLD,
      charSpacing: 1.5,
      margin: 0,
    });
    s.addText(d, {
      x: nx,
      y: y + 0.42,
      w: nw,
      h: 1.4,
      fontFace: SERIF,
      fontSize: 20,
      color: WHITE,
      lineSpacingMultiple: 1.15,
      valign: "top",
      margin: 0,
    });
  });
}

// ================================================================ 5 · ARRANCAMOS
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  head(s, "Para partir", "Qué necesito");

  const reqs = [
    "Accesos a Meta, Instagram y Facebook",
    "Acceso al ecommerce para píxel y catálogo",
    "Presupuesto de pauta confirmado",
    "Creativos del equipo creativo",
    "Un kickoff de 45 minutos",
  ];

  reqs.forEach((t, i) => {
    const y = 2.6 + i * 0.72;
    dot(s, M, y + 0.16, 0.2);
    s.addText(t, {
      x: M + 0.55,
      y,
      w: 6.1,
      h: 0.52,
      fontFace: SANS,
      fontSize: 17,
      color: INK,
      valign: "middle",
      margin: 0,
    });
  });

  // bloque de cierre
  const bx = 7.85;
  const bw = W - bx - M;
  s.addShape(pres.ShapeType.roundRect, {
    x: bx,
    y: 2.5,
    w: bw,
    h: 3.5,
    rectRadius: 0.08,
    fill: { color: WINE_DARK },
  });
  s.addText("Campañas activas en dos semanas.", {
    x: bx + 0.5,
    y: 2.95,
    w: bw - 1.0,
    h: 1.5,
    fontFace: SERIF,
    fontSize: 28,
    bold: true,
    color: WHITE,
    lineSpacingMultiple: 1.1,
    valign: "top",
    margin: 0,
  });
  s.addText("Nicolás  ·  Sherpas Studio", {
    x: bx + 0.5,
    y: 4.85,
    w: bw - 1.0,
    h: 0.32,
    fontFace: SANS,
    fontSize: 14,
    bold: true,
    color: GOLD,
    margin: 0,
  });
  s.addText("nicolas@sherpasstudio.cl\n+56 9 3125 6539", {
    x: bx + 0.5,
    y: 5.18,
    w: bw - 1.0,
    h: 0.62,
    fontFace: SANS,
    fontSize: 13,
    color: SOFT,
    lineSpacingMultiple: 1.2,
    margin: 0,
  });
}

// ---------------------------------------------------------------- salida
const out = process.argv[2] || "PobreVermut - Propuesta Campanas 2026.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK ->", out));
