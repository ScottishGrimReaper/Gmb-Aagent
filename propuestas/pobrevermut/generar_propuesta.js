/**
 * PobreVermut — Propuesta de campañas 2026
 * 5 slides. Diseño: fondo tinta, acento naranja aperitivo, tipografía sans.
 *
 * Uso: node generar_propuesta.js [ruta_salida.pptx]
 */

const pptxgen = require("pptxgenjs");

// ---------------------------------------------------------------- paleta
const INK = "0F0E0D"; // fondo dominante
const ORANGE = "EF5B25"; // acento — naranja aperitivo
const BONE = "F2EEE9"; // texto sobre tinta
const GREY = "8A807A"; // texto secundario sobre tinta
const DEEP = "140C07"; // texto sobre naranja

const DISPLAY = "Arial"; // titulares y numerales
const TEXT = "Calibri"; // texto corrido

const W = 13.333;
const M = 0.9;
const CW = W - M * 2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nicolas";
pres.title = "PobreVermut - Propuesta de campanas 2026";

// ---------------------------------------------------------------- helpers

/** Encabezado: índice + kicker naranja, y titular grande. */
function head(slide, index, kicker, title, sub, onOrange) {
  const accent = onOrange ? DEEP : ORANGE;
  const main = onOrange ? DEEP : BONE;
  const soft = onOrange ? DEEP : GREY;

  slide.addText(`${index}  —  ${kicker.toUpperCase()}`, {
    x: M,
    y: 0.82,
    w: CW,
    h: 0.3,
    fontFace: DISPLAY,
    fontSize: 11,
    bold: true,
    color: accent,
    charSpacing: 2.5,
    margin: 0,
  });
  slide.addText(title, {
    x: M,
    y: 1.24,
    w: CW,
    h: 0.8,
    fontFace: DISPLAY,
    fontSize: 38,
    bold: true,
    color: main,
    margin: 0,
  });
  if (sub) {
    slide.addText(sub, {
      x: M,
      y: 2.08,
      w: CW,
      h: 0.36,
      fontFace: TEXT,
      fontSize: 15,
      color: soft,
      margin: 0,
    });
  }
}

/** Punto sólido. */
function dot(slide, x, y, d, color) {
  slide.addShape(pres.ShapeType.ellipse, {
    x,
    y,
    w: d,
    h: d,
    fill: { color: color || ORANGE },
  });
}

// ================================================================ 1 · PORTADA
{
  const s = pres.addSlide();
  s.background = { color: INK };

  // el motivo: la rodaja de naranja
  dot(s, 9.85, 2.72, 2.6);

  s.addText("POBREVERMUT", {
    x: M,
    y: 0.95,
    w: 7.0,
    h: 0.34,
    fontFace: DISPLAY,
    fontSize: 12,
    bold: true,
    color: ORANGE,
    charSpacing: 6,
    margin: 0,
  });

  s.addText(
    [
      { text: "Campañas para", options: { color: BONE, breakLine: true } },
      { text: "vender online", options: { color: ORANGE } },
    ],
    {
      x: M,
      y: 3.0,
      w: 8.6,
      h: 2.1,
      fontFace: DISPLAY,
      fontSize: 54,
      bold: true,
      lineSpacingMultiple: 1.02,
      margin: 0,
    }
  );

  s.addText("Ya hay marca, comunidad y ecommerce. Ahora, a vender.", {
    x: M,
    y: 5.28,
    w: 8.6,
    h: 0.4,
    fontFace: TEXT,
    fontSize: 17,
    color: GREY,
    margin: 0,
  });

  s.addText("Propuesta  ·  Agosto 2026", {
    x: M,
    y: 6.28,
    w: 6.0,
    h: 0.3,
    fontFace: TEXT,
    fontSize: 13,
    color: GREY,
    margin: 0,
  });
  s.addText("Nicolás  ·  nicolas@sherpasstudio.cl", {
    x: M,
    y: 6.6,
    w: 6.0,
    h: 0.3,
    fontFace: TEXT,
    fontSize: 13,
    bold: true,
    color: BONE,
    margin: 0,
  });
}

// ================================================================ 2 · EL PLAN
{
  const s = pres.addSlide();
  s.background = { color: INK };
  head(
    s,
    "01",
    "El plan",
    "Tres fases, un objetivo",
    "Cada fase construye lo que la siguiente necesita."
  );

  const fases = [
    {
      n: "1",
      canal: "META",
      t: "Que nos\nconozcan",
      d: "Instalamos la marca y dejamos la medición andando.",
    },
    {
      n: "2",
      canal: "META",
      t: "Que\ncompren",
      d: "Campañas de venta directa al ecommerce.",
    },
    {
      n: "3",
      canal: "GOOGLE",
      t: "Que nos\nencuentren",
      d: "Capturamos a quien ya está buscando vermut.",
    },
  ];

  const cw = 3.4;
  const gap = (CW - cw * 3) / 2;
  fases.forEach((f, i) => {
    const x = M + i * (cw + gap);
    s.addText(f.n, {
      x,
      y: 2.85,
      w: cw,
      h: 1.15,
      fontFace: DISPLAY,
      fontSize: 72,
      bold: true,
      color: ORANGE,
      valign: "top",
      margin: 0,
    });
    s.addText(f.canal, {
      x,
      y: 4.12,
      w: cw,
      h: 0.28,
      fontFace: DISPLAY,
      fontSize: 10.5,
      bold: true,
      color: GREY,
      charSpacing: 2.5,
      margin: 0,
    });
    s.addText(f.t, {
      x,
      y: 4.46,
      w: cw,
      h: 1.15,
      fontFace: DISPLAY,
      fontSize: 25,
      bold: true,
      color: BONE,
      lineSpacingMultiple: 1.05,
      valign: "top",
      margin: 0,
    });
    s.addText(f.d, {
      x,
      y: 5.72,
      w: cw,
      h: 0.85,
      fontFace: TEXT,
      fontSize: 14,
      color: GREY,
      lineSpacingMultiple: 1.15,
      valign: "top",
      margin: 0,
    });
  });
}

// ================================================================ 3 · QUÉ HAGO
{
  const s = pres.addSlide();
  s.background = { color: INK };
  head(s, "02", "Mi trabajo", "Qué hago cada mes");

  const tareas = [
    "El plan de medios del mes",
    "El requerimiento de creativos al equipo creativo",
    "Configuración de campañas, píxel y catálogo",
    "Optimización durante todo el mes",
    "Reporte mensual con ventas, no con likes",
  ];

  tareas.forEach((t, i) => {
    const y = 2.82 + i * 0.8;
    s.addText(String(i + 1).padStart(2, "0"), {
      x: M,
      y,
      w: 0.72,
      h: 0.58,
      fontFace: DISPLAY,
      fontSize: 16,
      bold: true,
      color: ORANGE,
      valign: "middle",
      margin: 0,
    });
    s.addText(t, {
      x: M + 0.92,
      y,
      w: CW - 0.92,
      h: 0.58,
      fontFace: TEXT,
      fontSize: 21,
      color: BONE,
      valign: "middle",
      margin: 0,
    });
  });
}

// ================================================================ 4 · HONORARIO
{
  const s = pres.addSlide();
  s.background = { color: ORANGE };

  s.addText("03  —  HONORARIO", {
    x: M,
    y: 0.82,
    w: CW,
    h: 0.3,
    fontFace: DISPLAY,
    fontSize: 11,
    bold: true,
    color: DEEP,
    charSpacing: 2.5,
    margin: 0,
  });

  s.addText("$350.000", {
    x: M,
    y: 1.75,
    w: 9.5,
    h: 1.75,
    fontFace: DISPLAY,
    fontSize: 96,
    bold: true,
    color: DEEP,
    margin: 0,
  });
  s.addText("líquidos al mes", {
    x: M,
    y: 3.62,
    w: 8.0,
    h: 0.5,
    fontFace: TEXT,
    fontSize: 24,
    color: DEEP,
    margin: 0,
  });

  const notas = [
    [
      "LA PAUTA VA APARTE",
      "Entre $200.000 y $450.000 al mes según la fase, a nombre de PobreVermut.",
    ],
    [
      "UN PRIMER CICLO DE TRES MESES",
      "Partimos con tres meses de trabajo. Al cierre revisamos juntos los resultados y definimos cómo seguir.",
    ],
  ];
  const bw = 5.3;
  notas.forEach(([t, d], i) => {
    const x = M + i * (bw + 0.63);
    s.addText(t, {
      x,
      y: 5.05,
      w: bw,
      h: 0.3,
      fontFace: DISPLAY,
      fontSize: 10.5,
      bold: true,
      color: DEEP,
      charSpacing: 2,
      margin: 0,
    });
    s.addText(d, {
      x,
      y: 5.48,
      w: bw,
      h: 1.35,
      fontFace: DISPLAY,
      fontSize: 16,
      bold: true,
      color: DEEP,
      lineSpacingMultiple: 1.2,
      valign: "top",
      margin: 0,
    });
  });
}

// ================================================================ 5 · CIERRE
{
  const s = pres.addSlide();
  s.background = { color: INK };

  dot(s, 6.22, 1.85, 0.9);

  s.addText(
    [
      { text: "Trabajemos ", options: { color: BONE } },
      { text: "juntos.", options: { color: ORANGE } },
    ],
    {
      x: 1.5,
      y: 3.28,
      w: 10.33,
      h: 1.1,
      align: "center",
      fontFace: DISPLAY,
      fontSize: 52,
      bold: true,
      margin: 0,
    }
  );

  s.addText("Nicolás", {
    x: 1.5,
    y: 5.15,
    w: 10.33,
    h: 0.36,
    align: "center",
    fontFace: DISPLAY,
    fontSize: 16,
    bold: true,
    color: BONE,
    margin: 0,
  });
  s.addText("nicolas@sherpasstudio.cl  ·  +56 9 3125 6539", {
    x: 1.5,
    y: 5.54,
    w: 10.33,
    h: 0.36,
    align: "center",
    fontFace: TEXT,
    fontSize: 15,
    color: GREY,
    margin: 0,
  });
}

// ---------------------------------------------------------------- salida
const out = process.argv[2] || "PobreVermut - Propuesta Campanas 2026.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK ->", out));
