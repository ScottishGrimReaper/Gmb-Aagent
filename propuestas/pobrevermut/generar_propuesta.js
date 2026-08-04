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
  head(s, "01", "El plan", "Tres fases, un objetivo", "Cada fase construye lo que la siguiente necesita.");

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
    s.addText(f.meta, {
      x,
      y: 4.12,
      w: cw,
      h: 0.28,
      fontFace: DISPLAY,
      fontSize: 10.5,
      bold: true,
      color: GREY,
      charSpacing: 1.5,
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
    const y = 2.75 + i * 0.76;
    s.addText(String(i + 1).padStart(2, "0"), {
      x: M,
      y,
      w: 0.62,
      h: 0.52,
      fontFace: DISPLAY,
      fontSize: 14,
      bold: true,
      color: ORANGE,
      valign: "middle",
      margin: 0,
    });
    s.addText(t, {
      x: M + 0.7,
      y,
      w: 6.3,
      h: 0.52,
      fontFace: TEXT,
      fontSize: 18,
      color: BONE,
      valign: "middle",
      margin: 0,
    });
  });

  // lo que queda fuera
  const nx = 8.3;
  const nw = W - nx - M;
  s.addText("NO INCLUYE", {
    x: nx,
    y: 2.78,
    w: nw,
    h: 0.3,
    fontFace: DISPLAY,
    fontSize: 10.5,
    bold: true,
    color: ORANGE,
    charSpacing: 2.5,
    margin: 0,
  });
  ["Producción de creativos", "Community management", "La inversión en pauta"].forEach(
    (t, i) => {
      const y = 3.35 + i * 0.62;
      dot(s, nx, y + 0.19, 0.14, GREY);
      s.addText(t, {
        x: nx + 0.4,
        y,
        w: nw - 0.4,
        h: 0.5,
        fontFace: TEXT,
        fontSize: 16,
        color: GREY,
        valign: "middle",
        margin: 0,
      });
    }
  );
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
    y: 1.8,
    w: 6.4,
    h: 1.55,
    fontFace: DISPLAY,
    fontSize: 78,
    bold: true,
    color: DEEP,
    margin: 0,
  });
  s.addText("líquidos al mes", {
    x: M,
    y: 3.38,
    w: 6.4,
    h: 0.42,
    fontFace: TEXT,
    fontSize: 20,
    color: DEEP,
    margin: 0,
  });

  const filas = [
    ["Honorario bruto", "$412.979"],
    ["Retención 15,25%", "– $62.979"],
    ["Lo que recibo", "$350.000"],
  ];
  filas.forEach(([l, v], i) => {
    const y = 4.45 + i * 0.52;
    const last = i === filas.length - 1;
    s.addText(l, {
      x: M,
      y,
      w: 3.4,
      h: 0.42,
      fontFace: TEXT,
      fontSize: 14,
      bold: last,
      color: DEEP,
      valign: "middle",
      margin: 0,
    });
    s.addText(v, {
      x: M + 3.4,
      y,
      w: 2.2,
      h: 0.42,
      align: "right",
      fontFace: TEXT,
      fontSize: 14,
      bold: last,
      color: DEEP,
      valign: "middle",
      margin: 0,
    });
  });
  s.addText("Boleta de honorarios. La retención la declara PobreVermut.", {
    x: M,
    y: 6.12,
    w: 6.4,
    h: 0.34,
    fontFace: TEXT,
    fontSize: 12.5,
    italic: true,
    color: DEEP,
    margin: 0,
  });

  const nx = 7.75;
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
    const y = 1.95 + i * 2.6;
    s.addText(t, {
      x: nx,
      y,
      w: nw,
      h: 0.3,
      fontFace: DISPLAY,
      fontSize: 10.5,
      bold: true,
      color: DEEP,
      charSpacing: 2,
      margin: 0,
    });
    s.addText(d, {
      x: nx,
      y: y + 0.44,
      w: nw,
      h: 1.5,
      fontFace: DISPLAY,
      fontSize: 19,
      bold: true,
      color: DEEP,
      lineSpacingMultiple: 1.18,
      valign: "top",
      margin: 0,
    });
  });
}

// ================================================================ 5 · PARA PARTIR
{
  const s = pres.addSlide();
  s.background = { color: INK };
  head(s, "04", "Para partir", "Qué necesito");

  const reqs = [
    "Accesos a Meta, Instagram y Facebook",
    "Acceso al ecommerce para píxel y catálogo",
    "Presupuesto de pauta confirmado",
    "Creativos del equipo creativo",
    "Un kickoff de 45 minutos",
  ];

  reqs.forEach((t, i) => {
    const y = 2.75 + i * 0.76;
    dot(s, M, y + 0.19, 0.16);
    s.addText(t, {
      x: M + 0.48,
      y,
      w: 6.5,
      h: 0.52,
      fontFace: TEXT,
      fontSize: 17,
      color: BONE,
      valign: "middle",
      margin: 0,
    });
  });

  const nx = 8.3;
  const nw = W - nx - M;
  s.addText("Campañas activas\nen dos semanas.", {
    x: nx,
    y: 2.7,
    w: nw,
    h: 1.7,
    fontFace: DISPLAY,
    fontSize: 30,
    bold: true,
    color: ORANGE,
    lineSpacingMultiple: 1.08,
    valign: "top",
    margin: 0,
  });
  s.addText("Nicolás", {
    x: nx,
    y: 4.55,
    w: nw,
    h: 0.34,
    fontFace: DISPLAY,
    fontSize: 15,
    bold: true,
    color: BONE,
    margin: 0,
  });
  s.addText("nicolas@sherpasstudio.cl\n+56 9 3125 6539", {
    x: nx,
    y: 4.92,
    w: nw,
    h: 0.68,
    fontFace: TEXT,
    fontSize: 14,
    color: GREY,
    lineSpacingMultiple: 1.25,
    margin: 0,
  });
}

// ---------------------------------------------------------------- salida
const out = process.argv[2] || "PobreVermut - Propuesta Campanas 2026.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK ->", out));
