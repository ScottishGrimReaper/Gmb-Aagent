/**
 * PobreVermut — Propuesta de campañas 2026
 * 5 slides. Diseño según la identidad de PobreVermut:
 * bloques de color plano, titulares pesados en caja alta, texto en monoespaciada
 * y los activos reales de la marca (vaso de línea y bajada manuscrita).
 *
 * Uso: node generar_propuesta.js [ruta_salida.pptx]
 */

const path = require("path");
const pptxgen = require("pptxgenjs");

// ---------------------------------------------------------------- paleta de marca
const PURPLE = "3A1842"; // fondo dominante
const MAGENTA = "DE27A0"; // acento sobre púrpura
const LILAC = "BFA3C8"; // secundario sobre púrpura
const TEAL = "284E58"; // fondo del bloque de valor
const GREEN = "46BD7C"; // acento sobre teal
const WHITE = "FFFFFF";

const DISPLAY = "Arial"; // titulares — caja alta, pesados
const MONO = "Courier New"; // etiquetas y texto — caja alta, espaciada

const ART = path.join(__dirname, "marca");
const TAGLINE = path.join(ART, "tagline-blanco.png");

const W = 13.333;
const M = 0.9;
const CW = W - M * 2;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
pres.author = "Nicolas";
pres.title = "PobreVermut - Propuesta de campanas 2026";

// ---------------------------------------------------------------- helpers

/** Etiqueta monoespaciada en caja alta. */
function label(slide, text, o) {
  slide.addText(text.toUpperCase(), {
    x: o.x,
    y: o.y,
    w: o.w,
    h: o.h || 0.3,
    align: o.align,
    fontFace: MONO,
    fontSize: o.fontSize || 12,
    bold: true,
    color: o.color || MAGENTA,
    charSpacing: o.charSpacing === undefined ? 2 : o.charSpacing,
    valign: o.valign,
    lineSpacingMultiple: o.lineSpacingMultiple,
    margin: 0,
  });
}

/** Encabezado: etiqueta de sección + titular pesado en caja alta. */
function head(slide, index, kicker, title, sub, c) {
  c = c || {};
  label(slide, `${index}  —  ${kicker}`, {
    x: M,
    y: 0.82,
    w: CW,
    fontSize: 12,
    color: c.accent || MAGENTA,
  });
  slide.addText(title.toUpperCase(), {
    x: M,
    y: 1.26,
    w: CW,
    h: 0.8,
    fontFace: DISPLAY,
    fontSize: 36,
    bold: true,
    color: c.main || WHITE,
    margin: 0,
  });
  if (sub) {
    label(slide, sub, {
      x: M,
      y: 2.16,
      w: CW,
      h: 0.32,
      fontSize: 11,
      color: c.soft || LILAC,
      charSpacing: 1,
    });
  }
}

// ================================================================ 1 · PORTADA
{
  const s = pres.addSlide();
  s.background = { color: PURPLE };

  // el logotipo va fino y muy espaciado, como en la marca
  s.addText("POBREVERMUT", {
    x: M,
    y: 0.88,
    w: 7.5,
    h: 0.38,
    fontFace: DISPLAY,
    fontSize: 15,
    color: WHITE,
    charSpacing: 9,
    margin: 0,
  });

  s.addText(
    [
      { text: "CAMPAÑAS PARA", options: { color: WHITE, breakLine: true } },
      { text: "VENDER ONLINE", options: { color: MAGENTA } },
    ],
    {
      x: M,
      y: 2.65,
      w: 9.2,
      h: 2.2,
      fontFace: DISPLAY,
      fontSize: 54,
      bold: true,
      lineSpacingMultiple: 1.0,
      margin: 0,
    }
  );

  label(s, "Ya hay marca, comunidad y ecommerce. Ahora, a vender.", {
    x: M,
    y: 5.1,
    w: 9.2,
    h: 0.34,
    fontSize: 12,
    color: LILAC,
    charSpacing: 1,
  });

  label(s, "Propuesta  ·  Agosto 2026", {
    x: M,
    y: 6.24,
    w: 6.5,
    h: 0.3,
    fontSize: 11,
    color: LILAC,
    charSpacing: 1,
  });
  label(s, "Nicolás  ·  nicolas@sherpasstudio.cl", {
    x: M,
    y: 6.6,
    w: 6.5,
    h: 0.3,
    fontSize: 11,
    color: WHITE,
    charSpacing: 1,
  });
}

// ================================================================ 2 · EL PLAN
{
  const s = pres.addSlide();
  s.background = { color: PURPLE };
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
      canal: "Meta",
      t: "Que nos\nconozcan",
      d: "Instalamos la marca y dejamos la medición andando.",
    },
    {
      n: "2",
      canal: "Meta",
      t: "Que\ncompren",
      d: "Campañas de venta directa al ecommerce.",
    },
    {
      n: "3",
      canal: "Google",
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
      y: 2.92,
      w: cw,
      h: 1.15,
      fontFace: DISPLAY,
      fontSize: 72,
      bold: true,
      color: MAGENTA,
      valign: "top",
      margin: 0,
    });
    label(s, f.canal, {
      x,
      y: 4.18,
      w: cw,
      h: 0.28,
      fontSize: 11,
      color: MAGENTA,
      charSpacing: 3,
    });
    s.addText(f.t.toUpperCase(), {
      x,
      y: 4.54,
      w: cw,
      h: 1.1,
      fontFace: DISPLAY,
      fontSize: 23,
      bold: true,
      color: WHITE,
      lineSpacingMultiple: 1.05,
      valign: "top",
      margin: 0,
    });
    label(s, f.d, {
      x,
      y: 5.78,
      w: cw,
      h: 0.9,
      fontSize: 11,
      color: LILAC,
      charSpacing: 0.5,
      lineSpacingMultiple: 1.3,
      valign: "top",
    });
  });
}

// ================================================================ 3 · QUÉ HAGO
{
  const s = pres.addSlide();
  s.background = { color: PURPLE };
  head(s, "02", "Mi trabajo", "Qué hago cada mes");

  const tareas = [
    "El plan de medios del mes",
    "El requerimiento de creativos al equipo creativo",
    "Configuración de campañas, píxel y catálogo",
    "Optimización durante todo el mes",
    "Reporte mensual con ventas, no con likes",
  ];

  tareas.forEach((t, i) => {
    const y = 2.92 + i * 0.78;
    label(s, String(i + 1).padStart(2, "0"), {
      x: M,
      y,
      w: 0.75,
      h: 0.56,
      fontSize: 15,
      color: MAGENTA,
      charSpacing: 1,
      valign: "middle",
    });
    label(s, t, {
      x: M + 0.98,
      y,
      w: CW - 0.98,
      h: 0.56,
      fontSize: 15,
      color: WHITE,
      charSpacing: 0.8,
      valign: "middle",
    });
  });
}

// ================================================================ 4 · HONORARIO
{
  const s = pres.addSlide();
  s.background = { color: TEAL };

  label(s, "03  —  Honorario", { x: M, y: 0.82, w: CW, fontSize: 12, color: GREEN });

  s.addText("$300.000", {
    x: M,
    y: 1.75,
    w: 9.5,
    h: 1.75,
    fontFace: DISPLAY,
    fontSize: 96,
    bold: true,
    color: GREEN,
    margin: 0,
  });
  label(s, "líquidos al mes", {
    x: M,
    y: 3.68,
    w: 8.0,
    h: 0.42,
    fontSize: 18,
    color: WHITE,
    charSpacing: 3,
  });

  const notas = [
    [
      "La pauta va aparte",
      "Entre $200.000 y $450.000 al mes según la fase, a nombre de PobreVermut.",
    ],
    [
      "Un primer ciclo de tres meses",
      "Partimos con tres meses de trabajo. Al cierre revisamos juntos los resultados y definimos cómo seguir.",
    ],
  ];
  const bw = 5.3;
  notas.forEach(([t, d], i) => {
    const x = M + i * (bw + 0.63);
    label(s, t, { x, y: 5.05, w: bw, h: 0.3, fontSize: 11, color: GREEN, charSpacing: 2 });
    label(s, d, {
      x,
      y: 5.5,
      w: bw,
      h: 1.4,
      fontSize: 11.5,
      color: WHITE,
      charSpacing: 0.5,
      lineSpacingMultiple: 1.35,
      valign: "top",
    });
  });
}

// ================================================================ 5 · CIERRE
{
  const s = pres.addSlide();
  s.background = { color: PURPLE };

  s.addImage({ path: TAGLINE, x: 5.72, y: 1.15, w: 1.89, h: 2.06 });

  s.addText(
    [
      { text: "TRABAJEMOS ", options: { color: WHITE } },
      { text: "JUNTOS", options: { color: MAGENTA } },
    ],
    {
      x: 1.5,
      y: 3.78,
      w: 10.33,
      h: 0.95,
      align: "center",
      fontFace: DISPLAY,
      fontSize: 46,
      bold: true,
      margin: 0,
    }
  );

  label(s, "Nicolás", {
    x: 1.5,
    y: 5.28,
    w: 10.33,
    h: 0.32,
    align: "center",
    fontSize: 13,
    color: WHITE,
    charSpacing: 3,
  });
  label(s, "nicolas@sherpasstudio.cl  ·  +56 9 3125 6539", {
    x: 1.5,
    y: 5.68,
    w: 10.33,
    h: 0.32,
    align: "center",
    fontSize: 11,
    color: LILAC,
    charSpacing: 1,
  });
}

// ---------------------------------------------------------------- salida
const out = process.argv[2] || "PobreVermut - Propuesta Campanas 2026.pptx";
pres.writeFile({ fileName: out }).then(() => console.log("OK ->", out));
