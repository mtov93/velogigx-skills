const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak } = require("docx");
const fs = require("fs");

const B = {
  black: "1f2937", white: "FFFFFF",
  accent: "1a5c2e", accentLight: "effced",
  gray1: "f6f8f6", gray3: "999999", gray4: "666666",
  red: "C62828", redLight: "fff1f1", redBorder: "e53935",
  green: "2E7D32", greenLight: "E8F5E9",
  orange: "E65100", orangeLight: "FFF3E0",
  navy: "2b3a52",
  font: "Montserrat",
};
const CW = 9506;
const tb = { style: BorderStyle.SINGLE, size: 1, color: "EAEAEA" };
const cp = { top: 70, bottom: 70, left: 130, right: 130 };
const CH = "/home/claude/charts/";

// ---------- helpers ----------
function h1(text) {
  return new Paragraph({ spacing: { before: 260, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: B.accent, space: 8 } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, font: B.font, size: 30, color: B.accent })] });
}
function h2(text) {
  return new Paragraph({ spacing: { before: 320, after: 110 },
    children: [new TextRun({ text, bold: true, font: B.font, size: 24, color: B.black })] });
}
function p(runs, opts = {}) {
  const ch = typeof runs === "string"
    ? [new TextRun({ text: runs, font: B.font, size: 22, color: B.black })]
    : runs.map(r => typeof r === "string"
      ? new TextRun({ text: r, font: B.font, size: 22, color: B.black })
      : new TextRun({ font: B.font, size: 22, color: B.black, ...r }));
  return new Paragraph({ spacing: { after: 120, line: 290 }, ...opts, children: ch });
}
// lead-in en cursiva gris (enmarca la fase en lenguaje de negocio)
function leadin(text) {
  return new Paragraph({ spacing: { after: 180, line: 290 },
    children: [new TextRun({ text, italics: true, font: B.font, size: 21, color: B.gray4 })] });
}
// caption pequeño bajo un gráfico
function caption(text) {
  return new Paragraph({ spacing: { before: 40, after: 160 },
    children: [new TextRun({ text, italics: true, font: B.font, size: 17, color: B.gray3 })] });
}
// chip de fase (suave: fondo verde claro, texto verde)
function chip(text) {
  return new Table({ width: { size: 1450, type: WidthType.DXA }, columnWidths: [1450],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
      shading: { fill: B.accentLight, type: ShadingType.CLEAR },
      margins: { top: 70, bottom: 70, left: 160, right: 160 },
      width: { size: 1450, type: WidthType.DXA },
      children: [new Paragraph({ spacing:{after:0}, children: [new TextRun({ text, bold: true, font: B.font, size: 19, color: B.accent })] })],
    })] })] });
}
// callout con eyebrow en negrita + cuerpo, acento lateral, sin emojis
function callout(eyebrow, body, tone) {
  const map = {
    red:   { bar: B.redBorder,  bg: B.redLight,   eb: B.red },
    green: { bar: B.accent,     bg: B.accentLight,eb: B.accent },
    biz:   { bar: B.accent,     bg: B.gray1,      eb: B.accent },
    accent:{ bar: B.accent,     bg: B.gray1,      eb: B.accent },
    orange:{ bar: B.orange,     bg: B.orangeLight,eb: B.orange },
    gray:  { bar: B.navy,       bg: B.gray1,      eb: B.navy },
  }[tone] || { bar: B.navy, bg: B.gray1, eb: B.navy };
  const lines = Array.isArray(body) ? body : [body];
  const content = [];
  content.push(new Paragraph({ spacing: { after: 50 },
    children: [new TextRun({ text: eyebrow, bold: true, font: B.font, size: 18, color: map.eb })] }));
  lines.forEach((l, i) => {
    const runs = Array.isArray(l) ? l.map(r => typeof r === "string"
        ? new TextRun({ text: r, font: B.font, size: 19, color: B.black })
        : new TextRun({ font: B.font, size: 19, color: B.black, ...r }))
      : [new TextRun({ text: l, font: B.font, size: 19, color: B.black })];
    content.push(new Paragraph({ spacing: { after: i === lines.length-1 ? 0 : 70, line: 288 }, children: runs }));
  });
  return new Table({ width: { size: CW, type: WidthType.DXA }, columnWidths: [CW],
    rows: [new TableRow({ children: [new TableCell({
      borders: { left: { style: BorderStyle.SINGLE, size: 18, color: map.bar },
        top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
      shading: { fill: map.bg, type: ShadingType.CLEAR },
      margins: { top: 150, bottom: 150, left: 220, right: 200 },
      width: { size: CW, type: WidthType.DXA }, children: content,
    })] })] });
}
function tbl(headers, rows, colWidths) {
  const tr = [];
  tr.push(new TableRow({ tableHeader: true, children: headers.map((h, i) => new TableCell({
    borders: { bottom: { style: BorderStyle.SINGLE, size: 6, color: B.accent }, top:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
    width: { size: colWidths[i], type: WidthType.DXA },
    shading: { fill: B.accentLight, type: ShadingType.CLEAR },
    margins: cp, verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ spacing:{after:0}, children: [new TextRun({ text: h, bold: true, font: B.font, size: 17, color: B.accent })] })],
  })) }));
  rows.forEach((row, ri) => {
    const fill = ri % 2 === 0 ? B.white : B.gray1;
    tr.push(new TableRow({ children: row.map((cell, ci) => {
      const runs = Array.isArray(cell) ? cell.map(c => new TextRun({ font: B.font, size: 17, color: B.black, ...c }))
        : [new TextRun({ font: B.font, size: 17, color: B.black, ...(typeof cell === "object" ? cell : { text: String(cell) }) })];
      return new TableCell({
        borders: { bottom: tb, top:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        width: { size: colWidths[ci], type: WidthType.DXA },
        shading: { fill, type: ShadingType.CLEAR }, margins: cp, verticalAlign: VerticalAlign.CENTER,
        children: [new Paragraph({ spacing:{after:0}, children: runs })],
      });
    }) }));
  });
  return new Table({ width: { size: colWidths.reduce((a,b)=>a+b,0), type: WidthType.DXA }, columnWidths: colWidths, rows: tr });
}
function img(file, w, aspect) {
  return new Paragraph({ spacing: { before: 60, after: 60 }, alignment: AlignmentType.CENTER,
    children: [new ImageRun({ data: fs.readFileSync(CH + file), type: "png",
      transformation: { width: w, height: Math.round(w * aspect) } })] });
}
const FAIL = { text: "FALLA", bold: true, color: B.red };
const WARN = { text: "MEJORABLE", bold: true, color: B.orange };
const OK = { text: "OK", bold: true, color: B.green };
const CRIT = { text: "CRÍTICO", bold: true, color: B.red };
const ACEPT = { text: "ACEPTABLE", bold: true, color: B.green };
const sp = (n=200) => new Paragraph({ spacing: { before: n }, children: [] });
const pb = new Paragraph({ children: [new PageBreak()] });

// ================= COVER =================
const cover = [
  new Paragraph({ spacing: { before: 3600 }, children: [] }),
  new Paragraph({ children: [new ImageRun({ data: fs.readFileSync(CH + "logo.png"), type: "png", transformation: { width: 300, height: Math.round(300*0.2417) } })] }),
  new Paragraph({ spacing: { before: 60, after: 40 }, children: [new TextRun({ text: "MARKETING DIGITAL", font: B.font, color: B.gray3, size: 18 })] }),
  new Paragraph({ spacing: { before: 500 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: B.accent, space: 12 } },
    children: [ new TextRun({ text: "Mini ", bold: true, font: B.font, size: 46 }),
      new TextRun({ text: "Auditoría SEO & Performance", bold: true, font: B.font, size: 46, color: B.black }) ] }),
  new Paragraph({ spacing: { before: 220 }, children: [] }),
  new Table({ width: { size: 4400, type: WidthType.DXA }, columnWidths: [4400],
    rows: [new TableRow({ children: [new TableCell({
      borders: { left: { style: BorderStyle.SINGLE, size: 18, color: B.accent },
        top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
      shading: { fill: B.accentLight, type: ShadingType.CLEAR },
      margins: { top: 150, bottom: 150, left: 220, right: 200 },
      width: { size: 4400, type: WidthType.DXA },
      children: [
        new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: "GERENS", bold: true, font: B.font, size: 30, color: B.accent })] }),
        new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "gerens.pe", font: B.font, color: B.gray4, size: 22 })] }),
      ],
    })] })] }),
  new Paragraph({ spacing: { before: 300 }, children: [new TextRun({ text: "Junio 2026  ·  Documento confidencial", font: B.font, color: B.gray3 })] }),
];

const body = [];
// ===== INTRODUCCIÓN =====
body.push(h1("Introducción"));
body.push(p("La Escuela de Postgrado GĚRENS es una institución académica peruana fundada en 1998 y la primera escuela de postgrado no adscrita a una universidad en obtener licencia de la SUNEDU (2018). Está especializada en los sectores clave de la economía peruana \u2014 minería, recursos naturales y energía \u2014 con una oferta que incluye el MBA Minero (Maestría en Gestión Minera), el MBA STEM, educación ejecutiva y corporativa, y consultoría. Acumula más de 12,000 profesionales capacitados y cerca de 700 consultorías. Su sitio (gerens.pe) es bilingüe (español e inglés) y soporta inversión publicitaria en Google, Meta y LinkedIn."));
body.push(p("Esta auditoría evalúa cinco dimensiones del desempeño digital de gerens.pe \u2014 rendimiento técnico, contenido y on-page, visibilidad orgánica, autoridad de dominio y presencia en respuestas de IA \u2014 a partir de datos verificables de Google Lighthouse y DataForSEO, recolectados en junio de 2026."));
body.push(p([{ text: "Hallazgo principal: ", bold: true, color: B.accent },
  { text: "GĚRENS tiene una presencia orgánica mínima (1,295 visitas/mes) para una institución con más de 25 años, y un sitio crítico en velocidad: la home pesa 37 MB y tarda 7.4 s en mostrar el contenido principal. El poco tráfico que capta proviene de artículos de blog sin relación con matricularse y de búsquedas por su propio nombre \u2014 mientras sus páginas comerciales (MBA Minero, maestrías) no aparecen en Google y el mercado de postgrado (~33,000 búsquedas/mes) queda casi sin tocar.", bold: true }]));

// ===== RESUMEN EJECUTIVO =====
body.push(pb);
body.push(h1("Resumen Ejecutivo"));
body.push(p("Cuatro cifras clave:"));
body.push(img("kpis.png", 625, 0.225));
body.push(sp(120));
body.push(h2("Estado general por fase"));
body.push(tbl(["Fase", "Estado", "Hallazgo principal"], [
  ["1. Técnico", CRIT, "La home pesa 37 MB y tarda 7.4 s en mostrar el contenido principal."],
  ["2. Contenido", WARN, "El poco tráfico viene de blog informativo; las páginas de programa no posicionan."],
  ["3. Keywords", CRIT, "Presencia orgánica mínima; ausente en todas las búsquedas comerciales de posgrado."],
  ["4. Backlinks", WARN, "Autoridad modesta y razonable, pero no canalizada a las páginas que venden."],
  ["5. AI Overviews", WARN, "Google resume con IA en sus búsquedas clave y GĚRENS queda debajo del resumen."],
], [1700, 1750, 6056]));
body.push(sp(300));
body.push(h2("Scores de Google Lighthouse (home)"));
body.push(tbl(["Categoría", "Score", "Estado"], [
  ["Performance", { text: "14 / 100", bold: true }, CRIT],
  ["Accesibilidad", { text: "73 / 100", bold: true }, WARN],
  ["Best Practices", { text: "73 / 100", bold: true }, WARN],
  ["SEO", { text: "85 / 100", bold: true }, OK],
], [4000, 2753, 2753]));
body.push(sp(300));
body.push(callout("EL HALLAZGO CENTRAL", [
  [{text:"GĚRENS capta apenas "},{text:"~1,295 visitas/mes orgánicas",bold:true},{text:" \u2014 una cifra mínima para una institución de su trayectoria. Casi todo ese tráfico proviene de artículos de blog sin relación con el negocio (minas de oro, valoración, gerentes generales) o de búsquedas por su propio nombre."}],
  [{text:"Sus páginas comerciales \u2014 "},{text:"MBA Minero, MBA STEM y maestrías",bold:true},{text:" \u2014 no aparecen en Google. Y su competencia en buscador no son escuelas de negocio: son Scribd, el Estado, los diarios y los portales mineros."}],
], "biz"));
body.push(sp(200));
body.push(callout("LA OPORTUNIDAD EN PERSPECTIVA", [
  [{text:"El mercado de búsquedas comerciales de postgrado en Perú \u2014 MBA, maestrías, posgrado, diplomados, gestión pública \u2014 supera las "},{text:"33,000 búsquedas/mes",bold:true},{text:". GĚRENS capta efectivamente 0% de ese mercado fuera de su propia marca."}],
  "Corregir la velocidad del sitio, hacer que las páginas de programa posicionen y construir presencia en el nicho minero \u2014 donde GĚRENS debería ser la autoridad \u2014 convertiría una presencia casi nula en visibilidad real frente a quien decide dónde hacer un posgrado.",
], "green"));

// ===== FASE 01 — TÉCNICO =====
body.push(pb);
body.push(chip("FASE 01"));
body.push(h1("Problemas Técnicos — Prioridad Crítica"));
body.push(leadin("La velocidad del sitio es lo primero que experimenta un postulante, antes de ver un solo programa. GĚRENS invierte en publicidad en Google, Meta y LinkedIn para atraer tráfico; si la página tarda en cargar, una parte de esas personas \u2014 que ya costaron presupuesto \u2014 se va antes de ver la oferta académica. Y Google penaliza en su ranking a los sitios lentos. Esta sección mide cuán rápido y estable es gerens.pe frente al estándar que Google exige hoy."));
body.push(h2("Core Web Vitals y Métricas de Rendimiento"));
body.push(img("cwv.png", 600, 0.3176));
body.push(caption("Barra = valor del sitio; marca vertical = meta de Google. Menos es mejor."));
body.push(h2("Cómo leer estas métricas"));
body.push(tbl(["Indicador", "Qué mide (en simple)"], [
  ["LCP", "Cuánto tarda en aparecer el contenido principal de la página."],
  ["FCP", "Cuánto tarda en verse el primer elemento al abrir la página."],
  ["Speed Index", "Qué tan rápido se ve completa la pantalla."],
  ["TTI", "Cuánto tarda el sitio en volverse usable tras abrirlo."],
  ["CLS", "Cuánto se mueve la página mientras carga (genera clics por error)."],
  ["TTFB", "Cuánto tarda el servidor en responder la primera vez."],
  ["Performance", "Nota global de velocidad que Google le da al sitio."],
], [2200, 7306]));
body.push(sp(280));
body.push(tbl(["Métrica", "Valor", "Meta Google", "Estado"], [
  ["Performance Score", { text: "14 / 100", bold: true, color: B.red }, "\u2265 90", FAIL],
  ["LCP (Largest Contentful Paint)", { text: "7.4 s", bold: true, color: B.red }, "< 2.5 s", FAIL],
  ["FCP (First Contentful Paint)", { text: "2.1 s", bold: true, color: B.orange }, "< 1.8 s", WARN],
  ["Speed Index", { text: "8.1 s", bold: true, color: B.red }, "< 3.4 s", FAIL],
  ["TTI (Time to Interactive)", { text: "17.3 s", bold: true, color: B.red }, "< 3.8 s", FAIL],
  ["CLS (Cumulative Layout Shift)", { text: "0.26", bold: true, color: B.red }, "< 0.10", FAIL],
  ["TTFB (respuesta del servidor)", { text: "637 ms", bold: true, color: B.green }, "< 800 ms", OK],
], [3906, 1900, 1900, 1800]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"El contenido principal de la home tarda "},{text:"7.4 segundos en aparecer",bold:true},{text:", casi tres veces el umbral aceptable, y el sitio recién se vuelve usable a los 17 segundos. Una parte de quienes llegan abandona antes de ver la oferta."}],
  [{text:"Como GĚRENS lleva tráfico pagado (Google, Meta, LinkedIn) a estas páginas, ese presupuesto se desperdicia en visitantes que se van antes de cargar. Además, "},{text:"Google penaliza a los sitios lentos",bold:true},{text:" bajando la posición orgánica."}],
], "accent"));
body.push(h2("Peso de Página y Recursos (home)"));
body.push(tbl(["Métrica", "Valor", "Referencia", "Estado"], [
  ["Peso total de la página", { text: "37 MB", bold: true, color: B.red }, "< 2\u20133 MB", FAIL],
  ["Trabajo del hilo principal", { text: "7.9 s", bold: true, color: B.red }, "< 2.0 s", FAIL],
  ["Arranque de JavaScript", { text: "3.7 s", bold: true, color: B.red }, "< 0.6 s", FAIL],
], [3906, 1900, 1900, 1800]));
body.push(sp(280));
body.push(p([{text:"La página interna de la maestría carga mucho mejor (4 MB, LCP 2.7 s): ",},{text:"el problema está concentrado en la home",bold:true},{text:", que es justamente la puerta de entrada y la página a la que llega la publicidad."}]));
body.push(h2("Stack Tecnológico Detectado"));
body.push(tbl(["Componente", "Detalle", "Observación"], [
  ["Gestor de tags", "Google Tag Manager", "Orquesta múltiples scripts a la vez."],
  ["Analítica", "Yandex Metrica + Microsoft Clarity", "Dos herramientas de analítica en paralelo."],
  ["Publicidad", "Google Ads, LinkedIn Ads, Retargetly, Meta", "Cuatro plataformas de retargeting cargando."],
  ["CDN", "JSDelivr", "Recursos externos adicionales por descargar."],
], [2700, 3906, 2900]));
body.push(sp(280));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "Con una home de 37 MB y un LCP de 7.4 s, una porción significativa del tráfico que llega \u2014 orgánico y pagado \u2014 abandona antes de interactuar. Buena parte de ese peso viene de acumular varias herramientas de analítica y publicidad cargando a la vez. La velocidad condiciona el rendimiento de todo lo demás: SEO, contenido y publicidad terminan en una página que tarda demasiado en cargar.",
], "red"));

// ===== FASE 02 — CONTENIDO =====
body.push(pb);
body.push(chip("FASE 02"));
body.push(h1("Contenido y On-page — Tráfico que no es del Negocio"));
body.push(leadin("El sitio es el vendedor que trabaja las 24 horas. Si el contenido que atrae visitas no tiene relación con lo que se vende, ese tráfico se ve en los reportes pero no acerca a nadie a matricularse. Esta sección revisa de dónde viene realmente el tráfico de GĚRENS y si comunica su oferta a quien la está buscando."));
body.push(img("donut.png", 600, 0.3756));
body.push(sp(120));
body.push(h2("Composición del tráfico orgánico"));
body.push(tbl(["Página que atrae tráfico", "Búsqueda principal", "Tipo de contenido"], [
  ["Blog: principales mineras de oro", "minera oro / mineras de oro (1.3K)", { text: "Blog informativo", color: B.accent }],
  ["Blog: ¿Qué es valoración?", "valoración (1.6K)", { text: "Blog informativo", color: B.accent }],
  ["Blog: gerentes generales estratégicos", "gerentes generales (2.9K)", { text: "Blog informativo", color: B.accent }],
  ["Blog: ¿Qué es ser un comunicador?", "comunicadores (720)", { text: "Blog informativo", color: B.accent }],
  ["Home", "gerens (1.3K)", { text: "Marca", color: B.accent }],
], [3700, 3306, 2500]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"Las páginas que más tráfico atraen son artículos de blog de hace varios años sobre temas que no llevan a una matrícula: alguien que busca \u201Cqué es valoración\u201D o \u201Cminas de oro\u201D no está evaluando un MBA. Es tráfico que "},{text:"no alimenta el negocio",bold:true},{text:", y por debajo de él, la única página comercial con tracción es la home \u2014 y solo por el nombre de la marca."}],
], "accent"));
body.push(h2("Las páginas de programa no captan tráfico"));
body.push(p([{text:"GĚRENS "},{text:"sí tiene",bold:true},{text:" páginas comerciales construidas: el MBA Minero, el MBA STEM y decenas de cursos de educación ejecutiva y corporativa enfocados en minería (gestión de costos mineros, evaluación de proyectos mineros, perforación y voladura, relaves). El problema no es que falten páginas, sino que "},{text:"ninguna de ellas aparece entre las páginas que captan tráfico orgánico",bold:true},{text:". El esfuerzo editorial está en el blog informativo, desconectado de las páginas que venden."}]));
body.push(p([{text:"El sitio además es bilingüe (ES/EN); conviene verificar que la versión en inglés no genere contenido duplicado ni reste fuerza a las URLs principales en español."}]));
body.push(sp(120));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "El volumen de tráfico ya es bajo, y además es de baja relevancia comercial: la mayor parte no corresponde a personas buscando un posgrado. El contenido informativo puede ser un activo, pero hoy no está canalizado hacia las páginas de programa ni hacia las búsquedas de quien decide dónde matricularse.",
], "red"));

// ===== FASE 03 — KEYWORDS =====
body.push(pb);
body.push(chip("FASE 03"));
body.push(h1("Keywords — Presencia Comercial Casi Nula"));
body.push(leadin("Cada búsqueda en Google es un cliente potencial levantando la mano. Esta sección separa el poco tráfico que GĚRENS capta hoy del mercado comercial de posgrado que, teniendo los programas para atenderlo, aún no toca."));
body.push(h2("Métricas generales de posicionamiento"));
body.push(tbl(["Métrica", "Valor", "Análisis"], [
  ["Tráfico orgánico estimado (ETV)", { text: "1,295 visitas/mes", bold: true }, "Mínimo para una institución de 25+ años."],
  ["Keywords posicionadas", { text: "82", bold: true }, "Cobertura muy estrecha, mayormente de blog."],
  ["Keywords en Top 10", { text: "25", bold: true }, "Casi todas informativas, no comerciales."],
  ["Keywords en posición #1", { text: "1", bold: true }, "Es su propia marca (\u201Cgerens\u201D)."],
], [3900, 2306, 3300]));
body.push(sp(300));
body.push(h2("Las búsquedas comerciales del negocio — dónde rankea hoy"));
body.push(caption("Posición actual de gerens.pe en Google (Perú) para sus búsquedas comerciales clave."));
body.push(tbl(["Búsqueda comercial", "Vol./mes", "Posición GĚRENS", "Visitas/mes est."], [
  ["gerens (marca)", "1,300", { text: "#1", bold: true, color: B.green }, "~395"],
  ["escuelas de posgrado", "1,600", { text: "#26", bold: true, color: B.orange }, "~3"],
  ["maestría en gestión minera", "140", { text: "No posiciona", bold: true, color: B.red }, "~0"],
  ["gestión minera", "170", { text: "No posiciona", bold: true, color: B.red }, "~0"],
  ["posgrado", "4,400", { text: "No posiciona", bold: true, color: B.red }, "~0"],
  ["maestrías", "6,600", { text: "No posiciona", bold: true, color: B.red }, "~0"],
  ["mba", "6,600", { text: "No posiciona", bold: true, color: B.red }, "~0"],
  ["diplomados", "4,400", { text: "No posiciona", bold: true, color: B.red }, "~0"],
], [3906, 1500, 2400, 1700]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"GĚRENS domina su propia marca \u2014 está "},{text:"#1 en \u201Cgerens\u201D",bold:true},{text:" \u2014 pero está ausente en prácticamente todas las búsquedas comerciales del negocio. Lo más llamativo: incluso en "},{text:"\u201Cmaestría en gestión minera\u201D y \u201Cgestión minera\u201D",bold:true},{text:", el nicho donde GĚRENS debería ser la autoridad del país, no aparece en los resultados."}],
], "accent"));
body.push(h2("Keywords ausentes con intención comercial"));
body.push(p([{text:"Las búsquedas donde GĚRENS tiene programa pero no presencia. La columna \u201CVisitas/mes est.\u201D aplica un CTR conservador del 5% (el dominio hoy no posiciona en lo comercial; 5% asume alcanzar primera página, no Top 3). No se aplica ninguna tasa de conversión."}]));
body.push(tbl(["Keyword ausente", "Vol./mes", "Por qué debería aparecer", "Visitas/mes est."], [
  ["maestrías", "6,600", "Ofrece maestrías; término raíz del negocio.", "~330"],
  ["mba", "6,600", "Tiene MBA Minero y MBA STEM.", "~330"],
  ["posgrado", "4,400", "Es una escuela de postgrado licenciada.", "~220"],
  ["diplomados", "4,400", "Oferta amplia de diplomados ejecutivos.", "~220"],
  ["maestría en gestión pública", "3,600", "Cartera de gestión y management.", "~180"],
  ["doctorado", "2,400", "Nivel de posgrado colindante.", "~120"],
  ["escuelas de posgrado", "1,600", "Categoría directa; hoy en #26.", "~80"],
  ["maestría en finanzas", "880", "Tiene cursos de costos y finanzas.", "~44"],
  ["maestría en gestión de proyectos", "880", "Gestión de proyectos mineros.", "~44"],
  ["educación ejecutiva", "390", "Línea de negocio propia.", "~20"],
  ["maestría en project management", "320", "Programas de gestión de proyectos.", "~16"],
  ["gestión minera", "170", "Su especialidad central.", "~9"],
  ["maestría en gestión minera", "140", "Su programa estrella (MMBA).", "~7"],
  ["diplomado en minería", "70", "Múltiples diplomados mineros.", "~4"],
  ["maestría en minería", "70", "Su nicho de especialización.", "~4"],
], [3500, 1300, 3206, 1500]));
body.push(sp(220));
body.push(callout("OPORTUNIDAD EN TRÁFICO", [
  [{text:"Estas búsquedas suman "},{text:"~33,000 búsquedas/mes",bold:true},{text:". Capturarlas en primera página representa del orden de "},{text:"~1,650 visitas/mes",bold:true},{text:" \u2014 más del doble de todo el tráfico orgánico que GĚRENS capta hoy, y de intención mucho más comercial."}],
], "green"));
body.push(h2("Con quién compite hoy en Google"));
body.push(caption("Sitios que comparten más resultados de búsqueda con gerens.pe (keywords en común)."));
body.push(img("competidores.png", 545, 0.586));
body.push(sp(120));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"Los mayores \u201Ccompetidores\u201D de GĚRENS en Google son repositorios de documentos (Scribd, Studocu), el Estado (gob.pe, Minem) y diarios (Gestión, El Comercio) \u2014 "},{text:"ninguno es una escuela de postgrado",bold:true},{text:". Es la prueba de que su SEO compite en el carril informativo, no en el de la educación ejecutiva. ESAN también aparece, pero por sus artículos de blog, no por sus programas."}],
], "accent"));
body.push(h2("Con quién debería competir — las escuelas de negocio"));
body.push(caption("Valor del tráfico orgánico mensual: lo que costaría comprarlo en Google Ads."));
body.push(img("eficiencia.png", 600, 0.2589));
body.push(sp(120));
body.push(tbl(["Dominio", "Keywords", "Tráfico orgánico", "Valor del tráfico"], [
  ["gerens.pe", "82", "1,295 visitas/mes", { text: "$680", bold: true, color: B.red }],
  ["esan.edu.pe", "4,143", "173,843 visitas/mes", { text: "$96,692", bold: true, color: B.green }],
], [3000, 2000, 2706, 1800]));
body.push(sp(280));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"ESAN es el referente del mercado, no un competidor directo: muestra el tamaño del mercado de postgrado que GĚRENS hoy no toca. "},{text:"GĚRENS no necesita igualar a ESAN",bold:true},{text:", pero sí debería aparecer en las búsquedas comerciales de su propio nicho \u2014 la minería \u2014 donde tiene la oferta más especializada del país y, aun así, hoy es invisible."}],
], "accent"));
body.push(h2("La posición competitiva de un vistazo"));
body.push(tbl(["Dimensión", "Situación actual", "Estado", "Lectura"], [
  ["Volumen de tráfico", "1,295 visitas/mes", FAIL, "Mínimo para su trayectoria."],
  ["Calidad del tráfico", "Blog + marca", FAIL, "Las páginas de programa no captan."],
  ["Velocidad (LCP home)", "7.4 s", FAIL, "Casi 3x la meta de 2.5 s."],
  ["Peso de página (home)", "37 MB", FAIL, "~15x lo recomendable."],
  ["Autoridad de dominio", "146 dominios ref.", WARN, "Base razonable, sin canalizar."],
  ["Comerciales de marca", "#1 \u201Cgerens\u201D", OK, "Domina su propio nombre."],
  ["Comerciales del negocio", "Fuera del Top 100", FAIL, "Posgrado, maestrías, MBA: invisible."],
], [2700, 2500, 1600, 2706]));
body.push(sp(300));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "El problema de GĚRENS no es la calidad del tráfico, sino su ausencia: está fuera de las búsquedas que importan, incluso en su propio nicho minero, donde debería ser la autoridad. Cerrar la velocidad del sitio y posicionar las páginas de programa por las búsquedas comerciales es la palanca de mayor impacto del documento.",
], "red"));

// ===== FASE 04 — BACKLINKS =====
body.push(pb);
body.push(chip("FASE 04"));
body.push(h1("Backlinks — Autoridad Modesta y sin Canalizar"));
body.push(leadin("Los enlaces de otros sitios funcionan como recomendaciones: mientras más sitios relevantes enlacen, más confía Google. Esta sección mide la autoridad que GĚRENS ha acumulado y si esa fuerza está empujando a las páginas correctas."));
body.push(h2("Perfil de autoridad del dominio"));
body.push(tbl(["Métrica", "Valor", "Análisis"], [
  ["Dominios referentes", { text: "~146", bold: true }, "Base razonable para una escuela de nicho."],
  ["Backlinks totales", { text: "~1,260", bold: true }, "Volumen modesto, acumulado con los años."],
  ["Enlaces dofollow", { text: "~1,129 (90%)", bold: true }, "Buena proporción que transmite autoridad."],
  ["Domain Rank (DataForSEO, 0\u20131000)", { text: "~269", bold: true }, "Dominio de autoridad media."],
], [3906, 2300, 3300]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"La autoridad de GĚRENS es razonable para su tamaño, pero "},{text:"modesta frente a las escuelas grandes",bold:true},{text:" (ESAN supera los 470 dominios referentes). El problema no es solo cantidad: el respaldo apunta sobre todo a la home y al blog, no a las páginas de programa, que son las que necesitan competir por las búsquedas comerciales."}],
], "accent"));
body.push(sp(200));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "La autoridad acumulada hoy no trabaja para el negocio: si esa fuerza se canalizara mediante enlazado interno hacia las páginas del MBA Minero, las maestrías y los diplomados, podrían empezar a competir por las búsquedas comerciales que hoy se ceden por completo.",
], "red"));

// ===== FASE 05 — AI OVERVIEWS =====
body.push(pb);
body.push(chip("FASE 05"));
body.push(h1("AI Overviews — Google Responde y GĚRENS queda Debajo"));
body.push(leadin("Google ya responde muchas búsquedas con un resumen generado por IA, antes de mostrar los resultados. Si el sitio no es una de las fuentes que esa IA cita, desaparece de la primera respuesta que ve el usuario \u2014 aunque la página esté posicionada más abajo. Esta sección mide dónde está ocurriendo esto."));
body.push(h2("Búsquedas con AI Overview activo donde GĚRENS está debajo"));
body.push(tbl(["Búsqueda", "Vol./mes", "Posición GĚRENS", "AI Overview"], [
  ["mineras del perú", "3,600", "#12", { text: "Activo", bold: true, color: B.red }],
  ["gestión de riesgos", "2,900", "#11", { text: "Activo", bold: true, color: B.red }],
  ["trabajadores en minas", "2,400", "#16", { text: "Activo", bold: true, color: B.red }],
  ["crisis energética", "1,600", "#18", { text: "Activo", bold: true, color: B.red }],
  ["minas de oro", "1,300", "#8", { text: "Activo", bold: true, color: B.red }],
], [3906, 1700, 2200, 1700]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"En estas búsquedas, Google muestra arriba un resumen de IA que responde la consulta "},{text:"sin que el usuario haga clic",bold:true},{text:". GĚRENS aparece debajo, en página 1\u20132, pero el AI Overview se lleva la atención primero."}],
  [{text:"En \u201Cmineras del perú\u201D, el AI Overview se nutre de Wikipedia, el MINEM y medios especializados \u2014 "},{text:"no de GĚRENS",bold:true},{text:", que está en #12. Aproximadamente 8 de cada 10 de las keywords donde aparece ya tienen este formato activo."}],
], "accent"));
body.push(sp(200));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "El tráfico informativo del blog \u2014 lo único que hoy sostiene las cifras de GĚRENS \u2014 es justamente el más expuesto a los AI Overviews: a medida que Google responda más consultas con IA, ese volumen se erosionará. Es otra razón para reorientar la estrategia hacia las búsquedas comerciales de posgrado minero, menos absorbidas por la IA y de mucho mayor valor.",
], "red"));

// ===== CONCLUSIONES =====
body.push(pb);
body.push(h1("Conclusiones y Plan de Acción"));
body.push(h2("Principales hallazgos priorizados"));
body.push(tbl(["#", "Hallazgo", "Prioridad", "Impacto", "Plazo"], [
  ["1", "Home de 37 MB con LCP de 7.4 s", FAIL, { text: "Crítico", color: B.red, bold: true }, "1-3 meses"],
  ["2", "Presencia orgánica mínima (1,295 visitas/mes)", FAIL, { text: "Alto", color: B.orange, bold: true }, "3-6 meses"],
  ["3", "Páginas de programa (MBA Minero, maestrías) no posicionan", FAIL, { text: "Alto", color: B.orange, bold: true }, "1-3 meses"],
  ["4", "Ausente en búsquedas comerciales de posgrado", WARN, { text: "Alto", color: B.orange, bold: true }, "3-6 meses"],
  ["5", "Tráfico informativo expuesto a AI Overviews", WARN, { text: "Medio", color: B.orange, bold: true }, "Continuo"],
], [600, 4406, 1700, 1400, 1400]));
body.push(sp(300));
body.push(callout("PLAN DE ACCIÓN POR FASES", [
  [{text:"Mes 1-3 (Quick wins): ",bold:true},{text:"optimizar el peso y la velocidad de la home (objetivo LCP < 2.5 s, reducir el peso de 37 MB) y trabajar el on-page de las páginas de programa \u2014 MBA Minero, MBA STEM y maestrías \u2014 que hoy no posicionan."}],
  [{text:"Mes 4-6 (Nicho minero): ",bold:true},{text:"capturar las búsquedas donde GĚRENS debería ser la autoridad (gestión minera, maestría en gestión minera, diplomados mineros) y canalizar la autoridad existente hacia esas páginas mediante enlazado interno."}],
  [{text:"Mes 7-12 (Escala): ",bold:true},{text:"competir por los términos generales de postgrado (posgrado, maestrías, diplomados) y conectar el blog informativo con las páginas comerciales. Meta: pasar de ~1,295 a varios miles de visitas/mes de intención comercial."}],
  [{text:"Siguiente paso: ",bold:true},{text:"reunión de 30-45 min para revisar hallazgos, propuesta de roadmap priorizado y definición de las primeras acciones (velocidad de la home y on-page de programas)."}],
], "green"));

// Guardia: inserta un espaciador entre cualquier par de Tablas/callouts adyacentes
function spacerGuard(arr) {
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    out.push(arr[i]);
    if (arr[i] instanceof Table && arr[i + 1] instanceof Table) out.push(sp(200));
  }
  return out;
}

// ================= DOC =================
const doc = new Document({
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1200, right: 1200, bottom: 1200, left: 1200 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD", space: 6 } }, spacing: { after: 200 },
      children: [ new ImageRun({ data: fs.readFileSync(CH + "logo.png"), type: "png", transformation: { width: 118, height: Math.round(118*0.2417) } }),
        new TextRun({ text: "   ·   Auditoría SEO & Performance   ·   gerens.pe", font: B.font, size: 13, color: B.gray3 }) ] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [ new TextRun({ text: "Pág. ", font: B.font, size: 15, color: B.gray3 }),
        new TextRun({ children: [PageNumber.CURRENT], font: B.font, size: 15, color: B.gray3 }),
        new TextRun({ text: "   ·   Documento confidencial", font: B.font, size: 13, color: "CCCCCC" }) ] })] }) },
    children: spacerGuard([...cover, pb, ...body]),
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync("auditoria_gerens.docx", buf); console.log("OK docx", buf.length); });
