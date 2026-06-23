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
        new Paragraph({ spacing: { after: 30 }, children: [new TextRun({ text: "ESAN", bold: true, font: B.font, size: 30, color: B.accent })] }),
        new Paragraph({ spacing: { after: 0 }, children: [new TextRun({ text: "esan.edu.pe", font: B.font, color: B.gray4, size: 22 })] }),
      ],
    })] })] }),
  new Paragraph({ spacing: { before: 300 }, children: [new TextRun({ text: "Junio 2026  ·  Documento confidencial", font: B.font, color: B.gray3 })] }),
];

const body = [];
// ===== INTRODUCCIÓN =====
body.push(h1("Introducción"));
body.push(p("ESAN Graduate School of Business (Universidad ESAN) es una institución académica peruana fundada en 1963 \u2014 la primera escuela de postgrado en administración de negocios del mundo de habla hispana. Con sede en Santiago de Surco (Lima), su oferta incluye el MBA en sus distintas modalidades, más de 15 maestrías especializadas, un programa doctoral, educación ejecutiva (PEE, PAE, diplomas) y educación corporativa. Su MBA es reconocido como el N.°1 del Perú y se ubica entre los primeros de América Latina (QS, Financial Times), con acreditaciones internacionales AACSB y AMBA. Su presencia digital se reparte entre varias propiedades \u2014 el sitio comercial (esan.edu.pe), el portal institucional y de admisión (uesan.edu.pe), la biblioteca y el repositorio de investigación, además de un dominio de pregrado \u2014 e invierte de forma sostenida en publicidad en Google, Meta, LinkedIn y TikTok."));
body.push(p("Esta auditoría evalúa cinco dimensiones del desempeño digital de esan.edu.pe \u2014 rendimiento técnico, contenido y on-page, visibilidad orgánica, autoridad de dominio y presencia en respuestas de IA \u2014 a partir de datos verificables de Google Lighthouse y DataForSEO, recolectados en junio de 2026."));
body.push(p([{ text: "Hallazgo principal: ", bold: true, color: B.accent },
  { text: "ESAN es líder claro de su categoría \u2014 capta ~173,844 visitas/mes orgánicas y rankea #1 en \u201Cmba\u201D y #2 en \u201Cmaestrías\u201D \u2014 pero la mayor parte de ese tráfico proviene de artículos informativos de Conexión ESAN (tipo de cambio, recibos por honorarios, marcas), justo el contenido más expuesto a los AI Overviews de Google. Al mismo tiempo, varias búsquedas comerciales de alto valor (posgrado, diplomados, maestrías especializadas) las cede o las responde con un artículo de blog en lugar de su página de programa, y la velocidad del sitio \u2014 home de 24 MB, páginas de programa con LCP de ~5-6 s \u2014 erosiona el rendimiento del fuerte tráfico pagado que la institución dirige a esas mismas páginas.", bold: true }]));

// ===== RESUMEN EJECUTIVO =====
body.push(pb);
body.push(h1("Resumen Ejecutivo"));
body.push(p("Cuatro cifras clave:"));
body.push(img("kpis.png", 625, 0.225));
body.push(sp(120));
body.push(h2("Estado general por fase"));
body.push(tbl(["Fase", "Estado", "Hallazgo principal"], [
  ["1. Técnico", WARN, "La home pesa 24 MB; las páginas de programa cargan en ~5-6 s (LCP), por debajo del estándar de Google."],
  ["2. Contenido", WARN, "El grueso del tráfico viene de Conexión ESAN informativo; algunas páginas de programa ceden la búsqueda a un blog."],
  ["3. Keywords", ACEPT, "Liderazgo en \u201Cmba\u201D y \u201Cmaestrías\u201D, pero brechas en posgrado, diplomados y varias especializadas."],
  ["4. Backlinks", ACEPT, "Autoridad de dominio sólida (~478 dominios ref.), poco canalizada hacia las páginas de programa."],
  ["5. AI Overviews", CRIT, "Su mayor fuente de tráfico \u2014 el contenido informativo \u2014 es la más expuesta a los resúmenes de IA de Google."],
], [1700, 1750, 6056]));
body.push(sp(300));
body.push(h2("Scores de Google Lighthouse (home)"));
body.push(tbl(["Categoría", "Score", "Estado"], [
  ["Performance", { text: "50 / 100", bold: true }, WARN],
  ["Accesibilidad", { text: "89 / 100", bold: true }, OK],
  ["Best Practices", { text: "69 / 100", bold: true }, WARN],
  ["SEO", { text: "92 / 100", bold: true }, OK],
], [4000, 2753, 2753]));
body.push(sp(300));
body.push(callout("EL HALLAZGO CENTRAL", [
  [{text:"ESAN domina su categoría: capta "},{text:"~173,844 visitas/mes orgánicas",bold:true},{text:" con 4,143 keywords, está #1 en \u201Cmba\u201D, #2 en \u201Cmaestrías\u201D y #2 en \u201Cmaestría en finanzas\u201D. La fortaleza es real."}],
  [{text:"Pero la composición de ese tráfico es el punto a observar: la mayoría llega por artículos de Conexión ESAN sobre temas generales \u2014 tipo de cambio del dólar, recibos por honorarios, marcas, sueldo mínimo \u2014 "},{text:"el contenido que más rápido absorben los AI Overviews de Google",bold:true},{text:". Y en varias búsquedas comerciales (posgrado, diplomados, especialización, marketing) ESAN aparece con un blog o queda fuera del Top 10, en lugar de su página de programa."}],
], "biz"));
body.push(sp(200));
body.push(callout("LA OPORTUNIDAD EN PERSPECTIVA", [
  [{text:"El mercado de búsquedas comerciales de postgrado en Perú \u2014 MBA, maestrías, posgrado, diplomados, especializadas \u2014 supera las "},{text:"37,000 búsquedas/mes",bold:true},{text:". ESAN captura bien las cabezas que ya optimizó, pero cede una parte de las especializadas y de los términos genéricos de mayor volumen."}],
  "Endurecer la velocidad de las páginas de programa (donde aterriza el tráfico pagado), reorientar la autoridad del dominio hacia esas páginas y recuperar las búsquedas comerciales que hoy responde un artículo de blog consolidaría el liderazgo de ESAN justo donde se decide una matrícula \u2014 y blindaría su base de tráfico frente al avance de la IA en buscador.",
], "green"));

// ===== FASE 01 — TÉCNICO =====
body.push(pb);
body.push(chip("FASE 01"));
body.push(h1("Problemas Técnicos — Prioridad de Optimización"));
body.push(leadin("La velocidad del sitio es lo primero que experimenta un postulante, antes de ver un solo programa. ESAN invierte de forma intensiva en publicidad en Google, Meta, LinkedIn y TikTok para atraer tráfico; si la página tarda en cargar, una parte de esas personas \u2014 que ya costaron presupuesto \u2014 se va antes de ver la oferta académica. Y Google considera la velocidad como factor de ranking. Esta sección mide cuán rápido y estable es esan.edu.pe frente al estándar que Google exige hoy."));
body.push(h2("Velocidad y estabilidad del sitio"));
body.push(p([{text:"Google evalúa la experiencia con los "},{text:"Core Web Vitals",bold:true},{text:": las dos métricas clave son "},{text:"LCP",bold:true},{text:" (cuánto tarda en cargar el contenido principal) y "},{text:"CLS",bold:true},{text:" (cuánto se mueve la página al cargar). A ellas se suman métricas de laboratorio \u2014 FCP, Speed Index y TTI \u2014 que ayudan a diagnosticar de dónde viene la lentitud. Así se lee cada una:"}]));
body.push(sp(120));
body.push(tbl(["Indicador", "Qué mide (en simple)"], [
  ["LCP \u2014 Core Web Vital", "Cuánto tarda en aparecer el contenido principal de la página."],
  ["CLS \u2014 Core Web Vital", "Cuánto se mueve la página mientras carga (genera clics por error)."],
  ["FCP", "Cuánto tarda en verse el primer elemento al abrir la página."],
  ["Speed Index", "Qué tan rápido se ve completa la pantalla."],
  ["TTI", "Cuánto tarda el sitio en volverse usable tras abrirlo."],
  ["TTFB", "Cuánto tarda el servidor en responder la primera vez."],
  ["Performance", "Nota global de velocidad que Google le da al sitio."],
], [2600, 6906]));
body.push(sp(300));
body.push(img("cwv.png", 600, 0.3176));
body.push(caption("Métricas de tiempo, en segundos: barra = valor del sitio, marca vertical = meta de Google; menos es mejor. El CLS, al medirse en un puntaje y no en segundos, se reporta en la tabla siguiente."));
body.push(sp(280));
body.push(tbl(["Métrica", "Valor", "Meta Google", "Estado"], [
  ["Performance Score", { text: "50 / 100", bold: true, color: B.orange }, "\u2265 90", WARN],
  ["LCP (Largest Contentful Paint)", { text: "6.0 s", bold: true, color: B.red }, "< 2.5 s", FAIL],
  ["FCP (First Contentful Paint)", { text: "2.6 s", bold: true, color: B.orange }, "< 1.8 s", WARN],
  ["Speed Index", { text: "4.7 s", bold: true, color: B.orange }, "< 3.4 s", WARN],
  ["TTI (Time to Interactive)", { text: "6.6 s", bold: true, color: B.red }, "< 3.8 s", FAIL],
  ["CLS (Cumulative Layout Shift)", { text: "0.17", bold: true, color: B.orange }, "< 0.10", WARN],
  ["TTFB (respuesta del servidor)", { text: "1.56 s", bold: true, color: B.red }, "< 0.8 s", FAIL],
], [3906, 1900, 1900, 1800]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"El contenido principal de la home tarda "},{text:"6 segundos en aparecer",bold:true},{text:", más del doble del umbral aceptable, y el servidor demora 1.56 s solo en empezar a responder. Una parte de quienes llegan abandona antes de ver la oferta."}],
  [{text:"Como ESAN dirige tráfico pagado (Google, Meta, LinkedIn, TikTok) a estas páginas, ese presupuesto rinde menos cuando el visitante se va antes de cargar. Además, "},{text:"Google usa la velocidad como factor de ranking",bold:true},{text:", de modo que la lentitud también frena el crecimiento orgánico."}],
], "accent"));
body.push(h2("Peso de Página y Recursos (home)"));
body.push(tbl(["Métrica", "Valor", "Referencia", "Estado"], [
  ["Peso total de la home", { text: "24 MB", bold: true, color: B.red }, "< 2\u20133 MB", FAIL],
  ["Trabajo del hilo principal", { text: "2.9 s", bold: true, color: B.orange }, "< 2.0 s", WARN],
  ["Arranque de JavaScript", { text: "1.0 s", bold: true, color: B.orange }, "< 0.6 s", WARN],
], [3906, 1900, 1900, 1800]));
body.push(sp(280));
body.push(p([{text:"La página interna del MBA \u2014 donde aterriza buena parte de la publicidad \u2014 carga más liviana (3.8 MB) pero "},{text:"obtiene un Performance de 46/100, con LCP de 5.1 s y un CLS de 0.22",bold:true},{text:": el contenido se mueve mientras carga, lo que genera clics por error en una página de conversión. La lentitud no es solo de la home: es un patrón del sitio, y golpea justo donde se decide una matrícula."}]));
body.push(h2("Stack Tecnológico Detectado"));
body.push(tbl(["Componente", "Detalle", "Observación"], [
  ["Plataforma", "Joomla + módulo Google PageSpeed", "CMS con optimización de imágenes activada."],
  ["Gestor de tags", "Google Tag Manager", "Orquesta múltiples scripts a la vez."],
  ["Analítica", "GA + Hotjar + Microsoft Clarity + Yandex", "Cuatro herramientas de analítica en paralelo."],
  ["Publicidad", "Google Ads, Meta, LinkedIn, TikTok", "Cuatro plataformas de retargeting cargando."],
], [2700, 3906, 2900]));
body.push(sp(280));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "Con una home de 24 MB y páginas de programa con LCP de ~5-6 s, una porción del tráfico que llega \u2014 orgánico y, sobre todo, pagado \u2014 abandona antes de interactuar. Buena parte de ese peso viene de acumular cuatro herramientas de analítica y cuatro de publicidad cargando a la vez. Para una institución que invierte de forma sostenida en captación, la velocidad es la palanca de mayor retorno inmediato: mejora la conversión de cada visita pagada y el ranking orgánico al mismo tiempo.",
], "red"));

// ===== FASE 02 — CONTENIDO =====
body.push(pb);
body.push(chip("FASE 02"));
body.push(h1("Contenido y On-page — Un Tráfico que No es del Negocio"));
body.push(leadin("El sitio es el vendedor que trabaja las 24 horas. Si el contenido que atrae visitas no tiene relación con lo que se vende, ese tráfico se ve en los reportes pero no acerca a nadie a matricularse. Esta sección revisa de dónde viene realmente el tráfico de ESAN y si su oferta llega a quien la está buscando."));
body.push(img("donut.png", 600, 0.3756));
body.push(sp(120));
body.push(h2("Composición del tráfico orgánico"));
body.push(tbl(["Página que atrae tráfico", "Búsqueda principal", "Tipo de contenido"], [
  ["Conexión ESAN: factura vs. recibo por honorarios", "recibo por honorarios (246K)", { text: "Blog informativo", color: B.accent }],
  ["Conexión ESAN: tendencia del dólar", "tipo de cambio dólar (301K)", { text: "Blog informativo", color: B.accent }],
  ["Conexión ESAN: ranking de marcas", "marcas (110K)", { text: "Blog informativo", color: B.accent }],
  ["Conexión ESAN: aumento de sueldo mínimo", "sueldo mínimo peru (90K)", { text: "Blog informativo", color: B.accent }],
  ["Home", "esan (40K)", { text: "Marca", color: B.accent }],
], [3700, 3306, 2500]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"Las páginas que más tráfico atraen son artículos de Conexión ESAN sobre temas de coyuntura \u2014 el precio del dólar, cómo emitir un recibo por honorarios, el sueldo mínimo \u2014 que captan grandes volúmenes pero "},{text:"no corresponden a alguien evaluando una maestría",bold:true},{text:". Es un activo de marca y reputación valioso, pero por sí solo no alimenta la matrícula, y por debajo de él las páginas de programa aportan una fracción del tráfico."}],
], "accent"));
body.push(h2("Detalles on-page detectados"));
body.push(p([{text:"En la inspección del sitio aparecen dos hallazgos concretos y de bajo esfuerzo. El primero, en el encabezado de la home:"}]));
body.push(sp(120));
body.push(callout("LO QUE ENCUENTRA GOOGLE EN LA HOME", [
  [{text:"Etiqueta og:title: "},{text:"\u201CHome Esan\u201D",bold:true},{text:". Es el título que Google y las redes muestran al referenciar el sitio \u2014 un placeholder genérico, no el nombre de la marca ni su oferta. El title de la home tampoco incorpora las búsquedas comerciales de cabecera (MBA, maestrías, posgrado) que la propia ESAN ya trabaja."}],
], "accent"));
body.push(sp(200));
body.push(p([{text:"El segundo es estructural: "},{text:"la autoridad del dominio está repartida entre varias propiedades distintas",bold:true},{text:", cada una acumulando enlaces y señales por separado en lugar de concentrarlas en las páginas de programa que deben rankear."}]));
body.push(sp(120));
body.push(tbl(["Propiedad", "Rol", "Lectura SEO"], [
  ["esan.edu.pe", "Sitio comercial: programas + blog Conexión", "Donde deben ganar las páginas de matrícula."],
  ["uesan.edu.pe", "Portal institucional y admisión (pa.uesan.edu.pe)", "Capta enlaces y marca por separado."],
  ["biblioteca.uesan.edu.pe / repositorio", "Biblioteca, tesis e investigación", "Genera autoridad que no llega al sitio comercial."],
  ["Dominio de pregrado", "Oferta de pregrado bajo otra marca", "Diluye la señal de \u201CESAN\u201D entre dominios."],
], [3100, 3900, 3206]));
body.push(sp(300));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "El volumen de tráfico es alto, pero su intención comercial es baja: la mayor parte no corresponde a personas buscando un posgrado. El contenido de Conexión ESAN es un activo de marca real; el desafío es canalizarlo hacia las páginas de programa (enlazado interno, llamados a la acción contextuales) y afinar el on-page de esas páginas para que conviertan la enorme audiencia informativa en prospectos de matrícula.",
], "red"));

// ===== FASE 03 — KEYWORDS =====
body.push(pb);
body.push(chip("FASE 03"));
body.push(h1("Keywords — Liderazgo con Brechas Comerciales"));
body.push(leadin("Cada búsqueda en Google es un cliente potencial levantando la mano. Esta sección separa el tráfico que ESAN ya capta del mercado comercial de posgrado que, teniendo los programas para atenderlo, todavía cede a un blog propio o a la competencia."));
body.push(h2("Métricas generales de posicionamiento"));
body.push(tbl(["Métrica", "Valor", "Análisis"], [
  ["Tráfico orgánico estimado (ETV)", { text: "173,844 visitas/mes", bold: true }, "Líder claro de la categoría educativa en Perú."],
  ["Keywords posicionadas", { text: "4,143", bold: true }, "Cobertura amplia, con fuerte sesgo informativo."],
  ["Keywords en Top 10", { text: "~1,907", bold: true }, "Base sólida; muchas son de blog, no de programa."],
  ["Keywords en posición #1", { text: "135", bold: true }, "Incluye marca y cabezas comerciales (mba)."],
], [3900, 2306, 3300]));
body.push(sp(300));
body.push(h2("Las búsquedas comerciales del negocio — dónde rankea hoy"));
body.push(caption("Posición actual de esan.edu.pe en Google (Perú) para sus búsquedas comerciales clave."));
body.push(tbl(["Búsqueda comercial", "Vol./mes", "Posición ESAN", "Lectura"], [
  ["mba", "6,600", { text: "#1", bold: true, color: B.green }, "Página /mba. Capturada."],
  ["maestrías", "6,600", { text: "#2", bold: true, color: B.green }, "Página /maestrias. Capturada."],
  ["maestría en finanzas", "880", { text: "#2", bold: true, color: B.green }, "Página de programa. Capturada."],
  ["educación ejecutiva", "390", { text: "#3", bold: true, color: B.green }, "Página de sección; pero viene cayendo."],
  ["maestría en gestión pública", "3,600", { text: "#6", bold: true, color: B.orange }, "A un paso del Top 5; Pacífico está #5."],
  ["especialización", "1,900", { text: "#9", bold: true, color: B.orange }, "Detrás de Centrum (#1) y UPC (#2)."],
  ["doctorado", "2,400", { text: "#12", bold: true, color: B.orange }, "Fuera del Top 10 pese a tener el programa."],
  ["diplomados", "4,400", { text: "#38", bold: true, color: B.red }, "El hub /diploma casi no posiciona; UPC #6."],
  ["posgrado", "4,400", { text: "#53", bold: true, color: B.red }, "Solo la home; sin landing dedicada; UPC #21."],
  ["maestría en marketing", "390", { text: "#3 (blog)", bold: true, color: B.orange }, "Rankea un artículo de Conexión, no el programa."],
], [3300, 1300, 2200, 2706]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"ESAN gana las cabezas que ya optimizó \u2014 está "},{text:"#1 en \u201Cmba\u201D y #2 en \u201Cmaestrías\u201D",bold:true},{text:" \u2014 pero deja valor sobre la mesa en el resto. En \u201Cposgrado\u201D (4,400/mes) solo aparece la home en la posición #53 mientras UPC está #21; en \u201Cdiplomados\u201D (4,400/mes) su hub está en #38 frente al #6 de UPC; y en \u201Cmaestría en marketing\u201D "},{text:"Google muestra un artículo de Conexión ESAN en lugar de la página del programa",bold:true},{text:" \u2014 una canibalización que envía a la persona a leer, no a postular."}],
], "accent"));
body.push(h2("Oportunidades de captura comercial"));
body.push(p([{text:"Búsquedas de intención comercial donde ESAN "},{text:"tiene el programa pero cede la posición",bold:true},{text:" \u2014 queda fuera del Top 5 o Google muestra un artículo en lugar de la página de matrícula. Son las brechas con el mayor volumen detrás."}]));
body.push(tbl(["Búsqueda", "Vol./mes", "Hoy", "Por qué importa capturarla"], [
  ["posgrado", "4,400", { text: "#53", bold: true, color: B.red }, "Término de referencia de la categoría; hoy solo aparece la home. UPC está #21."],
  ["diplomados", "4,400", { text: "#38", bold: true, color: B.red }, "Cartera amplia de diplomas; el hub apenas posiciona. UPC está #6."],
  ["maestría en gestión pública", "3,600", { text: "#6", bold: true, color: B.orange }, "Tiene el programa líder del país; a un paso del Top 5."],
  ["postgrado", "2,400", { text: "#28", bold: true, color: B.red }, "Variante de alto volumen, sin landing dedicada."],
  ["doctorado", "2,400", { text: "#12", bold: true, color: B.orange }, "Programa doctoral acreditado AACSB, fuera del Top 10."],
  ["especialización", "1,900", { text: "#9", bold: true, color: B.orange }, "Detrás de Centrum (#1) y UPC (#2)."],
  ["escuelas de posgrado", "1,600", { text: "#51", bold: true, color: B.red }, "Búsqueda de quien todavía está eligiendo escuela."],
  ["diplomado en recursos humanos", "720", { text: "#7", bold: true, color: B.orange }, "Cerca del Top 5; diplomas de gestión de personas."],
  ["maestría en ciencia de datos", "590", { text: "#6", bold: true, color: B.orange }, "Vertical en alza, fuera del Top 5."],
  ["maestría en marketing digital", "170", { text: "#8", bold: true, color: B.orange }, "El programa correcto, pero cayó de #4 a #8."],
], [3300, 1150, 1100, 4256]));
body.push(sp(220));
body.push(callout("OPORTUNIDAD EN TRÁFICO", [
  [{text:"En conjunto, estas búsquedas representan "},{text:"~22,180 búsquedas/mes",bold:true},{text:" de intención comercial donde ESAN tiene el programa pero no la posición. Llevarlas al Top 5 significa del orden de "},{text:"~1,500 visitas/mes adicionales de alta intención",bold:true},{text:" \u2014 prospectos evaluando un programa, no lectores de coyuntura. Es tráfico de mucho mayor valor que el informativo: a estas personas las separa una sola decisión de la matrícula. (Aparte quedan \u201Cmaestría en marketing\u201D e \u201Cinteligencia artificial\u201D, tratadas como casos de canibalización y de competencia directa.)"}],
], "green"));
body.push(h2("Cómo compite ESAN, búsqueda por búsqueda"));
body.push(caption("Posición en Google (Perú) en las 10 búsquedas comerciales que deciden una matrícula, frente a sus tres rivales directos. Verde = Top 3 · naranja = 4-10 · gris = fuera del Top 10 o sin posicionar."));
body.push(tbl(["Búsqueda", "Vol./mes", "ESAN", "Pacífico", "UPC", "Centrum"], [
  ["mba", "6,600", { text: "#1", bold: true, color: B.green }, { text: "\u2014", color: B.gray4 }, { text: "#4", color: B.orange }, { text: "#2", color: B.green }],
  ["maestrías", "6,600", { text: "#2", bold: true, color: B.green }, { text: "#50", color: B.gray4 }, { text: "#4", color: B.orange }, { text: "\u2014", color: B.gray4 }],
  ["maestría en finanzas", "880", { text: "#2", bold: true, color: B.green }, { text: "#11", color: B.gray4 }, { text: "#7", color: B.orange }, { text: "#4", color: B.orange }],
  ["maestría en gestión pública", "3,600", { text: "#6", bold: true, color: B.orange }, { text: "#5", color: B.green }, { text: "#7", color: B.orange }, { text: "\u2014", color: B.gray4 }],
  ["doctorado", "2,400", { text: "#12", bold: true, color: B.gray4 }, { text: "\u2014", color: B.gray4 }, { text: "\u2014", color: B.gray4 }, { text: "\u2014", color: B.gray4 }],
  ["especialización", "1,900", { text: "#9", bold: true, color: B.orange }, { text: "#13", color: B.gray4 }, { text: "#2", color: B.green }, { text: "#1", color: B.green }],
  ["maestría en inteligencia artificial", "880", { text: "#5", bold: true, color: B.orange }, { text: "\u2014", color: B.gray4 }, { text: "#1", color: B.green }, { text: "\u2014", color: B.gray4 }],
  ["maestría en marketing", "390", { text: "#3 blog", bold: true, color: B.orange }, { text: "#63", color: B.gray4 }, { text: "\u2014", color: B.gray4 }, { text: "#4", color: B.orange }],
  ["diplomados", "4,400", { text: "#38", bold: true, color: B.gray4 }, { text: "#49", color: B.gray4 }, { text: "#6", color: B.orange }, { text: "\u2014", color: B.gray4 }],
  ["posgrado", "4,400", { text: "#53", bold: true, color: B.gray4 }, { text: "#63", color: B.gray4 }, { text: "#21", color: B.gray4 }, { text: "\u2014", color: B.gray4 }],
], [3100, 1250, 1250, 1350, 1100, 1456]));
body.push(sp(300));
body.push(caption("En cuántas de esas 10 búsquedas comerciales cada escuela aparece en Top 5 y en Top 10."));
body.push(img("visibilidad.png", 540, 0.354));
body.push(sp(160));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"El liderazgo de ESAN es real pero "},{text:"desigual",bold:true},{text:". Domina las cabezas de marca \u2014 #1 en \u201Cmba\u201D, #2 en \u201Cmaestrías\u201D y #2 en \u201Cfinanzas\u201D \u2014 donde nadie la alcanza. Pero pierde las búsquedas genéricas de alto volumen, las del que aún no eligió escuela: en \u201Cposgrado\u201D y \u201Cdiplomados\u201D (4,400/mes cada una) "},{text:"UPC rankea muy por encima",bold:true},{text:" (#21 y #6 vs. #53 y #38), y en \u201Cespecialización\u201D va detrás de Centrum (#1) y UPC (#2). En \u201Cinteligencia artificial\u201D, un vertical en alza, "},{text:"UPC está #1 y ESAN #5",bold:true},{text:". Medido por presencia comercial, UPC iguala a ESAN (ambas en 7 de 10 búsquedas en Top 10) y la supera justo donde se capta al prospecto indeciso."}],
], "accent"));
body.push(h2("Por qué la comparación NO es de tráfico total"));
body.push(p([{text:"Es tentador comparar el tráfico total de cada dominio, pero "},{text:"sería engañoso",bold:true},{text:": UPC y Pacífico son universidades completas (pregrado + posgrado + admisión), mientras que ESAN y Centrum son escuelas de posgrado. El tráfico total de las dos primeras incluye carreras, admisión y vida universitaria que nada tienen que ver con el negocio de ESAN. Por eso la comparación que importa es la de arriba \u2014 posición por posición en las mismas búsquedas \u2014 y no la del volumen agregado."}]));
body.push(sp(120));
body.push(tbl(["Dominio", "Tipo de institución", "Keywords", "Tráfico orgánico"], [
  ["upc.edu.pe", "Universidad (pregrado + posgrado)", "6,557", "946,675 visitas/mes"],
  ["up.edu.pe (Pacífico)", "Universidad (pregrado + posgrado)", "2,546", "388,562 visitas/mes"],
  ["esan.edu.pe", "Escuela de posgrado", "4,143", { text: "173,844 visitas/mes", bold: true, color: B.accent }],
  ["centrum.pucp.edu.pe", "Escuela de posgrado", "169", "15,420 visitas/mes"],
], [2700, 3400, 1500, 2406]));
body.push(sp(300));
body.push(h2("La posición competitiva de un vistazo"));
body.push(tbl(["Dimensión", "Situación actual", "Estado", "Lectura"], [
  ["Volumen de tráfico", "173,844 visitas/mes", OK, "Líder de la categoría educativa."],
  ["Calidad del tráfico", "Mayoría informativo", WARN, "Conexión ESAN > páginas de programa."],
  ["Velocidad (LCP home)", "6.0 s", FAIL, "~2.4x la meta de 2.5 s."],
  ["Velocidad (página MBA)", "5.1 s · CLS 0.22", FAIL, "La página de conversión también es lenta."],
  ["Autoridad de dominio", "~478 dominios ref.", OK, "Sólida; poco canalizada a programas."],
  ["Comerciales de cabecera", "#1 mba · #2 maestrías · #2 finanzas", OK, "Domina las cabezas de marca; nadie la alcanza."],
  ["Genéricas de alto volumen", "Posgrado #53 · Diplomados #38", WARN, "UPC rankea muy por encima (#21 · #6)."],
  ["Verticales emergentes", "IA #5 · marketing rankea un blog", WARN, "UPC lidera IA (#1); canibalización en marketing."],
], [2700, 2500, 1600, 2706]));
body.push(sp(300));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "ESAN no tiene un problema de presencia, sino de eficiencia comercial: lidera su categoría, pero parte de su enorme tráfico es informativo de bajo valor y varias búsquedas de matrícula las cede o las responde con un blog. Llevar las páginas de programa al Top 5 en posgrado, diplomados y las especializadas \u2014 y canalizar hacia ellas la autoridad que ya tiene \u2014 es la palanca de mayor impacto comercial del documento.",
], "red"));

// ===== FASE 04 — BACKLINKS =====
body.push(pb);
body.push(chip("FASE 04"));
body.push(h1("Backlinks — Autoridad Sólida y sin Canalizar"));
body.push(leadin("Los enlaces de otros sitios funcionan como recomendaciones: mientras más sitios relevantes enlacen, más confía Google. Esta sección mide la autoridad que ESAN ha acumulado y si esa fuerza está empujando a las páginas correctas."));
body.push(h2("Perfil de autoridad del dominio"));
body.push(tbl(["Métrica", "Valor", "Análisis"], [
  ["Dominios referentes (home)", { text: "~478", bold: true }, "Autoridad sólida para una institución educativa."],
  ["Backlinks totales (home)", { text: "~37,170", bold: true }, "Volumen alto acumulado en décadas."],
  ["Enlaces dofollow", { text: "~32,166 (87%)", bold: true }, "Buena proporción que transmite autoridad."],
  ["Domain Rank (DataForSEO, 0\u20131000)", { text: "~389", bold: true }, "Dominio de autoridad alta."],
], [3906, 2300, 3300]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"La autoridad de ESAN es alta y bien ganada. El punto de mejora no es la cantidad, sino "},{text:"hacia dónde apunta",bold:true},{text:": la fuerza se concentra en la home y en Conexión ESAN, mientras varias páginas de programa tienen muy pocos enlaces propios (la página /mba tiene ~5 dominios referentes; /maestrias y /doctorado, apenas 1). Esa autoridad no está empujando a las páginas que compiten por la matrícula."}],
], "accent"));
body.push(sp(200));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "ESAN ya tiene la autoridad que muchas escuelas quisieran; el trabajo es redistribuirla. Un enlazado interno deliberado desde los artículos de Conexión ESAN \u2014 que reciben el grueso de los enlaces externos \u2014 hacia las páginas de MBA, maestrías especializadas, posgrado y diplomados haría que esa fuerza ya acumulada empuje justo donde hoy hay brechas de posición.",
], "red"));

// ===== FASE 05 — AI OVERVIEWS =====
body.push(pb);
body.push(chip("FASE 05"));
body.push(h1("AI Overviews — La Base Informativa en Riesgo"));
body.push(leadin("Google ya responde muchas búsquedas con un resumen generado por IA, antes de mostrar los resultados. Si el sitio no es una de las fuentes que esa IA cita, su clic desaparece \u2014 aunque la página esté bien posicionada. Para ESAN esto es delicado: justamente su mayor fuente de tráfico es el contenido informativo más fácil de absorber por la IA."));
body.push(h2("Búsquedas con AI Overview activo que hoy dan tráfico a ESAN"));
body.push(tbl(["Búsqueda", "Vol./mes", "Tipo", "AI Overview"], [
  ["marcas", "110,000", "Conexión ESAN", { text: "Activo", bold: true, color: B.red }],
  ["sueldo mínimo peru", "90,500", "Conexión ESAN", { text: "Activo", bold: true, color: B.red }],
  ["valor económico agregado", "60,500", "Conexión ESAN", { text: "Activo", bold: true, color: B.red }],
  ["contratos", "33,100", "Conexión ESAN", { text: "Activo", bold: true, color: B.red }],
  ["design thinking", "33,100", "Conexión ESAN", { text: "Activo", bold: true, color: B.red }],
  ["mba", "6,600", "Programa", { text: "Activo", bold: true, color: B.red }],
  ["posgrado", "4,400", "Comercial", { text: "Activo", bold: true, color: B.red }],
  ["diplomados", "4,400", "Comercial", { text: "Activo", bold: true, color: B.red }],
], [3906, 1700, 2200, 1700]));
body.push(sp(300));
body.push(callout("QUÉ SIGNIFICA PARA EL NEGOCIO", [
  [{text:"Buena parte del tráfico estrella de ESAN llega por búsquedas que "},{text:"ya tienen un resumen de IA arriba",bold:true},{text:": marcas, sueldo mínimo, valor económico agregado, design thinking. A medida que Google responda más de estas consultas con su propia IA, ese volumen informativo \u2014 el que hoy sostiene buena parte de las 173,844 visitas \u2014 se erosiona, sin importar la posición de ESAN."}],
  [{text:"En las búsquedas comerciales también aparece el AI Overview (mba, posgrado, diplomados). Y cuando alguien compara escuelas (\u201Cmejores mba perú\u201D), Google muestra "},{text:"agregadores y rankings de terceros",bold:true},{text:", con ESAN representada por artículos de Conexión \u2014 no por su página de MBA."}],
], "accent"));
body.push(sp(200));
body.push(callout("IMPACTO ESTIMADO DE LA FASE", [
  "El tráfico informativo de Conexión ESAN \u2014 la mayor fuente de visitas del dominio \u2014 es justamente el más expuesto a los AI Overviews. Es la razón estratégica para no depender de él: reorientar el esfuerzo hacia las búsquedas comerciales de posgrado (menos absorbidas por la IA y de mucho mayor valor de matrícula) y asegurar que, cuando la IA cite fuentes en esas búsquedas, ESAN sea una de ellas, protege el negocio a futuro.",
], "red"));

// ===== CONCLUSIONES =====
body.push(pb);
body.push(h1("Conclusiones y Plan de Acción"));
body.push(h2("Principales hallazgos priorizados"));
body.push(tbl(["#", "Hallazgo", "Prioridad", "Impacto", "Plazo"], [
  ["1", "Contenido informativo (mayor fuente de tráfico) expuesto a AI Overviews", FAIL, { text: "Crítico", color: B.red, bold: true }, "Continuo"],
  ["2", "Páginas de programa lentas (home 24 MB; MBA LCP 5.1 s, CLS 0.22)", FAIL, { text: "Alto", color: B.orange, bold: true }, "1-3 meses"],
  ["3", "Posgrado (#53) y diplomados (#38) sin landing comercial posicionada", FAIL, { text: "Alto", color: B.orange, bold: true }, "1-3 meses"],
  ["4", "Canibalización: un blog rankea por \u201Cmaestría en marketing\u201D, no el programa", WARN, { text: "Medio", color: B.orange, bold: true }, "1-2 meses"],
  ["5", "Autoridad de dominio no canalizada hacia páginas de programa", WARN, { text: "Medio", color: B.orange, bold: true }, "2-4 semanas"],
], [600, 4406, 1700, 1400, 1400]));
body.push(sp(300));
body.push(callout("PLAN DE ACCIÓN POR FASES", [
  [{text:"Mes 1-3 (Quick wins): ",bold:true},{text:"optimizar el peso y la velocidad de la home (24 MB) y de las páginas de programa \u2014 objetivo LCP < 2.5 s y CLS < 0.10 en /mba y maestrías \u2014, corregir el og:title \u201CHome Esan\u201D y el title de la home, y resolver la canibalización en \u201Cmaestría en marketing\u201D (que rankee la página de programa, no el artículo)."}],
  [{text:"Mes 4-6 (Captura comercial): ",bold:true},{text:"crear/optimizar landings para posgrado, diplomados y las especializadas donde hoy hay brecha, y canalizar la autoridad del dominio mediante enlazado interno desde Conexión ESAN hacia esas páginas. Meta: ~1,500 visitas/mes adicionales de intención comercial."}],
  [{text:"Mes 7-12 (Blindaje y escala): ",bold:true},{text:"reducir la dependencia del tráfico informativo expuesto a IA, reforzar contenido de decisión (comparativas de programas, rankings, ROI) y asegurar presencia de ESAN como fuente citada en los AI Overviews de las búsquedas de posgrado."}],
  [{text:"Siguiente paso: ",bold:true},{text:"reunión de 30-45 min para revisar hallazgos, propuesta de roadmap priorizado y definición de las primeras acciones (velocidad de páginas de programa y captura de posgrado/diplomados)."}],
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
        new TextRun({ text: "   ·   Auditoría SEO & Performance   ·   esan.edu.pe", font: B.font, size: 13, color: B.gray3 }) ] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [ new TextRun({ text: "Pág. ", font: B.font, size: 15, color: B.gray3 }),
        new TextRun({ children: [PageNumber.CURRENT], font: B.font, size: 15, color: B.gray3 }),
        new TextRun({ text: "   ·   Documento confidencial", font: B.font, size: 13, color: "CCCCCC" }) ] })] }) },
    children: spacerGuard([...cover, pb, ...body]),
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync("auditoria_esan.docx", buf); console.log("OK docx", buf.length); });
