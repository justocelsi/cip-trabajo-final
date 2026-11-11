const pptxgen = require("pptxgenjs");

// Uso: npm install && node generar.js [salida.pptx]
const OUT = process.argv[2] || "borrador.pptx";

// Celeste y azul para Malvinas, tierra para la Campaña del Desierto, el sol como acento.
const THEME = {
  name: "Malvinas y Desierto",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B2433", lt1: "FFFFFF", dk2: "16324F", lt2: "EEF4FA",
    accent1: "4A8CC7", accent2: "A8502F", accent3: "F2B33D", accent4: "74ACDF",
    accent5: "F6E7DF", accent6: "6B7280", hlink: "2F6FA8", folHlink: "6B4E9B",
  },
};
const HEX = { celeste: "74ACDF", azul: "16324F", tierra: "A8502F", tierraSuave: "F6E7DF", celesteSuave: "EEF4FA", sol: "F2B33D", gris: "6B7280", texto: "1B2433" };

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "Malvinas y el «desierto»: dos relatos de la identidad argentina";
pres.author = "Justo Celsi";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };

const FOOT = "Cultura, Identidad y Política · ITBA · 2.º cuatrimestre de 2026";

pres.defineSlideMaster({
  title: "PORTADA",
  background: { color: HEX.azul },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 2.25, w: 8.8, h: 1.1, fontSize: 40, bold: true, color: "FFFFFF", valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 3.35, w: 8.8, h: 0.6, fontSize: 22, color: HEX.celeste, valign: "top", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENIDO",
  background: { color: "FFFFFF" },
  margin: [0.5, 0.5, 0.6, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontSize: 28, bold: true, color: HEX.azul, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.5, y: 5.2, w: 7.5, h: 0.3, fontSize: 10, color: HEX.gris, margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.5, h: 0.3, fontSize: 10, color: HEX.gris, align: "right" },
});
pres.defineSlideMaster({
  title: "CIERRE",
  background: { color: HEX.azul },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.4, w: 8.8, h: 0.8, fontSize: 30, bold: true, color: "FFFFFF", valign: "middle", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 5.2, w: 7.5, h: 0.3, fontSize: 10, color: HEX.celeste, margin: 0 } } },
  ],
});

function bandera(slide, x, y, w) {
  const h = w * 0.62, f = h / 3;
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: f, fill: { color: HEX.celeste }, line: { color: HEX.celeste }, objectName: "Bandera franja superior" });
  slide.addShape(pres.shapes.RECTANGLE, { x, y: y + f, w, h: f, fill: { color: "FFFFFF" }, line: { color: "FFFFFF" }, objectName: "Bandera franja blanca" });
  slide.addShape(pres.shapes.RECTANGLE, { x, y: y + 2 * f, w, h: f, fill: { color: HEX.celeste }, line: { color: HEX.celeste }, objectName: "Bandera franja inferior" });
  const s = f * 0.8;
  slide.addShape(pres.shapes.OVAL, { x: x + w / 2 - s / 2, y: y + f + (f - s) / 2, w: s, h: s, fill: { color: HEX.sol }, line: { color: "C8891A", width: 0.75 }, objectName: "Bandera sol" });
}

function chip(slide, x, y, label, color, txt = "FFFFFF") {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 1.55, h: 0.36, rectRadius: 0.18, fill: { color }, line: { color }, objectName: "Etiqueta " + label });
  slide.addText(label, { x, y, w: 1.55, h: 0.36, fontSize: 14, bold: true, color: txt, align: "center", valign: "middle", margin: 0, isTextBox: true });
}

function stat(slide, x, y, w, big, small, color) {
  slide.addText(big, { x, y, w, h: 0.75, fontFace: "Cambria", fontSize: 36, bold: true, color, margin: 0, valign: "bottom", isTextBox: true });
  slide.addText(small, { x, y: y + 0.8, w, h: 0.6, fontSize: 14, color: HEX.texto, margin: 0, valign: "top", isTextBox: true });
}

// 1. Portada
pres.addSection({ title: "Portada" });
{
  const s = pres.addSlide({ masterName: "PORTADA", sectionTitle: "Portada" });
  bandera(s, 0.6, 0.6, 1.9);
  s.addText("Malvinas y el «desierto»", { placeholder: "title" });
  s.addText("Dos relatos de la identidad argentina", { placeholder: "body" });
  s.addText([
    { text: "Eje 1: cultura, identidad y poder", options: { breakLine: true } },
    { text: "Justo Celsi, Santi y Santi · miércoles 25 de noviembre de 2026" },
  ], { x: 0.6, y: 4.3, w: 8.8, h: 0.8, fontSize: 16, color: "FFFFFF", margin: 0, paraSpaceAfter: 4, isTextBox: true });
  s.addNotes("Presentamos el trabajo final. Elegimos el eje de cultura, identidad y poder, y dos relatos que cualquier argentino conoce desde la escuela: Malvinas y la Campaña del Desierto.");
}

// 2. El caso y la pregunta
pres.addSection({ title: "Caso y pregunta" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Caso y pregunta" });
  s.addText("Dos relatos, un mismo territorio del sur", { placeholder: "title" });
  const hitos = [
    ["1833", "Gran Bretaña ocupa las Malvinas", HEX.azul],
    ["1878–1885", "Campaña del Desierto", HEX.tierra],
    ["1979", "La dictadura celebra su centenario", HEX.tierra],
    ["1982", "Guerra de Malvinas", HEX.azul],
    ["1994", "La Constitución reclama las islas y reconoce a los pueblos indígenas", HEX.gris],
    ["2026", "La selección muestra la bandera en el Mundial", HEX.azul],
  ];
  const x0 = 0.5, w = 9.0, n = hitos.length, step = w / n, ly = 1.75;
  s.addShape(pres.shapes.LINE, { x: x0 + step / 2, y: ly, w: step * (n - 1), h: 0, line: { color: "C9D3DE", width: 2 }, objectName: "Línea de tiempo" });
  hitos.forEach(([anio, txt, col], i) => {
    const cx = x0 + step * i + step / 2;
    s.addShape(pres.shapes.OVAL, { x: cx - 0.13, y: ly - 0.13, w: 0.26, h: 0.26, fill: { color: col }, line: { color: "FFFFFF", width: 1.5 }, objectName: "Hito " + anio });
    s.addText(anio, { x: cx - step / 2 + 0.05, y: 1.15, w: step - 0.1, h: 0.4, fontFace: "Cambria", fontSize: 16, bold: true, color: col, align: "center", margin: 0, isTextBox: true });
    s.addText(txt, { x: cx - step / 2 + 0.05, y: 2.0, w: step - 0.1, h: 1.0, fontSize: 13, color: HEX.texto, align: "center", valign: "top", margin: 0, isTextBox: true });
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 3.25, w: 9.0, h: 1.75, rectRadius: 0.12, fill: { color: HEX.celesteSuave }, line: { color: HEX.celesteSuave }, objectName: "Caja de la pregunta" });
  s.addText("Nuestra pregunta", { x: 0.8, y: 3.4, w: 8.4, h: 0.35, fontSize: 14, bold: true, color: HEX.azul, margin: 0, isTextBox: true });
  s.addText("¿Cómo dos relatos tan memorables para los argentinos, Malvinas y la Campaña del Desierto, hacen a la identidad argentina y a la vez son contradictorios?", { x: 0.8, y: 3.8, w: 8.4, h: 1.05, fontFace: "Cambria", fontSize: 20, italic: true, color: HEX.texto, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("Los dos relatos hablan del mismo sur. En 1833 Gran Bretaña ocupa las Malvinas; cincuenta años después el Estado conquista la Patagonia continental. En 1979 la dictadura celebra el centenario de esa conquista y en 1982 va a la guerra por las islas. La pregunta que guía el trabajo es cómo los dos relatos arman la identidad argentina y en qué se contradicen.");
}

// 3. Malvinas
pres.addSection({ title: "Los dos casos" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Los dos casos" });
  s.addText("Malvinas: la causa que une a casi todos", { placeholder: "title" });
  chip(s, 0.5, 1.15, "Malvinas", HEX.azul);
  stat(s, 0.5, 1.6, 2.8, "1833", "Gran Bretaña ocupa las islas y desaloja a las autoridades argentinas", HEX.azul);
  stat(s, 3.6, 1.6, 2.8, "74 días", "dura la guerra de 1982, del 2 de abril al 14 de junio", HEX.azul);
  stat(s, 6.7, 1.6, 2.8, "649", "soldados argentinos muertos, muchos de ellos conscriptos de 18 y 19 años", HEX.azul);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.35, w: 9.0, h: 1.6, fill: { color: HEX.celesteSuave }, line: { color: HEX.celesteSuave }, objectName: "Caja de contexto Malvinas" });
  s.addText([
    { text: "La guerra la decidió una dictadura en crisis. Tres días antes había reprimido una marcha en Plaza de Mayo, y el 2 de abril la plaza se llenó para apoyarla.", options: { bullet: true, breakLine: true } },
    { text: "Después vinieron la «desmalvinización» y el abandono de los excombatientes. Hoy el reclamo está en la Constitución, en el feriado del 2 de abril y en el mapa escolar.", options: { bullet: true, breakLine: true } },
    { text: "En 2013 los isleños votaron seguir siendo británicos.", options: { bullet: true } },
  ], { x: 0.7, y: 3.45, w: 8.6, h: 1.4, fontSize: 14, color: HEX.texto, margin: 0, paraSpaceAfter: 6, valign: "top", isTextBox: true });
  s.addNotes("Malvinas tiene tres fechas: la ocupación de 1833, la guerra de 1982 y el reclamo que sigue hoy. La guerra la hizo una dictadura que buscaba legitimidad. Después de la derrota hubo años de silencio, y hoy es casi la única causa que comparten todos los sectores políticos.");
}

// 4. Campaña del Desierto
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Los dos casos" });
  s.addText("La Campaña del Desierto: una memoria en disputa", { placeholder: "title" });
  chip(s, 0.5, 1.15, "Desierto", HEX.tierra);
  stat(s, 0.5, 1.6, 2.8, "1879", "el ejército de Roca llega al río Negro, la nueva frontera", HEX.tierra);
  stat(s, 3.6, 1.6, 2.8, "1885", "se rinde Sayhueque, uno de los últimos caciques en resistir", HEX.tierra);
  stat(s, 6.7, 1.6, 2.8, "1994", "la Constitución reconoce la preexistencia de los pueblos indígenas", HEX.tierra);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.35, w: 9.0, h: 1.6, fill: { color: HEX.tierraSuave }, line: { color: HEX.tierraSuave }, objectName: "Caja de contexto Desierto" });
  s.addText([
    { text: "Lo que se llamó «desierto» estaba habitado por mapuche, tehuelche, ranqueles y otros pueblos.", options: { bullet: true, breakLine: true } },
    { text: "Los prisioneros pasaron por la isla Martín García y muchos fueron repartidos como servicio doméstico o mano de obra.", options: { bullet: true, breakLine: true } },
    { text: "La tierra conquistada hizo posible la Argentina exportadora, que le vendía sobre todo a Gran Bretaña. Los monumentos a Roca se pintan y se discuten.", options: { bullet: true } },
  ], { x: 0.7, y: 3.45, w: 8.6, h: 1.4, fontSize: 14, color: HEX.texto, margin: 0, paraSpaceAfter: 6, valign: "top", isTextBox: true });
  s.addNotes("La Campaña del Desierto fue la ocupación militar de la Pampa y la Patagonia. El nombre ya dice algo: se lo llamó desierto, pero vivían pueblos indígenas. Hubo muertos, prisioneros y familias separadas. A diferencia de Malvinas, este relato no une: hoy se discute si fue un genocidio.");
}

// 5. Enfoque teórico
pres.addSection({ title: "Enfoque teórico" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Enfoque teórico" });
  s.addText("Con qué conceptos leemos los dos casos", { placeholder: "title" });
  const cards = [
    ["Stuart Hall (2003)", "La identidad se construye contra un afuera, y excluir es un acto de poder"],
    ["Guillermo O'Donnell (1993)", "El Estado es también la ley que llega, o no, a todo el territorio"],
    ["Eric Wolf (1982)", "Nombrar es poder: puede negar que existan otros"],
    ["Georges Balandier (1994)", "El poder se sostiene poniéndose en escena"],
    ["McCombs y Evatt (1995)", "Los medios influyen en qué temas importan y en cómo se los piensa"],
    ["Calvo y Aruguete (2020)", "En las redes se comparte lo que confirma lo que ya se cree"],
  ];
  const cw = 2.85, ch = 1.25, gx = 0.225, gy = 0.2, x0 = 0.5, y0 = 1.2;
  cards.forEach(([a, t], i) => {
    const col = i % 3, row = Math.floor(i / 3);
    const x = x0 + col * (cw + gx), y = y0 + row * (ch + gy);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cw, h: ch, rectRadius: 0.1, fill: { color: HEX.celesteSuave }, line: { color: HEX.celesteSuave }, objectName: "Tarjeta " + a });
    s.addText(a, { x: x + 0.15, y: y + 0.1, w: cw - 0.3, h: 0.35, fontSize: 14, bold: true, color: HEX.azul, margin: 0, isTextBox: true });
    s.addText(t, { x: x + 0.15, y: y + 0.45, w: cw - 0.3, h: 0.75, fontSize: 14, color: HEX.texto, margin: 0, valign: "top", isTextBox: true });
  });
  s.addText([
    { text: "Fuentes académicas adicionales: ", options: { bold: true, color: HEX.tierra } },
    { text: "Foucault (racismo de Estado), Guber y Lorenz (Malvinas), Delrio (la Patagonia después de la campaña)", options: { color: HEX.texto } },
  ], { x: 0.5, y: 4.2, w: 9.0, h: 0.75, fontSize: 14, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("Usamos seis textos de la materia y cuatro fuentes de afuera. Hall nos da la idea de que toda identidad necesita un afuera. O'Donnell distingue el Estado de la nación. Wolf muestra que nombrar es poder. Balandier explica el poder como escena. McCombs y Evatt, y Calvo y Aruguete, sirven para leer cómo los medios y las redes cuentan el conflicto.");
}

// 6. Hall
pres.addSection({ title: "Análisis" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Análisis" });
  s.addText("Dos «nosotros», dos afueras", { placeholder: "title" });
  const colW = 4.35;
  [[0.5, "Malvinas", HEX.azul, HEX.celesteSuave, "El afuera es el británico.", "Adentro entran todos, también los conscriptos de cualquier provincia."],
   [5.15, "Desierto", HEX.tierra, HEX.tierraSuave, "El afuera es el indígena, el «salvaje».", "Quedan afuera los pueblos que vivían en el territorio que se incorpora."]]
    .forEach(([x, lab, col, bg, a, b]) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.2, w: colW, h: 2.2, rectRadius: 0.1, fill: { color: bg }, line: { color: bg }, objectName: "Columna " + lab });
      chip(s, x + 0.2, 1.35, lab, col);
      s.addText([
        { text: a, options: { bullet: true, breakLine: true } },
        { text: b, options: { bullet: true } },
      ], { x: x + 0.2, y: 1.85, w: colW - 0.4, h: 1.45, fontSize: 16, color: HEX.texto, margin: 0, paraSpaceAfter: 8, valign: "top", isTextBox: true });
    });
  s.addText("«las identidades pueden funcionar como puntos de identificación y adhesión sólo debido a su capacidad de excluir, de omitir, de dejar “afuera”»", { x: 0.5, y: 3.6, w: 9.0, h: 0.95, fontFace: "Cambria", fontSize: 17, italic: true, color: HEX.azul, margin: 0, valign: "top", isTextBox: true });
  s.addText("Stuart Hall (2003), pp. 18–19", { x: 0.5, y: 4.6, w: 9.0, h: 0.3, fontSize: 12, color: HEX.gris, margin: 0, isTextBox: true });
  s.addNotes("Hall dice que la identidad no es una esencia: se arma contra un afuera. En Malvinas ese afuera es el británico y el nosotros incluye a todos. En el Desierto el afuera es el indígena, y el nosotros se arma dejándolo fuera.");
}

// 7. O'Donnell y Wolf
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Análisis" });
  s.addText("¿Cuenta la gente o cuenta el territorio?", { placeholder: "title" });
  const rows = [
    ["Malvinas", HEX.azul, HEX.celesteSuave, "Argentina pide tener en cuenta los «intereses» de los isleños, no sus deseos: manda el territorio. Así lo dice la Resolución 2065 de la ONU, de 1965."],
    ["Desierto", HEX.tierra, HEX.tierraSuave, "Para el Estado no había habitantes: lo llamó «desierto». Wolf escribe que nombrar lleva «la facultad de negar la existencia de categorías alternas» (pp. 468–469)."],
    ["Los dos", HEX.gris, "F1F2F4", "En 1833 el Estado tampoco controlaba la Patagonia continental. O'Donnell: el Estado es también la ley que llega, o no, a todo el territorio y a todos."],
  ];
  rows.forEach(([lab, col, bg, t], i) => {
    const y = 1.2 + i * 1.25;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 9.0, h: 1.1, fill: { color: bg }, line: { color: bg }, objectName: "Fila " + lab });
    chip(s, 0.7, y + 0.37, lab, col);
    s.addText(t, { x: 2.5, y: y + 0.1, w: 6.8, h: 0.9, fontSize: 15, color: HEX.texto, margin: 0, valign: "middle", isTextBox: true });
  });
  s.addNotes("En Malvinas la posición argentina pone el territorio por encima de la voluntad de los isleños. En el Desierto el Estado directamente negó que hubiera habitantes. Wolf ayuda a ver que llamarlo desierto ya es un acto de poder, y O'Donnell, que el Estado es la ley que efectivamente llega al territorio.");
}

// 8. La bandera y la prensa inglesa
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Análisis" });
  s.addText("La bandera en la cancha y la respuesta inglesa", { placeholder: "title" });
  bandera(s, 0.5, 1.25, 1.1);
  s.addText("Mundial 2026", { x: 1.8, y: 1.25, w: 3.0, h: 0.35, fontSize: 14, bold: true, color: HEX.azul, margin: 0, isTextBox: true });
  s.addText("Después de ganarle 2 a 1 a Inglaterra, el 15 de julio en Atlanta, los jugadores mostraron una bandera que decía «Las Malvinas son argentinas». Del lado británico se pidió que la FIFA los sancionara.", { x: 0.5, y: 2.0, w: 4.3, h: 1.6, fontSize: 14, color: HEX.texto, margin: 0, valign: "top", isTextBox: true });
  s.addText([
    { text: "McCombs y Evatt: ", options: { bold: true, color: HEX.azul } },
    { text: "los medios no sólo eligen el tema, también el modo en que se lo piensa.", options: { breakLine: true } },
    { text: "Calvo y Aruguete: ", options: { bold: true, color: HEX.azul } },
    { text: "en las redes cada comunidad comparte lo que confirma su posición." },
  ], { x: 0.5, y: 3.65, w: 4.3, h: 1.35, fontSize: 14, color: HEX.texto, margin: 0, paraSpaceAfter: 6, valign: "top", isTextBox: true });
  [1.2, 3.15].forEach((y, i) => {
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.2, y, w: 4.3, h: 1.8, rectRadius: 0.1, fill: { color: "F7F9FB" }, line: { color: "C9D3DE", width: 1, dashType: "dash" }, objectName: "Lugar para captura de X " + (i + 1) });
  });
  s.addNotes("En el Mundial 2026, después de ganarle a Inglaterra, los jugadores mostraron una bandera que decía «Las Malvinas son argentinas», y del lado británico se pidió que la FIFA los sancionara. El mismo hecho se cuenta de dos maneras: es lo que McCombs y Evatt llaman la segunda dimensión de la agenda. En las capturas de X se ve cómo cada lado comparte la versión que confirma lo que ya piensa.");
}

// 9. Balandier
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Análisis" });
  s.addText("El poder se pone en escena", { placeholder: "title" });
  const esc = [
    ["1979", "La dictadura celebra el centenario de la Conquista del Desierto", HEX.tierra],
    ["2 de abril de 1982", "La Plaza de Mayo, que tres días antes había sido reprimida, se llena para apoyar la guerra", HEX.azul],
    ["Hoy", "El acto escolar del 2 de abril, y los monumentos a Roca, pintados y defendidos", HEX.gris],
  ];
  esc.forEach(([t, d, col], i) => {
    const x = 0.5 + i * 3.075, w = 2.85;
    s.addShape(pres.shapes.OVAL, { x: x, y: 1.3, w: 0.5, h: 0.5, fill: { color: col }, line: { color: col }, objectName: "Escena " + (i + 1) });
    s.addText(String(i + 1), { x, y: 1.3, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: "FFFFFF", align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(t, { x, y: 1.95, w, h: 0.4, fontSize: 16, bold: true, color: col, margin: 0, isTextBox: true });
    s.addText(d, { x, y: 2.4, w, h: 1.1, fontSize: 14, color: HEX.texto, margin: 0, valign: "top", isTextBox: true });
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.7, w: 9.0, h: 1.3, fill: { color: HEX.celesteSuave }, line: { color: HEX.celesteSuave }, objectName: "Cita de Balandier" });
  s.addText("«La sociedad no se “sostiene” solamente por la coerción, ni por legitimar relaciones de fuerza, sino por el conjunto de transfiguraciones de las que es, a un mismo tiempo, objeto y ejecutora»", { x: 0.7, y: 3.8, w: 8.6, h: 0.85, fontFace: "Cambria", fontSize: 15, italic: true, color: HEX.azul, margin: 0, valign: "top", isTextBox: true });
  s.addText("Georges Balandier (1994), p. 41", { x: 0.7, y: 4.65, w: 8.6, h: 0.3, fontSize: 12, color: HEX.gris, margin: 0, isTextBox: true });
  s.addNotes("Balandier dice que el poder no se sostiene sólo con la fuerza, sino poniéndose en escena. Tres escenas: la dictadura celebrando el centenario en 1979, la plaza llena el 2 de abril de 1982 y, hoy, el acto escolar frente a los monumentos a Roca que se pintan.");
}

// 10. Conclusión
pres.addSection({ title: "Conclusión" });
{
  const s = pres.addSlide({ masterName: "CIERRE", sectionTitle: "Conclusión" });
  s.addText("Volvemos a la pregunta", { placeholder: "title" });
  s.addText("¿Cómo Malvinas y la Campaña del Desierto hacen a la identidad argentina y a la vez son contradictorios?", { x: 0.6, y: 1.25, w: 8.8, h: 0.8, fontFace: "Cambria", fontSize: 18, italic: true, color: HEX.celeste, margin: 0, valign: "top", isTextBox: true });
  const qs = [
    "Si toda identidad se arma contra un afuera, ¿la contradicción está en los relatos o en cómo funciona la identidad?",
    "¿Por qué un reclamo une a casi todos y el otro sigue en disputa?",
    "¿Qué cambia si la historia de la Patagonia se cuenta desde los pueblos que vivían ahí?",
  ];
  qs.forEach((q, i) => {
    const y = 2.25 + i * 0.95;
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.1, w: 0.45, h: 0.45, fill: { color: HEX.sol }, line: { color: HEX.sol }, objectName: "Tensión " + (i + 1) });
    s.addText(String(i + 1), { x: 0.6, y: y + 0.1, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: HEX.azul, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(q, { x: 1.3, y, w: 8.1, h: 0.75, fontSize: 16, color: "FFFFFF", margin: 0, valign: "middle", isTextBox: true });
  });
  s.addNotes("Volvemos a la pregunta y respondemos con lo que encontramos. Cerramos con las tres tensiones que quedan abiertas.");
}

// 11. Bibliografía
pres.addSection({ title: "Bibliografía" });
{
  const s = pres.addSlide({ masterName: "CONTENIDO", sectionTitle: "Bibliografía" });
  s.addText("Bibliografía", { placeholder: "title" });
  const obl = [
    "Balandier, G. (1994). El poder en escenas. Barcelona: Paidós.",
    "Calvo, E. y Aruguete, N. (2020). Fake news, trolls y otros encantos. Buenos Aires: Siglo XXI.",
    "Hall, S. (2003). Introducción: ¿quién necesita «identidad»? En S. Hall y P. du Gay (comps.), Cuestiones de identidad cultural (pp. 13–39). Buenos Aires: Amorrortu.",
    "McCombs, M. y Evatt, D. (1995). Los temas y los aspectos: explorando una nueva dimensión de la agenda setting. Comunicación y Sociedad, 8(1), 7–32.",
    "O'Donnell, G. (1993). Acerca del Estado, la democratización y algunos problemas conceptuales. Desarrollo Económico, 33(130), 163–184.",
    "Wolf, E. (2005 [1982]). Europa y la gente sin historia (2.ª ed.). Buenos Aires: Fondo de Cultura Económica.",
  ];
  const adi = [
    "Delrio, W. (2005). Memorias de expropiación. Bernal: Universidad Nacional de Quilmes.",
    "Foucault, M. (2000 [1976]). Defender la sociedad. Buenos Aires: Fondo de Cultura Económica.",
    "Guber, R. (2001). ¿Por qué Malvinas? Buenos Aires: Fondo de Cultura Económica.",
    "Lorenz, F. (2006). Las guerras por Malvinas. Buenos Aires: Edhasa.",
  ];
  const toRuns = (arr) => arr.map((t, i) => ({ text: t, options: { breakLine: i < arr.length - 1 } }));
  s.addText("De la materia", { x: 0.5, y: 1.1, w: 5.4, h: 0.3, fontSize: 14, bold: true, color: HEX.azul, margin: 0, isTextBox: true });
  s.addText(toRuns(obl), { x: 0.5, y: 1.45, w: 5.4, h: 3.6, fontSize: 11, color: HEX.texto, margin: 0, paraSpaceAfter: 5, valign: "top", isTextBox: true });
  s.addText("Adicionales", { x: 6.2, y: 1.1, w: 3.3, h: 0.3, fontSize: 14, bold: true, color: HEX.tierra, margin: 0, isTextBox: true });
  s.addText(toRuns(adi), { x: 6.2, y: 1.45, w: 3.3, h: 3.6, fontSize: 11, color: HEX.texto, margin: 0, paraSpaceAfter: 5, valign: "top", isTextBox: true });
}

(async () => {
  await pres.writeFile({ fileName: OUT });
    console.log("ok", OUT);
})();
