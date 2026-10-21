/**
 * Transparencia Tributaria — Fundación Apertura
 * Receptor de aportes del mapa (Google Apps Script vinculado a la planilla de carga).
 *
 * Qué hace:
 *  - Recibe los formularios de la página (errores/sugerencias y alta de localidades).
 *  - Registra todo en una planilla PRIVADA aparte («Transparencia Tributaria – Aportes (privado)»), que crea
 *    sola en el Drive de quien instala el script. Así los datos de contacto no quedan en la planilla pública.
 *  - Si es una localidad y la configuración lo permite, la da de alta en «localidades»
 *    con estado «aportada» (se ve en el mapa como «en revisión»), ubicándola con la API Georef.
 *  - Nunca modifica una localidad con estado «relevada»: esos aportes quedan para revisión manual.
 *  - Avisa por correo cada aporte nuevo (opcional).
 *
 * Cómo activarlo: ver la pestaña «Instrucciones» de la planilla o el instructivo entregado.
 */

// ===================== CONFIGURACIÓN =====================
const CONFIG = {
  EMAIL_AVISOS: "",            // correo que recibe un aviso por cada aporte (vacío = sin aviso). También se puede cargar en config → email_avisos
  AUTO_PUBLICAR: true,         // true = la localidad aportada aparece en el mapa enseguida, marcada «en revisión». También: config → auto_publicar (si/no)
  MAX_ENVIOS_POR_HORA: 5,      // por correo electrónico, para evitar abusos
  HOJA_LOCALIDADES: "localidades",
  HOJA_ACTIVIDADES: "actividades",
  HOJA_APORTES: "aportes",
  APORTES_ID: "",              // opcional: ID de una planilla propia para los aportes. Vacío = se crea sola la primera vez
  HOJA_CONFIG: "config"
};
// =========================================================

const COLS_APORTES = ["fecha", "aporte_id", "tipo", "estado_revision", "nombre", "email", "rol", "organizacion",
  "provincia", "localidad", "localidad_id", "actividad", "categoria", "comentario", "enlace", "detalle_json", "resultado"];

function doGet() {
  return json_({ ok: true, servicio: "Transparencia Tributaria — receptor de aportes", hora: new Date().toISOString() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = parse_(e);
    if (data.web_site) return json_({ ok: true, mensaje: "Recibido." });             // campo trampa: bots
    const err = validar_(data);
    if (err) return json_({ ok: false, error: err });
    if (!limite_(data.email)) return json_({ ok: false, error: "Recibimos varios envíos desde este correo en la última hora. Probá más tarde." });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cfg = leerConfig_(ss);
    const id = "AP-" + Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 900 + 100);
    let resultado = { ok: true, aporte_id: id, mensaje: "" };

    if (data.tipo === "localidad") {
      resultado = Object.assign(resultado, altaLocalidad_(ss, data, cfg, id));
    } else {
      resultado.mensaje = "Gracias. Recibimos tu comentario y lo vamos a revisar.";
    }
    registrarAporte_(ss, data, id, resultado);
    avisar_(cfg, data, id, resultado);
    return json_(resultado);
  } catch (ex) {
    return json_({ ok: false, error: "No pudimos registrar el envío: " + ex.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// ---------- Alta de localidad ----------
function altaLocalidad_(ss, d, cfg, aporteId) {
  const auto = String(cfg.auto_publicar || "").toLowerCase();
  const publicar = auto ? (auto === "si" || auto === "sí" || auto === "true") : CONFIG.AUTO_PUBLICAR;
  const sh = ss.getSheetByName(CONFIG.HOJA_LOCALIDADES);
  if (!sh) throw new Error("No existe la pestaña «localidades».");
  const values = sh.getDataRange().getValues();
  const head = values[0].map(h => String(h).trim());
  const col = name => head.indexOf(name);
  const prov = String(d.provincia || "").trim(), loc = String(d.localidad || "").trim();

  // ¿ya existe? (por id elegido en el formulario o por nombre + provincia)
  let fila = -1;
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if ((d.localidad_id && String(r[col("id")]) === d.localidad_id) ||
        (norm_(r[col("localidad")]) === norm_(loc) && norm_(r[col("provincia")]) === norm_(prov))) { fila = i; break; }
  }
  if (fila > 0) {
    const estado = String(values[fila][col("estado")] || "").toLowerCase();
    if (estado === "relevada") {
      return { publicado: false, localidad_id: values[fila][col("id")], mensaje: "Gracias. " + loc + " ya tiene datos relevados por la Fundación; tu aporte queda registrado para revisarlo y actualizar la información si corresponde." };
    }
  }
  if (!publicar) {
    return { publicado: false, mensaje: "Gracias. Recibimos los datos de " + loc + "; los vamos a revisar antes de publicarlos en el mapa." };
  }

  const geo = georef_(loc, prov);
  const hoy = Utilities.formatDate(new Date(), "America/Argentina/Buenos_Aires", "yyyy-MM-dd");
  const fila_ = fila > 0 ? values[fila].slice() : head.map(() => "");
  const set = (k, v) => { const c = col(k); if (c >= 0 && v !== undefined && v !== null && v !== "") fila_[c] = v; };
  const id = fila > 0 ? values[fila][col("id")] : nuevoId_(loc, values, col("id"));
  set("id", id); set("localidad", geo.nombre || loc); set("provincia", geo.provincia || prov);
  set("region", d.region); set("tipo", d.tipo_gobierno || "");
  fila_[col("estado")] = "aportada";
  set("ign_id", geo.id); set("lat", geo.lat); set("lon", geo.lon);
  set("adicional_oblig", numero_(d.adicional_oblig) ?? 0); set("adicional_detalle", d.adicional_detalle);
  // alícuotas: el formulario envía al_<clave> ya convertido a por mil
  head.forEach((h, c) => { if (h.indexOf("al_") === 0 && d[h] !== undefined && d[h] !== "") { const n = numero_(d[h]); if (n !== null) fila_[c] = n; } });
  const gen = numero_(d.al_general);
  if (gen !== null) head.forEach((h, c) => { if (h.indexOf("al_") === 0 && (fila_[c] === "" || fila_[c] === null) && h !== "al_financieras") fila_[c] = gen; });
  set("minimos", d.minimos); set("regimenes_especiales", d.regimenes_especiales);
  const norma = [d.ordenanza ? "Ordenanza " + d.ordenanza : "", d.periodo ? "ejercicio " + d.periodo : ""].filter(String).join(", ");
  set("normativa", (norma ? norma + ". " : "") + "Datos aportados por " + (d.rol ? d.rol.toLowerCase() : "un usuario") + "; en revisión por la Fundación.");
  // distancia a la localidad de referencia (config → localidad_referencia)
  const ref = values.find(r => String(r[col("id")]) === String(cfg.localidad_referencia || "ROS"));
  if (ref && geo.lat && col("km_rosario") >= 0 && fila_[col("km_rosario")] === "") {
    const R = 6371, rad = x => x * Math.PI / 180, la1 = rad(Number(ref[col("lat")])), la2 = rad(geo.lat);
    const a = Math.sin((la2 - la1) / 2) ** 2 + Math.cos(la1) * Math.cos(la2) * Math.sin(rad(geo.lon - Number(ref[col("lon")])) / 2) ** 2;
    fila_[col("km_rosario")] = Math.round(2 * R * Math.asin(Math.sqrt(a)));
  }
  set("unidad_valor", d.unidad_valor); set("fuente_url", d.enlace); set("fecha_relevamiento", hoy);
  if (col("aporte_id") >= 0) fila_[col("aporte_id")] = aporteId;

  if (fila > 0) sh.getRange(fila + 1, 1, 1, head.length).setValues([fila_]);
  else sh.appendRow(fila_);
  return { publicado: true, localidad_id: id, lat: geo.lat || "", lon: geo.lon || "", ign_id: geo.id || "",
           mensaje: "Gracias. " + (geo.nombre || loc) + " ya aparece en el mapa, marcada como «en revisión» hasta que la Fundación verifique los datos." };
}

function georef_(nombre, provincia) {
  const base = "https://apis.datos.gob.ar/georef/api/";
  for (const ep of ["municipios", "localidades"]) {
    try {
      const url = base + ep + "?nombre=" + encodeURIComponent(nombre) + (provincia ? "&provincia=" + encodeURIComponent(provincia) : "") + "&max=1&campos=id,nombre,centroide,provincia.nombre";
      const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (r.getResponseCode() !== 200) continue;
      const j = JSON.parse(r.getContentText()), it = (j[ep] || [])[0];
      if (it && it.centroide) return { id: ep === "municipios" ? it.id : "", nombre: it.nombre, provincia: it.provincia && it.provincia.nombre,
                                       lat: Math.round(it.centroide.lat * 10000) / 10000, lon: Math.round(it.centroide.lon * 10000) / 10000 };
    } catch (x) {}
  }
  return {};
}

function nuevoId_(nombre, values, cId) {
  const usados = new Set(values.slice(1).map(r => String(r[cId])));
  const base = norm_(nombre).replace(/[^a-z]/g, "").toUpperCase().slice(0, 3) || "LOC";
  if (!usados.has(base)) return base;
  for (let i = 2; i < 1000; i++) if (!usados.has(base + i)) return base + i;
  return base + Date.now();
}

// ---------- Registro y avisos ----------
function planillaAportes_() {
  const props = PropertiesService.getScriptProperties();
  const id = CONFIG.APORTES_ID || props.getProperty("APORTES_ID");
  if (id) { try { return SpreadsheetApp.openById(id); } catch (x) {} }
  const nueva = SpreadsheetApp.create("Transparencia Tributaria – Aportes (privado)");
  props.setProperty("APORTES_ID", nueva.getId());
  const hoja = nueva.getSheets()[0]; hoja.setName(CONFIG.HOJA_APORTES);
  return nueva;
}

function registrarAporte_(ss, d, id, res) {
  const priv = planillaAportes_();
  let sh = priv.getSheetByName(CONFIG.HOJA_APORTES);
  if (!sh) sh = priv.insertSheet(CONFIG.HOJA_APORTES);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLS_APORTES);
    sh.getRange(1, 1, 1, COLS_APORTES.length).setFontWeight("bold").setBackground("#293545").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  const detalle = Object.assign({}, d); delete detalle.web_site;
  sh.appendRow([new Date(), id, d.tipo === "localidad" ? "localidad" : "comentario", "nuevo", d.nombre, d.email, d.rol || "", d.organizacion || "",
    d.provincia || "", d.localidad || "", res.localidad_id || d.localidad_id || "", d.actividad || "", d.categoria || "",
    d.comentario || "", d.enlace || "", JSON.stringify(detalle), res.mensaje || ""]);
}

function avisar_(cfg, d, id, res) {
  const to = cfg.email_avisos || CONFIG.EMAIL_AVISOS;
  if (!to) return;
  const asunto = "[Transparencia Tributaria] " + (d.tipo === "localidad" ? "Nueva localidad: " : "Comentario: ") + (d.localidad || "");
  const cuerpo = [
    "Aporte " + id, "Tipo: " + d.tipo, "Localidad: " + (d.localidad || "") + " (" + (d.provincia || "") + ")",
    "De: " + d.nombre + " <" + d.email + "> · " + (d.rol || "") + " " + (d.organizacion || ""),
    d.categoria ? "Categoría: " + d.categoria : "", d.actividad ? "Actividad: " + d.actividad : "",
    d.comentario ? "\nComentario:\n" + d.comentario : "", d.enlace ? "\nEnlace: " + d.enlace : "",
    "\nResultado: " + (res.mensaje || ""), "\nRevisalo en la planilla privada de aportes: " + planillaAportes_().getUrl()
  ].filter(String).join("\n");
  try { MailApp.sendEmail(to, asunto, cuerpo, { replyTo: d.email }); } catch (x) {}
}

// ---------- Utilidades ----------
function parse_(e) {
  if (e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (x) {}
  }
  return (e && e.parameter) || {};
}

function validar_(d) {
  const mail = String(d.email || "").trim();
  if (!String(d.nombre || "").trim()) return "Falta el nombre.";
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail)) return "El correo electrónico no es válido.";
  if (!String(d.localidad || "").trim()) return "Falta la localidad.";
  if (d.tipo === "localidad") {
    if (!String(d.provincia || "").trim()) return "Falta la provincia.";
    if (numero_(d.al_general) === null) return "Falta la alícuota general.";
    if (!/^https?:\/\//i.test(String(d.enlace || ""))) return "Falta el enlace a la normativa.";
    if (numero_(d.al_general) > 100) return "La alícuota general parece demasiado alta: revisá la unidad (por mil o por ciento).";
  } else if (!String(d.comentario || "").trim()) return "Falta el comentario.";
  return "";
}

function limite_(email) {
  const c = CacheService.getScriptCache(), k = "tt_" + norm_(email).replace(/\s/g, "");
  const n = Number(c.get(k) || 0);
  if (n >= CONFIG.MAX_ENVIOS_POR_HORA) return false;
  c.put(k, String(n + 1), 3600);
  return true;
}

function leerConfig_(ss) {
  const sh = ss.getSheetByName(CONFIG.HOJA_CONFIG), o = {};
  if (!sh) return o;
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1] || "").trim(); });
  return o;
}

function numero_(v) {
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isFinite(v) ? v : null;
  let s = String(v).trim().replace(/[%‰\s$]/g, "");
  if (s.indexOf(",") >= 0) s = s.replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function norm_(s) {
  return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Prueba manual desde el editor: ejecutar probarConfiguracion y revisar el registro. */
function probarConfiguracion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("Pestañas: " + ss.getSheets().map(s => s.getName()).join(", "));
  Logger.log("Planilla privada de aportes: " + planillaAportes_().getUrl());
  Logger.log("Georef Rosario: " + JSON.stringify(georef_("Rosario", "Santa Fe")));
}
