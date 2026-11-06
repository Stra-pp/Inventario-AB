/* Envío automático de los correos de Inventario A&B, SIN abrir la app.
   Lo corre GitHub Actions (.github/workflows/correos-inventario.yml) varias
   veces entre 10:05 y 12:35 hora Ecuador. Carga el MISMO código de la app
   (index.html), así que los correos y los Excel salen idénticos a los de la
   app. Cada correo se marca como enviado en la configuración compartida
   (ultimoReporteDiario / ultimoReporteCostos / ultimoReporteSemanal), así que
   aunque corra varias veces —o alguien abra la app— sale una sola vez.

   Variables (secretos del repositorio):
     GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET  → app de Entra ID con
        permisos de APLICACIÓN Mail.Send y Sites.ReadWrite.All (o Sites.Selected).
     GRAPH_SENDER → buzón desde el que salen los correos.
   Opcional (ejecución manual): PRUEBA_PARA → envía los 3 correos como prueba
   solo a ese correo, con asunto "Prueba". */
const fs = require("fs"), path = require("path"), vm = require("vm");
const babel = require("@babel/core");
const ExcelJS = require("exceljs");
const { chromium } = require("playwright");

const req = k => { const v = process.env[k]; if (!v) { console.error(`Falta la variable ${k}`); process.exit(2); } return v; };
const TENANT = req("GRAPH_TENANT_ID"), CLIENT = req("GRAPH_CLIENT_ID"), SECRET = req("GRAPH_CLIENT_SECRET"), SENDER = req("GRAPH_SENDER");
const PRUEBA = (process.env.PRUEBA_PARA || "").trim();

async function token() {
  const r = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, { method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT, client_secret: SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }) });
  const j = await r.json(); if (!j.access_token) throw new Error("No se obtuvo token: " + JSON.stringify(j).slice(0, 300)); return j.access_token;
}

function cargarApp(extra) {
  const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
  const m = html.match(/<script type="text\/babel"[^>]*>([\s\S]*?)<\/script>/); if (!m) throw new Error("No se encontró el código de la app");
  const code = babel.transformSync(m[1], { babelrc: false, configFile: false, plugins: [[require.resolve("@babel/plugin-transform-react-jsx"), { runtime: "classic" }]] }).code;
  const P = () => new Proxy(function () {}, { get: (t, k) => k === "then" ? undefined : P(), apply: () => P(), construct: () => P() });
  const mem = {};
  const ctx = { console, setTimeout, clearTimeout, setInterval: () => 0, clearInterval() {}, Promise, Date, Math, JSON, URL, URLSearchParams, Intl,
    fetch, AbortController, TextEncoder, TextDecoder, Uint8Array, ArrayBuffer, Blob, Response, Headers,
    btoa: s => Buffer.from(s, "binary").toString("base64"), atob: s => Buffer.from(s, "base64").toString("binary"),
    React: { useState: v => [v, () => {}], useEffect() {}, useRef: () => ({}), useMemo: f => f(), useCallback: f => f, createElement: () => null, Fragment: "f" },
    ReactDOM: { createRoot: () => ({ render() {} }) }, document: P(), msal: P(), navigator: { userAgent: "inventario-ab-automatico" },
    localStorage: { getItem: k => (k in mem ? mem[k] : null), setItem: (k, v) => { mem[k] = String(v); }, removeItem: k => { delete mem[k]; } },
    sessionStorage: { getItem: () => null, setItem() {}, removeItem() {} },
    location: { href: "https://stra-pp.github.io/Inventario-AB/", origin: "https://stra-pp.github.io", pathname: "/Inventario-AB/", search: "", hash: "" },
    addEventListener() {}, removeEventListener() {}, ExcelJS, ...extra };
  ctx.window = ctx; ctx.globalThis = ctx; ctx.self = ctx;
  vm.createContext(ctx);
  vm.runInContext(code, ctx, { filename: "index.html" });
  return ctx;
}

(async () => {
  const br = await chromium.launch(); const pg = await br.newPage({ deviceScaleFactor: 2 });
  // Gráficos del Excel semanal: SVG → PNG con el navegador
  const svg2png = async (svg, w, h) => { await pg.setViewportSize({ width: w, height: h }); await pg.setContent(`<html><body style="margin:0">${svg}</body></html>`);
    return (await pg.screenshot({ clip: { x: 0, y: 0, width: w, height: h } })).toString("base64"); };
  const ctx = cargarApp({ INV_SVG2PNG: svg2png, __nuevoToken: token, __SENDER: SENDER });
  ctx.__T0 = await token();
  // Modo servidor: token de aplicación y envío desde el buzón configurado (no hay "/me").
  vm.runInContext(`MODE="graph"; graphToken=__T0;
    refreshToken = async function(){ try{ graphToken = await __nuevoToken(); return true; }catch(e){ return false; } };
    const __gfOrig = gfetch;
    gfetch = async function(p, ...a){ if(p==="/me/sendMail") p="/users/"+encodeURIComponent(__SENDER)+"/sendMail"; return __gfOrig(p, ...a); };`, ctx);
  const prueba = PRUEBA ? { asunto: "Prueba", para: PRUEBA } : undefined;
  let fallos = 0;
  // Avisar a la app que el servidor está activo (así la app no envía desde la cuenta de otros usuarios).
  if (!prueba) { try { await vm.runInContext("invMarcarServidor()", ctx); } catch (e) { console.error("No se pudo marcar el servidor: " + (e && e.message || e)); } }
  for (const [nombre, fn] of [["Correo de las 10:00 (cierre del día)", "invEnviarReporteDiario"], ["Correo de las 10:30 (costos)", "invEnviarReporteCostos"], ["Reporte semanal (lunes 10:30)", "invEnviarReporteSemanal"], ["Recordatorio de costos (lunes 11:00)", "invRecordatorioCostos"]]) {
    try { ctx.__p = prueba; const r = await vm.runInContext(`${fn}(__p)`, ctx); const dias = r && r.dias && r.dias.length ? ` (${r.dias.join(", ")})` : "";
      const espera = r && r.retenido ? ` · en espera de aprobación el ${r.dia}: ` + r.faltan.map(x => x.bodega + " " + x.estado).join("; ") : r && r.esperando ? ` · semana del ${r.semana} esperando que se publiquen todos sus días` : "";
      console.log(`${nombre}: ${r && r.enviado ? "ENVIADO" + dias : "nada que enviar"}${espera}`); }
    catch (e) { fallos++; console.error(`${nombre}: ERROR — ${e && e.message || e}`); }
  }
  await br.close();
  process.exit(fallos ? 1 : 0);
})().catch(e => { console.error(e); process.exit(1); });
