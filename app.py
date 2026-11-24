import time
import re
import requests
from flask import Flask, jsonify, request, send_file
from io import BytesIO
from datetime import date, timedelta
from reportlab.pdfgen import canvas
from openpyxl import Workbook
from math import isfinite
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32768
API_URL = "https://open.er-api.com/v6/latest/USD"
CACHE_SEGUNDOS = 3600  # 1 hora
_cache = {"tasas": None, "momento": 0, "fecha": None}
MONEDAS = {
    "USD": ("Dólar estadounidense", "🇺🇸"),
    "COP": ("Peso colombiano", "🇨🇴"),
    "EUR": ("Euro", "🇪🇺"),
    "MXN": ("Peso mexicano", "🇲🇽"),
    "GBP": ("Libra esterlina", "🇬🇧"),
    "BRL": ("Real brasileño", "🇧🇷"),
    "ARS": ("Peso argentino", "🇦🇷"),
    "CLP": ("Peso chileno", "🇨🇱"),
    "PEN": ("Sol peruano", "🇵🇪"),
    "JPY": ("Yen japonés", "🇯🇵"),
    "CAD": ("Dólar canadiense", "🇨🇦"),
    "CHF": ("Franco suizo", "🇨🇭"),
    "CNY": ("Yuan chino", "🇨🇳"),
    "AUD": ("Dólar australiano", "🇦🇺"),
    "UYU": ("Peso uruguayo", "🇺🇾"),
    "VES": ("Bolívar venezolano", "🇻🇪"),
    "PAB": ("Balboa panameño", "🇵🇦"),
    "CRC": ("Colón costarricense", "🇨🇷"),
    "DOP": ("Peso dominicano", "🇩🇴"),
    "INR": ("Rupia india", "🇮🇳"),
    "KRW": ("Won surcoreano", "🇰🇷"),
}
def cargar_tasas():
    """Devuelve (tasas, desactualizado). Usa caché de 1 hora."""
    ahora = time.time()
    if _cache["tasas"] and ahora - _cache["momento"] < CACHE_SEGUNDOS:
        return _cache["tasas"], False
    try:
        respuesta = requests.get(API_URL, timeout=10)
        respuesta.raise_for_status()
        datos = respuesta.json()
        if datos.get("result") != "success":
            raise ValueError("Respuesta inválida de la API")
        _cache["tasas"] = {k:v for k,v in datos["rates"].items() if re.fullmatch(r"[A-Z]{3}", k) and isinstance(v,(int,float)) and isfinite(v) and v>0}
        if not _cache["tasas"]: raise ValueError("Sin tasas válidas")
        _cache["momento"] = ahora
        _cache["fecha"] = datos.get("time_last_update_utc")
        return _cache["tasas"], False
    except Exception:
        # Si la API falla pero hay datos guardados, los usamos
        if _cache["tasas"]:
            return _cache["tasas"], True
        raise
# ---------------------------------------------------------------
# Página completa (HTML + CSS + JavaScript) en un solo bloque
# ---------------------------------------------------------------
PAGINA = r"""<!DOCTYPE html>
<html lang="es" data-tema="oscuro">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>AMVEXA · Conversor de monedas</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
*{box-sizing:border-box} :root{color-scheme:dark;--bg:#080b19;--ink:#f5f1ff;--muted:#a9acc9;--glass:rgba(27,31,59,.65);--line:rgba(192,178,255,.20);--gold:#e8cd99;--mint:#40f5d4;--field:rgba(7,10,25,.55)} [data-tema=claro]{color-scheme:light;--bg:#e9e7f6;--ink:#22213c;--muted:#555570;--glass:rgba(255,255,255,.67);--line:#b9b3d5;--gold:#805a20;--mint:#007d70;--field:#ffffffa0} html{scroll-behavior:smooth} body{margin:0;background:var(--bg);color:var(--ink);font:15px Inter,system-ui,sans-serif;min-height:100vh} body:before{content:'';position:fixed;inset:0;z-index:-1;background:radial-gradient(ellipse at 85% 10%,#7546c94d,transparent 45%),radial-gradient(ellipse at 35% 75%,#16c6b32b,transparent 50%),radial-gradient(ellipse at 0 0,#4c63c944,transparent 40%);animation:aurora 12s ease-in-out infinite alternate} @keyframes aurora{to{transform:scale(1.12) translateY(-2%)}} a{color:inherit;text-decoration:none} button,input,select{font:inherit;color:inherit} button{cursor:pointer;border:1px solid var(--line);border-radius:12px;background:var(--field);padding:12px 15px;transition:transform .2s,box-shadow .2s} button:hover{transform:translateY(-2px);box-shadow:0 0 20px #9974ff44} button:disabled{opacity:.5;cursor:wait} :focus-visible{outline:2px solid var(--mint);outline-offset:3px} .shell{display:grid;grid-template-columns:220px minmax(0,1fr);max-width:1700px;margin:auto}.sidebar{position:sticky;top:0;height:100vh;padding:32px 20px;border-right:1px solid var(--line);background:var(--glass);backdrop-filter:blur(24px);display:flex;flex-direction:column;gap:34px}.brand{display:flex;align-items:center;gap:10px;font-weight:700;letter-spacing:3px}.mark{width:46px;height:46px;flex:none}.brand small{display:block;font-size:8px;letter-spacing:2px;color:var(--gold);margin-top:5px}.nav{display:grid;gap:10px}.nav a{padding:14px 12px;border-radius:12px;color:var(--muted);border:1px solid transparent}.nav a:hover,.nav a:focus{border-color:var(--line);background:var(--field);color:var(--ink)}.signature{margin-top:auto;color:var(--muted);font-size:12px;line-height:1.8}.signature strong{color:var(--gold)}.app{padding:30px clamp(18px,3vw,48px);min-width:0}.cabecera,.titulo-fila,.fila-extra{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap}.cabecera{margin-bottom:24px}.eyebrow,.kicker{font-size:10px;letter-spacing:3px;text-transform:uppercase;color:var(--gold);font-weight:700}.acciones{display:flex;align-items:center;gap:12px;flex-wrap:wrap}.estado{font-size:11px;color:var(--mint)}.estado.error,.estado.aviso{color:var(--gold)}.hero{position:relative;overflow:hidden;min-height:260px;border:1px solid var(--line);border-radius:28px;background:linear-gradient(115deg,var(--glass),#7853bb26);padding:38px;display:flex;align-items:center;margin-bottom:24px;box-shadow:0 24px 70px #0002}.hero-copy{position:relative;z-index:1;max-width:60%}.hero h1{font-family:Georgia,serif;font-weight:400;font-size:clamp(32px,4vw,58px);line-height:1.05;margin:15px 0}.hero h1 em{color:var(--gold);font-weight:400}.hero p{color:var(--muted);line-height:1.7}.planet{position:absolute;right:1%;top:0;width:43%;height:100%;display:grid;place-items:center;pointer-events:none}.planet svg{width:100%;height:100%;filter:drop-shadow(0 0 20px #8c6cff88)}.orbit{transform-origin:180px 130px;animation:spin 22s linear infinite}.float{animation:float 4s ease-in-out infinite alternate}@keyframes spin{to{transform:rotate(360deg)}}@keyframes float{to{transform:translateY(-12px)}}.dashboard{display:grid;grid-template-columns:minmax(0,1.15fr) minmax(0,.85fr);gap:20px;align-items:start}.tarjeta{min-width:0;padding:24px;border:1px solid var(--line);border-radius:23px;background:var(--glass);backdrop-filter:blur(22px);box-shadow:inset 0 1px 0 #ffffff16,0 16px 40px #0002;scroll-margin-top:20px}.tarjeta h2{font-size:19px;font-weight:600;margin:5px 0 20px}.tarjeta small,#grafEstado{display:block;color:var(--muted);font-size:12px;line-height:1.6;margin-top:10px}.wide{grid-column:1/-1}.stack{display:grid;gap:20px;min-width:0}.campo,.inteligente{display:grid;gap:8px;margin-bottom:15px;min-width:0} label{color:var(--muted);font-size:12px;display:block} input,select{width:100%;min-width:0;background:var(--field);border:1px solid var(--line);border-radius:12px;padding:13px;outline-offset:2px} select option{background:var(--bg);color:var(--ink)} .grande input{font-size:30px;font-weight:600}.selectores{display:grid;grid-template-columns:minmax(0,1fr) 46px minmax(0,1fr);gap:10px;align-items:center}.intercambio{padding:10px;color:var(--gold);font-size:22px}.intercambio:active{transform:rotate(180deg)}.resultado{position:relative;overflow:hidden;border-radius:19px;padding:24px;background:linear-gradient(125deg,#7956cc35,#13bfa62b);border:1px solid var(--line);margin-top:8px}.res-pequeno,.res-tasa{color:var(--muted);font-size:12px}.res-grande{font-size:clamp(24px,3vw,42px);font-weight:600;color:var(--mint);margin:12px 0;overflow-wrap:anywhere}.copiar{margin-top:16px;background:linear-gradient(110deg,#e8cd99,#b58d54);color:#20172a;border:0;font-weight:700}.grid-rapidas{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}.rapida{display:grid;text-align:left;gap:8px;overflow-wrap:anywhere}.r-cod{color:var(--muted);font-size:12px}.r-val{font-weight:600;color:var(--mint)}.historial{padding:0;list-style:none;margin:0}.historial li{display:flex;justify-content:space-between;gap:10px;padding:14px 0;border-bottom:1px solid var(--line);cursor:pointer;overflow-wrap:anywhere}.historial strong{color:var(--mint)}.enlace{font-size:12px;padding:8px}.fila-extra select{flex:1;min-width:110px}.fila-extra button{color:var(--gold)}.chart-wrap{height:270px;margin-top:15px;position:relative} #multi{display:grid;gap:8px;margin-top:12px} #multi div,#neto{padding:12px;background:var(--field);border-radius:10px;color:var(--mint);overflow-wrap:anywhere} #neto{margin-top:12px}.valido{border-color:var(--mint)} footer{color:var(--muted);font-size:11px;padding:25px 0;text-align:center;line-height:1.7}.voice-row{display:flex;gap:8px;align-items:end}.voice-row .inteligente{flex:1}.voice-row button{margin-bottom:15px}.vacio{color:var(--muted)} @media(max-width:1050px){.shell{grid-template-columns:170px minmax(0,1fr)}.sidebar{padding:22px 12px}.dashboard{grid-template-columns:1fr}.wide{grid-column:auto}} @media(max-width:700px){.shell{display:block}.sidebar{position:relative;height:auto;padding:16px;gap:14px}.nav{display:flex;overflow:auto;gap:5px}.nav a{white-space:nowrap;padding:9px}.signature{display:none}.app{padding:16px}.hero{padding:24px;min-height:245px}.hero-copy{max-width:80%}.planet{opacity:.45;width:65%;right:-15%}.tarjeta{padding:17px}.selectores{grid-template-columns:1fr 40px 1fr}.selectores select{padding:10px 5px;font-size:12px}.chart-wrap{height:220px}} @media(prefers-reduced-motion:reduce){*,*:before{animation:none!important;transition:none!important;scroll-behavior:auto!important}}
</style>
</head>
<body><div class="shell"><aside class="sidebar"><a class="brand" href="#inicio"><svg class="mark" viewBox="0 0 50 50" fill="none" stroke="currentColor" style="color:var(--gold)"><rect x="2" y="2" width="46" height="46" rx="14"/><path d="M8 35L19 12L30 35M13 25H25M25 35V13L34 25L43 13V35" stroke-width="1.6"/></svg><span>AMVEXA<small>AURORA EDITION</small></span></a><nav class="nav" aria-label="Secciones"><a href="#inicio">◈ &nbsp; Panorama</a><a href="#conversion">⇄ &nbsp; Conversor</a><a href="#mercado">⌁ &nbsp; Mercados</a><a href="#herramientas">◇ &nbsp; Herramientas</a><a href="#actividad">◷ &nbsp; Actividad</a></nav><div class="signature">DISEÑO & TECNOLOGÍA<br><strong>Ana María</strong><br>AMVEXA / 03</div></aside><main class="app" id="inicio"><header class="cabecera"><div><div class="eyebrow">Tu espacio financiero</div><div style="margin-top:6px;color:var(--muted)">Panorama internacional</div></div><div class="acciones"><span id="estado" class="estado" role="status">Conectando…</span><button id="btnTema" aria-label="Cambiar tema">◐ Tema</button></div></header><section class="hero"><div class="hero-copy"><div class="eyebrow">01 / Global currency experience</div><h1>El valor no tiene <em>fronteras.</em></h1><p>Explora, convierte y comprende las divisas del mundo en un solo lugar.</p><a href="#conversion" style="color:var(--gold)">Explorar conversor →</a></div><div class="planet" aria-hidden="true"><svg viewBox="0 0 360 260"><defs><radialGradient id="sphere"><stop stop-color="#8567ff" stop-opacity=".35"/><stop offset="1" stop-color="#171d50" stop-opacity=".6"/></radialGradient><linearGradient id="wire"><stop stop-color="#a88bff"/><stop offset="1" stop-color="#40f5d4"/></linearGradient></defs><circle cx="180" cy="130" r="88" fill="url(#sphere)" stroke="url(#wire)"/><g fill="none" stroke="url(#wire)" opacity=".65"><ellipse cx="180" cy="130" rx="60" ry="88"/><ellipse cx="180" cy="130" rx="27" ry="88"/><ellipse cx="180" cy="130" rx="88" ry="35"/><ellipse cx="180" cy="130" rx="88" ry="65"/><path d="M92 130H268M180 42V218"/></g><g class="orbit" fill="none" stroke="#e8cd99"><ellipse cx="180" cy="130" rx="145" ry="47" transform="rotate(-27 180 130)" stroke-dasharray="4 7"/><circle cx="310" cy="65" r="4" fill="#e8cd99"/></g><g class="float" fill="#e8cd99" font-size="23" font-family="Georgia"><text x="42" y="75">€</text><text x="288" y="170">$</text><text x="65" y="213">¥</text></g></svg></div></section><div class="dashboard"><section class="tarjeta" id="conversion"><div class="kicker">02 / Exchange studio</div><h2>Conversión inteligente</h2><div class="voice-row"><div class="inteligente"><label for="busqueda">Escribe tu conversión</label><input id="busqueda" placeholder="250 USD a COP" autocomplete="off"></div><button id="voz" aria-label="Dictar conversión">🎙</button></div><small id="vozEstado" role="status">También puedes dictar tu consulta.</small><div class="campo grande"><label for="cantidad">Cantidad de origen</label><input id="cantidad" inputmode="decimal" value="100"></div><div class="selectores"><div class="campo"><label for="origen">Desde</label><select id="origen"></select></div><button id="btnCambiar" class="intercambio" aria-label="Intercambiar monedas">⇄</button><div class="campo"><label for="destino">Hacia</label><select id="destino"></select></div></div><div class="resultado" aria-live="polite"><div class="kicker">Valor estimado</div><div id="resTexto" class="res-pequeno">—</div><div id="resTotal" class="res-grande">—</div><div id="resTasa" class="res-tasa"></div><button id="btnCopiar" class="copiar">Copiar resultado</button></div></section><section class="tarjeta" id="mercado"><div class="kicker">03 / Market intelligence</div><h2>Evolución de divisas</h2><div class="fila-extra"><select id="periodo" aria-label="Periodo"><option value="30">30 días</option><option value="90">90 días</option><option value="365">1 año</option></select><button id="grafica">Consultar gráfica</button></div><div class="chart-wrap"><canvas id="grafico" aria-label="Gráfica histórica"></canvas></div><div id="grafEstado" role="status">Selecciona un par y consulta su evolución. Fuente: Frankfurter.</div></section><section class="tarjeta wide"><div class="kicker">04 / Global snapshot</div><h2>Equivalencias rápidas</h2><div id="rapidas" class="grid-rapidas"></div></section><section class="tarjeta" id="herramientas"><div class="kicker">05 / Smart tools</div><h2>Conversión múltiple</h2><label for="multiples">Selecciona varias monedas (Ctrl + clic en Windows)</label><select id="multiples" multiple size="6"></select><div id="multi"></div></section><div class="stack"><section class="tarjeta"><div class="kicker">06 / Net value</div><h2>Simulador de comisiones</h2><label for="comision">Comisión porcentual</label><input id="comision" type="number" min="0" max="100" step="0.1" value="0"><div id="neto" aria-live="polite">—</div><small>Comisión descontada del valor bruto.</small></section><section class="tarjeta"><div class="kicker">07 / Documents</div><h2>Exportar resultados</h2><div class="fila-extra"><button id="pdf">↓ Comprobante PDF</button><button id="excel">↓ Historial Excel</button></div><small id="exportEstado" role="status"></small></section></div><section class="tarjeta wide" id="actividad"><div class="kicker">08 / Activity</div><div class="titulo-fila"><h2>Historial de conversiones</h2><button id="btnLimpiar" class="enlace">Borrar</button></div><ul id="historial" class="historial"></ul></section></div><footer>AMVEXA 3.0 · Diseñado por Ana María · Tasas de open.er-api.com · Herramienta informativa, no bancaria.</footer></main></div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.9/dist/chart.umd.min.js"></script>
  <script>
    const $ = (id) => document.getElementById(id);
    const POPULARES = ["USD", "COP", "EUR", "MXN", "GBP", "BRL", "ARS", "CLP"];
    const CLAVE_HIST = "fxpro_historial";
    const CLAVE_TEMA = "fxpro_tema";
    // Palabras que la barra inteligente entiende
    const ALIAS = {
      dolar: "USD", dolares: "USD",
      euro: "EUR", euros: "EUR",
      yen: "JPY", yenes: "JPY",
      libra: "GBP", libras: "GBP",
      real: "BRL", reales: "BRL",
    };
    let tasas = {};
    let monedas = [];
    let mapa = {};
    let temporizador = null;
    /* ---------- Utilidades ---------- */
    function convertir(monto, de, a) {
      return monto * (tasas[a] / tasas[de]);
    }
    function formato(n) {
      const abs = Math.abs(n);
      const decimales = abs >= 1 ? 2 : abs >= 0.01 ? 4 : 6;
      return new Intl.NumberFormat("es-CO", {
        minimumFractionDigits: 2,
        maximumFractionDigits: decimales,
      }).format(n);
    }
    // Entiende "1.500,50", "1,500.50", "1500" y "1500,5"
    function leerNumero(texto) {
      let t = texto.trim().replace(/\s/g, "");
      const coma = t.lastIndexOf(",");
      const punto = t.lastIndexOf(".");
      if (coma > -1 && punto > -1) {
        if (coma > punto) t = t.replace(/\./g, "").replace(",", ".");
        else t = t.replace(/,/g, "");
      } else if (coma > -1) {
        const decimales = t.length - coma - 1;
        t = decimales <= 2 ? t.replace(",", ".") : t.replace(/,/g, "");
      } else if (punto > -1) {
        const decimales = t.length - punto - 1;
        if (decimales > 2) t = t.replace(/\./g, "");
      }
      const n = Number(t);
      return !t || !Number.isFinite(n) || n < 0 || n > 1e15 ? null : n;
    }
    /* ---------- Barra inteligente ---------- */
    function interpretarBusqueda(texto) {
      texto=texto.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/pesos? colombianos?/g,"cop").replace(/pesos? mexicanos?/g,"mxn").replace(/dolares? estadounidenses?/g,"usd").replace(/^\s*(convierte|convertir|cuanto es|cuanto son)\s+/i, "");
      const t = texto
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .trim();
      const m = t.match(
        /^([\d.,]+)\s*(?:(millones|millon|mil|k|m)\b)?\s*([a-z]+)\s*(?:a|en|to|->|=|hacia)\s*([a-z]+)$/
      );
      if (!m) return null;
      const [, num, mult, de, a] = m;
      let monto = leerNumero(num);
      if (monto === null) return null;
      if (mult === "k" || mult === "mil") monto *= 1e3;
      else if (mult) monto *= 1e6;
      const origen = ALIAS[de] || de.toUpperCase();
      const destino = ALIAS[a] || a.toUpperCase();
      if (!tasas[origen] || !tasas[destino]) return null;
      return { monto, origen, destino };
    }
    /* ---------- Pintar la pantalla ---------- */
    function actualizar() {
      const monto = leerNumero($("cantidad").value);
      const de = $("origen").value;
      const a = $("destino").value;
      if (monto === null || !tasas[de] || !tasas[a]) {
        $("resTexto").textContent = "Escribe una cantidad válida";
        $("resTotal").textContent = "—";
        $("resTasa").textContent = "";
        $("rapidas").innerHTML = "";
        actualizarExtras();
        return;
      }
      const total = convertir(monto, de, a);
      $("resTexto").textContent = `${formato(monto)} ${de} =`;
      $("resTotal").textContent = `${formato(total)} ${a}`;
      $("resTasa").textContent = `1 ${de} = ${formato(convertir(1, de, a))} ${a}`;
      pintarRapidas(monto, de);
      actualizarExtras();
    }
    function pintarRapidas(monto, de) {
      const cont = $("rapidas");
      cont.innerHTML = "";
      POPULARES.filter((c) => c !== de && tasas[c])
        .slice(0, 6)
        .forEach((c) => {
          const bandera = mapa[c] ? mapa[c].bandera : "🏳️";
          const boton = document.createElement("button");
          boton.className = "rapida";
          boton.innerHTML =
            `<span class="r-cod">${bandera} ${c}</span>` +
            `<span class="r-val">${formato(convertir(monto, de, c))}</span>`;
          boton.onclick = () => {
            $("destino").value = c;
            actualizar();
            programarGuardado();
          };
          cont.appendChild(boton);
        });
    }
    function llenarSelects() {
      const opciones = monedas
        .map((m) => `<option value="${m.codigo}">${m.bandera} ${m.codigo} — ${m.nombre}</option>`)
        .join("");
      $("origen").innerHTML = opciones;
      $("destino").innerHTML = opciones;
      $("origen").value = "USD";
      $("destino").value = "COP";
    }
    /* ---------- Historial ---------- */
    function leerHistorial() {
      try {
        const h=JSON.parse(localStorage.getItem(CLAVE_HIST)); return Array.isArray(h)?h.filter(x=>x&&Number.isFinite(x.monto)&&x.monto>=0&&/^[A-Z]{3}$/.test(x.de)&&/^[A-Z]{3}$/.test(x.a)).slice(0,6):[];
      } catch {
        return [];
      }
    }
    function guardarHistorial() {
      const monto = leerNumero($("cantidad").value);
      const de = $("origen").value;
      const a = $("destino").value;
      if (monto === null || !tasas[de] || !tasas[a]) return;
      const lista = leerHistorial();
      const primero = lista[0];
      const repetido = primero && primero.monto === monto && primero.de === de && primero.a === a;
      if (!repetido) {
        lista.unshift({ monto, de, a });
        localStorage.setItem(CLAVE_HIST, JSON.stringify(lista.slice(0, 6)));
      }
      pintarHistorial();
    }
    function programarGuardado() {
      clearTimeout(temporizador);
      temporizador = setTimeout(guardarHistorial, 1500);
    }
    function pintarHistorial() {
      const ul = $("historial");
      const lista = leerHistorial();
      ul.innerHTML = "";
      if (!lista.length) {
        ul.innerHTML = '<li class="vacio">Aún no hay conversiones.</li>';
        return;
      }
      lista.forEach((item) => {
        const total = tasas[item.de] && tasas[item.a] ? convertir(item.monto, item.de, item.a) : null;
        const li = document.createElement("li");
        li.innerHTML =
          `<span>${formato(item.monto)} ${item.de} → ${item.a}</span>` +
          `<strong>${total === null ? "—" : formato(total)}</strong>`;
        li.onclick = () => {
          $("cantidad").value = String(item.monto);
          $("origen").value = item.de;
          $("destino").value = item.a;
          actualizar();
        };
        ul.appendChild(li);
      });
    }
    /* ---------- Carga de tasas ---------- */
    async function cargar() {
      const estado = $("estado");
      try {
        const respuesta = await fetch("/api/tasas");
        const datos = await respuesta.json();
        if (!respuesta.ok) throw new Error(datos.error || "Error");
        tasas = datos.tasas;
        monedas = datos.monedas;
        mapa = Object.fromEntries(monedas.map((m) => [m.codigo, m]));
        if (!$("origen").options.length) llenarSelects();
        if (!$ ("multiples").options.length) { for(const m of monedas){const o=new Option(m.codigo+" · "+m.nombre,m.codigo);o.selected=["USD","COP","EUR"].includes(m.codigo);$("multiples").add(o);} }
        const fecha = datos.fecha
          ? new Date(datos.fecha).toLocaleString("es-CO", { dateStyle: "medium", timeStyle: "short" })
          : "";
        if (datos.desactualizado) {
          estado.textContent = "⚠ Tasas en caché";
          estado.className = "estado aviso";
        } else {
          estado.textContent = `● Actualizado · ${fecha}`;
          estado.className = "estado ok";
        }
        actualizar();
        pintarHistorial();
      } catch (error) {
        estado.textContent = "✖ Sin conexión";
        estado.className = "estado error";
      }
    }
    /* ---------- Tema ---------- */
    function aplicarTema(tema) {
      document.documentElement.dataset.tema = tema;
      localStorage.setItem(CLAVE_TEMA, tema);
    }
    /* ---------- Eventos ---------- */
    $("cantidad").addEventListener("input", () => {
      actualizar();
      programarGuardado();
    });
    $("origen").addEventListener("change", () => { actualizar(); programarGuardado(); });
    $("destino").addEventListener("change", () => { actualizar(); programarGuardado(); });
    $("btnCambiar").addEventListener("click", () => {
      const anterior = $("origen").value;
      $("origen").value = $("destino").value;
      $("destino").value = anterior;
      actualizar();
      programarGuardado();
    });
    $("busqueda").addEventListener("input", (e) => {
      const r = interpretarBusqueda(e.target.value);
      e.target.classList.toggle("valido", !!r);
      if (!r) return;
      $("cantidad").value = String(r.monto);
      $("origen").value = r.origen;
      $("destino").value = r.destino;
      actualizar();
      programarGuardado();
    });
    $("btnCopiar").addEventListener("click", () => {
      const boton = $("btnCopiar");
      navigator.clipboard.writeText($("resTotal").textContent).then(() => {
        boton.textContent = "¡Copiado!";
        setTimeout(() => (boton.textContent = "Copiar resultado"), 1500);
      });
    });
    $("btnLimpiar").addEventListener("click", () => {
      localStorage.removeItem(CLAVE_HIST);
      pintarHistorial();
    });
    $("btnTema").addEventListener("click", () => {
      const actual = document.documentElement.dataset.tema;
      aplicarTema(actual === "oscuro" ? "claro" : "oscuro");
    });
    /* ---------- Inicio ---------- */
    const temaGuardado = localStorage.getItem(CLAVE_TEMA);
    const prefiereClaro = window.matchMedia("(prefers-color-scheme: light)").matches;
    aplicarTema(temaGuardado || "oscuro");
    pintarHistorial();
    cargar();
    setInterval(cargar, 30 * 60 * 1000); // refresca cada 30 minutos

    function actualizarExtras(){
      const m=leerNumero($("cantidad").value), de=$("origen").value,a=$("destino").value,p=Number($("comision").value);
      $("multi").replaceChildren(); $("neto").textContent="—";
      if(m===null||!tasas[de]||!tasas[a])return;
      if(Number.isFinite(p)&&p>=0&&p<=100){const bruto=convertir(m,de,a);$("neto").textContent=`Comisión: ${formato(bruto*p/100)} ${a} · Neto: ${formato(bruto*(1-p/100))} ${a}`;}
      for(const o of $("multiples").selectedOptions){const d=document.createElement("div");d.textContent=`${o.value}: ${formato(convertir(m,de,o.value))}`;$("multi").append(d);}
    }
    $("comision").oninput=actualizarExtras; $("multiples").onchange=actualizarExtras;
    $("voz").onclick=()=>{const R=window.SpeechRecognition||window.webkitSpeechRecognition;if(!R){$("vozEstado").textContent="Navegador no compatible. Prueba Chrome o Edge.";return;}const r=new R();r.lang="es-CO";r.onresult=e=>{$("busqueda").value=e.results[0][0].transcript;$("busqueda").dispatchEvent(new Event("input"));$("vozEstado").textContent=interpretarBusqueda($("busqueda").value)?"Conversión reconocida":"No se reconoció el formato. Usa números y monedas.";};r.onerror=()=>{$("vozEstado").textContent="No se pudo usar el micrófono.";};try{r.start();$("vozEstado").textContent="Escuchando…";}catch{$("vozEstado").textContent="Micrófono no disponible.";}};
    let chart=null;
    $("grafica").onclick=async()=>{const b=$("grafica");b.disabled=true;$("grafEstado").textContent="Consultando…";if(chart){chart.destroy();chart=null;}try{const de=$("origen").value,a=$("destino").value;const r=await fetch(`/api/historia?de=${encodeURIComponent(de)}&a=${encodeURIComponent(a)}&dias=${$("periodo").value}`);const d=await r.json();if(!r.ok)throw Error(d.error);if(!window.Chart)throw Error("No se pudo cargar Chart.js");chart=new Chart($("grafico"),{type:"line",data:{labels:d.fechas,datasets:[{label:`${de} → ${a}`,data:d.valores,borderColor:"#54dccb",backgroundColor:"rgba(84,220,203,.14)",fill:true,pointRadius:1,tension:.2}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:"index",intersect:false}}});$("grafEstado").textContent=`${de} → ${a} · Datos históricos de Frankfurter`; }catch(e){$("grafEstado").textContent=e.message;}finally{b.disabled=false;}};
    async function exportar(tipo){try{const monto=leerNumero($("cantidad").value),de=$("origen").value,a=$("destino").value;if(monto===null||!tasas[de]||!tasas[a])throw Error("Conversión inválida");const r=await fetch("/api/exportar/"+tipo,{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({monto,de,a,comision:Number($("comision").value),historial:leerHistorial()})});if(!r.ok)throw Error("No se pudo exportar");const url=URL.createObjectURL(await r.blob());const link=document.createElement("a");link.href=url;link.download=tipo==="pdf"?"AMVEXA.pdf":"AMVEXA.xlsx";link.click();setTimeout(()=>URL.revokeObjectURL(url),1000);$("exportEstado").textContent="Archivo generado";}catch(e){$("exportEstado").textContent=e.message;}}
    $("pdf").onclick=()=>exportar("pdf");$("excel").onclick=()=>exportar("xlsx");

  </script>
</body>
</html>"""
@app.route("/")
def index():
    return PAGINA
@app.route("/api/tasas")
def api_tasas():
    try:
        tasas, desactualizado = cargar_tasas()
    except Exception:
        return jsonify({"error": "No se pudieron obtener las tasas"}), 503
    orden = list(MONEDAS)
    monedas = []
    for codigo in tasas:
        nombre, bandera = MONEDAS.get(codigo, (codigo, "🏳️"))
        monedas.append({"codigo": codigo, "nombre": nombre, "bandera": bandera})
    # Las más populares primero, el resto por orden alfabético
    monedas.sort(
        key=lambda m: (
            orden.index(m["codigo"]) if m["codigo"] in orden else 999,
            m["codigo"],
        )
    )
    return jsonify(
        {
            "tasas": tasas,
            "monedas": monedas,
            "fecha": _cache["fecha"],
            "desactualizado": desactualizado,
        }
    )

@app.get("/api/historia")
def historia():
    de, a = request.args.get("de", ""), request.args.get("a", "")
    dias = request.args.get("dias", "30")
    if not re.fullmatch(r"[A-Z]{3}", de) or not re.fullmatch(r"[A-Z]{3}", a) or dias not in {"30", "90", "365"}:
        return jsonify(error="Parámetros inválidos"), 400
    fin = date.today()
    inicio = fin - timedelta(days=int(dias))
    if de == a:
        return jsonify(fechas=[inicio.isoformat(), fin.isoformat()], valores=[1, 1])
    try:
        r = requests.get(f"https://api.frankfurter.dev/v1/{inicio}..{fin}", params={"base":de,"symbols":a}, timeout=12)
        r.raise_for_status()
        datos = r.json()["rates"]
        puntos = [(f, v[a]) for f,v in sorted(datos.items()) if a in v]
        if not puntos: raise ValueError()
        return jsonify(fechas=[p[0] for p in puntos], valores=[p[1] for p in puntos])
    except (requests.RequestException, ValueError, KeyError, TypeError):
        return jsonify(error="Histórico no disponible para este par o conexión."), 503

@app.post("/api/exportar/<tipo>")
def exportar(tipo):
    if tipo not in {"pdf", "xlsx"}: return jsonify(error="Formato inválido"), 400
    try:
        d = request.get_json()
        tasas, antigua = cargar_tasas()
        def fila(x):
            m = float(x["monto"]); de = x["de"]; a = x["a"]
            if not isfinite(m) or not 0 <= m <= 1e15 or de not in tasas or a not in tasas: raise ValueError()
            return [m, de, a, m * tasas[a] / tasas[de]]
        actual = fila(d)
        p = float(d.get("comision", 0))
        if not isfinite(p) or not 0 <= p <= 100: raise ValueError()
        salida = BytesIO()
        if tipo == "pdf":
            c = canvas.Canvas(salida)
            c.setTitle("AMVEXA - Comprobante")
            c.setFont("Helvetica-Bold", 24); c.drawString(50, 780, "AMVEXA")
            c.setFont("Helvetica", 12)
            lineas = ["Comprobante informativo - no es una transaccion", f"Cantidad: {actual[0]:,.2f} {actual[1]}", f"Resultado: {actual[3]:,.2f} {actual[2]}", f"Comision: {p:g}%", f"Neto: {actual[3]*(1-p/100):,.2f} {actual[2]}", f"Tasas: {_cache['fecha']}", f"Cache vencida: {'Si' if antigua else 'No'}", "Desarrollado por Ana Maria"]
            for i,linea in enumerate(lineas): c.drawString(50, 740-i*25, linea)
            c.save()
        else:
            w = Workbook(); h = w.active; h.title = "Conversiones"
            h.append(["Cantidad", "Origen", "Destino", "Resultado recalculado"])
            h.append(actual)
            for x in d.get("historial", [])[:6]: h.append(fila(x))
            h.append(["Tasas consultadas", str(_cache['fecha'])]); h.append(["Cache vencida", antigua]); h.freeze_panes="A2"
            for col in "ABCD": h.column_dimensions[col].width=25
            w.save(salida)
        salida.seek(0)
        return send_file(salida, as_attachment=True, download_name="AMVEXA."+tipo, mimetype="application/pdf" if tipo=="pdf" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except (ValueError, KeyError, TypeError, AttributeError): return jsonify(error="Datos inválidos"), 400
    except Exception: return jsonify(error="Exportación no disponible"), 503


if __name__ == "__main__":
    app.run(debug=False)
