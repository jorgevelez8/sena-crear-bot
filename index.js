'use strict';
require('dotenv').config();

const { Bot, InputFile } = require('grammy');
const Groq      = require('groq-sdk');
const axios     = require('axios');
const XLSX      = require('xlsx');
const fs        = require('fs');
const path      = require('path');
const http      = require('http');

// ── Clientes API ──────────────────────────────────────────
const bot       = new Bot(process.env.TELEGRAM_TOKEN);
const groq      = new Groq({ apiKey: process.env.GROQ_API_KEY });
const Anthropic = require('@anthropic-ai/sdk');
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ══════════════════════════════════════════════════════════
// ── SESIONES — Redis (persistente) con fallback en memoria
// ══════════════════════════════════════════════════════════
// PUNTO 2: Sesiones sobreviven reinicios de Render.
// Si UPSTASH_REDIS_REST_URL no está configurado, usa Map en memoria.

let _redis = null;
try {
  if (process.env.UPSTASH_REDIS_REST_URL) {
    const { Redis } = require('@upstash/redis');
    _redis = new Redis({
      url:   process.env.UPSTASH_REDIS_REST_URL,
      token: process.env.UPSTASH_REDIS_REST_TOKEN,
    });
    console.log('✅ Redis Upstash conectado — sesiones persistentes');
  } else {
    console.warn('⚠️  UPSTASH_REDIS_REST_URL no configurado — sesiones en memoria (se pierden al reiniciar)');
  }
} catch (e) {
  console.error('Redis init error:', e.message);
}

const _planMem = new Map();
const _dxMem   = new Map();
const TTL       = 172800; // 48 horas en segundos

async function getSesion(chatId) {
  if (_redis) {
    try {
      const d = await _redis.get(`plan:${chatId}`);
      if (d) return typeof d === 'string' ? JSON.parse(d) : d;
    } catch (e) { console.error('redis get plan:', e.message); }
  }
  if (!_planMem.has(chatId)) _planMem.set(chatId, { paso: 0, datos: {} });
  return _planMem.get(chatId);
}

async function setSesion(chatId, sesion) {
  _planMem.set(chatId, sesion);
  if (_redis) {
    try { await _redis.setex(`plan:${chatId}`, TTL, JSON.stringify(sesion)); }
    catch (e) { console.error('redis set plan:', e.message); }
  }
}

async function delSesion(chatId) {
  _planMem.delete(chatId);
  if (_redis) {
    try { await _redis.del(`plan:${chatId}`); } catch {}
  }
}

async function getDxSesion(chatId) {
  if (_redis) {
    try {
      const d = await _redis.get(`dx:${chatId}`);
      if (d) return typeof d === 'string' ? JSON.parse(d) : d;
    } catch (e) { console.error('redis get dx:', e.message); }
  }
  if (!_dxMem.has(chatId)) _dxMem.set(chatId, { paso: 0, datos: { scores: Array(40).fill(null) } });
  return _dxMem.get(chatId);
}

async function setDxSesion(chatId, sesion) {
  _dxMem.set(chatId, sesion);
  if (_redis) {
    try { await _redis.setex(`dx:${chatId}`, TTL, JSON.stringify(sesion)); }
    catch (e) { console.error('redis set dx:', e.message); }
  }
}

async function delDxSesion(chatId) {
  _dxMem.delete(chatId);
  if (_redis) {
    try { await _redis.del(`dx:${chatId}`); } catch {}
  }
}

async function hasDxSesion(chatId) {
  if (_dxMem.has(chatId)) return true;
  if (_redis) {
    try { return (await _redis.exists(`dx:${chatId}`)) > 0; } catch {}
  }
  return false;
}

// ══════════════════════════════════════════════════════════
// ── PREGUNTAS (~41, cubre formato oficial SENA) ───────────
// ══════════════════════════════════════════════════════════
const PREGUNTAS = [
  // === Datos del beneficiario ===
  {
    key: 'nombre',
    msg: '¡Hola! 👋 Vamos a crear el *Plan de Negocio SENA Línea CREAR*.\nPuedes responder con voz 🎤 o escribiendo.\n\n¿Cuál es el *nombre completo* del beneficiario?',
  },
  { key: 'tipoDoc',      msg: '¿*Tipo de documento*?\n_(CC Cédula · CE Cédula Extranjería · PA Pasaporte · TI Tarjeta Identidad)_' },
  { key: 'numDoc',       msg: '¿*Número de documento*?' },
  { key: 'genero',       msg: '¿*Género*?\n_(Masculino / Femenino / Otro)_' },
  { key: 'departamento', msg: '¿En qué *departamento* está el proyecto?' },
  { key: 'municipio',    msg: '¿En qué *municipio*?' },
  {
    key: 'grupoPoblacional',
    msg: '¿A qué *grupo poblacional* pertenece?\n_(Ej: Víctima de violencia, Desplazado, Campesino, Indígena, Afrocolombiano, LGBTI, Discapacidad...)_',
  },
  { key: 'nombreProyecto', msg: '¿Cuál es el *nombre del proyecto*?' },
  { key: 'tipoProyecto',   msg: '¿Es *Economía Popular* o *Economía Campesina*?' },
  {
    key: 'sector',
    msg: '¿En qué *sector* trabaja?\n_(Ej: Agricultura, Comercio, Artesanías, Turismo, Manufactura, Alimentos...)_',
  },
  {
    key: 'ciiu',
    msg: '¿Cuál es la *actividad económica* (código CIIU)?\n_(Ej: 0111 Cultivo de cereales · 4711 Tienda alimentos · 1411 Confección ropa...)_',
  },
  { key: 'asociativo',  msg: '¿El proyecto es *asociativo* (varias personas juntas)?\n_Responda: SI o NO_' },
  {
    key: 'numPersonas',
    msg: '¿*Cuántas personas* conforman el grupo asociativo?',
    soloSi: d => /^s/i.test(d.asociativo || ''),
    numerico: true,
  },
  { key: 'lugarOps', msg: '¿Ya *cuenta con un lugar de operaciones*?\n_Responda: SI o NO_' },

  // === Sección 1 — Cliente ===
  {
    key: 'clienteCarac',
    msg: '¿*Quién le compra*?\nDescríbame al cliente: dónde vive, estrato, edad, género, cuánto gana, si trabaja...',
  },
  {
    key: 'clienteCual',
    msg: '¿Cuáles son los *gustos, preferencias y cualidades* de ese cliente?\n_(Qué le gusta, qué valora, cómo decide comprar)_',
  },

  // === Sección 2 — Problema ===
  {
    key: 'problema',
    msg: '¿Qué *problema o necesidad* resuelve su negocio?\n¿Por qué sus clientes lo necesitan?',
  },

  // === Sección 3 — Competencia ===
  {
    key: 'competidor1',
    msg: '¿Quién es su *primer competidor*?\nDígame: nombre, dónde está, qué vende, a qué precio, qué ventajas tiene y qué desventajas comparado con usted.',
  },
  { key: 'competidor2', msg: '¿Hay un *segundo competidor*?\n_(Si no hay, diga "no hay")_' },
  { key: 'competidor3', msg: '¿Y un *tercer competidor*?\n_(Si no hay, diga "no hay")_' },

  // === Sección 4 — Descripción y propuesta de valor ===
  { key: 'descripcion', msg: '¿En qué *consiste su proyecto*?\nCuénteme con sus propias palabras qué va a hacer y qué lo hace diferente.' },
  { key: 'pvNuestro',   msg: 'Propuesta de valor:\n*"Nuestro producto/servicio es..."*\n_(Complete la frase: describa qué ofrece)_' },
  { key: 'pvAyuda',     msg: '*"...que ayuda a..."*\n_(¿A quién ayuda? Tipo de persona o empresa)_' },
  { key: 'pvQue',       msg: '*"...a que..."*\n_(¿Qué logran sus clientes con su producto?)_' },
  { key: 'pvMediante',  msg: '*"...mediante..."*\n_(¿Cómo lo logran? ¿Qué hace diferente su negocio?)_' },

  // === Sección 5 — Productos o servicios ===
  { key: 'prod1Nombre',     msg: '¿Cuál es el *nombre del producto o servicio principal*?' },
  { key: 'prod1Desc',       msg: '¿Cómo lo *describiría*? ¿Qué es exactamente?' },
  { key: 'prod1Unidad',     msg: '¿Cuál es la *unidad de medida*?\n_(Ej: Kilogramo, Litro, Unidad, Hora, Docena, Porción...)_' },
  { key: 'prod1Precio',     msg: '¿A qué *precio* lo vende?\n_(Número en pesos, ej: 15000 o "quince mil pesos")_', numerico: true },
  { key: 'prod1UnidadesMes',msg: '¿Cuántas *unidades vende al mes* aproximadamente?\n_(Número)_', numerico: true },
  { key: 'prod1Costo',      msg: '¿Cuánto le *cuesta producir* cada unidad?\n_(Materias primas e insumos — número en pesos)_', numerico: true },

  // === Sección 10 — Costos fijos ===
  {
    key: 'costosFijosDesc',
    msg: '¿Cuáles son sus *costos fijos mensuales*?\nCuénteme los gastos que paga todos los meses aunque no venda nada: arriendo, servicios, internet, transporte...',
  },
  { key: 'costosFijosTotal', msg: '¿Cuánto suman esos costos fijos *en total al mes*?\n_(Número en pesos)_', numerico: true },

  // === Sección 11 — Mano de obra ===
  // PUNTO 5: Alimenta D16 del MODELO FINANCIERO (filas 388-393 del template)
  {
    key: 'manoObraTotal',
    msg: '¿Paga *sueldos o jornales* a empleados?\n¿Cuánto paga en total al mes en sueldos?\n_(Si trabaja solo sin sueldo fijo, diga *cero*)_',
    numerico: true,
  },

  // === Sección 15 — Gastos de administración y ventas ===
  // PUNTO 5: Alimenta D20 del MODELO FINANCIERO (filas 357-366 del template, col M y U)
  {
    key: 'gastosAdmin',
    msg: '¿Cuánto gasta al mes en *administración y ventas*?\n_(Publicidad, internet, papelería, transporte de ventas, contador, delivery...)_\n_Si no tiene, diga *cero*_',
    numerico: true,
  },

  // === Sección 7 — Permisos y licencias ===
  // PUNTO 5: Alimenta D18 del MODELO FINANCIERO (Z184 del template)
  {
    key: 'permisosTotal',
    msg: '¿Cuánto *invirtió o necesita invertir* en permisos, licencias o registros?\n_(RUT, Cámara de Comercio, INVIMA, sanidad, registro de marca...)_\n_Si no tiene, diga *cero*_',
    numerico: true,
  },

  // === Sección 12 — Inversión ===
  { key: 'inversion',    msg: '¿Cuánto necesita *invertir en total*?\n_(Maquinaria, equipos, adecuaciones, materias primas iniciales — número)_', numerico: true },
  { key: 'aportePropio', msg: '¿Cuánto puede *aportar usted* de esa inversión?\n_(Mínimo el 10% del total — número)_', numerico: true },
  {
    key: 'inversionDesc',
    msg: '¿En qué va a *usar el dinero del Fondo Emprender*?\n_(Ej: máquina de coser $500.000, materias primas $300.000, adecuaciones $200.000)_',
  },

  // === Impacto ===
  { key: 'impactoEco',    msg: '¿Cómo *beneficia económicamente* este proyecto a su comunidad o región?\n_(Empleos que genera, ingresos, encadenamientos...)_' },
  { key: 'impactoSocial', msg: '¿Qué *impacto social* tiene el proyecto?\n_(A cuántas familias ayuda, qué cambia en su comunidad)_' },
];

// ── Claves de preguntas opcionales (se pueden saltar con "saltar"/"s") ──────
// Todas las demás son obligatorias para el formato SENA oficial
const CLAVES_OPCIONALES = new Set([
  'genero', 'grupoPoblacional', 'ciiu', 'asociativo', 'numPersonas',
  'lugarOps', 'clienteCual', 'competidor2', 'competidor3',
  'pvNuestro', 'pvAyuda', 'pvQue', 'pvMediante',
  'prod1Desc', 'prod1Unidad',
  'costosFijosDesc', 'gastosAdmin', 'permisosTotal',
  // manoObraTotal y costosFijosTotal NO son opcionales — alimentan D16/D17 del MODELO FINANCIERO
  'impactoEco', 'impactoSocial',
]);

// Valor por defecto cuando se salta una pregunta opcional
const DEFAULTS_OPCIONALES = {
  asociativo:  'NO',
  prod1Unidad: 'Unidad',
};

// ── Campos obligatorios para /listo ───────────────────────
// PUNTO 3: /listo bloqueado si alguno falta
const CAMPOS_REQUERIDOS = [
  { key: 'nombre',          label: 'Nombre completo' },
  { key: 'numDoc',          label: 'Número de documento' },
  { key: 'departamento',    label: 'Departamento' },
  { key: 'municipio',       label: 'Municipio' },
  { key: 'nombreProyecto',  label: 'Nombre del proyecto' },
  { key: 'tipoProyecto',    label: 'Tipo de proyecto' },
  { key: 'sector',          label: 'Sector' },
  { key: 'clienteCarac',    label: 'Descripción del cliente' },
  { key: 'problema',        label: 'Problema que resuelve' },
  { key: 'prod1Nombre',     label: 'Nombre del producto' },
  { key: 'prod1Precio',     label: 'Precio' },
  { key: 'prod1UnidadesMes',label: 'Unidades por mes' },
  { key: 'prod1Costo',      label: 'Costo por unidad' },
  { key: 'inversion',       label: 'Inversión total' },
  { key: 'aportePropio',    label: 'Aporte propio' },
  { key: 'inversionDesc',   label: 'Uso del dinero FE' },
];

function validarCamposObligatorios(datos) {
  const faltantes = CAMPOS_REQUERIDOS
    .filter(c => !datos[c.key] || datos[c.key].trim() === '')
    .map(c => c.label);
  return faltantes; // [] = todo bien
}

// ── Helpers básicos ───────────────────────────────────────
function limpiarNombre(texto) {
  return texto
    .replace(/^(mi nombre (completo )?es|me llamo|yo me llamo|soy|yo soy|me dicen)\s+/i, '')
    .trim();
}

function getPregunta(paso, datos) {
  for (let i = paso; i < PREGUNTAS.length; i++) {
    const p = PREGUNTAS[i];
    if (!p.soloSi || p.soloSi(datos)) return { p, i };
  }
  return null;
}

function fmt(n) {
  return '$' + Math.round(Number(n) || 0).toLocaleString('es-CO');
}

function maletin(n) {
  const p = Number(n) || 1;
  return p >= 6 ? 10000000 : p >= 4 ? 7000000 : p >= 2 ? 5000000 : 2000000;
}

function getResumenItems(sesion) {
  const items = [];
  let num = 1;
  for (const p of PREGUNTAS) {
    if (p.soloSi && !p.soloSi(sesion.datos)) continue;
    const label = p.msg.match(/\*([^*]+)\*/)?.[1] || p.key;
    items.push({ num, key: p.key, label, valor: sesion.datos[p.key] || '' });
    num++;
  }
  return items;
}

async function mostrarResumen(ctx, sesion) {
  const items = getResumenItems(sesion);
  const lineas = ['📋 *Revisa las respuestas antes de generar el plan:*\n'];
  for (const it of items) {
    const val = it.valor.length > 55 ? it.valor.slice(0, 55) + '…' : it.valor;
    lineas.push(`*${it.num}.* ${it.label}: _${val || '–'}_`);
  }
  lineas.push('\n✅ Escribe */listo* para generar el plan completo');
  lineas.push('✏️ Escribe el *número* de lo que quieras corregir _(ej: 3)_');
  await ctx.reply(lineas.join('\n'), { parse_mode: 'Markdown' });
}

function getLastAnsweredIndex(sesion) {
  for (let i = sesion.paso - 1; i >= 0; i--) {
    if (sesion.datos[PREGUNTAS[i]?.key] !== undefined) return i;
  }
  return -1;
}

// ══════════════════════════════════════════════════════════
// ── PARSEO CON CLAUDE ─────────────────────────────────────
// ══════════════════════════════════════════════════════════

// PUNTO 4: extraerNumero devuelve null en fallo (no 0 silencioso)
async function extraerNumero(texto) {
  const t = String(texto).trim();
  // Fast-path: ya es número
  const clean = t.replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  if (/^\d+(\.\d+)?$/.test(clean)) return parseFloat(clean);
  // Expresiones de cero legítimas
  if (/^(cero|nada|no\s+ten|no\s+hay|ninguno|0)$/i.test(t)) return 0;
  // Llamar Haiku
  try {
    const msg = await anthropic.messages.create({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 15,
      messages:   [{ role: 'user', content: `Número del texto (solo dígitos enteros, sin puntos ni comas): "${texto}"` }],
    });
    const n = Number(msg.content[0].text.replace(/[^\d]/g, ''));
    return isNaN(n) ? null : n; // null = no se pudo → caller pide repetir
  } catch {
    return null; // API caída → null para que caller informe al usuario
  }
}

async function parsearCompetidor(texto) {
  if (!texto || /^no\s*hay/i.test(texto.trim())) return null;
  try {
    const msg = await anthropic.messages.create({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 200,
      messages:   [{
        role: 'user',
        content: `Del siguiente texto sobre un competidor extrae y responde SOLO un JSON con estos campos (strings): nombre, localizacion, producto, precio, ventajas, desventajas. Si no encuentras un campo deja "". Sin texto extra.\nTexto: "${texto}"`,
      }],
    });
    const raw = msg.content[0].text.trim().replace(/```[\w]*\n?/g, '').replace(/```/g, '');
    return JSON.parse(raw);
  } catch {
    return { nombre: texto.slice(0, 60), localizacion: '', producto: '', precio: '', ventajas: '', desventajas: '' };
  }
}

async function parsearCostosFijos(texto, total) {
  if (!texto) return [{ descripcion: 'Costos fijos', valorMensual: total }];
  try {
    const msg = await anthropic.messages.create({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages:   [{
        role: 'user',
        content: `Lista de costos fijos del texto. Solo JSON array sin texto extra:\n[{"descripcion":"...","valorMensual":0},...]\nTexto: "${texto}"\nTotal mensual: ${total}`,
      }],
    });
    const raw = msg.content[0].text.trim().replace(/```[\w]*\n?/g, '').replace(/```/g, '');
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [{ descripcion: texto.slice(0, 50), valorMensual: total }];
  } catch {
    return [{ descripcion: texto.slice(0, 60), valorMensual: total }];
  }
}

async function parsearInversion(texto, total) {
  if (!texto) return [{ descripcion: 'Inversión', cantidad: 1, valor: total }];
  try {
    const msg = await anthropic.messages.create({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages:   [{
        role: 'user',
        content: `Lista de inversiones del texto. Solo JSON array sin texto extra:\n[{"descripcion":"...","cantidad":1,"valor":0},...]\nTexto: "${texto}"\nTotal: ${total}`,
      }],
    });
    const raw = msg.content[0].text.trim().replace(/```[\w]*\n?/g, '').replace(/```/g, '');
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [{ descripcion: texto.slice(0, 50), cantidad: 1, valor: total }];
  } catch {
    return [{ descripcion: texto.slice(0, 60), cantidad: 1, valor: total }];
  }
}

async function generarPlanCompleto(datos) {
  const d = datos;
  const p   = Number(d.prod1Precio)      || 0;
  const u   = Number(d.prod1UnidadesMes) || 0;
  const cv  = Number(d.prod1Costo)       || 0;
  const cf  = Number(d.costosFijosTotal) || 0;
  const inv = Number(d.inversion)        || 0;
  const ap  = Number(d.aportePropio)     || 0;
  const np  = Number(d.numPersonas)      || 1;

  const prompt = `Eres un asesor del SENA Colombia experto en planes de negocio para la Línea CREAR Especial (víctimas del conflicto, campesinos, desplazados). Redacta en español claro, formal y empático. Máximo 200 palabras por sección. Responde SOLO con el JSON sin texto adicional ni bloques de código.

DATOS DEL BENEFICIARIO:
- Nombre: ${d.nombre} (${d.tipoDoc} ${d.numDoc}) — ${d.genero}
- Grupo: ${d.grupoPoblacional}
- Ubicación: ${d.municipio}, ${d.departamento}
- Proyecto: ${d.nombreProyecto} (${d.tipoProyecto} · ${d.sector} · CIIU: ${d.ciiu})
- Modalidad: ${/^s/i.test(d.asociativo || '') ? `Asociativo ${np} personas` : 'Individual'}
- Descripción: ${d.descripcion}
- Propuesta de valor: Nuestro ${d.pvNuestro} ayuda a ${d.pvAyuda} a que ${d.pvQue} mediante ${d.pvMediante}
- Producto: ${d.prod1Nombre} (${d.prod1Unidad}) · Precio: $${p.toLocaleString('es-CO')} · ${u} unidades/mes
- Cliente: ${d.clienteCarac}
- Problema que resuelve: ${d.problema}
- Competidor 1: ${d.competidor1 || 'no especificado'}
- Costo unitario: $${cv.toLocaleString('es-CO')} · Costos fijos: $${cf.toLocaleString('es-CO')}/mes
- Inversión total: $${inv.toLocaleString('es-CO')} · Aporte propio: $${ap.toLocaleString('es-CO')}
- Uso del dinero: ${d.inversionDesc}
- Impacto económico: ${d.impactoEco}
- Impacto social: ${d.impactoSocial}

{"descripcionNegocio":"...","mercadoObjetivo":"...","analisisCompetencia":"...","estrategiaComercial":"...","planOperativo":"...","analisisRiesgos":"...","justificacionInversion":"..."}`;

  const msg = await anthropic.messages.create({
    model:      'claude-haiku-4-5-20251001',
    max_tokens: 3500,
    messages:   [{ role: 'user', content: prompt }],
  });

  const raw = msg.content[0].text.trim()
    .replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  return JSON.parse(raw);
}

// ══════════════════════════════════════════════════════════
// ── EXCEL OFICIAL SENA ────────────────────────────────────
// ══════════════════════════════════════════════════════════
// PUNTO 5: Incluye mano de obra (D16), permisos (D18) y G&A (D20)
async function generarExcelOficial(datos, costosFijosItems, inversionItems, comps) {
  const wb = XLSX.readFile(path.join(__dirname, 'template.xlsx'), { cellFormula: true, cellStyles: true });
  const ws = wb.Sheets['PROYECTO'];

  function set(addr, value) {
    if (value === null || value === undefined || value === '') return;
    const t = typeof value === 'number' ? 'n' : 's';
    ws[addr] = { t, v: value, w: String(value) };
  }

  const d   = datos;
  const p1  = Number(d.prod1Precio)      || 0;
  const u1  = Number(d.prod1UnidadesMes) || 0;
  const cv1 = Number(d.prod1Costo)       || 0;
  const inv = Number(d.inversion)        || 0;
  const ap  = Number(d.aportePropio)     || 0;
  const np  = Number(d.numPersonas)      || 1;
  const u1a2 = Math.round(u1 * 1.10);
  const u1a3 = Math.round(u1 * 1.20);

  // ── Datos básicos ──
  set('B10',  d.nombre || '');
  set('T10',  d.tipoDoc || 'CC');
  set('AC10', d.numDoc || '');
  set('B14',  d.nombreProyecto || '');
  set('T14',  d.departamento || '');
  set('AC14', d.municipio || '');
  set('B18',  d.tipoProyecto || '');
  set('L18',  d.sector || '');
  set('Z18',  d.ciiu || '');
  set('R21', /^s/i.test(d.asociativo || '') ? 'SI' : 'NO');
  if (/^s/i.test(d.asociativo || '')) set('AN21', np);
  set('R23', /^s/i.test(d.lugarOps || '') ? 'SI' : 'NO');

  // ── Cliente ──
  set('A61', d.clienteCarac || '');
  set('S61', d.clienteCual  || '');

  // ── Problema ──
  set('A77', d.problema || '');

  // ── Competidores (filas 94-96) ──
  for (let i = 0; i < 3; i++) {
    const comp = comps[i];
    if (!comp) continue;
    const r = 94 + i;
    set(`A${r}`,  comp.nombre       || '');
    set(`I${r}`,  comp.localizacion || '');
    set(`R${r}`,  comp.producto     || '');
    set(`X${r}`,  comp.precio       || '');
    set(`AC${r}`, comp.ventajas     || '');
    set(`AK${r}`, comp.desventajas  || '');
  }

  // ── Descripción y propuesta de valor ──
  set('I102', d.descripcion || '');
  set('G105', d.pvNuestro   || '');
  set('G106', d.pvAyuda     || '');
  set('G107', d.pvQue       || '');
  set('G108', d.pvMediante  || '');

  // ── Productos (fila 116) ──
  set('A116', d.prod1Nombre || '');
  set('K116', d.prod1Desc   || '');
  set('Y116', d.prod1Unidad || '');

  // ── Precios (fila 229) — solo Q229 (año 1); W229/AC229 son fórmulas ──
  set('Q229', p1);

  // ── Unidades mensuales (filas 241-252) ──
  for (let mes = 0; mes < 12; mes++) {
    const row = 241 + mes;
    set(`G${row}`,  u1);
    set(`S${row}`,  u1a2);
    set(`AE${row}`, u1a3);
  }

  // ── Costos fijos (filas 297-304): A=desc · T=valor mensual · AA=12 meses ──
  // AE = fórmula T*AA → NO escribir
  for (let i = 0; i < Math.min(costosFijosItems.length, 8); i++) {
    const row  = 297 + i;
    const item = costosFijosItems[i];
    set(`A${row}`,  item.descripcion || '');
    set(`T${row}`,  Number(item.valorMensual) || 0);
    set(`AA${row}`, 12);
  }

  // ── Costos variables producto 1 (fila 318): A=desc · K=unidad · R=valor unit · AA=cant ──
  // AE = fórmula R*AA → NO escribir
  set('A318', 'Materias primas e insumos');
  set('K318', 'Unidad');
  set('R318', cv1);
  set('AA318', 1);

  // ── % Participación producto 1 (fila 349) ──
  set('Q349', 100);

  // ── PUNTO 5A: Mano de obra (fila 389) — alimenta D16 del MODELO FINANCIERO ──
  // D16 = AK389*AP389. AK389 = IF(AH389="SI", AB389*(1+I383), AB389)
  // AB389=sueldo mensual · AH389="NO"(sin prestaciones para emprendedor informal) · AP389=12 meses
  {
    const manoObra = Number(d.manoObraTotal) || 0;
    if (manoObra > 0) {
      set('A389',  d.manoObraDesc || 'Empleado/Operario');
      set('G389',  'Empleo directo');
      set('AB389', manoObra);
      set('AH389', 'NO'); // sin factor prestacional (autoempleado informal)
      set('AP389', 12);   // 12 meses · D16 = AK389*AP389 = manoObra*12
    }
  }

  // ── PUNTO 5B: Permisos y licencias (fila 184, col Z) — alimenta D18 ──
  // Z190 = SUM(Z184:AF189). Z183=Costo($) es encabezado. Z184=primer ítem.
  {
    const permisos = Number(d.permisosTotal) || 0;
    if (permisos > 0) {
      set('Z184', permisos); // costo total en permiso/licencia fila 1
    }
  }

  // ── PUNTO 5C: Gastos de administración y ventas (fila 358) — alimenta D20 ──
  // Y367 = SUM(Y357:AD366). Y358 = M358*U358. M=valor mensual · U=meses(12).
  {
    const gastos = Number(d.gastosAdmin) || 0;
    if (gastos > 0) {
      set('A358', 'Gastos de administración y ventas');
      set('M358', gastos);
      set('U358', 12); // Y358 = M358*U358 = gastos*12 (auto-calcula)
    }
  }

  // ── Inversiones fijas (filas 403-422) ──
  // Fila 402 = encabezado. T=valor unit · AA=cant. AD=fórmula T*AA → NO escribir
  // AM="X" FE · AO="X" emprendedor. AE448/449/450 = fórmulas → NO escribir
  {
    const items   = inversionItems.slice(0, 19);
    const sumItems = items.reduce((s, it) => s + (Number(it.valor) || 0), 0);
    const diff     = inv - sumItems;

    for (let i = 0; i < items.length; i++) {
      const row  = 403 + i;
      const item = items[i];
      let valor  = Number(item.valor) || 0;
      if (i === items.length - 1 && Math.abs(diff) > 0) valor += diff;
      set(`A${row}`,  item.descripcion || '');
      set(`T${row}`,  valor);
      set(`AA${row}`, Number(item.cantidad) || 1);
      set(`AM${row}`, 'X'); // FE por defecto
    }

    // Aporte propio: fila extra marcada como emprendedor
    if (ap > 0 && items.length < 20) {
      const rowAp = 403 + items.length;
      set(`A${rowAp}`,  'Capital propio del emprendedor');
      set(`T${rowAp}`,  ap);
      set(`AA${rowAp}`, 1);
      set(`AO${rowAp}`, 'X'); // emprendedor
      // Reducir primer ítem FE para que el total cuadre
      const cell0 = ws['T403'];
      if (cell0 && cell0.v) {
        ws['T403'] = { t: 'n', v: Math.max(0, Number(cell0.v) - ap), w: String(Math.max(0, Number(cell0.v) - ap)) };
      }
    }
  }

  // ── Avances ──
  set('G459', 'En proceso');
  set('G460', 'Identificado');

  // ── Impacto ──
  set('K487', d.impactoEco    || '');
  set('K489', d.impactoSocial || '');

  return wb;
}

// ── Transcripción Groq Whisper ────────────────────────────
async function transcribir(fileId) {
  const info = await bot.api.getFile(fileId);
  const url  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${info.file_path}`;
  const resp = await axios.get(url, { responseType: 'arraybuffer' });

  const tmp = `/tmp/audio_${Date.now()}.ogg`;
  fs.writeFileSync(tmp, Buffer.from(resp.data));

  try {
    const result = await groq.audio.transcriptions.create({
      file:     fs.createReadStream(tmp),
      model:    'whisper-large-v3',
      language: 'es',
    });
    return result.text.trim();
  } finally {
    try { fs.unlinkSync(tmp); } catch {}
  }
}

// ── Finalizar: generar Excels y enviar ───────────────────
async function finalizar(ctx, sesion) {
  await ctx.reply('🎉 *¡Plan completado!* Procesando con IA… ⏳\n_Esto puede tardar 30-60 segundos._', { parse_mode: 'Markdown' });

  const d   = sesion.datos;
  const p   = Number(d.prod1Precio)      || 0;
  const u   = Number(d.prod1UnidadesMes) || 0;
  const cv  = Number(d.prod1Costo)       || 0;
  const cf  = Number(d.costosFijosTotal) || 0;
  const inv = Number(d.inversion)        || 0;
  const ap  = Number(d.aportePropio)     || 0;
  const np  = Number(d.numPersonas)      || 1;

  const [comp1, comp2, comp3, costosFijosItems, inversionItems] = await Promise.all([
    parsearCompetidor(d.competidor1),
    parsearCompetidor(d.competidor2),
    parsearCompetidor(d.competidor3),
    parsearCostosFijos(d.costosFijosDesc, cf),
    parsearInversion(d.inversionDesc, inv),
  ]);
  const comps = [comp1, comp2, comp3];

  let planIA = null;
  try {
    planIA = await generarPlanCompleto(d);
  } catch (e) {
    console.error('Claude plan error:', e.message);
    await ctx.reply('⚠️ No pude generar las secciones narrativas, pero el Excel oficial sí va a salir.');
  }

  const fecha  = new Date().toISOString().slice(0, 10);
  const nombre = (d.nombre || 'beneficiario').replace(/\s+/g, '_');

  let wbOficial;
  try {
    wbOficial = await generarExcelOficial(d, costosFijosItems, inversionItems, comps);
  } catch (e) {
    console.error('Excel oficial error:', e.message);
    await ctx.reply('❌ No pude generar el Excel oficial SENA. Revisa el template.xlsx en el servidor.');
    return;
  }

  const fnOficial  = `PROYECTO_CREAR_${nombre}_${fecha}.xlsx`;
  const tmpOficial = `/tmp/${fnOficial}`;
  XLSX.writeFile(wbOficial, tmpOficial);

  await ctx.replyWithDocument(new InputFile(tmpOficial, fnOficial), {
    caption: [
      '📊 *Excel Oficial SENA — Línea CREAR Especial*',
      `👤 ${d.nombre || ''}`,
      `🏭 ${d.nombreProyecto || ''}`,
      `📅 ${fecha}`,
      '',
      '_Abrir en Excel para ver fórmulas del Modelo Financiero_',
    ].join('\n'),
    parse_mode: 'Markdown',
  });
  try { fs.unlinkSync(tmpOficial); } catch {}

  if (planIA) {
    const ven1 = p * u * 12;
    const ven2 = Math.round(p * 1.04 * (u * 1.1) * 12);
    const ven3 = Math.round(p * 1.08 * (u * 1.2) * 12);
    const mc   = p > 0 ? (p - cv) / p : 0;
    const cf12 = cf * 12;
    const ptoEq = mc > 0 ? Math.ceil(cf12 / (mc * p)) : 0;
    const fe    = Math.max(0, inv - ap);
    const apPct = inv > 0 ? (ap / inv * 100) : 0;
    const kit   = maletin(np);

    const wb2 = XLSX.utils.book_new();

    const ws1 = XLSX.utils.aoa_to_sheet([
      ['PLAN DE NEGOCIO — SENA LÍNEA CREAR ESPECIAL'],
      ['Bot de voz · ' + fecha],
      [],
      ['BENEFICIARIO'],
      ['Nombre completo',     d.nombre          || ''],
      ['Tipo / N° documento', (d.tipoDoc || 'CC') + ' ' + (d.numDoc || '')],
      ['Género',              d.genero          || ''],
      ['Municipio',           d.municipio       || ''],
      ['Departamento',        d.departamento    || ''],
      ['Grupo poblacional',   d.grupoPoblacional || ''],
      [],
      ['PROYECTO'],
      ['Nombre del proyecto', d.nombreProyecto  || ''],
      ['Tipo de proyecto',    d.tipoProyecto    || ''],
      ['Sector / CIIU',       (d.sector || '') + ' · ' + (d.ciiu || '')],
      ['Modalidad',           /^s/i.test(d.asociativo || '') ? `Asociativo ${np} personas` : 'Individual'],
      ['Maletín de formación', fmt(kit)],
      ['Lugar de operaciones', /^s/i.test(d.lugarOps || '') ? 'SÍ' : 'NO'],
      [],
      ['PROPUESTA DE VALOR'],
      ['Descripción',  d.descripcion || ''],
      ['Nuestro:',     d.pvNuestro   || ''],
      ['Ayuda a:',     d.pvAyuda     || ''],
      ['A que:',       d.pvQue       || ''],
      ['Mediante:',    d.pvMediante  || ''],
      [],
      ['PRODUCTO PRINCIPAL'],
      ['Nombre',       d.prod1Nombre || ''],
      ['Descripción',  d.prod1Desc   || ''],
      ['Unidad',       d.prod1Unidad || ''],
    ]);
    ws1['!cols'] = [{ wch: 28 }, { wch: 65 }];
    XLSX.utils.book_append_sheet(wb2, ws1, 'DATOS');

    const ws2 = XLSX.utils.aoa_to_sheet([
      ['MODELO FINANCIERO'],
      [],
      ['VENTAS'],
      ['Producto',                          d.prod1Nombre || ''],
      ['Precio Año 1',                      p],
      ['Unidades/mes',                      u],
      ['Ventas Año 1',                      ven1],
      ['Ventas Año 2 (+4% precio +10% u.)', ven2],
      ['Ventas Año 3 (+8% precio +20% u.)', ven3],
      [],
      ['COSTOS'],
      ['Costo variable/unidad',             cv],
      ['Costos fijos/mes',                  cf],
      ['Mano de obra/mes',                  Number(d.manoObraTotal) || 0],
      ['Gastos admin y ventas/mes',         Number(d.gastosAdmin)   || 0],
      ['Permisos y licencias (total)',      Number(d.permisosTotal) || 0],
      [],
      ['RESULTADOS AÑO 1'],
      ['Margen bruto',                      ven1 - cv * u * 12],
      ['EBITDA',                            ven1 - cv * u * 12 - cf12],
      ['Punto de equilibrio (unid./año)',   ptoEq],
      [],
      ['INVERSIÓN Y FONDO EMPRENDER'],
      ['Inversión total',                   inv],
      ['Aporte emprendedor',                ap],
      ['% Aporte',                          apPct.toFixed(1) + '% ' + (apPct >= 10 ? '✓ CUMPLE' : '✗ NO CUMPLE (mín. 10%)')],
      ['Solicitado al Fondo Emprender',     fe],
      ['Maletín de formación',              kit],
    ]);
    ws2['!cols'] = [{ wch: 40 }, { wch: 22 }];
    XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO FINANCIERO');

    const secciones = [
      ['DESCRIPCIÓN DEL NEGOCIO',      planIA.descripcionNegocio],
      ['MERCADO OBJETIVO',              planIA.mercadoObjetivo],
      ['ANÁLISIS DE COMPETENCIA',       planIA.analisisCompetencia],
      ['ESTRATEGIA COMERCIAL',          planIA.estrategiaComercial],
      ['PLAN OPERATIVO',                planIA.planOperativo],
      ['ANÁLISIS DE RIESGOS',           planIA.analisisRiesgos],
      ['JUSTIFICACIÓN DE LA INVERSIÓN', planIA.justificacionInversion],
    ];
    const filas = [['PLAN NARRATIVO — SENA LÍNEA CREAR'], ['Generado con IA · ' + fecha], []];
    for (const [titulo, contenido] of secciones) {
      filas.push([titulo]);
      filas.push([contenido || '']);
      filas.push([]);
    }
    const ws3 = XLSX.utils.aoa_to_sheet(filas);
    ws3['!cols'] = [{ wch: 120 }];
    XLSX.utils.book_append_sheet(wb2, ws3, 'PLAN NARRATIVO');

    const fnPlan  = `PlanNarrado_${nombre}_${fecha}.xlsx`;
    const tmpPlan = `/tmp/${fnPlan}`;
    XLSX.writeFile(wb2, tmpPlan);
    await ctx.replyWithDocument(new InputFile(tmpPlan, fnPlan), {
      caption: '📝 *Plan Narrativo + Modelo Financiero*\n_Secciones redactadas por IA para el expediente_',
      parse_mode: 'Markdown',
    });
    try { fs.unlinkSync(tmpPlan); } catch {}
  }

  // Guardar plan con PIN — accesible desde cualquier chat con el PIN
  if (_redis) {
    try {
      const pin = String(Math.floor(1000 + Math.random() * 9000));
      await _redis.setex(`plan:pin:${pin}`, 30 * 24 * 3600, JSON.stringify(datos));
      await ctx.reply(
        `🔑 *PIN para retomar este plan:* \`${pin}\`\n\n` +
        '_Guárdalo. Cualquier dinamizador puede usarlo en /retomar para corregir observaciones de SENA (válido 30 días)._',
        { parse_mode: 'Markdown' }
      );
    } catch (e) { console.error('redis save plan pin:', e.message); }
  }

  await ctx.reply('✨ Listo. El Excel oficial SENA ya tiene los datos listos para SharePoint.\nUsa /nuevo para el siguiente beneficiario.');
}

// ══════════════════════════════════════════════════════════
// ── PROCESAR RESPUESTA ────────────────────────────────────
// ══════════════════════════════════════════════════════════
// PUNTO 4: error de extraerNumero → pregunta al usuario, no guarda 0
// FIX 2: saltar preguntas opcionales + progreso con tiempo estimado
async function procesarRespuesta(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);
  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) return;

  const { p, i } = actual;

  // ── Manejo de "saltar" en preguntas opcionales ──
  const esSkip = /^(saltar|s|no\s*s[eé]|no\s*aplica|omitir|pasar)$/i.test(texto.trim());
  if (esSkip) {
    if (!CLAVES_OPCIONALES.has(p.key)) {
      await ctx.reply(
        '⚠️ Esta pregunta es *obligatoria* para el formato SENA. Por favor respóndela.',
        { parse_mode: 'Markdown' }
      );
      return;
    }
    sesion.datos[p.key] = DEFAULTS_OPCIONALES[p.key] || '';
    sesion.paso = i + 1;
    await setSesion(chatId, sesion);
    await ctx.reply('⏭️ _Pregunta omitida._', { parse_mode: 'Markdown' });
    const sig = getPregunta(sesion.paso, sesion.datos);
    if (sig) {
      const restantes = PREGUNTAS.length - sesion.paso;
      const minutos   = Math.ceil(restantes * 0.5);
      const progreso  = `_(${sesion.paso}/${PREGUNTAS.length} · ~${minutos} min)_`;
      await ctx.reply(`${sig.p.msg}\n\n${progreso}`, { parse_mode: 'Markdown' });
    } else {
      sesion.modo = 'revisando';
      await setSesion(chatId, sesion);
      await mostrarResumen(ctx, sesion);
    }
    return;
  }

  let valor = p.key === 'nombre' ? limpiarNombre(texto) : texto;

  if (p.numerico) {
    const n = await extraerNumero(texto);
    if (n === null) {
      await ctx.reply(
        '⚠️ No entendí ese número. Escríbalo en dígitos _(ej: 150000)_ o repítalo más claro.',
        { parse_mode: 'Markdown' }
      );
      return; // no avanzar — el usuario debe repetir
    }
    valor = String(n);
  }

  sesion.datos[p.key] = valor;
  sesion.paso = i + 1;
  await setSesion(chatId, sesion); // persistir

  const preview = valor.length > 120 ? valor.slice(0, 120) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });

  const sig = getPregunta(sesion.paso, sesion.datos);
  if (sig) {
    const restantes = PREGUNTAS.length - sesion.paso;
    const minutos   = Math.ceil(restantes * 0.5);
    const progreso  = `_(${sesion.paso}/${PREGUNTAS.length} · ~${minutos} min)_`;
    await ctx.reply(`${sig.p.msg}\n\n${progreso}`, { parse_mode: 'Markdown' });
  } else {
    sesion.modo = 'revisando';
    await setSesion(chatId, sesion);
    await mostrarResumen(ctx, sesion);
  }
}

// ── Corrección desde resumen ──────────────────────────────
async function aplicarCorreccion(ctx, sesion, chatId, texto) {
  const key   = sesion.corrigiendoKey;
  const p     = PREGUNTAS.find(q => q.key === key);
  let valor   = key === 'nombre' ? limpiarNombre(texto) : texto;

  if (p && p.numerico) {
    const n = await extraerNumero(texto);
    if (n === null) {
      await ctx.reply(
        '⚠️ No entendí ese número. Escríbalo en dígitos _(ej: 150000)_ o repítalo más claro.',
        { parse_mode: 'Markdown' }
      );
      return;
    }
    valor = String(n);
  }

  sesion.datos[key] = valor;
  sesion.modo = 'revisando';
  await setSesion(chatId, sesion); // persistir

  const preview = valor.length > 80 ? valor.slice(0, 80) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });
  await mostrarResumen(ctx, sesion);
}

// ══════════════════════════════════════════════════════════
// ── HANDLERS COMPARTIDOS TEXTO + VOZ ─────────────────────
// ══════════════════════════════════════════════════════════

async function manejarConfirmacionViabilidad(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);
  sesion.modo  = 'revisando';
  await setSesion(chatId, sesion);
  if (/^(continuar|si|sí|generar|ok|dale)$/i.test(texto.trim())) {
    await ctx.reply('Generando el plan con los datos actuales…');
    await finalizar(ctx, sesion);
    await delSesion(chatId);
  } else {
    await ctx.reply('Volviste al resumen. Corrige los datos y vuelve a escribir /listo.');
    await mostrarResumen(ctx, sesion);
  }
}

async function manejarPinRetomar(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);
  const pin    = texto.trim().replace(/\D/g, '');

  if (!pin || pin.length !== 4) {
    await ctx.reply('Escribe el *PIN de 4 dígitos* que recibiste al generar el plan.', { parse_mode: 'Markdown' });
    return;
  }

  let datosGuardados = null;
  if (_redis) {
    try {
      const raw = await _redis.get(`plan:pin:${pin}`);
      if (raw) datosGuardados = typeof raw === 'string' ? JSON.parse(raw) : raw;
    } catch (e) { console.error('redis get pin:', e.message); }
  }

  if (!datosGuardados) {
    await ctx.reply(
      `PIN *${pin}* no encontrado o expirado.\n\nLos planes se guardan 30 días. Si pasó más tiempo, usa /nuevo para empezar de cero.`,
      { parse_mode: 'Markdown' }
    );
    sesion.modo = undefined;
    await setSesion(chatId, sesion);
    return;
  }

  sesion.datos = datosGuardados;
  sesion.paso  = PREGUNTAS.length;
  sesion.modo  = 'revisando';
  await setSesion(chatId, sesion);
  await ctx.reply(`✅ Plan de *${datosGuardados.nombre || 'beneficiario'}* cargado.`, { parse_mode: 'Markdown' });
  await mostrarResumen(ctx, sesion);
}

// ══════════════════════════════════════════════════════════
// ── COMANDOS ──────────────────────────────────────────────
// ══════════════════════════════════════════════════════════
bot.command('start', async ctx => {
  await delSesion(ctx.chat.id);
  await ctx.reply(
    '👋 *Bienvenido al Bot SENA · Línea CREAR Especial*\n\n' +
    '*Plan de Negocio:*\n' +
    '📝 /nuevo — Crear plan de negocio\n' +
    '🔁 /retomar — Recargar plan anterior con PIN\n' +
    '↩️ /atras — Corregir respuesta anterior\n' +
    '📋 /resumen — Ver y corregir todas las respuestas\n' +
    '✅ /listo — Generar el Excel oficial SENA\n' +
    '🔄 /reiniciar — Cancelar y empezar de cero\n\n' +
    '*Diagnóstico Empresarial:*\n' +
    '🏥 /dx — Iniciar diagnóstico (40 preguntas 0/1/2)\n' +
    '📊 /dx_resumen — Ver puntajes actuales\n' +
    '✅ /dx_listo — Generar Excel del diagnóstico\n\n' +
    '─────────────────────',
    { parse_mode: 'Markdown' }
  );
  const primera = getPregunta(0, {});
  if (primera) {
    await ctx.reply(
      primera.p.msg + `\n\n_(Pregunta 1 de ${PREGUNTAS.length})_`,
      { parse_mode: 'Markdown' }
    );
  }
});

bot.command('nuevo', async ctx => {
  await delSesion(ctx.chat.id);
  await ctx.reply('🆕 Iniciando nuevo plan de negocio...');
  const primera = getPregunta(0, {});
  if (primera) {
    await ctx.reply(
      primera.p.msg + `\n\n_(Pregunta 1 de ${PREGUNTAS.length})_`,
      { parse_mode: 'Markdown' }
    );
  }
});

bot.command('reiniciar', async ctx => {
  await delSesion(ctx.chat.id);
  await ctx.reply('🔄 Plan cancelado. Usa /nuevo para empezar de cero.');
});

// FIX 4: /retomar — carga el último plan generado de un beneficiario por cédula
bot.command('retomar', async ctx => {
  const chatId = ctx.chat.id;
  if (!_redis) {
    await ctx.reply('⚠️ /retomar requiere Redis. Configura UPSTASH_REDIS_REST_URL en Render.');
    return;
  }
  const sesion = await getSesion(chatId);
  sesion.modo = 'esperando_pin_retomar';
  await setSesion(chatId, sesion);
  await ctx.reply(
    '🔁 *Retomar plan existente*\n\n¿Cuál es el *PIN de 4 dígitos* del plan?\n_(Lo recibiste al generar el plan — válido 30 días)_',
    { parse_mode: 'Markdown' }
  );
});

bot.command('atras', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);
  if (sesion.modo === 'revisando') {
    sesion.modo = undefined;
    await setSesion(chatId, sesion);
    await ctx.reply('↩️ Volviste al flujo de preguntas. Escribe /resumen cuando termines.');
    return;
  }
  const lastIdx = getLastAnsweredIndex(sesion);
  if (lastIdx < 0) {
    await ctx.reply('Ya estás en la primera pregunta.');
    return;
  }
  delete sesion.datos[PREGUNTAS[lastIdx].key];
  sesion.paso = lastIdx;
  await setSesion(chatId, sesion);
  const p = getPregunta(sesion.paso, sesion.datos);
  if (p) {
    await ctx.reply(
      `↩️ *Volvamos atrás:*\n\n${p.p.msg}\n\n_(${sesion.paso + 1}/${PREGUNTAS.length})_`,
      { parse_mode: 'Markdown' }
    );
  }
});

bot.command('resumen', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);
  if (!sesion || Object.keys(sesion.datos).length === 0) {
    await ctx.reply('No hay ningún plan en curso. Usa /nuevo para empezar.');
    return;
  }
  sesion.modo = 'revisando';
  await setSesion(chatId, sesion);
  await mostrarResumen(ctx, sesion);
});

// PUNTO 3: /listo protegido con validación de campos obligatorios
bot.command('listo', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = await getSesion(chatId);

  if (!sesion || sesion.modo !== 'revisando') {
    await ctx.reply('Primero completa el cuestionario. Usa /nuevo para empezar.');
    return;
  }

  const faltantes = validarCamposObligatorios(sesion.datos);
  if (faltantes.length > 0) {
    await ctx.reply(
      '⚠️ *Faltan estos datos obligatorios:*\n\n' +
      faltantes.map(f => `• ${f}`).join('\n') +
      '\n\nEscribe el *número* de la respuesta a corregir.',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  // Verificar aporte mínimo del 10%
  const inv = Number(sesion.datos.inversion) || 0;
  const ap  = Number(sesion.datos.aportePropio) || 0;
  if (inv > 0 && ap / inv < 0.10) {
    await ctx.reply(
      `⚠️ El aporte propio (*${fmt(ap)}*) es menor al *10%* de la inversión total (*${fmt(inv)}*).\n\n` +
      `Mínimo requerido: *${fmt(Math.ceil(inv * 0.10))}*\n\n` +
      'Corrige el aporte propio (escribe el número correspondiente en el resumen) o la inversión total.',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  // ── FIX 3: Alerta de viabilidad financiera ──────────────
  const p1  = Number(sesion.datos.prod1Precio)       || 0;
  const u1  = Number(sesion.datos.prod1UnidadesMes)  || 0;
  const cv1 = Number(sesion.datos.prod1Costo)        || 0;
  const cf1 = Number(sesion.datos.costosFijosTotal)  || 0;
  const mo1 = Number(sesion.datos.manoObraTotal)     || 0;
  const ga1 = Number(sesion.datos.gastosAdmin)       || 0;

  const ingresoMes   = p1 * u1;
  const cvMes        = cv1 * u1;
  const margenMes    = ingresoMes - cvMes;
  const costosFijMes = cf1 + mo1 + ga1;
  const ebitdaMes    = margenMes - costosFijMes;

  if (p1 > 0 && u1 > 0 && ebitdaMes < 0) {
    sesion.modo = 'confirmando_viabilidad';
    await setSesion(chatId, sesion);
    await ctx.reply(
      '⚠️ *Alerta financiera antes de generar el plan:*\n\n' +
      `Ingresos al mes:       *${fmt(ingresoMes)}*\n` +
      `Costo variable/mes:    *${fmt(cvMes)}*\n` +
      `Costos fijos/mes:      *${fmt(costosFijMes)}*\n` +
      `─────────────────────\n` +
      `EBITDA mensual:        *${fmt(ebitdaMes)}* ❌\n\n` +
      'Con estos números el negocio *pierde dinero cada mes*. ' +
      'SENA puede rechazar el plan en la evaluación financiera.\n\n' +
      '¿Quiere *corregir los datos* o *continuar de todas formas*?\n' +
      '_Responda: *corregir* o *continuar*_',
      { parse_mode: 'Markdown' }
    );
    return;
  }

  await finalizar(ctx, sesion);
  await delSesion(chatId);
});

bot.command('estado', async ctx => {
  const sesion = await getSesion(ctx.chat.id);
  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) {
    await ctx.reply('No hay ningún plan en curso. Usa /nuevo para empezar.');
    return;
  }
  const pct = Math.round(sesion.paso / PREGUNTAS.length * 100);
  await ctx.reply(
    `📋 *Progreso:* ${sesion.paso}/${PREGUNTAS.length} preguntas (${pct}%)\n\n` +
    `*Siguiente:* ${actual.p.msg.replace(/\*|_/g, '').slice(0, 80)}…`,
    { parse_mode: 'Markdown' }
  );
});

// ── Mensajes de voz / audio ───────────────────────────────
bot.on(['message:voice', 'message:audio'], async ctx => {
  const chatId = ctx.chat.id;

  // ── Flujo diagnóstico ──
  if (await hasDxSesion(chatId)) {
    const dxSesion = await getDxSesion(chatId);
    if (dxSesion.paso >= DX_TOTAL) {
      await ctx.reply('El diagnóstico está completo. Usa /dx_listo para generar el Excel.');
      return;
    }
    const msg = await ctx.reply('🎤 Transcribiendo…');
    try {
      const fileId = ctx.message.voice?.file_id || ctx.message.audio?.file_id;
      const texto  = await transcribir(fileId);
      await ctx.api.deleteMessage(chatId, msg.message_id).catch(() => {});
      await dxProcesarRespuesta(ctx, texto);
    } catch (e) {
      console.error('Transcripción dx fallida:', e.message);
      await ctx.reply('❌ No pude escuchar. Intenta de nuevo o escribe 0, 1 o 2.');
    }
    return;
  }

  // ── Flujo plan de negocio ──
  const sesion       = await getSesion(chatId);
  const enRevision   = sesion.modo === 'revisando';
  const enCorreccion = sesion.modo === 'corrigiendo';
  const actual       = getPregunta(sesion.paso, sesion.datos);

  // Modos que esperan texto libre — transcribir voz y despachar a la función correcta
  if (sesion.modo === 'confirmando_viabilidad' || sesion.modo === 'esperando_pin_retomar') {
    const msg = await ctx.reply('🎤 Transcribiendo…');
    try {
      const fileId = ctx.message.voice?.file_id || ctx.message.audio?.file_id;
      const texto  = await transcribir(fileId);
      await ctx.api.deleteMessage(chatId, msg.message_id).catch(() => {});
      if (sesion.modo === 'confirmando_viabilidad') {
        await manejarConfirmacionViabilidad(ctx, texto);
      } else {
        await manejarPinRetomar(ctx, texto);
      }
    } catch (e) {
      console.error('Transcripción fallida:', e.message);
      await ctx.reply('❌ No pude escuchar. Por favor escribe la respuesta.');
    }
    return;
  }

  if (!actual && !enCorreccion) {
    await ctx.reply(enRevision
      ? '📋 Estás revisando el plan. Escribe el número a corregir o */listo*.'
      : 'Usa /nuevo para el plan de negocio o /dx para el diagnóstico.',
      { parse_mode: 'Markdown' });
    return;
  }

  const msg = await ctx.reply('🎤 Transcribiendo…');
  try {
    const fileId = ctx.message.voice?.file_id || ctx.message.audio?.file_id;
    const texto  = await transcribir(fileId);
    await ctx.api.deleteMessage(chatId, msg.message_id).catch(() => {});
    if (enCorreccion) {
      await aplicarCorreccion(ctx, sesion, chatId, texto);
    } else {
      await procesarRespuesta(ctx, texto);
    }
  } catch (e) {
    console.error('Transcripción fallida:', e.message);
    await ctx.reply('❌ No pude escuchar bien. Intenta de nuevo o escribe la respuesta.');
  }
});

// ── Mensajes de texto ─────────────────────────────────────
bot.on('message:text', async ctx => {
  if (ctx.message.text.startsWith('/')) return;
  const chatId = ctx.chat.id;

  // ── Flujo diagnóstico ──
  if (await hasDxSesion(chatId)) {
    const dxSesion = await getDxSesion(chatId);
    if (dxSesion.paso >= DX_TOTAL) {
      await ctx.reply('El diagnóstico está completo. Usa /dx_listo para generar el Excel.');
      return;
    }
    await dxProcesarRespuesta(ctx, ctx.message.text);
    return;
  }

  // ── Flujo plan de negocio ──
  const sesion = await getSesion(chatId);

  if (sesion.modo === 'confirmando_viabilidad') {
    await manejarConfirmacionViabilidad(ctx, ctx.message.text);
    return;
  }

  if (sesion.modo === 'esperando_pin_retomar') {
    await manejarPinRetomar(ctx, ctx.message.text);
    return;
  }

  if (sesion.modo === 'revisando') {
    const n = parseInt(ctx.message.text.trim(), 10);
    if (!isNaN(n) && n >= 1) {
      const items = getResumenItems(sesion);
      const item  = items.find(it => it.num === n);
      if (item) {
        sesion.modo = 'corrigiendo';
        sesion.corrigiendoKey = item.key;
        await setSesion(chatId, sesion);
        const pregunta = PREGUNTAS.find(p => p.key === item.key);
        await ctx.reply(
          `✏️ *Corrigiendo #${n} — ${item.label}*\n\n${pregunta.msg}\n\n_Responde con voz 🎤 o escribe el nuevo valor:_`,
          { parse_mode: 'Markdown' }
        );
      } else {
        await ctx.reply(`No encontré la pregunta ${n}. Escribe un número del 1 al ${items.length} o */listo*.`, { parse_mode: 'Markdown' });
      }
    } else {
      await ctx.reply('Escribe el *número* de la pregunta a corregir, o */listo* para generar el plan.', { parse_mode: 'Markdown' });
    }
    return;
  }

  if (sesion.modo === 'corrigiendo') {
    await aplicarCorreccion(ctx, sesion, chatId, ctx.message.text);
    return;
  }

  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) {
    await ctx.reply('Usa /nuevo para el plan de negocio o /dx para el diagnóstico.');
    return;
  }
  await procesarRespuesta(ctx, ctx.message.text);
});

// ══════════════════════════════════════════════════════════
// ── FLUJO DE DIAGNÓSTICO EMPRESARIAL (40 preguntas 0/1/2) ─
// ══════════════════════════════════════════════════════════

const DX_BASICOS = [
  { key: 'dxNombre',      msg: '¿Nombre completo del emprendedor?' },
  { key: 'dxCedula',      msg: '¿Número de cédula?' },
  { key: 'dxNegocio',     msg: '¿Nombre de la Unidad Productiva o negocio?' },
  { key: 'dxCiiu',        msg: '¿Sector económico (CIIU)?\n_(Ej: 1411 Confección · 0111 Cultivos · 4711 Tienda)_' },
  { key: 'dxTipo',        msg: '¿Tipo de iniciativa?\n_(Individual / Familiar / Asociativa)_' },
  { key: 'dxTiempo',      msg: '¿Cuánto tiempo lleva funcionando el negocio?\n_(Ej: 2 años, 8 meses)_' },
  { key: 'dxEmpleados',   msg: '¿Cuántos empleados tiene actualmente?' },
  { key: 'dxDireccion',   msg: '¿Dirección de la Unidad Productiva?' },
  { key: 'dxTelefono',    msg: '¿Teléfono o celular de contacto?' },
  { key: 'dxDinamizador', msg: '¿Nombre del dinamizador que aplica el diagnóstico?' },
];

const DX_PREGUNTAS = [
  // Área 1 — Legal y Administrativa
  { area: 1, num: 1,  msg: '¿Para la planeación estratégica y fijación de objetivos tiene en cuenta al *cliente interno* y las tendencias del mercado?' },
  { area: 1, num: 2,  msg: '¿Es suficiente el número de trabajadores para el cumplimiento de las actividades diarias?' },
  { area: 1, num: 3,  msg: '¿Cuenta con *reemplazos* en los momentos de ausencia prolongada de algún trabajador?' },
  { area: 1, num: 4,  msg: '¿La Unidad Productiva está *bancarizada* para la gestión de ventas?' },
  { area: 1, num: 5,  msg: '¿El emprendedor cuenta con *asignación de salario*?' },
  { area: 1, num: 6,  msg: '¿El emprendedor y sus empleados están afiliados a *seguridad social*?' },
  { area: 1, num: 7,  msg: '¿Tiene *estructura organizacional* definida? (organigrama, manual de funciones, perfiles)' },
  { area: 1, num: 8,  msg: '¿Identifica y usa *elementos de protección personal*? ¿La UP está señalizada?' },
  { area: 1, num: 9,  msg: '¿Ha implementado *herramientas de innovación* para mejorar competitivamente?' },
  { area: 1, num: 10, msg: '¿Conoce los *permisos y licencias* que requiere? ¿Está comprometida con el medio ambiente?' },
  // Área 2 — Comercial
  { area: 2, num: 1,  msg: '¿Tiene claro qué *productos o servicios* ofrece al mercado? (todos ellos)' },
  { area: 2, num: 2,  msg: '¿Conoce la *participación de cada producto* en el total de ventas?' },
  { area: 2, num: 3,  msg: '¿Se fundamenta en datos para hacer *cambios en los productos o servicios*?' },
  { area: 2, num: 4,  msg: '¿Tiene definido el *diferenciador* de su producto frente a la competencia?' },
  { area: 2, num: 5,  msg: '¿Tiene identificado y caracterizado a su *cliente objetivo y potencial*?' },
  { area: 2, num: 6,  msg: '¿Investiga *tendencias del sector* a nivel nacional e internacional?' },
  { area: 2, num: 7,  msg: '¿Tiene en cuenta el *análisis de la competencia* en sus procesos? (precio, calidad, fortalezas)' },
  { area: 2, num: 8,  msg: '¿Tiene *meta en ventas mensuales* y hace seguimiento a su cumplimiento?' },
  { area: 2, num: 9,  msg: '¿Cuenta con *catálogo o portafolio* de productos/servicios?' },
  { area: 2, num: 10, msg: '¿Cuenta con un *plan de fidelización* para los clientes?' },
  // Área 3 — Técnica-Operativa
  { area: 3, num: 1,  msg: '¿Identifica y cuantifica los *riesgos de producción* en sus procesos?' },
  { area: 3, num: 2,  msg: '¿Existen *fichas técnicas* de máquinas, equipos y productos?' },
  { area: 3, num: 3,  msg: '¿La *infraestructura física* es adecuada para el desarrollo de la actividad?' },
  { area: 3, num: 4,  msg: '¿Los puestos de trabajo y maquinaria están distribuidos para *optimizar tiempos*?' },
  { area: 3, num: 5,  msg: '¿Ha replicado algún *avance de la competencia* en su propia empresa?' },
  { area: 3, num: 6,  msg: '¿Tiene un manejo adecuado de los *residuos* que genera la empresa?' },
  { area: 3, num: 7,  msg: '¿El proceso de producción/servicio está *definido*? (etapas, tiempos, controles)' },
  { area: 3, num: 8,  msg: '¿El producto puede *adaptarse* rápidamente a las tendencias del mercado?' },
  { area: 3, num: 9,  msg: '¿Realiza *control de calidad* a su producto o servicio?' },
  { area: 3, num: 10, msg: '¿Planea las compras de materia prima y tiene *control de inventarios*?' },
  // Área 4 — Financiera y Contable
  { area: 4, num: 1,  msg: '¿Separa los *gastos personales* de los gastos del negocio?' },
  { area: 4, num: 2,  msg: '¿Tiene proceso definido para *fijar el precio* de sus productos/servicios?' },
  { area: 4, num: 3,  msg: '¿Conoce su *margen de rentabilidad*?' },
  { area: 4, num: 4,  msg: '¿Identifica si la empresa está *generando utilidad*?' },
  { area: 4, num: 5,  msg: '¿Tiene identificadas las *fluctuaciones de demanda* por temporadas o variaciones del mercado?' },
  { area: 4, num: 6,  msg: '¿Lleva *registro de ingresos*? (ventas, cuentas por cobrar, capital de trabajo)' },
  { area: 4, num: 7,  msg: '¿Lleva *registros contables*? (compras, gastos, cuentas por pagar, proveedores)' },
  { area: 4, num: 8,  msg: '¿Sabe cuánto le cuesta producir su bien o servicio? (*costos variables*)' },
  { area: 4, num: 9,  msg: '¿*Reinvierte ganancias* para el mejoramiento de la Unidad Productiva?' },
  { area: 4, num: 10, msg: '¿Identifica y cumple las *normas tributarias y contables*?' },
];

const DX_ESCALA = '\n\n🔢 *0* = No/Nunca · *1* = A veces/En proceso · *2* = Sí/Siempre';
const DX_TOTAL  = DX_BASICOS.length + DX_PREGUNTAS.length; // 50

// PUNTO 4: clasificar012 devuelve null en fallo (no 0 silencioso)
async function clasificar012(texto) {
  const t = texto.trim();
  if (t === '0') return 0;
  if (t === '1') return 1;
  if (t === '2') return 2;
  try {
    const msg = await anthropic.messages.create({
      model:      'claude-haiku-4-5-20251001',
      max_tokens: 5,
      messages:   [{ role: 'user', content: `Clasifica en 0=No/Nunca, 1=A veces/En proceso, 2=Sí/Siempre. Solo el número.\nRespuesta: "${texto}"` }],
    });
    const n = parseInt(msg.content[0].text.trim(), 10);
    if (isNaN(n)) return null; // no pudo clasificar → caller pide repetir
    return Math.min(2, Math.max(0, n));
  } catch {
    return null; // API caída → null para que caller informe al usuario
  }
}

async function generarExcelDx(datos) {
  const wb    = XLSX.readFile(path.join(__dirname, 'dx_template.xlsx'), { cellFormula: true, cellStyles: true });
  const wsDx  = wb.Sheets['Diagnóstico'];
  const wsDup = wb.Sheets['Datos de la Unidad Productiva'];

  function setD(ws, addr, value) {
    if (value === null || value === undefined || value === '') return;
    const t = typeof value === 'number' ? 'n' : 's';
    ws[addr] = { t, v: value, w: String(value) };
  }

  const fecha = new Date().toLocaleDateString('es-CO');

  setD(wsDx, 'B4', datos.dxNegocio    || '');
  setD(wsDx, 'B5', datos.dxNombre     || '');
  setD(wsDx, 'B7', datos.dxDinamizador || '');
  setD(wsDx, 'B8', fecha);

  const scores = datos.scores || Array(40).fill(0);
  for (let i = 0; i < 10; i++) setD(wsDx, `B${11 + i}`, scores[i] ?? 0);
  for (let i = 0; i < 10; i++) setD(wsDx, `B${23 + i}`, scores[10 + i] ?? 0);
  for (let i = 0; i < 10; i++) setD(wsDx, `B${35 + i}`, scores[20 + i] ?? 0);
  for (let i = 0; i < 10; i++) setD(wsDx, `B${47 + i}`, scores[30 + i] ?? 0);

  setD(wsDup, 'B7',  datos.dxNombre      || '');
  setD(wsDup, 'B8',  datos.dxCedula      || '');
  setD(wsDup, 'B10', datos.dxNegocio     || '');
  setD(wsDup, 'B11', datos.dxCiiu        || '');
  setD(wsDup, 'B13', datos.dxTiempo      || '');
  setD(wsDup, 'B14', datos.dxEmpleados   || '');
  setD(wsDup, 'B15', datos.dxDireccion   || '');
  setD(wsDup, 'B16', datos.dxTelefono    || '');
  setD(wsDup, 'B23', datos.dxDinamizador || '');

  return wb;
}

function dxPreguntaMsg(paso) {
  if (paso < DX_BASICOS.length) {
    const p = DX_BASICOS[paso];
    return `${p.msg}\n\n_(Paso ${paso + 1} de ${DX_TOTAL})_`;
  }
  const qi = paso - DX_BASICOS.length;
  const p  = DX_PREGUNTAS[qi];
  const cambioArea = qi === 0 || DX_PREGUNTAS[qi - 1].area !== p.area;
  const cabecera   = cambioArea ? `\n*📋 ÁREA ${p.area}*\n\n` : '';
  return `${cabecera}*${p.area}.${p.num}* ${p.msg}${DX_ESCALA}\n\n_(Pregunta ${paso + 1} de ${DX_TOTAL})_`;
}

function dxResumenMsg(datos) {
  const scores = datos.scores || [];
  const s1 = scores.slice(0, 10).reduce((a, b) => a + (b || 0), 0);
  const s2 = scores.slice(10, 20).reduce((a, b) => a + (b || 0), 0);
  const s3 = scores.slice(20, 30).reduce((a, b) => a + (b || 0), 0);
  const s4 = scores.slice(30, 40).reduce((a, b) => a + (b || 0), 0);
  const total = s1 + s2 + s3 + s4;
  const pct   = (total / 80 * 100).toFixed(0);
  const nivel = total <= 20 ? 'Inicio' : total <= 40 ? 'Básico' : total <= 60 ? 'Intermedio' : 'Avanzado';

  return [
    '📊 *Resumen del Diagnóstico*\n',
    `👤 ${datos.dxNombre || '—'} · ${datos.dxNegocio || '—'}`,
    '',
    `Área 1 Legal/Adm:    *${s1}/20*`,
    `Área 2 Comercial:    *${s2}/20*`,
    `Área 3 Técn-Oper:    *${s3}/20*`,
    `Área 4 Financiera:   *${s4}/20*`,
    `─────────────────`,
    `Total: *${total}/80* (${pct}%) → _${nivel}_`,
    '',
    '✅ Escribe */dx_listo* para generar el Excel oficial',
  ].join('\n');
}

// PUNTO 4: clasificar012 null → pide repetir, no guarda 0 silencioso
async function dxProcesarRespuesta(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = await getDxSesion(chatId);
  const paso   = sesion.paso;

  if (paso >= DX_TOTAL) {
    await ctx.reply('El diagnóstico ya está completo. Usa /dx_listo para generar el Excel.');
    return;
  }

  if (paso < DX_BASICOS.length) {
    sesion.datos[DX_BASICOS[paso].key] = texto;
    sesion.paso++;
    await setDxSesion(chatId, sesion);
  } else {
    const qi    = paso - DX_BASICOS.length;
    const score = await clasificar012(texto);
    if (score === null) {
      await ctx.reply(
        '❌ No entendí esa respuesta. Por favor responda *0*, *1* o *2*:\n\n' +
        '0 = No / Nunca\n1 = A veces / En proceso\n2 = Sí / Siempre',
        { parse_mode: 'Markdown' }
      );
      return; // no avanzar
    }
    sesion.datos.scores[qi] = score;
    sesion.paso++;
    await setDxSesion(chatId, sesion);
    await ctx.reply(`✅ *${score}*`, { parse_mode: 'Markdown' });
  }

  if (sesion.paso >= DX_TOTAL) {
    await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
  } else {
    await ctx.reply(dxPreguntaMsg(sesion.paso), { parse_mode: 'Markdown' });
  }
}

// ── Comandos diagnóstico ──────────────────────────────────
bot.command('dx', async ctx => {
  const chatId = ctx.chat.id;
  await delDxSesion(chatId);
  await ctx.reply(
    '🏥 *Diagnóstico Empresarial SENA*\n\n' +
    'Vamos a evaluar 4 áreas de tu negocio con 40 preguntas.\n' +
    'Puedes responder con voz 🎤 o escribiendo.\n\n' +
    '*Escala:*\n' +
    '*0* = No / Nunca / No implementado\n' +
    '*1* = A veces / En proceso / Parcial\n' +
    '*2* = Sí / Siempre / Completamente\n\n' +
    '─────────────────────\n' +
    'Primero necesito algunos datos básicos:',
    { parse_mode: 'Markdown' }
  );
  // Inicializar sesión en Redis
  const sesion = await getDxSesion(chatId);
  await ctx.reply(dxPreguntaMsg(0), { parse_mode: 'Markdown' });
});

bot.command('dx_listo', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = await getDxSesion(chatId);
  if (!sesion || sesion.paso < DX_BASICOS.length) {
    await ctx.reply('Primero completa el diagnóstico con /dx');
    return;
  }

  await ctx.reply('📊 Generando el Excel oficial del diagnóstico… ⏳');

  let wb;
  try {
    wb = await generarExcelDx(sesion.datos);
  } catch (e) {
    console.error('Error dx excel:', e.message);
    await ctx.reply('❌ No pude generar el Excel. Revisa el archivo dx_template.xlsx en el servidor.');
    return;
  }

  const nombre = (sesion.datos.dxNombre || 'beneficiario').replace(/\s+/g, '_');
  const fecha  = new Date().toISOString().slice(0, 10);
  const fn     = `Diagnostico_${nombre}_${fecha}.xlsx`;
  const tmp    = `/tmp/${fn}`;
  XLSX.writeFile(wb, tmp);

  await ctx.replyWithDocument(new InputFile(tmp, fn), {
    caption: [
      '📋 *Diagnóstico Empresarial SENA*',
      `👤 ${sesion.datos.dxNombre || ''}`,
      `🏭 ${sesion.datos.dxNegocio || ''}`,
      `📅 ${fecha}`,
    ].join('\n'),
    parse_mode: 'Markdown',
  });
  try { fs.unlinkSync(tmp); } catch {}

  await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
  await delDxSesion(chatId);
  await ctx.reply('✨ Listo. Usa /dx para un nuevo diagnóstico o /nuevo para el plan de negocio.');
});

bot.command('dx_resumen', async ctx => {
  const sesion = await getDxSesion(ctx.chat.id);
  if (!sesion || sesion.paso === 0) {
    await ctx.reply('No hay diagnóstico en curso. Usa /dx para empezar.');
    return;
  }
  await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
});

// ── Health check ──────────────────────────────────────────
const PORT = process.env.PORT || 3000;
http.createServer((_, res) => res.end('OK')).listen(PORT);

console.log('🤖 Bot SENA CREAR v4 (Redis + validación + costos completos) iniciando…');
bot.start();
