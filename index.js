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

// ── Estado en memoria  chatId → { paso, datos, modo } ────
const sesiones = new Map();

// ── Preguntas (~38, cubre formato oficial SENA) ───────────
const PREGUNTAS = [
  // === Datos del beneficiario ===
  {
    key: 'nombre',
    msg: '¡Hola! 👋 Vamos a crear el *Plan de Negocio SENA Línea CREAR*.\nPuedes responder con voz 🎤 o escribiendo.\n\n¿Cuál es el *nombre completo* del beneficiario?',
  },
  { key: 'tipoDoc',       msg: '¿*Tipo de documento*?\n_(CC Cédula · CE Cédula Extranjería · PA Pasaporte · TI Tarjeta Identidad)_' },
  { key: 'numDoc',        msg: '¿*Número de documento*?' },
  { key: 'genero',        msg: '¿*Género*?\n_(Masculino / Femenino / Otro)_' },
  { key: 'departamento',  msg: '¿En qué *departamento* está el proyecto?' },
  { key: 'municipio',     msg: '¿En qué *municipio*?' },
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
  { key: 'asociativo',   msg: '¿El proyecto es *asociativo* (varias personas juntas)?\n_Responda: SI o NO_' },
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
  { key: 'prod1Nombre', msg: '¿Cuál es el *nombre del producto o servicio principal*?' },
  { key: 'prod1Desc',   msg: '¿Cómo lo *describiría*? ¿Qué es exactamente?' },
  { key: 'prod1Unidad', msg: '¿Cuál es la *unidad de medida*?\n_(Ej: Kilogramo, Litro, Unidad, Hora, Docena, Porción...)_' },
  { key: 'prod1Precio',      msg: '¿A qué *precio* lo vende?\n_(Número en pesos, ej: 15000 o "quince mil pesos")_', numerico: true },
  { key: 'prod1UnidadesMes', msg: '¿Cuántas *unidades vende al mes* aproximadamente?\n_(Número)_', numerico: true },
  { key: 'prod1Costo',       msg: '¿Cuánto le *cuesta producir* cada unidad?\n_(Materias primas e insumos — número en pesos)_', numerico: true },

  // === Sección 10 — Costos fijos ===
  {
    key: 'costosFijosDesc',
    msg: '¿Cuáles son sus *costos fijos mensuales*?\nCuénteme los gastos que paga todos los meses aunque no venda nada: arriendo, servicios, internet, transporte...',
  },
  { key: 'costosFijosTotal', msg: '¿Cuánto suman esos costos fijos *en total al mes*?\n_(Número en pesos)_', numerico: true },

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

// ── Helpers básicos ───────────────────────────────────────
function limpiarNombre(texto) {
  return texto
    .replace(/^(mi nombre (completo )?es|me llamo|yo me llamo|soy|yo soy|me dicen)\s+/i, '')
    .trim();
}

function getSesion(chatId) {
  if (!sesiones.has(chatId)) sesiones.set(chatId, { paso: 0, datos: {} });
  return sesiones.get(chatId);
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

// ── extraerNumero: verbal → número via Claude ─────────────
// Resuelve "un millón quinientos mil" → 1500000
async function extraerNumero(texto) {
  const clean = String(texto).replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
  if (/^\d+(\.\d+)?$/.test(clean)) return parseFloat(clean) || 0;
  try {
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 15,
      messages: [{ role: 'user', content: `Número del texto (solo dígitos enteros, sin puntos ni comas): "${texto}"` }],
    });
    return Number(msg.content[0].text.replace(/[^\d]/g, '')) || 0;
  } catch { return 0; }
}

// ── parsearCompetidor: texto libre → campos del Excel ─────
async function parsearCompetidor(texto) {
  if (!texto || /^no\s*hay/i.test(texto.trim())) return null;
  try {
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 200,
      messages: [{
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

// ── parsearCostosFijos: texto → lista {desc, valorMensual} ─
async function parsearCostosFijos(texto, total) {
  if (!texto) return [{ descripcion: 'Costos fijos', valorMensual: total }];
  try {
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages: [{
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

// ── parsearInversion: texto → lista {desc, cantidad, valor} ─
async function parsearInversion(texto, total) {
  if (!texto) return [{ descripcion: 'Inversión', cantidad: 1, valor: total }];
  try {
    const msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 300,
      messages: [{
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

// ── Generar Plan Narrativo con Claude ─────────────────────
async function generarPlanCompleto(datos) {
  const d = datos;
  const p   = Number(d.prod1Precio)     || 0;
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

// ── Llenar Excel oficial SENA ─────────────────────────────
async function generarExcelOficial(datos, costosFijosItems, inversionItems, comps) {
  const templatePath = path.join(__dirname, 'template.xlsx');
  const wb = XLSX.readFile(templatePath, { cellFormula: true, cellStyles: true });
  const ws = wb.Sheets['PROYECTO'];

  // Escribe en la celda top-left de cada merge (preserva merges y estilos)
  function set(addr, value) {
    if (value === null || value === undefined || value === '') return;
    const t = typeof value === 'number' ? 'n' : 's';
    ws[addr] = { t, v: value, w: String(value) };
  }

  const d = datos;
  const p1   = Number(d.prod1Precio)     || 0;
  const u1   = Number(d.prod1UnidadesMes) || 0;
  const cv1  = Number(d.prod1Costo)       || 0;
  const cf   = Number(d.costosFijosTotal) || 0;
  const inv  = Number(d.inversion)        || 0;
  const ap   = Number(d.aportePropio)     || 0;
  const np   = Number(d.numPersonas)      || 1;

  // Precios años 2 y 3 (+4% y +8%)
  const p1a2 = Math.round(p1 * 1.04);
  const p1a3 = Math.round(p1 * 1.08);
  // Unidades años 2 y 3 (+10% y +20%)
  const u1a2 = Math.round(u1 * 1.10);
  const u1a3 = Math.round(u1 * 1.20);

  // ── Datos básicos ──
  set('B10', d.nombre || '');
  set('T10', d.tipoDoc || 'CC');
  set('AC10', d.numDoc || '');
  set('B14', d.nombreProyecto || '');
  set('T14', d.departamento || '');
  set('AC14', d.municipio || '');
  set('B18', d.tipoProyecto || '');
  set('L18', d.sector || '');
  set('Z18', d.ciiu || '');
  set('R21', /^s/i.test(d.asociativo || '') ? 'SI' : 'NO');
  if (/^s/i.test(d.asociativo || '')) set('AN21', np);
  set('R23', /^s/i.test(d.lugarOps || '') ? 'SI' : 'NO');

  // ── Sección 1 — Cliente ──
  set('A61', d.clienteCarac || '');
  set('S61', d.clienteCual || '');

  // ── Sección 2 — Problema ──
  set('A77', d.problema || '');

  // ── Sección 3 — Competidores ──
  // Filas 94, 95, 96 · cols: A(nombre) I(loc) R(prod) X(precio) AC(ventajas) AK(desventajas)
  for (let i = 0; i < 3; i++) {
    const comp = comps[i];
    if (!comp) continue;
    const r = 94 + i;
    set(`A${r}`, comp.nombre || '');
    set(`I${r}`, comp.localizacion || '');
    set(`R${r}`, comp.producto || '');
    set(`X${r}`, comp.precio || '');
    set(`AC${r}`, comp.ventajas || '');
    set(`AK${r}`, comp.desventajas || '');
  }

  // ── Sección 4 — Descripción y propuesta de valor ──
  set('I102', d.descripcion || '');
  set('G105', d.pvNuestro || '');
  set('G106', d.pvAyuda || '');
  set('G107', d.pvQue || '');
  set('G108', d.pvMediante || '');

  // ── Sección 5 — Productos ──
  // Tabla de nombres (fila 116)
  set('A116', d.prod1Nombre || '');
  set('K116', d.prod1Desc || '');
  set('Y116', d.prod1Unidad || '');
  set('AF116', d.prod1Unidad || '');

  // ── Sección 9 — Precios proyectados ──
  // Fila 229: precios producto 1 para años 1, 2, 3
  set('A229', d.prod1Nombre || 'Producto 1');
  set('Q229', p1);
  set('W229', p1a2);
  set('AC229', p1a3);

  // ── Sección 9 — Unidades mensuales (filas 241-252) ──
  // Col G = año 1 · Col S = año 2 · Col AE = año 3
  for (let mes = 0; mes < 12; mes++) {
    const row = 241 + mes;
    set(`G${row}`, u1);
    set(`S${row}`, u1a2);
    set(`AE${row}`, u1a3);
  }

  // ── Sección 10 — Costos fijos (filas 297-304) ──
  // Cols: A=descripcion · T=valor mensual
  for (let i = 0; i < Math.min(costosFijosItems.length, 8); i++) {
    const row = 297 + i;
    const item = costosFijosItems[i];
    set(`A${row}`, item.descripcion || '');
    set(`T${row}`, Number(item.valorMensual) || 0);
  }

  // ── Sección 10 — Costos variables producto 1 (fila 318) ──
  // A318=descripcion · T318=costo unitario
  set('A318', 'Materias primas e insumos');
  set('T318', cv1);

  // ── Sección 10 — % Participación producto 1 (fila 349) ──
  set('Q349', 100);  // col Q para producto 1, 100% si solo hay 1 producto

  // ── Sección 12 — Inversiones fijas (filas 402-421) ──
  // Cols: A=descripcion · AA=cantidad · AD=valor · AO=aporte FE
  for (let i = 0; i < Math.min(inversionItems.length, 8); i++) {
    const row = 402 + i;
    const item = inversionItems[i];
    set(`A${row}`,  item.descripcion || '');
    set(`AA${row}`, Number(item.cantidad) || 1);
    set(`AD${row}`, Number(item.valor) || 0);
  }

  // ── Sección 13 — Valor del proyecto ──
  set('AE448', inv);
  set('AE449', ap);
  set('AE450', Math.max(0, inv - ap));

  // ── Sección 14 — Avances ──
  set('G459', 'En proceso');   // Legal
  set('G460', 'Identificado'); // Comercial

  // ── Sección 16 — Impacto ──
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

  const d = sesion.datos;
  const p   = Number(d.prod1Precio)     || 0;
  const u   = Number(d.prod1UnidadesMes) || 0;
  const cv  = Number(d.prod1Costo)       || 0;
  const cf  = Number(d.costosFijosTotal) || 0;
  const inv = Number(d.inversion)        || 0;
  const ap  = Number(d.aportePropio)     || 0;
  const np  = Number(d.numPersonas)      || 1;

  // 1. Parsear competidores, costos e inversión en paralelo
  const [comp1, comp2, comp3, costosFijosItems, inversionItems] = await Promise.all([
    parsearCompetidor(d.competidor1),
    parsearCompetidor(d.competidor2),
    parsearCompetidor(d.competidor3),
    parsearCostosFijos(d.costosFijosDesc, cf),
    parsearInversion(d.inversionDesc, inv),
  ]);
  const comps = [comp1, comp2, comp3];

  // 2. Generar plan narrativo
  let planIA = null;
  try {
    planIA = await generarPlanCompleto(d);
  } catch (e) {
    console.error('Claude plan error:', e.message);
    await ctx.reply('⚠️ No pude generar las secciones narrativas, pero el Excel oficial sí va a salir.');
  }

  // 3. Llenar Excel oficial SENA
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

  const fnOficial = `PROYECTO_CREAR_${nombre}_${fecha}.xlsx`;
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

  // 4. Si hay plan narrativo, enviar también el plan completo
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

    // Hoja 1 — Datos del beneficiario
    const ws1 = XLSX.utils.aoa_to_sheet([
      ['PLAN DE NEGOCIO — SENA LÍNEA CREAR ESPECIAL'],
      ['Bot de voz · ' + fecha],
      [],
      ['BENEFICIARIO'],
      ['Nombre completo',     d.nombre        || ''],
      ['Tipo / N° documento', (d.tipoDoc || 'CC') + ' ' + (d.numDoc || '')],
      ['Género',              d.genero        || ''],
      ['Municipio',           d.municipio     || ''],
      ['Departamento',        d.departamento  || ''],
      ['Grupo poblacional',   d.grupoPoblacional || ''],
      [],
      ['PROYECTO'],
      ['Nombre del proyecto', d.nombreProyecto || ''],
      ['Tipo de proyecto',    d.tipoProyecto  || ''],
      ['Sector / CIIU',       (d.sector || '') + ' · ' + (d.ciiu || '')],
      ['Modalidad',           /^s/i.test(d.asociativo || '') ? `Asociativo ${np} personas` : 'Individual'],
      ['Maletín de formación', fmt(kit)],
      ['Lugar de operaciones', /^s/i.test(d.lugarOps || '') ? 'SÍ' : 'NO'],
      [],
      ['PROPUESTA DE VALOR'],
      ['Descripción',         d.descripcion || ''],
      ['Nuestro:',            d.pvNuestro   || ''],
      ['Ayuda a:',            d.pvAyuda     || ''],
      ['A que:',              d.pvQue       || ''],
      ['Mediante:',           d.pvMediante  || ''],
      [],
      ['PRODUCTO PRINCIPAL'],
      ['Nombre',       d.prod1Nombre || ''],
      ['Descripción',  d.prod1Desc   || ''],
      ['Unidad',       d.prod1Unidad || ''],
    ]);
    ws1['!cols'] = [{ wch: 28 }, { wch: 65 }];
    XLSX.utils.book_append_sheet(wb2, ws1, 'DATOS');

    // Hoja 2 — Modelo financiero
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

    // Hoja 3 — Plan narrativo
    const secciones = [
      ['DESCRIPCIÓN DEL NEGOCIO',     planIA.descripcionNegocio],
      ['MERCADO OBJETIVO',             planIA.mercadoObjetivo],
      ['ANÁLISIS DE COMPETENCIA',      planIA.analisisCompetencia],
      ['ESTRATEGIA COMERCIAL',         planIA.estrategiaComercial],
      ['PLAN OPERATIVO',               planIA.planOperativo],
      ['ANÁLISIS DE RIESGOS',          planIA.analisisRiesgos],
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

    const fnPlan = `PlanNarrado_${nombre}_${fecha}.xlsx`;
    const tmpPlan = `/tmp/${fnPlan}`;
    XLSX.writeFile(wb2, tmpPlan);
    await ctx.replyWithDocument(new InputFile(tmpPlan, fnPlan), {
      caption: '📝 *Plan Narrativo + Modelo Financiero*\n_Secciones redactadas por IA para el expediente_',
      parse_mode: 'Markdown',
    });
    try { fs.unlinkSync(tmpPlan); } catch {}
  }

  await ctx.reply('✨ Listo. El Excel oficial SENA ya tiene los datos listos para SharePoint.\nUsa /nuevo para el siguiente beneficiario.');
}

// ── Procesar respuesta y avanzar ──────────────────────────
async function procesarRespuesta(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = getSesion(chatId);
  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) return;

  const { p, i } = actual;
  let valor = p.key === 'nombre' ? limpiarNombre(texto) : texto;

  // Campos numéricos: parsear números verbales con Claude
  if (p.numerico) {
    valor = String(await extraerNumero(texto));
  }

  sesion.datos[p.key] = valor;
  sesion.paso = i + 1;

  const preview = valor.length > 120 ? valor.slice(0, 120) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });

  const sig = getPregunta(sesion.paso, sesion.datos);
  if (sig) {
    const progreso = `_(${sesion.paso}/${PREGUNTAS.length})_`;
    await ctx.reply(`${sig.p.msg}\n\n${progreso}`, { parse_mode: 'Markdown' });
  } else {
    sesion.modo = 'revisando';
    await mostrarResumen(ctx, sesion);
  }
}

// ── Comandos ──────────────────────────────────────────────
bot.command('start', async ctx => {
  sesiones.delete(ctx.chat.id);
  await ctx.reply(
    '👋 *Bienvenido al Bot SENA · Línea CREAR Especial*\n\n' +
    '*Plan de Negocio:*\n' +
    '📝 /nuevo — Crear plan de negocio (38 preguntas)\n' +
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
  sesiones.delete(ctx.chat.id);
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
  sesiones.delete(ctx.chat.id);
  await ctx.reply('🔄 Plan cancelado. Usa /nuevo para empezar de cero.');
});

bot.command('atras', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = getSesion(chatId);
  if (sesion.modo === 'revisando') {
    sesion.modo = undefined;
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
  const sesion = sesiones.get(chatId);
  if (!sesion || Object.keys(sesion.datos).length === 0) {
    await ctx.reply('No hay ningún plan en curso. Usa /nuevo para empezar.');
    return;
  }
  sesion.modo = 'revisando';
  await mostrarResumen(ctx, sesion);
});

bot.command('listo', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = sesiones.get(chatId);
  if (!sesion || sesion.modo !== 'revisando') {
    await ctx.reply('Primero completa el cuestionario. Usa /nuevo para empezar.');
    return;
  }
  await finalizar(ctx, sesion);
  sesiones.delete(chatId);
});

bot.command('estado', async ctx => {
  const sesion = getSesion(ctx.chat.id);
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

// ── Corrección desde resumen ──────────────────────────────
async function aplicarCorreccion(ctx, sesion, texto) {
  const key   = sesion.corrigiendoKey;
  const p     = PREGUNTAS.find(q => q.key === key);
  let valor   = key === 'nombre' ? limpiarNombre(texto) : texto;
  if (p && p.numerico) valor = String(await extraerNumero(texto));
  sesion.datos[key] = valor;
  sesion.modo = 'revisando';
  const preview = valor.length > 80 ? valor.slice(0, 80) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });
  await mostrarResumen(ctx, sesion);
}

// ── Mensajes de voz / audio ───────────────────────────────
bot.on(['message:voice', 'message:audio'], async ctx => {
  const chatId = ctx.chat.id;

  // ── Flujo diagnóstico ──
  if (dxSesiones.has(chatId)) {
    const dxSesion = dxSesiones.get(chatId);
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
  const sesion       = getSesion(chatId);
  const enRevision   = sesion.modo === 'revisando';
  const enCorreccion = sesion.modo === 'corrigiendo';
  const actual       = getPregunta(sesion.paso, sesion.datos);

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
      await aplicarCorreccion(ctx, sesion, texto);
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
  if (dxSesiones.has(chatId)) {
    const dxSesion = dxSesiones.get(chatId);
    if (dxSesion.paso >= DX_TOTAL) {
      await ctx.reply('El diagnóstico está completo. Usa /dx_listo para generar el Excel.');
      return;
    }
    await dxProcesarRespuesta(ctx, ctx.message.text);
    return;
  }

  // ── Flujo plan de negocio ──
  const sesion = getSesion(chatId);

  if (sesion.modo === 'revisando') {
    const n = parseInt(ctx.message.text.trim(), 10);
    if (!isNaN(n) && n >= 1) {
      const items = getResumenItems(sesion);
      const item  = items.find(it => it.num === n);
      if (item) {
        sesion.modo = 'corrigiendo';
        sesion.corrigiendoKey = item.key;
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
    await aplicarCorreccion(ctx, sesion, ctx.message.text);
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

const dxSesiones = new Map(); // chatId → { paso, datos }

// ── Datos básicos del diagnóstico (10 preguntas) ──────────
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

// ── 40 preguntas diagnóstico (escala 0/1/2) ───────────────
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

// ── clasificar012: texto libre → 0, 1 o 2 (Claude) ───────
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
    return isNaN(n) ? 0 : Math.min(2, Math.max(0, n));
  } catch { return 0; }
}

// ── Generar Excel de Diagnóstico ──────────────────────────
async function generarExcelDx(datos) {
  const templatePath = path.join(__dirname, 'dx_template.xlsx');
  const wb  = XLSX.readFile(templatePath, { cellFormula: true, cellStyles: true });
  const wsDx  = wb.Sheets['Diagnóstico'];
  const wsDup = wb.Sheets['Datos de la Unidad Productiva'];

  function setD(ws, addr, value) {
    if (value === null || value === undefined || value === '') return;
    const t = typeof value === 'number' ? 'n' : 's';
    ws[addr] = { t, v: value, w: String(value) };
  }

  const fecha = new Date().toLocaleDateString('es-CO');

  // ── Hoja Diagnóstico — cabecera ──
  setD(wsDx, 'B4', datos.dxNegocio    || '');
  setD(wsDx, 'B5', datos.dxNombre     || '');
  setD(wsDx, 'B7', datos.dxDinamizador || '');
  setD(wsDx, 'B8', fecha);

  // ── Hoja Diagnóstico — puntajes (col B) ──
  const scores = datos.scores || Array(40).fill(0);
  // Área 1: filas 11-20
  for (let i = 0; i < 10; i++) setD(wsDx, `B${11 + i}`, scores[i]);
  // Área 2: filas 23-32
  for (let i = 0; i < 10; i++) setD(wsDx, `B${23 + i}`, scores[10 + i]);
  // Área 3: filas 35-44
  for (let i = 0; i < 10; i++) setD(wsDx, `B${35 + i}`, scores[20 + i]);
  // Área 4: filas 47-56
  for (let i = 0; i < 10; i++) setD(wsDx, `B${47 + i}`, scores[30 + i]);

  // ── Hoja DUP — datos básicos ──
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

// ── Helpers diagnóstico ───────────────────────────────────
function getDxSesion(chatId) {
  if (!dxSesiones.has(chatId)) {
    dxSesiones.set(chatId, { paso: 0, datos: { scores: Array(40).fill(null) } });
  }
  return dxSesiones.get(chatId);
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
  const pct = (total / 80 * 100).toFixed(0);
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

async function dxProcesarRespuesta(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = getDxSesion(chatId);
  const paso   = sesion.paso;

  if (paso >= DX_TOTAL) {
    await ctx.reply('El diagnóstico ya está completo. Usa /dx_listo para generar el Excel.');
    return;
  }

  if (paso < DX_BASICOS.length) {
    // Datos básicos: guardar como texto
    sesion.datos[DX_BASICOS[paso].key] = texto;
  } else {
    // Pregunta diagnóstico: clasificar como 0/1/2
    const qi    = paso - DX_BASICOS.length;
    const score = await clasificar012(texto);
    sesion.datos.scores[qi] = score;
    await ctx.reply(`✅ *${score}*`, { parse_mode: 'Markdown' });
  }

  sesion.paso++;

  if (sesion.paso >= DX_TOTAL) {
    // Fin del cuestionario
    await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
  } else {
    // Siguiente pregunta
    const msg = dxPreguntaMsg(sesion.paso);
    await ctx.reply(msg, { parse_mode: 'Markdown' });
  }
}

// ── Comandos diagnóstico ──────────────────────────────────
bot.command('dx', async ctx => {
  const chatId = ctx.chat.id;
  dxSesiones.delete(chatId);
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
  const sesion = getDxSesion(chatId);
  await ctx.reply(dxPreguntaMsg(0), { parse_mode: 'Markdown' });
});

bot.command('dx_listo', async ctx => {
  const chatId = ctx.chat.id;
  const sesion = dxSesiones.get(chatId);
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

  // Mostrar resumen de puntajes
  await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
  dxSesiones.delete(chatId);
  await ctx.reply('✨ Listo. Usa /dx para un nuevo diagnóstico o /nuevo para el plan de negocio.');
});

bot.command('dx_resumen', async ctx => {
  const sesion = dxSesiones.get(ctx.chat.id);
  if (!sesion) {
    await ctx.reply('No hay diagnóstico en curso. Usa /dx para empezar.');
    return;
  }
  await ctx.reply(dxResumenMsg(sesion.datos), { parse_mode: 'Markdown' });
});

// ── Health check (Render necesita puerto HTTP abierto) ────
const PORT = process.env.PORT || 3000;
http.createServer((_, res) => res.end('OK')).listen(PORT);

console.log('🤖 Bot SENA CREAR v3 (Plan + Diagnóstico) iniciando…');
bot.start();
