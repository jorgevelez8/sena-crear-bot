'use strict';
require('dotenv').config();

const { Bot, InputFile } = require('grammy');
const Groq  = require('groq-sdk');
const axios = require('axios');
const XLSX  = require('xlsx');
const fs    = require('fs');
const http  = require('http');

// ── Clientes API ──────────────────────────────────────────
const bot       = new Bot(process.env.TELEGRAM_TOKEN);
const groq      = new Groq({ apiKey: process.env.GROQ_API_KEY });
const Anthropic = require('@anthropic-ai/sdk');
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Estado en memoria  chatId → { paso, datos } ──────────
const sesiones = new Map();

// ── Preguntas del plan (orden del wizard) ─────────────────
const PREGUNTAS = [
  {
    key: 'nombre',
    msg: '¡Hola! 👋 Vamos a crear el *Plan de Negocio* paso a paso.\nPuedes responder con voz 🎤 o escribiendo.\n\n¿Cuál es el *nombre completo* del beneficiario?',
  },
  { key: 'numDoc',         msg: '¿Cuál es el *número de cédula*?' },
  { key: 'municipio',      msg: '¿En qué *municipio* está el proyecto?' },
  { key: 'departamento',   msg: '¿En qué *departamento*?' },
  {
    key: 'grupo',
    msg: '¿A qué *grupo poblacional* pertenece?\n_(Ej: Víctima de violencia, Desplazado, Campesino, Indígena, Afrocolombiano...)_',
  },
  { key: 'nombreProyecto', msg: '¿Cuál es el *nombre del proyecto*?' },
  { key: 'tipoProyecto',   msg: '¿Es *Economía Popular* o *Economía Campesina*?' },
  { key: 'asociativo',     msg: '¿El proyecto es *individual* o *asociativo* (varias personas)?' },
  {
    key: 'numPersonas',
    msg: '¿Cuántas *personas* conforman el grupo?',
    soloSi: d => /asoc/i.test(d.asociativo || ''),
  },
  {
    key: 'sector',
    msg: '¿En qué *sector* trabaja?\n_(Ej: Agricultura, Comercio, Artesanías, Turismo, Manufactura...)_',
  },
  { key: 'descripcion',  msg: '¿En qué *consiste el negocio*? Cuénteme con sus propias palabras.' },
  { key: 'cliente',      msg: '¿*Quién le compra*? Describa a su cliente.' },
  { key: 'problema',     msg: '¿Qué *problema o necesidad* resuelve su negocio?' },
  { key: 'producto1',    msg: '¿Cuál es el *producto o servicio principal*?' },
  { key: 'precio1',      msg: '¿A qué *precio* lo vende?\n_(Solo el número en pesos, ej: 15000)_' },
  { key: 'unidadesMes',  msg: '¿Cuántas *unidades vende al mes* aproximadamente?\n_(Solo el número)_' },
  {
    key: 'costoUnit',
    msg: '¿Cuánto le *cuesta producir* cada unidad?\n_(Materias primas e insumos — solo número)_',
  },
  {
    key: 'costosFijos',
    msg: '¿Cuánto paga de *gastos fijos al mes* en total?\n_(Arriendo, servicios, internet — solo número)_',
  },
  {
    key: 'inversion',
    msg: '¿Cuánto necesita *invertir en total* para el proyecto?\n_(Maquinaria, adecuaciones, permisos — número)_',
  },
  {
    key: 'aporte',
    msg: '¿Cuánto puede *aportar usted* de esa inversión?\n_(Mínimo el 10% del total — número)_',
  },
];

// ── Helpers ───────────────────────────────────────────────
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

// Devuelve la lista de preguntas respondidas con número de display
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

// ── Plan completo con Claude ──────────────────────────────
async function generarPlanCompleto(datos) {
  const p   = Number(datos.precio1)     || 0;
  const u   = Number(datos.unidadesMes) || 0;
  const cv  = Number(datos.costoUnit)   || 0;
  const cf  = Number(datos.costosFijos) || 0;
  const inv = Number(datos.inversion)   || 0;
  const ap  = Number(datos.aporte)      || 0;
  const np  = Number(datos.numPersonas) || 1;

  const prompt = `Eres un asesor del SENA Colombia experto en planes de negocio para la Línea CREAR Especial (víctimas del conflicto, campesinos, desplazados). Redacta en español claro, formal y empático. Máximo 200 palabras por sección. Responde SOLO con el JSON sin texto adicional ni bloques de código.

DATOS DEL BENEFICIARIO:
- Nombre: ${datos.nombre} | Grupo: ${datos.grupo}
- Ubicación: ${datos.municipio}, ${datos.departamento}
- Proyecto: ${datos.nombreProyecto} (${datos.tipoProyecto} · ${datos.sector})
- Modalidad: ${/asoc/i.test(datos.asociativo || '') ? `Asociativo — ${np} personas` : 'Individual'}
- Negocio: ${datos.descripcion}
- Producto/servicio: ${datos.producto1} a $${p.toLocaleString('es-CO')} · ${u} unidades/mes
- Cliente objetivo: ${datos.cliente}
- Problema que resuelve: ${datos.problema}
- Costo unitario: $${cv.toLocaleString('es-CO')} · Costos fijos: $${cf.toLocaleString('es-CO')}/mes
- Inversión total: $${inv.toLocaleString('es-CO')} · Aporte propio: $${ap.toLocaleString('es-CO')}

{
  "descripcionNegocio": "Narrativa del negocio: propuesta de valor, origen y contexto del emprendedor.",
  "mercadoObjetivo": "Perfil del cliente ideal, frecuencia de compra, tamaño de mercado local estimado.",
  "analisisCompetencia": "Competidores directos e indirectos en la zona. Ventajas diferenciales del proyecto.",
  "estrategiaComercial": "Canales de venta, estrategia de precios, promoción y fidelización de clientes.",
  "planOperativo": "Proceso productivo paso a paso, proveedores clave, capacidad instalada y horarios.",
  "analisisRiesgos": "3 riesgos principales con probabilidad, impacto y plan de mitigación para cada uno.",
  "justificacionInversion": "Desglose de la inversión por rubro (maquinaria, insumos, adecuaciones, capital de trabajo) con justificación de cada uno."
}`;

  const msg = await anthropic.messages.create({
    model:      'claude-haiku-4-5-20251001',
    max_tokens: 3500,
    messages:   [{ role: 'user', content: prompt }],
  });

  const raw = msg.content[0].text.trim()
    .replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
  return JSON.parse(raw);
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

// ── Generar Excel ─────────────────────────────────────────
function generarExcel(datos, planIA = null) {
  const p   = Number(datos.precio1)     || 0;
  const u   = Number(datos.unidadesMes) || 0;
  const cv  = Number(datos.costoUnit)   || 0;
  const cf  = Number(datos.costosFijos) || 0;
  const inv = Number(datos.inversion)   || 0;
  const ap  = Number(datos.aporte)      || 0;
  const np  = Number(datos.numPersonas) || 1;

  const ven1   = p * u * 12;
  const ven2   = Math.round(p * 1.04 * (u * 1.1) * 12);
  const ven3   = Math.round(p * 1.08 * (u * 1.2) * 12);
  const cv1    = cv * u * 12;
  const cf1    = cf * 12;
  const margen = ven1 - cv1;
  const ebitda = margen - cf1;
  const mc     = p > 0 ? (p - cv) / p : 0;
  const ptoEq  = mc > 0 ? Math.ceil(cf1 / (mc * p)) : 0;
  const fe     = Math.max(0, inv - ap);
  const apPct  = inv > 0 ? (ap / inv * 100) : 0;
  const kit    = maletin(np);

  const wb = XLSX.utils.book_new();

  // Hoja 1 — Proyecto
  const ws1 = XLSX.utils.aoa_to_sheet([
    ['PLAN DE NEGOCIO — SENA LÍNEA CREAR ESPECIAL'],
    ['Generado por bot de voz · ' + new Date().toLocaleDateString('es-CO')],
    [],
    ['INFORMACIÓN DEL BENEFICIARIO'],
    ['Nombre completo',     datos.nombre        || ''],
    ['Número de cédula',    datos.numDoc        || ''],
    ['Municipio',           datos.municipio     || ''],
    ['Departamento',        datos.departamento  || ''],
    ['Grupo poblacional',   datos.grupo         || ''],
    [],
    ['INFORMACIÓN DEL PROYECTO'],
    ['Nombre del proyecto', datos.nombreProyecto || ''],
    ['Tipo de proyecto',    datos.tipoProyecto  || ''],
    ['Sector / Actividad',  datos.sector        || ''],
    ['¿Asociativo?',        datos.asociativo    || 'Individual'],
    ['Número de personas',  np],
    ['Maletín de formación', fmt(kit)],
    [],
    ['COMPONENTE COMERCIAL'],
    ['Descripción del negocio',      datos.descripcion || ''],
    ['Cliente objetivo',             datos.cliente     || ''],
    ['Problema que resuelve',        datos.problema    || ''],
    ['Producto / servicio principal', datos.producto1  || ''],
  ]);
  ws1['!cols'] = [{ wch: 30 }, { wch: 65 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'PROYECTO');

  // Hoja 2 — Modelo Financiero
  const ws2 = XLSX.utils.aoa_to_sheet([
    ['MODELO FINANCIERO — LÍNEA CREAR'],
    [],
    ['VENTAS'],
    ['Producto',                        datos.producto1 || ''],
    ['Precio unitario (Año 1)',          p],
    ['Unidades / mes',                  u],
    ['Ventas Año 1',                    ven1],
    ['Ventas Año 2 (+4% precio, +10% cant.)', ven2],
    ['Ventas Año 3 (+8% precio, +20% cant.)', ven3],
    [],
    ['COSTOS'],
    ['Costo variable por unidad',       cv],
    ['Total costos variables Año 1',    cv1],
    ['Costos fijos mensuales',          cf],
    ['Total costos fijos Año 1',        cf1],
    [],
    ['RESULTADOS'],
    ['Margen bruto Año 1',              margen],
    ['EBITDA Año 1',                    ebitda],
    ['Punto de equilibrio (unid./año)', ptoEq],
    [],
    ['INVERSIÓN Y FONDO EMPRENDER'],
    ['Inversión total requerida',       inv],
    ['Aporte del emprendedor',          ap],
    ['% de aporte',                     apPct.toFixed(1) + '%  ' + (apPct >= 10 ? '✓ CUMPLE' : '✗ NO CUMPLE (mín. 10%)')],
    ['Solicitado al Fondo Emprender',   fe],
    ['Maletín de formación',            kit],
  ]);
  ws2['!cols'] = [{ wch: 42 }, { wch: 22 }];
  XLSX.utils.book_append_sheet(wb, ws2, 'MODELO FINANCIERO');

  // Hoja 3 — Plan narrativo generado por IA (solo si Claude respondió)
  if (planIA) {
    const secciones = [
      ['DESCRIPCIÓN DEL NEGOCIO',     planIA.descripcionNegocio],
      ['MERCADO OBJETIVO',             planIA.mercadoObjetivo],
      ['ANÁLISIS DE COMPETENCIA',      planIA.analisisCompetencia],
      ['ESTRATEGIA COMERCIAL',         planIA.estrategiaComercial],
      ['PLAN OPERATIVO',               planIA.planOperativo],
      ['ANÁLISIS DE RIESGOS',          planIA.analisisRiesgos],
      ['JUSTIFICACIÓN DE LA INVERSIÓN', planIA.justificacionInversion],
    ];
    const filas = [
      ['PLAN DE NEGOCIO COMPLETO — SENA LÍNEA CREAR ESPECIAL'],
      ['Generado con IA · ' + new Date().toLocaleDateString('es-CO')],
      [],
    ];
    for (const [titulo, contenido] of secciones) {
      filas.push([titulo]);
      filas.push([contenido || '']);
      filas.push([]);
    }
    const ws3 = XLSX.utils.aoa_to_sheet(filas);
    ws3['!cols'] = [{ wch: 120 }];
    XLSX.utils.book_append_sheet(wb, ws3, 'PLAN NARRATIVO');
  }

  return wb;
}

// ── Procesar respuesta y avanzar al siguiente paso ────────
async function procesarRespuesta(ctx, texto) {
  const chatId = ctx.chat.id;
  const sesion = getSesion(chatId);
  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) return;

  const { p, i } = actual;
  const valor = p.key === 'nombre' ? limpiarNombre(texto) : texto;
  sesion.datos[p.key] = valor;
  sesion.paso = i + 1;

  // Confirmación
  const preview = valor.length > 120 ? valor.slice(0, 120) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });

  // Siguiente pregunta o pasar a revisión
  const sig = getPregunta(sesion.paso, sesion.datos);
  if (sig) {
    const progreso = `_(${sesion.paso}/${PREGUNTAS.length})_`;
    await ctx.reply(`${sig.p.msg}\n\n${progreso}`, { parse_mode: 'Markdown' });
  } else {
    sesion.modo = 'revisando';
    await mostrarResumen(ctx, sesion);
  }
}

// ── Finalizar: generar y enviar Excel ────────────────────
async function finalizar(ctx, sesion) {
  await ctx.reply('🎉 *¡Plan completado!* Generando el plan de negocio con IA… ⏳', { parse_mode: 'Markdown' });

  let planIA = null;
  try {
    planIA = await generarPlanCompleto(sesion.datos);
  } catch (e) {
    console.error('Claude API error:', e.message);
    await ctx.reply('⚠️ No pude generar las secciones narrativas, pero el Excel con los datos está listo.');
  }

  const wb     = generarExcel(sesion.datos, planIA);
  const nombre = (sesion.datos.nombre || 'beneficiario').replace(/\s+/g, '_');
  const fecha  = new Date().toISOString().slice(0, 10);
  const fn     = `PlanNegocio_${nombre}_${fecha}.xlsx`;
  const tmp    = `/tmp/${fn}`;

  XLSX.writeFile(wb, tmp);

  await ctx.replyWithDocument(new InputFile(tmp, fn), {
    caption: [
      '📊 *Plan de Negocio SENA Línea CREAR*',
      `👤 ${sesion.datos.nombre || ''}`,
      `🏭 ${sesion.datos.nombreProyecto || ''}`,
      `📅 ${fecha}`,
    ].join('\n'),
    parse_mode: 'Markdown',
  });

  try { fs.unlinkSync(tmp); } catch {}

  await ctx.reply('✨ Listo. Usa /nuevo para registrar el siguiente beneficiario.');
}

// ── Comandos ──────────────────────────────────────────────
bot.command('start', async ctx => {
  sesiones.delete(ctx.chat.id);
  await ctx.reply(
    '👋 *Bienvenido al Bot SENA · Línea CREAR*\n\n' +
    'Te voy a guiar para crear el *Plan de Negocio* paso a paso.\n' +
    'Puedes responder con voz 🎤 o escribiendo ✍️\n\n' +
    '↩️ /atras — Corregir la respuesta anterior\n' +
    '📋 /resumen — Ver y corregir todas las respuestas\n' +
    '✅ /listo — Generar el plan (al final)\n' +
    '🔄 /reiniciar — Cancelar y empezar de cero\n\n' +
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

// ── Guardar corrección (modo corrigiendo) ─────────────────
async function aplicarCorreccion(ctx, sesion, texto) {
  const key   = sesion.corrigiendoKey;
  const valor = key === 'nombre' ? limpiarNombre(texto) : texto;
  sesion.datos[key] = valor;
  sesion.modo = 'revisando';
  const preview = valor.length > 80 ? valor.slice(0, 80) + '…' : valor;
  await ctx.reply(`✅ _"${preview}"_`, { parse_mode: 'Markdown' });
  await mostrarResumen(ctx, sesion);
}

// ── Mensajes de voz / audio ───────────────────────────────
bot.on(['message:voice', 'message:audio'], async ctx => {
  const sesion    = getSesion(ctx.chat.id);
  const enRevision = sesion.modo === 'revisando';
  const enCorreccion = sesion.modo === 'corrigiendo';
  const actual    = getPregunta(sesion.paso, sesion.datos);

  if (!actual && !enCorreccion) {
    await ctx.reply(enRevision
      ? '📋 Estás revisando el plan. Escribe el número a corregir o */listo*.'
      : 'Usa /nuevo para iniciar un plan.',
      { parse_mode: 'Markdown' });
    return;
  }

  const msg = await ctx.reply('🎤 Transcribiendo…');
  try {
    const fileId = ctx.message.voice?.file_id || ctx.message.audio?.file_id;
    const texto  = await transcribir(fileId);
    await ctx.api.deleteMessage(ctx.chat.id, msg.message_id).catch(() => {});
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
  const sesion = getSesion(chatId);

  // Modo revisando: espera número de pregunta
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

  // Modo corrigiendo: guarda texto y vuelve al resumen
  if (sesion.modo === 'corrigiendo') {
    await aplicarCorreccion(ctx, sesion, ctx.message.text);
    return;
  }

  // Flujo normal
  const actual = getPregunta(sesion.paso, sesion.datos);
  if (!actual) {
    await ctx.reply('Usa /nuevo para iniciar un plan.');
    return;
  }
  await procesarRespuesta(ctx, ctx.message.text);
});

// ── Health check (Render necesita un puerto HTTP abierto) ─
const PORT = process.env.PORT || 3000;
http.createServer((_, res) => res.end('OK')).listen(PORT);

// ── Arranque ──────────────────────────────────────────────
console.log('🤖 Bot SENA CREAR iniciando…');
bot.start();
