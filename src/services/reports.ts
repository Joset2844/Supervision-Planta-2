import { CatalogoItem, Maquina, TurnoMaquina } from '../types';
import { buscarEnCatalogo, resolverModelo, etiquetaPliego, AREAS_ORDEN, turnosDeMaquina, nombreTarjeta } from './excel';

export function formatearFechaTurno(fechaIso: string): string {
  const partes = (fechaIso || '').split('-');
  if (partes.length !== 3) {
    return new Date().toLocaleDateString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }
  return `${partes[2]}/${partes[1]}/${partes[0]}`;
}

export function etiquetaTurno(tipo: string): string {
  return tipo === 'noche' ? 'Turno Noche' : 'Turno Día';
}

export function bloqueMaquinaWsp(m: Maquina, t: TurnoMaquina, catalogo: CatalogoItem[], nombreMostrar?: string) {
  const nombre = nombreMostrar || (m.multiOp ? nombreTarjeta(m, t.instancia || 1) : m.nombre);
  const lineas: string[] = [];
  if (t.motivoParada) {
    lineas.push(`▪️ *${nombre}*  ⏸ _${t.motivoParada}_${t.observaciones ? '\n     📝 ' + t.observaciones : ''}`);
    return { lineas, total: 0, merma: 0 };
  }
  if (!t.registros || t.registros.length === 0) {
    const match = buscarEnCatalogo(catalogo, t.op);
    const opTxt = t.op ? t.op : 'sin OP asignada';
    const descTxt = match ? (resolverModelo(match, t.modeloIdx).descripcion ? ' — ' + resolverModelo(match, t.modeloIdx).descripcion : '') : '';
    lineas.push(`▪️ *${nombre}*  ·  0 u\n     ${opTxt}${descTxt}${t.observaciones ? '\n     📝 ' + t.observaciones : ''}`);
    return { lineas, total: 0, merma: 0 };
  }

  const grupos: Record<string, any> = {};
  const ordenGrupos: string[] = [];
  const opsConGrupo: string[] = [];

  t.registros.forEach((r) => {
    const opKey = r.op || t.op || '';
    const descKey = r.descripcion || '';
    const pliegoKey = r.pliego || '';
    const claveGrupo = `${opKey}||${descKey}||${pliegoKey}`;
    if (!grupos[claveGrupo]) {
      const matchG = buscarEnCatalogo(catalogo, opKey);
      const descBaseG = descKey || (matchG ? resolverModelo(matchG, t.modeloIdx).descripcion : '');
      grupos[claveGrupo] = {
        op: opKey,
        descripcion: descBaseG + (pliegoKey ? ' - ' + etiquetaPliego(pliegoKey) : ''),
        total: 0,
        merma: 0,
        etq: etiquetaPliego(pliegoKey),
        obs: [] as string[],
      };
      ordenGrupos.push(claveGrupo);
      if (opsConGrupo.indexOf(opKey) === -1) opsConGrupo.push(opKey);
    }
    grupos[claveGrupo].total += r.cantidad || 0;
    grupos[claveGrupo].merma += r.merma || 0;
    const o = typeof r.observacion === 'undefined' ? t.observaciones || '' : r.observacion;
    if (o && grupos[claveGrupo].obs.indexOf(o) === -1) grupos[claveGrupo].obs.push(o);
  });

  let totalMaquina = 0;
  let mermaMaquina = 0;
  const detalle = ordenGrupos.map((claveGrupo) => {
    const g = grupos[claveGrupo];
    totalMaquina += g.total;
    mermaMaquina += g.merma;
    const opTxt = g.op ? g.op : 'sin OP asignada';
    const descTxt = g.descripcion ? ' — ' + g.descripcion : '';
    const mermaTxt = g.merma > 0 ? `  (merma ${g.merma})` : '';
    const obsWsp = (g.etq ? [g.etq] : []).concat(g.obs);
    const obsTxt = obsWsp.length ? '\n       📝 ' + obsWsp.join(' / ') : '';
    return `     ${opTxt}${descTxt}  *${g.total} u*${mermaTxt}${obsTxt}`;
  });

  if (t.op && opsConGrupo.indexOf(t.op) === -1) {
    const matchActual = buscarEnCatalogo(catalogo, t.op);
    const descActual = matchActual ? resolverModelo(matchActual, t.modeloIdx).descripcion : '';
    detalle.push(`     ${t.op}${descActual ? ' — ' + descActual : ''}  *0 u*`);
  }
  if (t.observaciones) detalle.push(`     📝 (pendiente) ${t.observaciones}`);

  lineas.push(`▪️ *${nombre}*  ·  ${totalMaquina} u${mermaMaquina > 0 ? ` (merma ${mermaMaquina})` : ''}`);
  return { lineas: lineas.concat(detalle), total: totalMaquina, merma: mermaMaquina };
}

export function generarReporteProduccion(maquinas: Maquina[], turno: TurnoMaquina[], catalogo: CatalogoItem[]): string {
  const fecha = new Date().toLocaleString('es-PE', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
  const activas = maquinas
    .filter((m) => m.activa)
    .sort((a, b) => {
      const getR = (area: string) => (area === 'vasos_envases' ? 0 : area === 'papel_carton' ? 1 : 2);
      const ra = getR(a.area);
      const rb = getR(b.area);
      if (ra !== rb) return ra - rb;
      return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
    });

  let l: string[] = [];
  l.push('🏭 *REPORTE DE TURNO*');
  l.push('🗓️ ' + fecha);
  l.push('━━━━━━━━━━━━━━━');
  if (activas.length === 0) {
    l.push('');
    l.push('No hay máquinas activas.');
    return l.join('\n');
  }

  let totalGeneral = 0;
  let mermaGeneral = 0;

  AREAS_ORDEN.forEach((areaDef) => {
    const deArea = activas.filter((m) => (m.area || '') === areaDef.key);
    if (deArea.length === 0) return;
    l.push('');
    l.push(`🔹 *${areaDef.label}*`);
    deArea.forEach((m) => {
      turnosDeMaquina(m, turno).forEach((t) => {
        const nom = m.multiOp ? nombreTarjeta(m, t.instancia || 1) : m.nombre;
        const b = bloqueMaquinaWsp(m, t, catalogo, nom);
        totalGeneral += b.total;
        mermaGeneral += b.merma;
        l.push('');
        l = l.concat(b.lineas);
      });
    });
  });

  l.push('');
  l.push('━━━━━━━━━━━━━━━');
  l.push(`✅ *Total turno: ${totalGeneral} u*${mermaGeneral > 0 ? `   ⚠️ Merma: ${mermaGeneral} u` : ''}`);
  return l.join('\n');
}

export function generarReporteActivacion(maquinas: Maquina[], turno: TurnoMaquina[], catalogo: CatalogoItem[], fechaIso: string, tipo: string): string {
  const activas = maquinas
    .filter((m) => m.activa)
    .sort((a, b) => {
      const getR = (area: string) => (area === 'vasos_envases' ? 0 : area === 'papel_carton' ? 1 : 2);
      const ra = getR(a.area);
      const rb = getR(b.area);
      if (ra !== rb) return ra - rb;
      return a.nombre.localeCompare(b.nombre, 'es', { sensitivity: 'base' });
    });

  let l: string[] = [];
  l.push(`🟢 *ACTIVACIÓN DE MÁQUINAS* · ${formatearFechaTurno(fechaIso)}${tipo ? ' · ' + etiquetaTurno(tipo) : ''}`);
  if (activas.length === 0) {
    l.push('No hay máquinas activas.');
    return l.join('\n');
  }
  let detenidas = 0;
  let totalTarjetas = 0;

  AREAS_ORDEN.forEach((areaDef) => {
    const deArea = activas.filter((m) => (m.area || '') === areaDef.key);
    if (deArea.length === 0) return;
    l.push(`🔹 *${areaDef.label}* (${deArea.length})`);
    deArea.forEach((m) => {
      turnosDeMaquina(m, turno).forEach((t) => {
        totalTarjetas++;
        const nom = m.multiOp ? nombreTarjeta(m, t.instancia || 1) : m.nombre;
        const match = buscarEnCatalogo(catalogo, t.op);
        const desc = match ? resolverModelo(match, t.modeloIdx).descripcion : '';
        const partes: string[] = [];
        if (t.motivoParada) {
          detenidas++;
          partes.push(`⏸ _${t.motivoParada}_`);
        }
        if (t.op) {
          partes.push(`🔖 ${t.op}${t.pliego ? ' (' + etiquetaPliego(t.pliego) + ')' : ''}${desc ? ' · ' + desc : ''}`);
        } else if (!t.motivoParada) {
          partes.push('⚠️ _sin OP_');
        }
        if (t.observaciones) partes.push(`📝 ${t.observaciones}`);
        l.push(`▪️ *${nom}*: ${partes.join(' · ')}`);
      });
    });
  });
  l.push(`✅ *${totalTarjetas} activa(s)*${detenidas ? ` · ⏸ ${detenidas} detenida(s)` : ''}`);
  return l.join('\n');
}

export async function copiarAlPortapapeles(texto: string): Promise<boolean> {
  let copiado = false;
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(texto);
      copiado = true;
    }
  } catch {}
  if (!copiado) {
    try {
      const tmp = document.createElement('textarea');
      tmp.value = texto;
      tmp.style.position = 'fixed';
      tmp.style.opacity = '0';
      document.body.appendChild(tmp);
      tmp.focus();
      tmp.select();
      copiado = document.execCommand('copy');
      document.body.removeChild(tmp);
    } catch {}
  }
  return copiado;
}
