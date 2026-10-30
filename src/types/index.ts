export interface CatalogoModelo {
  modelo?: string;
  descripcion: string;
  sap?: string;
  cantidadTotal?: string | number;
}

export interface CatalogoItem {
  op: string;
  descripcion: string;
  cliente: string;
  cantidadTotal?: string | number;
  sap?: string;
  descripcionBase?: string;
  modelos: CatalogoModelo[];
}

export interface ListaItem {
  id: string;
  op: string;
  descripcion: string;
  sap: string;
  cliente: string;
  observacion?: string;
  tipoEmpaque: 'caja' | 'paquete';
  unidadesXCaja: number;
  cantidadCajas: number;
  totalUnidades: number;
  fecha: string;
}

export interface Maquina {
  id: string;
  nombre: string;
  activa: boolean;
  area: 'vasos_envases' | 'papel_carton' | '';
  codigo: string;
  bloqueada: boolean;
  tipoProducto: 'pp' | 'pt' | '';
  multiOp?: boolean;
}

export interface RegistroTurno {
  id: string;
  op: string;
  descripcion: string;
  sap?: string;
  pliego?: number | string;
  cantidad: number;
  merma: number;
  observacion?: string;
  fecha: string;
}

export interface TurnoMaquina {
  maquinaId: string;
  instancia?: number;
  op: string;
  objetivo?: number | null;
  motivoParada: string;
  observaciones: string;
  personal?: string;
  modeloIdx: number;
  pliego?: number | string;
  registros: RegistroTurno[];
}

export interface FilaTurnoExcel {
  maquina: string;
  codigoMaquina: string;
  tipoProducto: string;
  area: string;
  op: string;
  sap: string;
  producto: string;
  total: number;
  merma: number;
  observaciones: string;
}

export interface HistorialTurnoLocal {
  id: string;
  fecha: string;
  fechaTurno?: string;
  tipoTurno?: string;
  resumen: string;
  filas: FilaTurnoExcel[];
}

export interface PlantillaItem {
  maquinaId: string;
  op: string;
  objetivo?: number | null;
}

export interface MetaTurno {
  tipo: 'dia' | 'noche' | '';
  fecha: string;
}

export interface SupabaseConfig {
  url: string;
  key: string;
}

export interface SupabaseSession {
  access_token: string;
  refresh_token: string;
  user_id: string;
  email: string;
  usuario: string;
}

export interface TurnoDBRow {
  fecha: string;
  tipo: 'dia' | 'noche' | string;
  filas: FilaTurnoExcel[];
  turno?: TurnoMaquina[];
  created_by?: string;
  created_by_email?: string;
  updated_at?: string;
  creado_en?: string;
}

export interface TurnoVersionDBRow {
  id: string;
  fecha: string;
  tipo: string;
  filas: FilaTurnoExcel[];
  turno?: TurnoMaquina[];
  created_by_email?: string;
  updated_at?: string;
  archivado_en: string;
  accion?: string;
}

export type TabType = 'inventario' | 'maquinas' | 'produccion' | 'resumen' | 'historialdb';

