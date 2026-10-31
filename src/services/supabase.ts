import { SupabaseConfig, SupabaseSession, TurnoDBRow, TurnoVersionDBRow } from '../types';
import { STORAGE_KEYS } from './storage';

export const DEFAULT_SB_URL = 'https://skoksthbuyrhekfchmug.supabase.co';
export const DEFAULT_SB_KEY =
  'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNrb2tzdGhidXlyaGVrZmNobXVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTAyMTEwNjUsImV4cCI6MjEwNTc4NzA2NX0.1vrW7PVIzPvi5a4kIIc-1EoPIYsZcOTeWVq0vXBr5f4';
export const SB_DOMINIO_USUARIO = 'planta.local';

export function getSupabaseConfig(forcePrompt = false): SupabaseConfig | null {
  let url = '';
  let key = '';
  try {
    url = localStorage.getItem(STORAGE_KEYS.SB_URL) || (import.meta.env.VITE_SUPABASE_URL as string) || DEFAULT_SB_URL;
    key = localStorage.getItem(STORAGE_KEYS.SB_KEY) || (import.meta.env.VITE_SUPABASE_ANON_KEY as string) || DEFAULT_SB_KEY;
  } catch {}

  if (forcePrompt || !url || !key) {
    const u = prompt('URL del proyecto Supabase (https://xxxx.supabase.co):', url);
    if (u === null) return null;
    const k = prompt('Anon public key de Supabase:', key);
    if (k === null) return null;
    url = u.trim().replace(/\/+$/, '');
    key = k.trim();
    if (!url || !key) {
      alert('URL y clave son obligatorias');
      return null;
    }
    try {
      localStorage.setItem(STORAGE_KEYS.SB_URL, url);
      localStorage.setItem(STORAGE_KEYS.SB_KEY, key);
    } catch {}
  }

  return { url: url.replace(/\/+$/, ''), key };
}

export function getStoredSession(): SupabaseSession | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEYS.SB_SESION);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (s && !s.usuario && s.email) {
      s.usuario = s.email.split('@')[0];
    }
    return s;
  } catch {
    return null;
  }
}

export function saveSession(s: SupabaseSession | null) {
  try {
    if (s) {
      localStorage.setItem(STORAGE_KEYS.SB_SESION, JSON.stringify(s));
    } else {
      localStorage.removeItem(STORAGE_KEYS.SB_SESION);
    }
  } catch {}
}

export function sbUsuarioAEmail(usuario: string): string {
  const u = usuario.trim().toLowerCase();
  if (u.includes('@')) {
    return u;
  }
  return u.replace(/\s+/g, '') + '@' + SB_DOMINIO_USUARIO;
}

export async function sbLogin(cfg: SupabaseConfig, usuario: string, pass: string): Promise<SupabaseSession> {
  const email = sbUsuarioAEmail(usuario);
  const resp = await fetch(cfg.url + '/auth/v1/token?grant_type=password', {
    method: 'POST',
    headers: { apikey: cfg.key, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password: pass }),
  });
  const data = await resp.json();
  if (!resp.ok) {
    throw new Error(data.error_description || data.msg || data.message || `HTTP ${resp.status}`);
  }
  const nombreUsuario = data.user.email.split('@')[0];
  const session: SupabaseSession = {
    access_token: data.access_token,
    refresh_token: data.refresh_token,
    user_id: data.user.id,
    email: data.user.email,
    usuario: nombreUsuario,
  };
  saveSession(session);
  return session;
}

export async function sbSignUp(cfg: SupabaseConfig, usuario: string, pass: string): Promise<SupabaseSession> {
  let email = sbUsuarioAEmail(usuario);
  // If no domain was specified and default domain is .local, ensure signup uses .com to satisfy Supabase validation
  if (!usuario.includes('@') && email.endsWith('.local')) {
    email = usuario.trim().toLowerCase().replace(/\s+/g, '') + '@planta.com';
  }
  const resp = await fetch(cfg.url + '/auth/v1/signup', {
    method: 'POST',
    headers: { apikey: cfg.key, 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password: pass }),
  });
  const data = await resp.json();
  if (!resp.ok) {
    throw new Error(data.error_description || data.msg || data.message || `HTTP ${resp.status}`);
  }
  if (data.access_token) {
    const nombreUsuario = data.user.email.split('@')[0];
    const session: SupabaseSession = {
      access_token: data.access_token,
      refresh_token: data.refresh_token,
      user_id: data.user.id,
      email: data.user.email,
      usuario: nombreUsuario,
    };
    saveSession(session);
    return session;
  }
  // Try logging in with the credentials
  return sbLogin(cfg, email, pass);
}

let sbRenovandoPromesa: Promise<SupabaseSession | null> | null = null;
export async function sbRenovarSesion(cfg: SupabaseConfig, session: SupabaseSession): Promise<SupabaseSession | null> {
  if (!session || !session.refresh_token) {
    throw new Error('sin refresh_token');
  }
  if (sbRenovandoPromesa) return sbRenovandoPromesa;

  sbRenovandoPromesa = (async () => {
    try {
      const resp = await fetch(cfg.url + '/auth/v1/token?grant_type=refresh_token', {
        method: 'POST',
        headers: { apikey: cfg.key, 'Content-Type': 'application/json' },
        body: JSON.stringify({ refresh_token: session.refresh_token }),
      });
      const data = await resp.json();
      if (!resp.ok) {
        throw new Error(data.error_description || data.msg || `HTTP ${resp.status}`);
      }
      const newSession: SupabaseSession = {
        access_token: data.access_token,
        refresh_token: data.refresh_token,
        user_id: data.user.id,
        email: data.user.email,
        usuario: data.user.email.split('@')[0],
      };
      saveSession(newSession);
      return newSession;
    } finally {
      sbRenovandoPromesa = null;
    }
  })();

  return sbRenovandoPromesa;
}

export async function sbFetch<T = any>(
  cfg: SupabaseConfig,
  ruta: string,
  opts?: { method?: string; json?: any; prefer?: string },
  session?: SupabaseSession | null,
  reintentado = false
): Promise<T> {
  const options = opts || {};
  const headers: Record<string, string> = {
    apikey: cfg.key,
    Authorization: 'Bearer ' + (session ? session.access_token : cfg.key),
  };
  if (options.json) headers['Content-Type'] = 'application/json';
  if (options.prefer) headers['Prefer'] = options.prefer;

  const resp = await fetch(`${cfg.url}/rest/v1/${ruta}`, {
    method: options.method || 'GET',
    headers,
    cache: 'no-store',
    body: options.json ? JSON.stringify(options.json) : undefined,
  });

  if (!resp.ok) {
    const text = await resp.text();
    if (resp.status === 401 && session && !reintentado) {
      try {
        const renewed = await sbRenovarSesion(cfg, session);
        if (renewed) {
          return sbFetch<T>(cfg, ruta, opts, renewed, true);
        }
      } catch {
        saveSession(null);
        throw new Error('Sesión vencida, inicia sesión de nuevo');
      }
    }
    if (resp.status === 401) {
      saveSession(null);
    }
    throw new Error(`HTTP ${resp.status} ${text}`);
  }

  if (resp.status === 204) {
    return null as any;
  }

  const responseText = await resp.text();
  if (!responseText || !responseText.trim()) {
    return null as any;
  }

  try {
    return JSON.parse(responseText);
  } catch {
    return responseText as any;
  }
}
