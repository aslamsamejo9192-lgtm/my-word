import { createClient, type SupabaseClient } from '@supabase/supabase-js'

/**
 * Optional Supabase client. Created ONLY when both env vars are set;
 * otherwise the app runs fully local on IndexedDB.
 *
 * NEVER put a service-role key in the frontend. If Supabase is used,
 * Row Level Security must be enabled so users only access their own
 * documents. No Supabase account or keys are bundled with this app.
 */
const url = import.meta.env.VITE_SUPABASE_URL as string | undefined
const anonKey = import.meta.env.VITE_SUPABASE_ANON_KEY as string | undefined

export const isSupabaseConfigured = Boolean(url && anonKey)

export const supabase: SupabaseClient | null =
  isSupabaseConfigured && url && anonKey ? createClient(url, anonKey) : null
