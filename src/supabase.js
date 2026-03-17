import { createClient } from '@supabase/supabase-js'

export const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

// ── Auth ──────────────────────────────────────────────────────────────
export const authSignIn = (email, password) =>
  supabase.auth.signInWithPassword({ email, password })

export const authSignUp = (email, password, fullName) =>
  supabase.auth.signUp({
    email, password,
    options: { data: { full_name: fullName } }
  })

export const authSignOut = () => supabase.auth.signOut()

// ── Profile ───────────────────────────────────────────────────────────
export const getProfile = async (userId) => {
  const { data, error } = await supabase
    .from('profiles').select('*').eq('id', userId).single()
  return { data, error }
}

export const touchLastSeen = (userId) =>
  supabase.from('profiles')
    .update({ last_seen: new Date().toISOString() })
    .eq('id', userId)

// ── Scenarios ─────────────────────────────────────────────────────────
export const fetchScenarios = async (userId) => {
  const { data, error } = await supabase
    .from('scenarios').select('*').eq('user_id', userId)
    .order('updated_at', { ascending: false })
  return { data: data || [], error }
}

export const upsertScenario = async (userId, scen) => {
  const row = {
    user_id:     userId,
    name:        scen.name,
    description: scen.description || '',
    overrides:   scen.overrides   || {},
    total_capex: scen.totalCapex  || null,
    delta_pct:   scen.deltaPct    || null,
  }
  if (scen.dbId) {
    const { data, error } = await supabase
      .from('scenarios').update(row)
      .eq('id', scen.dbId).eq('user_id', userId)
      .select().single()
    return { data, error }
  }
  const { data, error } = await supabase
    .from('scenarios').insert(row).select().single()
  return { data, error }
}

export const removeScenario = (userId, dbId) =>
  supabase.from('scenarios').delete()
    .eq('id', dbId).eq('user_id', userId)

// ── Audit log ─────────────────────────────────────────────────────────
export const pushLog = (entry) =>
  supabase.from('audit_log').insert(entry)

export const fetchLog = async (userId, limit = 150) => {
  const { data, error } = await supabase
    .from('audit_log').select('*').eq('user_id', userId)
    .order('created_at', { ascending: false }).limit(limit)
  return { data: data || [], error }
}
