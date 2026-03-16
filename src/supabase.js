import { createClient } from '@supabase/supabase-js'

const supabaseUrl  = import.meta.env.VITE_SUPABASE_URL
const supabaseKey  = import.meta.env.VITE_SUPABASE_ANON_KEY

export const supabase = createClient(supabaseUrl, supabaseKey)

// ── Auth ──────────────────────────────────────────────────────────────
export const signIn = (email, password) =>
  supabase.auth.signInWithPassword({ email, password })

export const signInMagicLink = (email) =>
  supabase.auth.signInWithOtp({ email, options: { emailRedirectTo: window.location.origin } })

export const signUp = (email, password, fullName) =>
  supabase.auth.signUp({ email, password, options: { data: { full_name: fullName } } })

export const signOut = () => supabase.auth.signOut()
export const getSession = () => supabase.auth.getSession()

// ── Profile ───────────────────────────────────────────────────────────
export const getProfile = async (userId) => {
  const { data, error } = await supabase
    .from('profiles').select('*').eq('id', userId).single()
  return { data, error }
}

export const updateLastSeen = (userId) =>
  supabase.from('profiles').update({ last_seen: new Date().toISOString() }).eq('id', userId)

// ── Scenarios ─────────────────────────────────────────────────────────
export const getScenarios = async (userId) => {
  const { data, error } = await supabase
    .from('scenarios').select('*').eq('user_id', userId)
    .order('updated_at', { ascending: false })
  return { data: data || [], error }
}

export const saveScenario = async (userId, scenario) => {
  const payload = {
    user_id: userId, name: scenario.name,
    description: scenario.description || '',
    overrides: scenario.overrides,
    total_capex: scenario.totalCapex || null,
    delta_pct: scenario.deltaPct || null,
  }
  if (scenario.id) {
    const { data, error } = await supabase
      .from('scenarios').update(payload)
      .eq('id', scenario.id).eq('user_id', userId).select().single()
    return { data, error }
  }
  const { data, error } = await supabase
    .from('scenarios').insert(payload).select().single()
  return { data, error }
}

export const deleteScenario = async (userId, scenarioId) => {
  const { error } = await supabase
    .from('scenarios').delete().eq('id', scenarioId).eq('user_id', userId)
  return { error }
}

// ── Audit Log ─────────────────────────────────────────────────────────
export const logAdjustment = async ({
  userId, userEmail, projectId, projectName,
  macroName, categoria, field, valueBefore, valueAfter, scenarioId
}) => {
  const { error } = await supabase.from('audit_log').insert({
    user_id: userId, user_email: userEmail,
    project_id: projectId, project_name: projectName,
    macro_name: macroName, categoria, field,
    value_before: valueBefore !== undefined ? { v: valueBefore } : null,
    value_after:  valueAfter  !== undefined ? { v: valueAfter  } : null,
    scenario_id: scenarioId || null,
  })
  return { error }
}

export const getAuditLog = async (userId, limit = 100) => {
  const { data, error } = await supabase
    .from('audit_log').select('*').eq('user_id', userId)
    .order('created_at', { ascending: false }).limit(limit)
  return { data: data || [], error }
}

export const getAdminAuditLog = async (limit = 500) => {
  const { data, error } = await supabase
    .from('audit_log').select('*')
    .order('created_at', { ascending: false }).limit(limit)
  return { data: data || [], error }
}
