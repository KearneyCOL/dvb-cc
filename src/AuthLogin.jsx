import { useState, useEffect } from "react"
import { authSignIn, authSignUp } from "./supabase"

const C = {
  red:"#E8182A", ink:"#1A1917", soft:"#6B6A68",
  border:"#E5E3DF", surf:"#F7F6F3", card:"#FFFFFF"
}

const inp = {
  width:"100%", padding:"11px 14px", borderRadius:10,
  border:`1.5px solid ${C.border}`, background:C.card,
  fontSize:14, color:C.ink, outline:"none",
  fontFamily:"'Outfit',system-ui,sans-serif",
  boxSizing:"border-box", transition:"border-color .15s",
}

export default function AuthLogin({ onAuth }) {
  const [mode,     setMode]     = useState("login")   // login | signup
  const [email,    setEmail]    = useState("")
  const [password, setPassword] = useState("")
  const [name,     setName]     = useState("")
  const [loading,  setLoading]  = useState(false)
  const [error,    setError]    = useState("")

  // Cargar Outfit si no está cargada
  useEffect(() => {
    if (document.querySelector('link[href*="Outfit"]')) return
    const l = document.createElement("link")
    l.rel  = "stylesheet"
    l.href = "https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800;900&display=swap"
    document.head.appendChild(l)
  }, [])

  const submit = async () => {
    setError(""); setLoading(true)
    try {
      if (mode === "signup") {
        if (!name.trim()) throw new Error("Ingresa tu nombre completo")
        if (password.length < 6) throw new Error("La contraseña debe tener mínimo 6 caracteres")
        const { data, error } = await authSignUp(email, password, name.trim())
        if (error) throw error
        if (data?.session) {
          onAuth(data.session)
        } else {
          // Supabase creó el usuario pero sin sesión (confirmación activada)
          setError("⚠️  Revisa tu correo y confirma la cuenta, o desactiva 'Confirm email' en Supabase → Authentication → Settings")
        }
      } else {
        const { data, error } = await authSignIn(email, password)
        if (error) {
          if (error.message.includes("Invalid login")) throw new Error("Correo o contraseña incorrectos")
          if (error.message.includes("Email not confirmed")) throw new Error("Confirma tu correo antes de ingresar, o desactiva la confirmación en Supabase")
          throw error
        }
        onAuth(data.session)
      }
    } catch (e) {
      setError(e.message || "Error inesperado")
    }
    setLoading(false)
  }

  const onKey = (e) => { if (e.key === "Enter") submit() }

  return (
    <div style={{ minHeight:"100vh", background:C.surf, display:"flex",
      alignItems:"center", justifyContent:"center", padding:24,
      fontFamily:"'Outfit',system-ui,sans-serif" }}>
      <div style={{ width:"100%", maxWidth:400 }}>

        {/* Logo */}
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <div style={{ display:"inline-flex", alignItems:"center",
            justifyContent:"center", width:52, height:52,
            borderRadius:14, background:C.red, marginBottom:16 }}>
            <span style={{ color:"#fff", fontWeight:900, fontSize:15,
              letterSpacing:"-.02em" }}>DVB</span>
          </div>
          <div style={{ fontSize:22, fontWeight:800, color:C.ink,
            letterSpacing:"-.02em", marginBottom:4 }}>
            DVB Command Center
          </div>
          <div style={{ fontSize:13, color:C.soft }}>
            Claro Colombia · CAPEX 2026 · Kearney
          </div>
        </div>

        {/* Card */}
        <div style={{ background:C.card, borderRadius:18,
          padding:"28px 32px", border:`1px solid ${C.border}`,
          boxShadow:"0 4px 24px rgba(0,0,0,0.06)" }}>

          {/* Tabs */}
          <div style={{ display:"flex", gap:4, marginBottom:24,
            background:C.surf, borderRadius:10, padding:4 }}>
            {[["login","Ingresar"],["signup","Registrarme"]].map(([m,l]) => (
              <button key={m} onClick={() => { setMode(m); setError("") }}
                style={{ flex:1, padding:"8px 0", borderRadius:8, border:"none",
                  cursor:"pointer", fontSize:13, fontWeight:600,
                  fontFamily:"'Outfit',system-ui,sans-serif",
                  background: mode===m ? C.card : "transparent",
                  color:      mode===m ? C.ink  : C.soft,
                  boxShadow:  mode===m ? "0 1px 4px rgba(0,0,0,0.08)" : "none",
                  transition:"all .15s" }}>
                {l}
              </button>
            ))}
          </div>

          <div style={{ display:"flex", flexDirection:"column", gap:10 }}>
            {mode === "signup" && (
              <input value={name} onChange={e=>setName(e.target.value)}
                placeholder="Nombre completo" style={inp} onKeyDown={onKey}
                onFocus={e=>e.target.style.borderColor=C.red}
                onBlur={e=>e.target.style.borderColor=C.border}/>
            )}
            <input value={email} onChange={e=>setEmail(e.target.value)}
              placeholder="Correo electrónico" type="email"
              style={inp} onKeyDown={onKey}
              onFocus={e=>e.target.style.borderColor=C.red}
              onBlur={e=>e.target.style.borderColor=C.border}/>
            <input value={password} onChange={e=>setPassword(e.target.value)}
              placeholder={mode==="signup" ? "Contraseña (mín. 6 caracteres)" : "Contraseña"}
              type="password" style={inp} onKeyDown={onKey}
              onFocus={e=>e.target.style.borderColor=C.red}
              onBlur={e=>e.target.style.borderColor=C.border}/>

            {error && (
              <div style={{ fontSize:12, color:"#B91C1C", padding:"10px 12px",
                background:"#FEF2F2", borderRadius:8,
                border:"1px solid #FECACA", lineHeight:1.5 }}>
                {error}
              </div>
            )}

            <button onClick={submit}
              disabled={loading || !email || !password}
              style={{ padding:"12px 0", borderRadius:10, border:"none",
                background: (!email||!password||loading) ? "#E5E3DF" : C.red,
                color:      (!email||!password||loading) ? C.soft    : "#fff",
                fontSize:14, fontWeight:700, cursor: (!email||!password||loading) ? "not-allowed":"pointer",
                fontFamily:"'Outfit',system-ui,sans-serif",
                transition:"all .15s", marginTop:4 }}>
              {loading ? "..." : mode==="signup" ? "Crear cuenta" : "Ingresar"}
            </button>
          </div>

          {mode === "login" && (
            <div style={{ textAlign:"center", marginTop:18,
              fontSize:12, color:C.soft }}>
              ¿Aún no tienes cuenta?{" "}
              <span onClick={() => { setMode("signup"); setError("") }}
                style={{ color:C.red, cursor:"pointer", fontWeight:700 }}>
                Regístrate aquí
              </span>
            </div>
          )}
        </div>

        <div style={{ textAlign:"center", marginTop:20,
          fontSize:11, color:C.soft }}>
          Kearney × Claro Colombia · Confidencial
        </div>
      </div>
    </div>
  )
}
