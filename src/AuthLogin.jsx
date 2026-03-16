import { useState } from "react"
import { signIn, signInMagicLink, signUp } from "./supabase"

const T = {
  red:"#E8182A", ink:"#1A1917", inkSoft:"#6B6A68",
  border:"#E5E3DF", surface:"#F7F6F3", card:"#FFFFFF"
}

export default function AuthLogin({ onAuth }) {
  const [mode, setMode]       = useState("login")   // login | magic | signup
  const [email, setEmail]     = useState("")
  const [password, setPassword] = useState("")
  const [name, setName]       = useState("")
  const [loading, setLoading] = useState(false)
  const [error, setError]     = useState("")
  const [sent, setSent]       = useState(false)

  const inp = {
    width:"100%", padding:"10px 14px", borderRadius:10,
    border:`1.5px solid ${T.border}`, background:T.card,
    fontSize:14, color:T.ink, outline:"none",
    fontFamily:"'Outfit',system-ui", boxSizing:"border-box",
  }

  const handleSubmit = async () => {
    setError(""); setLoading(true)
    try {
      if (mode === "magic") {
        const { error } = await signInMagicLink(email)
        if (error) throw error
        setSent(true)
      } else if (mode === "signup") {
        const { error } = await signUp(email, password, name)
        if (error) throw error
        setSent(true)
      } else {
        const { data, error } = await signIn(email, password)
        if (error) throw error
        onAuth(data.session)
      }
    } catch (e) {
      setError(e.message || "Error de autenticación")
    }
    setLoading(false)
  }

  return (
    <div style={{minHeight:"100vh",background:T.surface,display:"flex",
      alignItems:"center",justifyContent:"center",padding:24}}>
      <div style={{width:"100%",maxWidth:400}}>

        {/* Logo / título */}
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{display:"inline-flex",alignItems:"center",justifyContent:"center",
            width:48,height:48,borderRadius:14,background:T.red,marginBottom:16}}>
            <span style={{color:"#fff",fontWeight:900,fontSize:15,letterSpacing:"-.02em"}}>DVB</span>
          </div>
          <div style={{fontSize:22,fontWeight:800,color:T.ink,letterSpacing:"-.02em"}}>
            DVB Command Center
          </div>
          <div style={{fontSize:13,color:T.inkSoft,marginTop:4}}>
            Claro Colombia · CAPEX 2026 · Kearney
          </div>
        </div>

        {/* Card */}
        <div style={{background:T.card,borderRadius:18,padding:"28px 32px",
          border:`1px solid ${T.border}`,boxShadow:"0 4px 24px rgba(0,0,0,0.06)"}}>

          {sent ? (
            <div style={{textAlign:"center",padding:"16px 0"}}>
              <div style={{fontSize:36,marginBottom:12}}>📬</div>
              <div style={{fontSize:16,fontWeight:700,color:T.ink,marginBottom:8}}>
                {mode==="magic" ? "Revisa tu correo" : "Cuenta creada"}
              </div>
              <div style={{fontSize:13,color:T.inkSoft,lineHeight:1.6}}>
                {mode==="magic"
                  ? `Enviamos un link de acceso a ${email}`
                  : `Confirma tu correo en ${email} para activar tu cuenta`}
              </div>
              <button onClick={()=>{setSent(false);setMode("login")}}
                style={{marginTop:20,padding:"8px 20px",borderRadius:99,
                  border:`1.5px solid ${T.border}`,background:"transparent",
                  color:T.inkSoft,fontSize:13,cursor:"pointer"}}>
                Volver
              </button>
            </div>
          ) : (<>
            {/* Tabs */}
            <div style={{display:"flex",gap:4,marginBottom:24,
              background:T.surface,borderRadius:10,padding:4}}>
              {[["login","Contraseña"],["magic","Magic Link"]].map(([m,l])=>(
                <button key={m} onClick={()=>{setMode(m);setError("")}}
                  style={{flex:1,padding:"7px 0",borderRadius:8,border:"none",
                    cursor:"pointer",fontSize:13,fontWeight:600,
                    background:mode===m?T.card:"transparent",
                    color:mode===m?T.ink:T.inkSoft,
                    boxShadow:mode===m?"0 1px 4px rgba(0,0,0,0.08)":"none",
                    transition:"all .15s"}}>
                  {l}
                </button>
              ))}
            </div>

            <div style={{display:"flex",flexDirection:"column",gap:12}}>
              {mode==="signup" && (
                <input value={name} onChange={e=>setName(e.target.value)}
                  placeholder="Nombre completo" style={inp}/>
              )}
              <input value={email} onChange={e=>setEmail(e.target.value)}
                placeholder="Correo electrónico" type="email" style={inp}
                onKeyDown={e=>e.key==="Enter"&&handleSubmit()}/>
              {mode!=="magic" && (
                <input value={password} onChange={e=>setPassword(e.target.value)}
                  placeholder="Contraseña" type="password" style={inp}
                  onKeyDown={e=>e.key==="Enter"&&handleSubmit()}/>
              )}

              {error && (
                <div style={{fontSize:12,color:T.red,padding:"8px 12px",
                  background:"#FFF1F1",borderRadius:8,border:"1px solid #FDD8DA"}}>
                  {error}
                </div>
              )}

              <button onClick={handleSubmit} disabled={loading||!email}
                style={{padding:"11px 0",borderRadius:10,border:"none",
                  background:(!email||loading)?"#E5E3DF":T.red,
                  color:(!email||loading)?T.inkSoft:"#fff",
                  fontSize:14,fontWeight:700,cursor:(!email||loading)?"not-allowed":"pointer",
                  transition:"all .15s",marginTop:4}}>
                {loading ? "..." : mode==="magic" ? "Enviar Magic Link"
                  : mode==="signup" ? "Crear cuenta" : "Ingresar"}
              </button>
            </div>

            <div style={{textAlign:"center",marginTop:20,fontSize:12,color:T.inkSoft}}>
              {mode==="login" ? (<>
                ¿No tienes cuenta?{" "}
                <span onClick={()=>{setMode("signup");setError("")}}
                  style={{color:T.red,cursor:"pointer",fontWeight:600}}>
                  Regístrate
                </span>
              </>) : (<>
                ¿Ya tienes cuenta?{" "}
                <span onClick={()=>{setMode("login");setError("")}}
                  style={{color:T.red,cursor:"pointer",fontWeight:600}}>
                  Ingresar
                </span>
              </>)}
            </div>
          </>)}
        </div>

        <div style={{textAlign:"center",marginTop:20,fontSize:11,color:T.inkSoft}}>
          Kearney × Claro Colombia · Confidencial
        </div>
      </div>
    </div>
  )
}
