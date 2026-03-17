import { useState, useMemo, useCallback, useEffect, useRef } from "react";
import "./index.css";
import * as XLSX from "xlsx";
import { supabase, getProfile, touchLastSeen, fetchScenarios,
         upsertScenario, removeScenario, pushLog, fetchLog,
         authSignOut } from "./supabase";
import AuthLogin from "./AuthLogin";

/* ══════════════════════════════════════════════════════════════════════════
   DVB COMMAND CENTER · Kearney × Claro Colombia · CAPEX 2026
   Design: Outfit · #F7F6F3 bg · Kearney Red #E8182A · cards blancas
   
   ESTRUCTURA JERÁRQUICA:
   N1 Categoría → N2 Dominio → N3 Programa → N4 Proyecto → N5 Componentes
══════════════════════════════════════════════════════════════════════════ */

const GS = `
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800;900&display=swap');
*,*::before,*::after{box-sizing:border-box;font-family:'Outfit',system-ui,sans-serif;}
body{margin:0;background:#F7F6F3;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:transparent;}
::-webkit-scrollbar-thumb{background:#D4D2CC;border-radius:99px;}

@keyframes fadeUp{from{opacity:0;transform:translateY(18px)}to{opacity:1;transform:translateY(0)}}
@keyframes fadeIn{from{opacity:0}to{opacity:1}}
@keyframes scaleIn{from{opacity:0;transform:scale(.96)}to{opacity:1;transform:scale(1)}}
@keyframes slideRight{from{opacity:0;transform:translateX(32px)}to{opacity:1;transform:translateX(0)}}
@keyframes ticker{from{transform:translateX(0)}to{transform:translateX(-50%)}}
@keyframes pulseDot{0%,100%{box-shadow:0 0 0 0 rgba(45,200,100,0)}50%{box-shadow:0 0 0 5px rgba(45,200,100,.12)}}
@keyframes barGrow{from{transform:scaleX(0);transform-origin:left}to{transform:scaleX(1)}}

.fu {animation:fadeUp .5s cubic-bezier(.22,1,.36,1) both;}
.fu1{animation:fadeUp .5s .08s cubic-bezier(.22,1,.36,1) both;}
.fu2{animation:fadeUp .5s .16s cubic-bezier(.22,1,.36,1) both;}
.fu3{animation:fadeUp .5s .24s cubic-bezier(.22,1,.36,1) both;}
.fi {animation:fadeIn .22s ease both;}
.si {animation:scaleIn .28s cubic-bezier(.22,1,.36,1) both;}
.sr {animation:slideRight .36s cubic-bezier(.22,1,.36,1) both;}

.hover-lift{transition:transform .2s cubic-bezier(.22,1,.36,1),box-shadow .2s ease;}
.hover-lift:hover{transform:translateY(-3px);box-shadow:0 16px 40px rgba(0,0,0,0.1)!important;}
.btn-primary{transition:all .18s cubic-bezier(.22,1,.36,1);}
.btn-primary:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(232,24,42,0.34)!important;}
.btn-ghost{transition:all .15s;}
.btn-ghost:hover{background:rgba(255,255,255,0.16)!important;}
.tab-pill{transition:all .2s cubic-bezier(.4,0,.2,1);}
.row-hover{transition:background .1s;}
.row-hover:hover{background:#FFF8F7!important;}
.tree-row{transition:background .1s;cursor:pointer;}
.tree-row:hover{background:#FFF8F7!important;}
.macro-card{transition:transform .18s cubic-bezier(.22,1,.36,1),box-shadow .18s ease,border-color .14s;}
.macro-card:hover{transform:translateY(-2px);box-shadow:0 12px 32px rgba(0,0,0,0.09)!important;}
.comp-hd{cursor:pointer;transition:background .12s;}
.comp-hd:hover{background:#FAFAF8!important;}

.tick-wrap{overflow:hidden;white-space:nowrap;}
.tick-inner{display:inline-block;animation:ticker 90s linear infinite;}
.ldot{width:7px;height:7px;border-radius:50%;background:#2DC878;display:inline-block;animation:pulseDot 2s ease infinite;}
`;

/* ── Tokens de diseño ─────────────────────────────────────────────────────── */
const T = {
  red:"#E8182A", redDk:"#B5111F", redBg:"#FFF1F1", redSoft:"#FDD8DA", redXsoft:"#FFF8F8",
  ink:"#111110", inkMid:"#6B6860", inkSoft:"#9C9A95", inkXsoft:"#C8C6C0",
  surface:"#F7F6F3", card:"#FFFFFF", border:"#E8E6E0", borderSm:"#F0EEE9",
  gold:"#C8941C", goldBg:"#FFFBEB", goldBdr:"#FDE68A",
  blue:"#2563EB", blueBg:"#EFF6FF", blueBdr:"#BFDBFE",
  green:"#059669", greenBg:"#ECFDF5", greenBdr:"#A7F3D0",
  violet:"#7C3AED", violetBg:"#F5F3FF", violetBdr:"#DDD6FE",
  orange:"#D97706", orangeBg:"#FFFBEB", orangeBdr:"#FCD34D",
  cyan:"#0891B2", cyanBg:"#ECFEFF", cyanBdr:"#A5F3FC",
  gray:"#6B7280", grayBg:"#F9FAFB", grayBdr:"#E5E7EB",
  rose:"#BE123C", roseBg:"#FFE4E6", roseBdr:"#FECDD3",
};

/* ── Configuración de Categorías ─────────────────────────────────────────── */
const CATEGORIA_CFG = {
  "Red":                 {c:"#2563EB", bg:"#EFF6FF", bdr:"#BFDBFE", icon:"📡"},
  "Enterprise y B2B":    {c:"#7C3AED", bg:"#F5F3FF", bdr:"#DDD6FE", icon:"🏢"},
  "Tecnología":          {c:"#0891B2", bg:"#ECFEFF", bdr:"#A5F3FC", icon:"💻"},
  "Inversión Clientes":  {c:"#E8182A", bg:"#FFF1F1", bdr:"#FDD8DA", icon:"👥"},
  "Corporativo":         {c:"#6B7280", bg:"#F9FAFB", bdr:"#E5E7EB", icon:"🏛️"},
};

/* ── Configuración de Tipos de Inversión ─────────────────────────────────── */
const TIPO_CFG = {
  "Crecimiento":     {c:"#047857", bg:"#D1FAE5", bdr:"#A7F3D0", label:"Crecimiento"},
  "Modernización":   {c:"#B45309", bg:"#FEF3C7", bdr:"#FCD34D", label:"Modernización"},
  "Operación":       {c:"#6B7280", bg:"#F3F4F6", bdr:"#E5E7EB", label:"Operación"},
  "Regulatorio":     {c:"#DC2626", bg:"#FEE2E2", bdr:"#FECACA", label:"Regulatorio"},
  "Transformación":  {c:"#6D28D9", bg:"#EDE9FE", bdr:"#DDD6FE", label:"Transformación"},
  "Evolución":       {c:"#2563EB", bg:"#DBEAFE", bdr:"#BFDBFE", label:"Evolución"},
  "Comercial":       {c:"#BE123C", bg:"#FFE4E6", bdr:"#FECDD3", label:"Comercial"},
};

/* ── Configuración de Dominios ─────────────────────────────────────────────── */
const DOMINIO_CFG = {
  "Acceso Fijo":           {c:"#2563EB", icon:"🔌"},
  "Acceso Móvil":          {c:"#059669", icon:"📶"},
  "Transporte y Core":     {c:"#7C3AED", icon:"🌐"},
  "Data Center":           {c:"#0891B2", icon:"🖥️"},
  "Soluciones Enterprise": {c:"#C8941C", icon:"🏢"},
  "Servicios DC B2B":      {c:"#D97706", icon:"☁️"},
  "Gobierno y Verticales": {c:"#059669", icon:"🏛️"},
  "Wholesale":             {c:"#6B7280", icon:"📊"},
  "BSS":                   {c:"#7C3AED", icon:"💼"},
  "OSS":                   {c:"#2563EB", icon:"⚙️"},
  "Digital":               {c:"#E8182A", icon:"📱"},
  "Infraestructura IT":    {c:"#0891B2", icon:"🖥️"},
  "Evolución Plataformas": {c:"#059669", icon:"🔄"},
  "Consumidor":            {c:"#E8182A", icon:"👤"},
  "Enterprise":            {c:"#7C3AED", icon:"🏢"},
  "Soporte":               {c:"#6B7280", icon:"🔧"},
};

/* ══════════════════════════════════════════════════════════════════════════
   DATA — ESTRUCTURA COMPLETA CAPEX 2026
   Categoría → Dominio → Tipo → Programa → Proyecto
══════════════════════════════════════════════════════════════════════════ */

const DATA = [
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: RED
  // ═══════════════════════════════════════════════════════════════════════════
  
  // ─── DOMINIO: ACCESO FIJO ───────────────────────────────────────────────────
  
  // Programa: Expansión Cobertura
  {categoria:"Red", dominio:"Acceso Fijo", tipo:"Crecimiento",
   programa:"Expansión Cobertura", P_base:45000000, proyectos:[
    {id:"AF.EC.01", n:"FTTH Greenfield", m:"HHP", prio:"CRECIMIENTO", P_base:25000000,
     dt:"expansion_geo", pt:{objetivo:200000, actual:120000}, df:"benchmark_amx", pf:{pa:160, pb:160},
     componentes:"Backbone F1, OLT, Red distribución F2/F3, Acometida",
     driver_q:"HHP objetivo menos HHP actual", driver_p:"Benchmark LatAm",
     unit_econ:"USD 160/HHP", racional:"Market share, ARPU fibra"},
    {id:"AF.EC.02", n:"FTTH Infill", m:"HHP", prio:"CRECIMIENTO", P_base:12000000,
     dt:"expansion_geo", pt:{objetivo:150000, actual:100000}, df:"benchmark_amx", pf:{pa:80, pb:80},
     componentes:"Extensiones F2/F3, Splitters, Drop a edificio",
     driver_q:"Edificios sin cobertura en huella", driver_p:"Benchmark LatAm",
     unit_econ:"USD 80/HHP", racional:"Penetración, costo servicio"},
    {id:"AF.EC.03", n:"FTTH Red Neutra", m:"HHP", prio:"CRECIMIENTO", P_base:8000000,
     dt:"expansion_geo", pt:{objetivo:200000, actual:100000}, df:"mercado", pf:{pa:40, pb:40},
     componentes:"IRU o arriendo infraestructura terceros, Equipos activos propios",
     driver_q:"HHP objetivo vía wholesale", driver_p:"Contratos IRU",
     unit_econ:"USD 40/HHP", racional:"Asset-light, time-to-market"},
  ]},
  
  // Programa: Expansión Capacidad
  {categoria:"Red", dominio:"Acceso Fijo", tipo:"Crecimiento",
   programa:"Expansión Capacidad", P_base:8500000, proyectos:[
    {id:"AF.EK.01", n:"Puertos OLT y Agregación", m:"puerto", prio:"CRECIMIENTO", P_base:5000000,
     dt:"activaciones", pt:{brutas:200000, churn:40000}, df:"benchmark_amx", pf:{pa:25, pb:25},
     componentes:"Chasis OLT, Tarjetas PON, Uplinks, NMS",
     driver_q:"Altas netas por ratio split", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 25/puerto", racional:"Capacidad ahead of demand"},
    {id:"AF.EK.02", n:"CPE Nuevos Clientes", m:"cliente", prio:"CRECIMIENTO", P_base:3500000,
     dt:"activaciones", pt:{brutas:100000, churn:20000}, df:"benchmark_amx", pf:{pa:45, pb:45},
     componentes:"ONT GPON/XGS-PON, Router WiFi 6, STB",
     driver_q:"Altas brutas fijo", driver_p:"Contratos volumen",
     unit_econ:"USD 45/cliente", racional:"Habilitación servicio"},
  ]},
  
  // Programa: Obligaciones Regulatorias (Fijo)
  {categoria:"Red", dominio:"Acceso Fijo", tipo:"Regulatorio",
   programa:"Obligaciones Regulatorias", P_base:5000000, proyectos:[
    {id:"AF.OR.01", n:"Servicio Universal Fijo", m:"sitio", prio:"REGULATORIO", P_base:5000000,
     dt:"cobertura_crc", pt:{comprometido:200, ejecutado:50}, df:"benchmark_crc", pf:{pa:33333, pb:33333},
     componentes:"Despliegue zonas rurales, Escuelas, Centros comunitarios",
     driver_q:"Compromisos CRC vigentes", driver_p:"Tarifas reguladas",
     unit_econ:"Por obligación", racional:"Cumplimiento licencia"},
  ]},
  
  // Programa: Modernización Tecnológica (Fijo)
  {categoria:"Red", dominio:"Acceso Fijo", tipo:"Modernización",
   programa:"Modernización Tecnológica", P_base:18000000, proyectos:[
    {id:"AF.MT.01", n:"Migración HFC a Fibra", m:"suscriptor", prio:"EBITDA", P_base:12000000,
     dt:"migraciones", pt:{base_migrable:200000, tasa:0.5}, df:"benchmark_amx", pf:{pa:120, pb:120},
     componentes:"Retiro coaxial, Nueva acometida FTTH, Swap CPE",
     driver_q:"Suscriptores HFC a migrar", driver_p:"Costo por migración",
     unit_econ:"USD 120/suscriptor", racional:"OPEX, churn, velocidad"},
    {id:"AF.MT.02", n:"Upgrade GPON a XGS-PON", m:"OLT", prio:"EBITDA", P_base:4000000,
     dt:"obsolescencia", pt:{total_activos:500, pct_eol:0.3, meta_pct:0.1}, df:"benchmark_amx", pf:{pa:8000, pb:8000},
     componentes:"Tarjetas OLT 10G, Actualización NMS, ONTs compatibles",
     driver_q:"OLTs en capacidad o EoL", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 8K/OLT", racional:"10Gbps simétrico, enterprise"},
    {id:"AF.MT.03", n:"Refresh CPE y ONT", m:"CPE", prio:"PMO", P_base:2000000,
     dt:"obsolescencia", pt:{total_activos:200000, pct_eol:0.15, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:35, pb:35},
     componentes:"ONT nueva generación, Router WiFi 6/7",
     driver_q:"Porcentaje equipos EoL por base instalada", driver_p:"Contratos volumen",
     unit_econ:"USD 35/CPE", racional:"Costo soporte, WiFi upgrade"},
  ]},
  
  // Programa: Operación (Fijo)
  {categoria:"Red", dominio:"Acceso Fijo", tipo:"Operación",
   programa:"Operación", P_base:8000000, proyectos:[
    {id:"AF.OP.01", n:"Mantenimiento Planta Externa", m:"km", prio:"PMO", P_base:5000000,
     dt:"mantenimiento", pt:{activos:6250, frecuencia:1}, df:"historico", pf:{pa:800, pb:800},
     componentes:"Reparación fibra, Empalmes, Obra civil menor",
     driver_q:"Km red por tasa fallas histórica", driver_p:"Histórico mas IPC",
     unit_econ:"USD 800/km/año", racional:"SLA, NPS"},
    {id:"AF.OP.02", n:"Inventario y Reposición", m:"activo", prio:"PMO", P_base:3000000,
     dt:"mantenimiento", pt:{activos:100000, frecuencia:1}, df:"historico", pf:{pa:30, pb:30},
     componentes:"Stock repuestos red, CPE reposición por fallas",
     driver_q:"Porcentaje base instalada", driver_p:"Histórico mas IPC",
     unit_econ:"2-3% base activos", racional:"MTTR, continuidad"},
  ]},
  
  // ─── DOMINIO: ACCESO MÓVIL ──────────────────────────────────────────────────
  
  // Programa: Expansión Cobertura (Móvil)
  {categoria:"Red", dominio:"Acceso Móvil", tipo:"Crecimiento",
   programa:"Expansión Cobertura", P_base:52000000, proyectos:[
    {id:"AM.EC.01", n:"Nuevos Sitios 4G", m:"sitio", prio:"CRECIMIENTO", P_base:19500000,
     dt:"expansion_geo", pt:{objetivo:500, actual:200}, df:"benchmark_amx", pf:{pa:65000, pb:65000},
     componentes:"Torre o Mástil, Radios LTE, Antenas, Energía, TX, Obra civil",
     driver_q:"Gap cobertura poblacional 4G", driver_p:"Benchmark LatAm",
     unit_econ:"USD 65K/sitio", racional:"Cobertura 4G, zonas sin servicio"},
    {id:"AM.EC.02", n:"Nuevos Sitios 5G", m:"sitio", prio:"CRECIMIENTO", P_base:28500000,
     dt:"expansion_geo", pt:{objetivo:400, actual:100}, df:"benchmark_amx", pf:{pa:95000, pb:95000},
     componentes:"Radios 5G NR, Antenas MIMO masivo, Energía reforzada, TX fibra",
     driver_q:"Plan cobertura 5G greenfield", driver_p:"Benchmark LatAm",
     unit_econ:"USD 95K/sitio", racional:"Liderazgo 5G, nuevas zonas"},
    {id:"AM.EC.03", n:"Small Cells y DAS", m:"nodo", prio:"CRECIMIENTO", P_base:4000000,
     dt:"expansion_geo", pt:{objetivo:200, actual:40}, df:"benchmark_amx", pf:{pa:25000, pb:25000},
     componentes:"Small cells indoor/outdoor, DAS activo/pasivo, Backhaul",
     driver_q:"Hotspots indoor y densificación urbana", driver_p:"Benchmark LatAm",
     unit_econ:"USD 25K/nodo", racional:"Indoor, enterprise, venues"},
  ]},
  
  // Programa: Expansión Capacidad (Móvil)
  {categoria:"Red", dominio:"Acceso Móvil", tipo:"Crecimiento",
   programa:"Expansión Capacidad", P_base:21000000, proyectos:[
    {id:"AM.EK.01", n:"Carriers Adicionales 4G", m:"carrier", prio:"CRECIMIENTO", P_base:12000000,
     dt:"congestion", pt:{total_activos:5000, pct_congestion:0.20, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:12000, pb:12000},
     componentes:"Módulos radio LTE, Antenas multi-banda, Reconfiguración",
     driver_q:"Sitios 4G sobre umbral congestión", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 12K/carrier", racional:"QoE, descongestión 4G"},
    {id:"AM.EK.02", n:"Carriers Adicionales 5G", m:"carrier", prio:"CRECIMIENTO", P_base:9000000,
     dt:"congestion", pt:{total_activos:1000, pct_congestion:0.25, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:18000, pb:18000},
     componentes:"Módulos radio 5G NR, MIMO adicional, SW features",
     driver_q:"Sitios 5G sobre umbral congestión", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 18K/carrier", racional:"Capacidad 5G, nuevos servicios"},
  ]},
  
  // Programa: Obligaciones Regulatorias (Móvil)
  {categoria:"Red", dominio:"Acceso Móvil", tipo:"Regulatorio",
   programa:"Obligaciones Regulatorias", P_base:58000000, proyectos:[
    {id:"AM.OR.01", n:"Cobertura Mandatoria 4G", m:"sitio", prio:"REGULATORIO", P_base:40000000,
     dt:"cobertura_crc", pt:{comprometido:100, ejecutado:20}, df:"benchmark_crc", pf:{pa:500000, pb:500000},
     componentes:"Sitios rurales, Cobertura vías nacionales, Municipios",
     driver_q:"Compromisos licencia AWS, 700MHz y otros", driver_p:"Fórmula CRC",
     unit_econ:"Por obligación", racional:"Cumplimiento espectro 4G"},
    {id:"AM.OR.02", n:"Cobertura Mandatoria 5G", m:"sitio", prio:"REGULATORIO", P_base:18000000,
     dt:"cobertura_crc", pt:{comprometido:50, ejecutado:10}, df:"benchmark_crc", pf:{pa:450000, pb:450000},
     componentes:"Sitios 5G ciudades principales, Zonas prioritarias",
     driver_q:"Compromisos licencia 3.5GHz y mmWave", driver_p:"Fórmula CRC",
     unit_econ:"Por obligación", racional:"Cumplimiento espectro 5G"},
  ]},
  
  // Programa: Modernización Tecnológica (Móvil)
  {categoria:"Red", dominio:"Acceso Móvil", tipo:"Modernización",
   programa:"Modernización Tecnológica", P_base:85000000, proyectos:[
    {id:"AM.MT.01", n:"Upgrade 4G a 5G", m:"sitio", prio:"EBITDA", P_base:45000000,
     dt:"migraciones", pt:{base_migrable:2000, tasa:0.5}, df:"benchmark_amx", pf:{pa:45000, pb:45000},
     componentes:"Radios 5G en sitio existente, Antenas MIMO, SW, Refuerzo energía",
     driver_q:"Sitios 4G elegibles para 5G", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 45K/sitio", racional:"Reusar infraestructura existente"},
    {id:"AM.MT.02", n:"Refresh Equipos 4G EoL", m:"sitio", prio:"PMO", P_base:15000000,
     dt:"obsolescencia", pt:{total_activos:3000, pct_eol:0.20, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:25000, pb:25000},
     componentes:"Reemplazo radios 4G sin soporte, Actualización SW",
     driver_q:"Porcentaje equipos 4G End-of-Life", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 25K/sitio", racional:"Seguridad, evitar OPEX soporte"},
    {id:"AM.MT.03", n:"Refresh Equipos 5G EoL", m:"sitio", prio:"PMO", P_base:7000000,
     dt:"obsolescencia", pt:{total_activos:500, pct_eol:0.20, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:35000, pb:35000},
     componentes:"Reemplazo radios 5G sin soporte, Actualización SW",
     driver_q:"Porcentaje equipos 5G End-of-Life", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 35K/sitio", racional:"Seguridad, soporte vendor"},
    {id:"AM.MT.04", n:"Swap de Vendor RAN", m:"sitio", prio:"EBITDA", P_base:12000000,
     dt:"obsolescencia", pt:{total_activos:500, pct_eol:0.40, meta_pct:0.05}, df:"mercado", pf:{pa:60000, pb:60000},
     componentes:"Equipos nuevo vendor, Desmonte, Integración, Testing",
     driver_q:"Sitios según estrategia multi-vendor", driver_p:"Contrato nuevo vendor",
     unit_econ:"USD 60K/sitio", racional:"De-riesgo supply chain, TCO"},
    {id:"AM.MT.05", n:"Reubicación de Sitios", m:"sitio", prio:"PMO", P_base:6000000,
     dt:"mantenimiento", pt:{activos:120, frecuencia:1}, df:"mercado", pf:{pa:50000, pb:50000},
     componentes:"Construcción nuevo sitio, Desmonte anterior, TX, Permisos",
     driver_q:"Arriendos terminados o en riesgo", driver_p:"Estimación por sitio",
     unit_econ:"USD 50K/sitio", racional:"Continuidad cobertura"},
  ]},
  
  // Programa: Operación (Móvil)
  {categoria:"Red", dominio:"Acceso Móvil", tipo:"Operación",
   programa:"Operación", P_base:35000000, proyectos:[
    {id:"AM.OP.01", n:"Mantenimiento Sitios RAN", m:"sitio", prio:"PMO", P_base:21000000,
     dt:"mantenimiento", pt:{activos:7000, frecuencia:1}, df:"historico", pf:{pa:3000, pb:3000},
     componentes:"Preventivo, Correctivo, Inspecciones periódicas",
     driver_q:"Total sitios por ciclos mantenimiento", driver_p:"Histórico mas IPC",
     unit_econ:"USD 3K/sitio/año", racional:"Uptime, SLA red"},
    {id:"AM.OP.02", n:"Energía y Climatización", m:"sitio", prio:"PMO", P_base:8000000,
     dt:"mantenimiento", pt:{activos:4000, frecuencia:1}, df:"historico", pf:{pa:2000, pb:2000},
     componentes:"Rectificadores, Bancos baterías, HVAC, Generadores backup",
     driver_q:"Sitios por perfil antigüedad sistemas", driver_p:"Histórico mas IPC",
     unit_econ:"USD 2K/sitio/año", racional:"Disponibilidad, eficiencia"},
    {id:"AM.OP.03", n:"Torres e Infraestructura Física", m:"sitio", prio:"PMO", P_base:6000000,
     dt:"mantenimiento", pt:{activos:3000, frecuencia:1}, df:"historico", pf:{pa:2000, pb:2000},
     componentes:"Mantenimiento estructural torres, Repuestos, Seguridad",
     driver_q:"Sitios por ciclos inspección", driver_p:"Histórico mas IPC",
     unit_econ:"USD 2K/sitio/año", racional:"Seguridad, normativa"},
  ]},
  
  // ─── DOMINIO: TRANSPORTE Y CORE ─────────────────────────────────────────────
  
  // Programa: Expansión Backbone
  {categoria:"Red", dominio:"Transporte y Core", tipo:"Crecimiento",
   programa:"Expansión Backbone", P_base:32000000, proyectos:[
    {id:"TC.EB.01", n:"Fibra Nacional Long-haul", m:"km", prio:"CRECIMIENTO", P_base:15000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:1000, crecimiento_pct:0.30}, df:"benchmark_amx", pf:{pa:25000, pb:25000},
     componentes:"Fibra óptica, Equipos DWDM, Amplificadores, Regeneradores",
     driver_q:"Rutas nuevas mas upgrade capacidad", driver_p:"Benchmark LatAm",
     unit_econ:"USD 25K/km", racional:"Latencia, redundancia, capacidad"},
    {id:"TC.EB.02", n:"Anillos Metropolitanos", m:"anillo", prio:"CRECIMIENTO", P_base:12000000,
     dt:"expansion_geo", pt:{objetivo:10, actual:4}, df:"benchmark_amx", pf:{pa:2000000, pb:2000000},
     componentes:"Fibra metro, Nodos agregación, DWDM/ROADM metro",
     driver_q:"Ciudades sin anillo redundante", driver_p:"Benchmark LatAm",
     unit_econ:"USD 2M/anillo", racional:"Resiliencia, enterprise-ready"},
    {id:"TC.EB.03", n:"Capacidad DWDM Lambdas", m:"lambda", prio:"CRECIMIENTO", P_base:5000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:200, crecimiento_pct:0.50}, df:"benchmark_amx", pf:{pa:50000, pb:50000},
     componentes:"Transponders, Muxponders, Amplificadores",
     driver_q:"Rutas sobre umbral utilización", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 50K/lambda", racional:"Headroom tráfico"},
  ]},
  
  // Programa: Infraestructura 5G
  {categoria:"Red", dominio:"Transporte y Core", tipo:"Crecimiento",
   programa:"Infraestructura 5G", P_base:35000000, proyectos:[
    {id:"TC.I5.01", n:"Transporte 5G Xhaul", m:"sitio", prio:"CRECIMIENTO", P_base:8000000,
     dt:"expansion_geo", pt:{objetivo:800, actual:200}, df:"benchmark_amx", pf:{pa:12000, pb:12000},
     componentes:"Fronthaul fibra/eCPRI, Midhaul, Equipos timing/sync PTP",
     driver_q:"Sitios 5G nuevos mas upgrades", driver_p:"Benchmark LatAm",
     unit_econ:"USD 12K/sitio", racional:"Habilitación arquitectura 5G"},
    {id:"TC.I5.02", n:"Core 5G Standalone", m:"programa", prio:"CRECIMIENTO", P_base:15000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:1, crecimiento_pct:1}, df:"benchmark_amx", pf:{pa:15000000, pb:15000000},
     componentes:"Funciones cloud-native AMF, SMF, UPF, Service mesh",
     driver_q:"Roadmap 5G SA", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 15M programa", racional:"Network slicing, 5G enterprise"},
    {id:"TC.I5.03", n:"Telco Cloud y vRAN", m:"nodo", prio:"CRECIMIENTO", P_base:12000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:10, crecimiento_pct:0.50}, df:"benchmark_amx", pf:{pa:500000, pb:500000},
     componentes:"Servidores COTS, Software NFV/CNF, Orquestador MANO",
     driver_q:"Roadmap virtualización red", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 500K/nodo", racional:"Agilidad, eficiencia OPEX"},
  ]},
  
  // Programa: Conectividad Internacional
  {categoria:"Red", dominio:"Transporte y Core", tipo:"Crecimiento",
   programa:"Conectividad Internacional", P_base:15000000, proyectos:[
    {id:"TC.CI.01", n:"Capacidad IRU Cables Submarinos", m:"Gbps", prio:"CRECIMIENTO", P_base:9000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:500, crecimiento_pct:0.40}, df:"mercado", pf:{pa:45000, pb:45000},
     componentes:"Derechos uso cables existentes, Cross-connects en NAPs",
     driver_q:"Crecimiento tráfico internacional", driver_p:"Mercado IRU/año",
     unit_econ:"Variable USD/Gbps", racional:"Latencia contenido, redundancia"},
    {id:"TC.CI.02", n:"Participación Cables Submarinos", m:"proyecto", prio:"CRECIMIENTO", P_base:6000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:1, crecimiento_pct:1}, df:"mercado", pf:{pa:6000000, pb:6000000},
     componentes:"Equity en consorcio, Landing station, Backhaul terrestre",
     driver_q:"Rutas estratégicas nuevos cables", driver_p:"Términos consorcio",
     unit_econ:"Por proyecto", racional:"Capacidad asegurada largo plazo"},
  ]},
  
  // Programa: Modernización Core y TX
  {categoria:"Red", dominio:"Transporte y Core", tipo:"Modernización",
   programa:"Modernización Core y TX", P_base:8000000, proyectos:[
    {id:"TC.MC.01", n:"Refresh Routers Core IP/MPLS", m:"nodo", prio:"EBITDA", P_base:8000000,
     dt:"obsolescencia", pt:{total_activos:40, pct_eol:0.40, meta_pct:0.05}, df:"benchmark_amx", pf:{pa:500000, pb:500000},
     componentes:"Routers backbone, Líneas 400G, Software",
     driver_q:"Equipos EoL mas necesidades capacidad", driver_p:"Acuerdo marco vendor",
     unit_econ:"USD 500K/nodo", racional:"400G readiness, seguridad"},
  ]},
  
  // Programa: Operación (TX/Core)
  {categoria:"Red", dominio:"Transporte y Core", tipo:"Operación",
   programa:"Operación", P_base:12000000, proyectos:[
    {id:"TC.OP.01", n:"IRUs y Compartición Regulatoria", m:"contrato", prio:"PMO", P_base:7000000,
     dt:"mantenimiento", pt:{activos:100, frecuencia:1}, df:"historico", pf:{pa:70000, pb:70000},
     componentes:"Pagos Internexa, Compartición mandatoria infraestructura",
     driver_q:"Capacidad contratada vigente", driver_p:"Contratos existentes",
     unit_econ:"Por contrato", racional:"Continuidad red"},
    {id:"TC.OP.02", n:"Mantenimiento TX y Core", m:"equipo", prio:"PMO", P_base:5000000,
     dt:"mantenimiento", pt:{activos:500, frecuencia:1}, df:"historico", pf:{pa:10000, pb:10000},
     componentes:"Mantenimiento equipos DWDM, Routers, Transmisión óptica",
     driver_q:"Inventario equipos activos", driver_p:"Histórico mas IPC",
     unit_econ:"5% base activos/año", racional:"Uptime, SLA"},
  ]},
  
  // ─── DOMINIO: DATA CENTER ───────────────────────────────────────────────────
  
  // Programa: Expansión Capacidad DC
  {categoria:"Red", dominio:"Data Center", tipo:"Crecimiento",
   programa:"Expansión Capacidad DC", P_base:25000000, proyectos:[
    {id:"DC.EC.01", n:"Whitespace y Racks", m:"kW", prio:"CRECIMIENTO", P_base:10000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:500, crecimiento_pct:0.30}, df:"mercado", pf:{pa:15000, pb:15000},
     componentes:"Nuevo espacio físico, Racks, PDUs, Cableado estructurado",
     driver_q:"Proyección utilización mas pipeline", driver_p:"Benchmark mercado DC",
     unit_econ:"USD 15K/kW IT", racional:"Revenue cloud/enterprise, workloads internos"},
    {id:"DC.EC.02", n:"Energía y Refrigeración", m:"MW", prio:"CRECIMIENTO", P_base:12000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:2, crecimiento_pct:0.50}, df:"mercado", pf:{pa:8000000, pb:8000000},
     componentes:"UPS, Transformadores, Generadores, Chillers, CRAH/CRAC",
     driver_q:"MW adicionales requeridos", driver_p:"Benchmark mercado DC",
     unit_econ:"USD 8M/MW", racional:"PUE, capacidad crítica"},
    {id:"DC.EC.03", n:"Edge Computing Propio", m:"nodo", prio:"CRECIMIENTO", P_base:3000000,
     dt:"expansion_geo", pt:{objetivo:20, actual:5}, df:"benchmark_amx", pf:{pa:200000, pb:200000},
     componentes:"Micro DCs, Servidores edge, Conectividad baja latencia",
     driver_q:"Casos uso edge propios", driver_p:"Benchmark edge",
     unit_econ:"USD 200K/nodo", racional:"5G MEC, IoT, gaming"},
  ]},
  
  // Programa: Operación DC
  {categoria:"Red", dominio:"Data Center", tipo:"Operación",
   programa:"Operación", P_base:4500000, proyectos:[
    {id:"DC.OP.01", n:"Mantenimiento y Refresh Infra DC", m:"m2", prio:"PMO", P_base:4500000,
     dt:"mantenimiento", pt:{activos:30000, frecuencia:1}, df:"historico", pf:{pa:150, pb:150},
     componentes:"Mantenimiento preventivo/correctivo, Refresh UPS/HVAC/Generadores EoL",
     driver_q:"m2 bajo gestión mas equipos EoL", driver_p:"Histórico mas IPC",
     unit_econ:"USD 150/m2/año", racional:"Uptime, certificación Tier"},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: ENTERPRISE Y B2B
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Programa: Soluciones Conectividad
  {categoria:"Enterprise y B2B", dominio:"Soluciones Enterprise", tipo:"Crecimiento",
   programa:"Soluciones Conectividad", P_base:18000000, proyectos:[
    {id:"EB.SC.01", n:"Redes Dedicadas y SD-WAN", m:"contrato", prio:"CRECIMIENTO", P_base:10000000,
     dt:"pipeline_b2b", pt:{pipeline:200, win_rate:0.50}, df:"mercado", pf:{pa:100000, pb:100000},
     componentes:"Fibra dedicada, MPLS, SD-WAN, Equipos CPE enterprise",
     driver_q:"Pipeline comercial cerrado", driver_p:"Cotización por proyecto",
     unit_econ:"Por contrato", racional:"Revenue B2B, contratos multi-año"},
    {id:"EB.SC.02", n:"Conectividad Multi-Cloud", m:"POP", prio:"CRECIMIENTO", P_base:5000000,
     dt:"pipeline_b2b", pt:{pipeline:100, win_rate:0.50}, df:"mercado", pf:{pa:50000, pb:50000},
     componentes:"Direct Connect AWS/Azure/OCI/GCP, Cloud on-ramps, Cross-connects",
     driver_q:"Pipeline hybrid cloud", driver_p:"Costo por POP",
     unit_econ:"USD 50K/POP", racional:"Habilitación multi-cloud"},
    {id:"EB.SC.03", n:"5G Privado y Network Slicing", m:"despliegue", prio:"CRECIMIENTO", P_base:3000000,
     dt:"pipeline_b2b", pt:{pipeline:30, win_rate:0.40}, df:"mercado", pf:{pa:250000, pb:250000},
     componentes:"RAN dedicado/compartido, Core slice, Integración IT cliente",
     driver_q:"Pipeline 5G enterprise", driver_p:"Cotización por solución",
     unit_econ:"Por despliegue", racional:"Industria 4.0, verticales"},
  ]},
  
  // Programa: Servicios Data Center
  {categoria:"Enterprise y B2B", dominio:"Servicios DC B2B", tipo:"Crecimiento",
   programa:"Servicios Data Center", P_base:15000000, proyectos:[
    {id:"EB.SD.01", n:"Colocation y Housing", m:"rack", prio:"CRECIMIENTO", P_base:6000000,
     dt:"pipeline_b2b", pt:{pipeline:750, win_rate:0.50}, df:"mercado", pf:{pa:8000, pb:8000},
     componentes:"Build-out racks/jaulas/salas, Conectividad, Energía dedicada",
     driver_q:"Pipeline comercial DC", driver_p:"Costo build-out por rack",
     unit_econ:"USD 8K/rack", racional:"Revenue recurrente, bajo churn"},
    {id:"EB.SD.02", n:"Cloud Privado y Managed Services", m:"contrato", prio:"CRECIMIENTO", P_base:5000000,
     dt:"pipeline_b2b", pt:{pipeline:50, win_rate:0.50}, df:"mercado", pf:{pa:100000, pb:100000},
     componentes:"Infraestructura IaaS dedicada, DRaaS, Backup, Monitoreo 24/7",
     driver_q:"Pipeline servicios gestionados", driver_p:"Cotización por solución",
     unit_econ:"Por contrato", racional:"ARPU premium, stickiness"},
    {id:"EB.SD.03", n:"Edge Computing Empresarial", m:"nodo", prio:"CRECIMIENTO", P_base:4000000,
     dt:"pipeline_b2b", pt:{pipeline:40, win_rate:0.40}, df:"mercado", pf:{pa:150000, pb:150000},
     componentes:"Nodos edge en premisas cliente o co-ubicados",
     driver_q:"Pipeline edge B2B", driver_p:"Costo por nodo",
     unit_econ:"USD 150K/nodo", racional:"IoT industrial, real-time analytics"},
  ]},
  
  // Programa: Gobierno y Verticales
  {categoria:"Enterprise y B2B", dominio:"Gobierno y Verticales", tipo:"Crecimiento",
   programa:"Gobierno y Verticales", P_base:12000000, proyectos:[
    {id:"EB.GV.01", n:"Conectividad Social MINTIC", m:"proyecto", prio:"CRECIMIENTO", P_base:8000000,
     dt:"pipeline_b2b", pt:{pipeline:10, win_rate:0.50}, df:"mercado", pf:{pa:800000, pb:800000},
     componentes:"Despliegue zonas rurales, Escuelas, Hospitales, Centros digitales",
     driver_q:"Licitaciones adjudicadas", driver_p:"Valor contrato licitación",
     unit_econ:"Por proyecto", racional:"Posicionamiento, créditos USO"},
    {id:"EB.GV.02", n:"Verticales Estratégicos", m:"proyecto", prio:"CRECIMIENTO", P_base:4000000,
     dt:"pipeline_b2b", pt:{pipeline:20, win_rate:0.40}, df:"mercado", pf:{pa:500000, pb:500000},
     componentes:"Soluciones Oil and Gas, Energía, Minería, Nube soberana gobierno",
     driver_q:"Pipeline sectorial", driver_p:"Cotización por proyecto",
     unit_econ:"Por proyecto", racional:"Alto valor, contratos largo plazo"},
  ]},
  
  // Programa: Capacidad Mayorista
  {categoria:"Enterprise y B2B", dominio:"Wholesale", tipo:"Crecimiento",
   programa:"Capacidad Mayorista", P_base:5000000, proyectos:[
    {id:"EB.WS.01", n:"Servicios Wholesale", m:"Gbps", prio:"CRECIMIENTO", P_base:5000000,
     dt:"crecimiento_trafico", pt:{capacidad_actual:200, crecimiento_pct:0.50}, df:"mercado", pf:{pa:50000, pb:50000},
     componentes:"Capacidad IP/MPLS carriers, Tránsito OTTs, Dark fiber, IRUs",
     driver_q:"Contratos wholesale vigentes mas pipeline", driver_p:"Precios mercado",
     unit_econ:"Variable por Gbps", racional:"Monetización activos infraestructura"},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: TECNOLOGÍA
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Programa: Transformación BSS
  {categoria:"Tecnología", dominio:"BSS", tipo:"Transformación",
   programa:"Transformación BSS", P_base:45000000, proyectos:[
    {id:"TI.TB.01", n:"Modernización CRM", m:"programa", prio:"TRANSFORMACIÓN", P_base:25000000,
     dt:"migraciones", pt:{base_migrable:1, tasa:1}, df:"mercado", pf:{pa:25000000, pb:25000000},
     componentes:"Nueva plataforma CRM, Migración datos, Integraciones canales",
     driver_q:"Alcance funcional del programa", driver_p:"Contrato vendor/SI",
     unit_econ:"USD 25M programa", racional:"CX, omnicanalidad, eficiencia"},
    {id:"TI.TB.02", n:"Billing Convergente", m:"programa", prio:"TRANSFORMACIÓN", P_base:20000000,
     dt:"migraciones", pt:{base_migrable:1, tasa:1}, df:"mercado", pf:{pa:20000000, pb:20000000},
     componentes:"Rating engine, Charging, Facturación unificada, Revenue assurance",
     driver_q:"Alcance funcional del programa", driver_p:"Contrato vendor/SI",
     unit_econ:"USD 20M programa", racional:"Agilidad comercial, time-to-market"},
  ]},
  
  // Programa: Transformación OSS
  {categoria:"Tecnología", dominio:"OSS", tipo:"Transformación",
   programa:"Transformación OSS", P_base:10000000, proyectos:[
    {id:"TI.TO.01", n:"Orquestación y Automatización Red", m:"programa", prio:"TRANSFORMACIÓN", P_base:10000000,
     dt:"migraciones", pt:{base_migrable:1, tasa:1}, df:"mercado", pf:{pa:10000000, pb:10000000},
     componentes:"Orquestador SDN/NFV, Closed-loop automation, Inventario unificado",
     driver_q:"Casos de uso automatización priorizados", driver_p:"Contrato vendor/SI",
     unit_econ:"USD 10M programa", racional:"Reducción OPEX, zero-touch"},
  ]},
  
  // Programa: Plataformas Digitales
  {categoria:"Tecnología", dominio:"Digital", tipo:"Transformación",
   programa:"Plataformas Digitales", P_base:13000000, proyectos:[
    {id:"TI.PD.01", n:"Canales Digitales App y Web", m:"año", prio:"TRANSFORMACIÓN", P_base:5000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"mercado", pf:{pa:5000000, pb:5000000},
     componentes:"App Mi Claro, Portal web, E-commerce, Chatbot IA, Autoservicio",
     driver_q:"Roadmap digital anual", driver_p:"Costo fábrica software",
     unit_econ:"USD 5M/año", racional:"Digital mix, reducción costo servicio"},
    {id:"TI.PD.02", n:"Data y Analytics", m:"programa", prio:"TRANSFORMACIÓN", P_base:8000000,
     dt:"migraciones", pt:{base_migrable:1, tasa:1}, df:"mercado", pf:{pa:8000000, pb:8000000},
     componentes:"Data lake enterprise, Plataforma BI, AI/ML platform",
     driver_q:"Casos de uso analytics priorizados", driver_p:"Plataforma mas vendor",
     unit_econ:"USD 8M programa", racional:"Monetización datos, decisiones"},
  ]},
  
  // Programa: Sostenimiento IT
  {categoria:"Tecnología", dominio:"Infraestructura IT", tipo:"Operación",
   programa:"Sostenimiento IT", P_base:22000000, proyectos:[
    {id:"TI.SI.01", n:"Infraestructura y Licenciamiento", m:"año", prio:"PMO", P_base:12000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"historico", pf:{pa:12000000, pb:12000000},
     componentes:"Refresh servidores/storage, Licencias SW enterprise, Cloud IaaS/PaaS",
     driver_q:"Porcentaje EoL mas consumo cloud proyectado", driver_p:"Benchmark mas contratos",
     unit_econ:"Por ciclo refresh", racional:"Performance, compliance, soporte"},
    {id:"TI.SI.02", n:"Ciberseguridad", m:"año", prio:"PMO", P_base:5000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"mercado", pf:{pa:5000000, pb:5000000},
     componentes:"Herramientas seguridad, SOC 24/7, Compliance, Pentesting, DR",
     driver_q:"Landscape amenazas mas regulación", driver_p:"Benchmark mercado",
     unit_econ:"USD 5M/año", racional:"Gestión riesgo, compliance"},
    {id:"TI.SI.03", n:"Mantenimiento Aplicativo", m:"año", prio:"PMO", P_base:5000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"historico", pf:{pa:5000000, pb:5000000},
     componentes:"Corrección bugs, Parches seguridad, Mejoras menores, Soporte L2/L3",
     driver_q:"Portfolio aplicaciones críticas", driver_p:"Porcentaje valor desarrollo",
     unit_econ:"15-20% valor app/año", racional:"Estabilidad, deuda técnica"},
  ]},
  
  // Programa: Evolución Plataformas
  {categoria:"Tecnología", dominio:"Evolución Plataformas", tipo:"Evolución",
   programa:"Evolución Plataformas", P_base:6000000, proyectos:[
    {id:"TI.EP.01", n:"Mejoras BSS CRM y Billing", m:"año", prio:"EVOLUCIÓN", P_base:4000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"mercado", pf:{pa:4000000, pb:4000000},
     componentes:"Nuevas features CRM, Productos en billing, APIs, Integraciones",
     driver_q:"Backlog features priorizados", driver_p:"Costo fábrica software",
     unit_econ:"USD 4M/año", racional:"Mejora continua, competitividad"},
    {id:"TI.EP.02", n:"Mejoras Canales Digitales", m:"año", prio:"EVOLUCIÓN", P_base:2000000,
     dt:"mantenimiento", pt:{activos:1, frecuencia:1}, df:"mercado", pf:{pa:2000000, pb:2000000},
     componentes:"Features app/web, UX improvements, Integraciones ecosistema",
     driver_q:"Releases planificados por año", driver_p:"Costo fábrica software",
     unit_econ:"USD 2M/año", racional:"CX, engagement, NPS digital"},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: INVERSIÓN CLIENTES
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Programa: Adquisición Consumidor
  {categoria:"Inversión Clientes", dominio:"Consumidor", tipo:"Comercial",
   programa:"Adquisición Consumidor", P_base:75000000, proyectos:[
    {id:"IC.AC.01", n:"Activaciones Fijo Hogar", m:"activación", prio:"CRECIMIENTO", P_base:32000000,
     dt:"activaciones", pt:{brutas:400000, churn:100000}, df:"benchmark_amx", pf:{pa:80, pb:80},
     componentes:"Instalación, Provisioning, CPE primera línea, Comisión canal",
     driver_q:"Altas brutas fijo proyectadas", driver_p:"Costo por activación",
     unit_econ:"USD 80/activación", racional:"Crecimiento base fijo"},
    {id:"IC.AC.02", n:"Activación y Subsidio Móvil", m:"alta", prio:"CRECIMIENTO", P_base:43000000,
     dt:"activaciones", pt:{brutas:500000, churn:70000}, df:"benchmark_amx", pf:{pa:100, pb:100},
     componentes:"Subsidio terminales postpago, SIMs, Activación prepago, Comisiones",
     driver_q:"Altas brutas móvil proyectadas", driver_p:"Costo promedio por alta",
     unit_econ:"USD 100/alta promedio", racional:"Mix postpago, market share"},
  ]},
  
  // Programa: Retención Consumidor
  {categoria:"Inversión Clientes", dominio:"Consumidor", tipo:"Comercial",
   programa:"Retención Consumidor", P_base:28000000, proyectos:[
    {id:"IC.RC.01", n:"Programas Retención y Lealtad", m:"save", prio:"RETENCIÓN", P_base:20000000,
     dt:"migraciones", pt:{base_migrable:400000, tasa:0.25}, df:"benchmark_amx", pf:{pa:50, pb:50},
     componentes:"Ofertas retención, Upgrade equipos, Programa loyalty, Beneficios",
     driver_q:"Suscriptores alto valor en riesgo churn", driver_p:"Costo por save",
     unit_econ:"USD 50/save", racional:"Reducción churn, lifetime value"},
    {id:"IC.RC.02", n:"Migraciones ARPU-Accretive", m:"upgrade", prio:"CRECIMIENTO", P_base:8000000,
     dt:"migraciones", pt:{base_migrable:400000, tasa:0.20}, df:"benchmark_amx", pf:{pa:30, pb:30},
     componentes:"Migración cobre/HFC a fibra, Upgrade velocidad/datos",
     driver_q:"Base elegible para migración", driver_p:"Costo por upgrade",
     unit_econ:"USD 30/upgrade", racional:"ARPU uplift, satisfacción"},
  ]},
  
  // Programa: Inversión B2B
  {categoria:"Inversión Clientes", dominio:"Enterprise", tipo:"Comercial",
   programa:"Inversión B2B", P_base:12000000, proyectos:[
    {id:"IC.IB.01", n:"Adquisición Enterprise y PyME", m:"deal", prio:"CRECIMIENTO", P_base:8000000,
     dt:"pipeline_b2b", pt:{pipeline:200, win_rate:0.50}, df:"mercado", pf:{pa:80000, pb:80000},
     componentes:"POCs, Diseño solución, Onboarding, Adquisición PyME simplificada",
     driver_q:"Deals enterprise mas PyMEs nuevas", driver_p:"Porcentaje TCV mas costo PyME",
     unit_econ:"5-8% TCV", racional:"Crecimiento revenue B2B"},
    {id:"IC.IB.02", n:"Renovación Contratos B2B", m:"contrato", prio:"RETENCIÓN", P_base:4000000,
     dt:"pipeline_b2b", pt:{pipeline:300, win_rate:0.70}, df:"mercado", pf:{pa:19000, pb:19000},
     componentes:"Incentivos renovación, Upgrade servicios, Cross-sell, Upsell",
     driver_q:"ACV contratos por renovar", driver_p:"Porcentaje costo retención sobre ACV",
     unit_econ:"2-3% ACV", racional:"Protección y expansión revenue"},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: CORPORATIVO
  // ═══════════════════════════════════════════════════════════════════════════
  
  // Programa: Soporte Corporativo
  {categoria:"Corporativo", dominio:"Soporte", tipo:"Operación",
   programa:"Soporte Corporativo", P_base:6000000, proyectos:[
    {id:"CO.SC.01", n:"Instalaciones y Sedes", m:"m2", prio:"PMO", P_base:4000000,
     dt:"mantenimiento", pt:{activos:40000, frecuencia:1}, df:"historico", pf:{pa:100, pb:100},
     componentes:"Adecuaciones oficinas, Mantenimiento preventivo, Seguridad física",
     driver_q:"m2 mas número de sedes", driver_p:"Costo/m2 mas contratos servicios",
     unit_econ:"USD 100/m2/año", racional:"Experiencia empleado, activos"},
    {id:"CO.SC.02", n:"Flota Vehicular", m:"vehículo", prio:"PMO", P_base:2000000,
     dt:"mantenimiento", pt:{activos:72, frecuencia:1}, df:"mercado", pf:{pa:28000, pb:28000},
     componentes:"Flota técnica instaladores/mantenimiento, Flota administrativa",
     driver_q:"Vehículos en ciclo reemplazo", driver_p:"Costo promedio vehículo",
     unit_econ:"USD 28K/vehículo", racional:"Operaciones campo, SLA técnico"},
  ]},
];

/* ══════════════════════════════════════════════════════════════════════════
   DRIVERS DE CANTIDAD Y PRECIO
══════════════════════════════════════════════════════════════════════════ */

const DRV_N2 = {
  obsolescencia: {
    label:"Obsolescencia EoL/EoS", icon:"⚠️", color:"#D97706", tag:"Demanda",
    desc:"Activos fuera de soporte que deben reemplazarse",
    params:[{k:"total_activos",l:"Activos en parque",p:false},{k:"pct_eol",l:"% EoL actual",p:true},{k:"meta_pct",l:"Meta EoL ≤",p:true}],
    calcQ:d=>Math.round((d.total_activos||0)*Math.max(0,(d.pct_eol||0)-(d.meta_pct||0))),
  },
  congestion: {
    label:"Congestión PRB >80%", icon:"📶", color:"#E8182A", tag:"Demanda",
    desc:"Sitios congestionados que superan el umbral de calidad",
    params:[{k:"total_activos",l:"Total sitios red",p:false},{k:"pct_congestion",l:"% congestionados",p:true},{k:"meta_pct",l:"Meta ≤",p:true}],
    calcQ:d=>Math.round((d.total_activos||0)*Math.max(0,(d.pct_congestion||0)-(d.meta_pct||0))),
  },
  cobertura_crc: {
    label:"Obligación Regulatoria CRC", icon:"⚖️", color:"#E8182A", tag:"Regulatorio",
    desc:"Compromisos legales de cobertura pendientes de ejecutar",
    params:[{k:"comprometido",l:"Comprometido CRC",p:false},{k:"ejecutado",l:"Ya ejecutado",p:false}],
    calcQ:d=>Math.max(0,(d.comprometido||0)-(d.ejecutado||0)),
  },
  activaciones: {
    label:"Activaciones Netas Plan", icon:"📈", color:"#059669", tag:"Demanda",
    desc:"Nuevos clientes proyectados menos churn esperado",
    params:[{k:"brutas",l:"Activaciones brutas",p:false},{k:"churn",l:"Churn estimado",p:false}],
    calcQ:d=>Math.max(0,(d.brutas||0)-(d.churn||0)),
  },
  migraciones: {
    label:"Migraciones de Tecnología", icon:"🔄", color:"#2563EB", tag:"Demanda",
    desc:"Clientes o equipos a migrar según tasa de adopción",
    params:[{k:"base_migrable",l:"Base migrable",p:false},{k:"tasa",l:"Tasa migración",p:true}],
    calcQ:d=>Math.round((d.base_migrable||0)*(d.tasa||0)),
  },
  crecimiento_trafico: {
    label:"Crecimiento de Tráfico", icon:"🌐", color:"#2563EB", tag:"Demanda",
    desc:"Capacidad adicional requerida por crecimiento de tráfico",
    params:[{k:"capacidad_actual",l:"Capacidad actual",p:false},{k:"crecimiento_pct",l:"Crecimiento %",p:true}],
    calcQ:d=>Math.round((d.capacidad_actual||0)*(d.crecimiento_pct||0)),
  },
  expansion_geo: {
    label:"Expansión de Cobertura", icon:"🗺️", color:"#059669", tag:"Demanda",
    desc:"Gap entre cobertura objetivo y cobertura actual",
    params:[{k:"objetivo",l:"Objetivo (HP/km²)",p:false},{k:"actual",l:"Actual cubierto",p:false}],
    calcQ:d=>Math.max(0,(d.objetivo||0)-(d.actual||0)),
  },
  pipeline_b2b: {
    label:"Pipeline B2B / Empresas", icon:"🏢", color:"#C8941C", tag:"Demanda",
    desc:"Oportunidades calificadas ponderadas por win-rate",
    params:[{k:"pipeline",l:"Pipeline calificado",p:false},{k:"win_rate",l:"Win-rate histórico",p:true}],
    calcQ:d=>Math.round((d.pipeline||0)*(d.win_rate||0)),
  },
  mantenimiento: {
    label:"Plan Mantenimiento PMO", icon:"🔧", color:"#6B6860", tag:"Operación",
    desc:"Intervenciones de mantenimiento programadas al año",
    params:[{k:"activos",l:"Activos en parque",p:false},{k:"frecuencia",l:"Intervenciones/año",p:false}],
    calcQ:d=>Math.round((d.activos||0)*(d.frecuencia||1)),
  },
};

const DRV_FIN = {
  benchmark_amx:{label:"Benchmark Grupo AMX", color:"#2563EB", desc:"Precio de referencia del grupo AMX para la misma categoría de inversión en LatAm"},
  historico:{label:"Histórico Contractual Claro", color:"#6B6860", desc:"Promedio ponderado de contratos ejecutados por Claro en los últimos 3 años"},
  benchmark_crc:{label:"Valor Regulatorio CRC", color:"#E8182A", desc:"Precio máximo o referencia establecido por resolución regulatoria CRC vigente"},
  mercado:{label:"Precio de Mercado", color:"#C8941C", desc:"Mejor precio obtenido en proceso competitivo abierto o cotización spot"},
};

/* ══════════════════════════════════════════════════════════════════════════
   HELPERS
══════════════════════════════════════════════════════════════════════════ */

const fu = v => {
  if(v===undefined||v===null||isNaN(v)) return "—";
  const abs=Math.abs(v);
  if(abs>=1e9)  return `$${(v/1e9).toFixed(2)}B`;
  if(abs>=1e6)  return `$${(v/1e6).toFixed(2)}M`;
  if(abs>=1e3)  return `$${(v/1e3).toFixed(1)}K`;
  return `$${v.toFixed(0)}`;
};

const fn = v => {
  if(v===undefined||v===null||isNaN(v)) return "—";
  if(Math.abs(v)>=1e6) return `${(v/1e6).toFixed(1)}M`;
  if(Math.abs(v)>=1e3) return `${(v/1e3).toFixed(1)}K`;
  return String(Math.round(v));
};

const dp = (a,b)=> b>0 ? (a-b)/b*100 : 0;
const sg = v => v>=0 ? "+" : "";
const cc = v => v<-0.5 ? T.green : v>0.5 ? T.red : T.inkMid;

const calcQ = (dk, pt) => {
  const D = DRV_N2[dk];
  return D?.calcQ ? D.calcQ(pt) : 0;
};

const calcCap = (p, ov) => {
  if(!ov) return p.P_base;
  if(ov._capex != null) return ov._capex;
  const dk = ov.dt || p.dt;
  const pt = ov.pt ? {...p.pt,...ov.pt} : {...p.pt};
  const pf = ov.pf ? {...p.pf,...ov.pf} : {...p.pf};
  const Q = calcQ(dk, pt);
  const P = pf.pb;
  if(Q>0 && P>0) return Q*P;
  return p.P_base;
};

/* ══════════════════════════════════════════════════════════════════════════
   UI COMPONENTS
══════════════════════════════════════════════════════════════════════════ */

function NoiseSVG(){
  return(
    <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",pointerEvents:"none",opacity:.04}} xmlns="http://www.w3.org/2000/svg">
      <filter id="nz"><feTurbulence type="fractalNoise" baseFrequency=".65" numOctaves="3" stitchTiles="stitch"/></filter>
      <rect width="100%" height="100%" filter="url(#nz)"/>
    </svg>
  );
}

function PBar({pct,color=T.red,h=4}){
  return(
    <div style={{height:h,background:T.borderSm,borderRadius:99,overflow:"hidden"}}>
      <div style={{height:"100%",width:`${Math.min(100,Math.max(0,pct))}%`,background:color,borderRadius:99,animation:"barGrow .6s cubic-bezier(.22,1,.36,1) both"}}/>
    </div>
  );
}

function TipoBadge({tipo, sm}){
  const cfg = TIPO_CFG[tipo] || TIPO_CFG["Crecimiento"];
  return(
    <span style={{
      display:"inline-flex",alignItems:"center",padding:sm?"2px 8px":"3px 10px",
      borderRadius:99,background:cfg.bg,border:`1px solid ${cfg.bdr}`,
      fontSize:sm?8.5:10,fontWeight:700,color:cfg.c,
      letterSpacing:".05em",textTransform:"uppercase",whiteSpace:"nowrap"
    }}>
      {cfg.label}
    </span>
  );
}

function CategoriaBadge({cat}){
  const cfg = CATEGORIA_CFG[cat] || CATEGORIA_CFG["Red"];
  return(
    <span style={{
      display:"inline-flex",alignItems:"center",gap:4,padding:"3px 10px",
      borderRadius:99,background:cfg.bg,border:`1px solid ${cfg.bdr}`,
      fontSize:10,fontWeight:700,color:cfg.c,whiteSpace:"nowrap"
    }}>
      <span>{cfg.icon}</span>
      <span>{cat}</span>
    </span>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   MAIN APP COMPONENT
══════════════════════════════════════════════════════════════════════════ */

export default function App() {
  const [session, setSession] = useState(null);
  const [profile, setProfile] = useState(null);
  const [authReady, setAuthReady] = useState(false);
  const [scenarios, setScenarios] = useState([]);
  const [activeScen, setActiveScen] = useState("base");
  const [overrides, setOverrides] = useState({});
  const [auditLog, setAuditLog] = useState([]);
  
  // Vista actual
  const [view, setView] = useState("dashboard"); // dashboard, detail
  const [selectedProg, setSelectedProg] = useState(null);
  const [selectedProy, setSelectedProy] = useState(null);
  
  // Filtros
  const [filterCat, setFilterCat] = useState("all");
  const [filterDom, setFilterDom] = useState("all");
  const [filterTipo, setFilterTipo] = useState("all");
  
  // Auth
  const initUser = useCallback(async (s) => {
    if (!s) { setSession(null); setProfile(null); return; }
    setSession(s);
    const p = await getProfile(s.user.id);
    setProfile(p);
    await touchLastSeen(s.user.id);
    const { data: scen } = await fetchScenarios(s.user.id);
    setScenarios(scen || []);
    const { data: logs } = await fetchLog(s.user.id);
    setAuditLog(logs || []);
  }, []);
  
  useEffect(() => {
    const style = document.createElement("style");
    style.textContent = GS;
    document.head.appendChild(style);
    return () => document.head.removeChild(style);
  }, []);
  
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session: s } }) => {
      initUser(s);
      setAuthReady(true);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_e, s) => {
      initUser(s);
    });
    return () => subscription.unsubscribe();
  }, [initUser]);
  
  if (!authReady) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.surface}}>
      <div style={{textAlign:"center"}}>
        <div style={{fontSize:32,marginBottom:12}}>⏳</div>
        <div style={{color:T.inkMid,fontSize:14}}>Cargando...</div>
      </div>
    </div>
  );
  
  if (!session) return <AuthLogin />;
  
  // Handlers
  const handleSignOut = async () => {
    await authSignOut();
    setSession(null);
  };
  
  const handleChange = (id, patch) => {
    setOverrides(p => ({...p, [id]: {...(p[id]||{}), ...patch}}));
  };
  
  // Cálculos
  const totals = useMemo(() => {
    let base = 0, dvb = 0;
    const byCat = {}, byDom = {}, byTipo = {}, byProg = {};
    
    DATA.forEach(prog => {
      const progKey = `${prog.categoria}|${prog.dominio}|${prog.programa}`;
      let progBase = 0, progDvb = 0;
      
      prog.proyectos.forEach(p => {
        const cap = calcCap(p, overrides[p.id]);
        progBase += p.P_base;
        progDvb += cap;
      });
      
      byProg[progKey] = { base: progBase, dvb: progDvb, ...prog };
      
      base += progBase;
      dvb += progDvb;
      
      byCat[prog.categoria] = byCat[prog.categoria] || { base: 0, dvb: 0 };
      byCat[prog.categoria].base += progBase;
      byCat[prog.categoria].dvb += progDvb;
      
      byDom[prog.dominio] = byDom[prog.dominio] || { base: 0, dvb: 0 };
      byDom[prog.dominio].base += progBase;
      byDom[prog.dominio].dvb += progDvb;
      
      byTipo[prog.tipo] = byTipo[prog.tipo] || { base: 0, dvb: 0 };
      byTipo[prog.tipo].base += progBase;
      byTipo[prog.tipo].dvb += progDvb;
    });
    
    return { base, dvb, byCat, byDom, byTipo, byProg };
  }, [overrides]);
  
  // Datos filtrados
  const filteredData = useMemo(() => {
    return DATA.filter(prog => {
      if (filterCat !== "all" && prog.categoria !== filterCat) return false;
      if (filterDom !== "all" && prog.dominio !== filterDom) return false;
      if (filterTipo !== "all" && prog.tipo !== filterTipo) return false;
      return true;
    });
  }, [filterCat, filterDom, filterTipo]);
  
  // Listas únicas para filtros
  const categorias = [...new Set(DATA.map(d => d.categoria))];
  const dominios = [...new Set(DATA.map(d => d.dominio))];
  const tipos = [...new Set(DATA.map(d => d.tipo))];
  
  const delta = dp(totals.dvb, totals.base);
  
  return (
    <div style={{minHeight:"100vh", background:T.surface}}>
      
      {/* ══════ HEADER ══════ */}
      <header style={{
        background:`linear-gradient(135deg, ${T.redDk}, ${T.red})`,
        padding:"16px 24px",
        position:"sticky", top:0, zIndex:100,
        boxShadow:"0 4px 20px rgba(0,0,0,0.15)"
      }}>
        <NoiseSVG />
        <div style={{position:"relative", zIndex:1, display:"flex", justifyContent:"space-between", alignItems:"center"}}>
          <div>
            <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:4}}>
              <span style={{fontSize:20}}>📊</span>
              <span style={{fontSize:18, fontWeight:800, color:"#fff", letterSpacing:"-.02em"}}>
                DVB Command Center
              </span>
            </div>
            <div style={{fontSize:11, color:"rgba(255,255,255,0.7)", fontWeight:500}}>
              Kearney × Claro Colombia · CAPEX 2026
            </div>
          </div>
          
          <div style={{display:"flex", alignItems:"center", gap:16}}>
            {/* Totales */}
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:8, color:"rgba(255,255,255,0.5)", textTransform:"uppercase", letterSpacing:".1em"}}>
                CAPEX Total
              </div>
              <div style={{fontSize:22, fontWeight:900, color:"#fff"}}>
                {fu(totals.dvb)}
              </div>
              {Math.abs(delta) > 0.1 && (
                <div style={{fontSize:10, fontWeight:700, color: delta < 0 ? "#86EFAC" : "#FCA5A5"}}>
                  {sg(delta)}{delta.toFixed(1)}% vs base
                </div>
              )}
            </div>
            
            {/* User */}
            <div style={{display:"flex", alignItems:"center", gap:8}}>
              <div style={{
                width:36, height:36, borderRadius:"50%",
                background:"rgba(255,255,255,0.15)",
                display:"flex", alignItems:"center", justifyContent:"center",
                fontSize:14, color:"#fff", fontWeight:700
              }}>
                {(profile?.full_name || session.user.email)[0].toUpperCase()}
              </div>
              <button onClick={handleSignOut} className="btn-ghost" style={{
                background:"rgba(255,255,255,0.1)",
                border:"1px solid rgba(255,255,255,0.2)",
                borderRadius:8, padding:"6px 12px",
                color:"rgba(255,255,255,0.8)", fontSize:11, fontWeight:600,
                cursor:"pointer", fontFamily:"'Outfit',system-ui"
              }}>
                Salir
              </button>
            </div>
          </div>
        </div>
      </header>
      
      {/* ══════ FILTROS ══════ */}
      <div style={{
        background:T.card,
        borderBottom:`1px solid ${T.border}`,
        padding:"12px 24px",
        display:"flex", gap:12, alignItems:"center", flexWrap:"wrap"
      }}>
        <span style={{fontSize:11, fontWeight:700, color:T.inkMid}}>Filtros:</span>
        
        <select value={filterCat} onChange={e => setFilterCat(e.target.value)} style={{
          padding:"6px 12px", borderRadius:8, border:`1px solid ${T.border}`,
          fontSize:11, fontWeight:600, background:T.card, color:T.ink,
          cursor:"pointer", fontFamily:"'Outfit',system-ui"
        }}>
          <option value="all">Todas las Categorías</option>
          {categorias.map(c => <option key={c} value={c}>{CATEGORIA_CFG[c]?.icon} {c}</option>)}
        </select>
        
        <select value={filterDom} onChange={e => setFilterDom(e.target.value)} style={{
          padding:"6px 12px", borderRadius:8, border:`1px solid ${T.border}`,
          fontSize:11, fontWeight:600, background:T.card, color:T.ink,
          cursor:"pointer", fontFamily:"'Outfit',system-ui"
        }}>
          <option value="all">Todos los Dominios</option>
          {dominios.map(d => <option key={d} value={d}>{DOMINIO_CFG[d]?.icon} {d}</option>)}
        </select>
        
        <select value={filterTipo} onChange={e => setFilterTipo(e.target.value)} style={{
          padding:"6px 12px", borderRadius:8, border:`1px solid ${T.border}`,
          fontSize:11, fontWeight:600, background:T.card, color:T.ink,
          cursor:"pointer", fontFamily:"'Outfit',system-ui"
        }}>
          <option value="all">Todos los Tipos</option>
          {tipos.map(t => <option key={t} value={t}>{t}</option>)}
        </select>
        
        <div style={{marginLeft:"auto", fontSize:11, color:T.inkSoft}}>
          {filteredData.reduce((s,p) => s + p.proyectos.length, 0)} proyectos · {filteredData.length} programas
        </div>
      </div>
      
      {/* ══════ KPIs POR CATEGORÍA ══════ */}
      <div style={{padding:"20px 24px"}}>
        <div style={{display:"grid", gridTemplateColumns:"repeat(5, 1fr)", gap:12}}>
          {categorias.map(cat => {
            const cfg = CATEGORIA_CFG[cat];
            const data = totals.byCat[cat] || { base: 0, dvb: 0 };
            const d = dp(data.dvb, data.base);
            const pct = totals.dvb > 0 ? (data.dvb / totals.dvb * 100) : 0;
            
            return (
              <div key={cat} className="macro-card hover-lift fu" style={{
                background:T.card,
                borderRadius:14,
                padding:"14px 16px",
                border:`1.5px solid ${filterCat === cat ? cfg.c : T.border}`,
                cursor:"pointer",
                boxShadow:"0 2px 8px rgba(0,0,0,0.04)"
              }} onClick={() => setFilterCat(filterCat === cat ? "all" : cat)}>
                <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:10}}>
                  <div style={{
                    width:32, height:32, borderRadius:8, background:cfg.bg,
                    display:"flex", alignItems:"center", justifyContent:"center",
                    fontSize:16
                  }}>
                    {cfg.icon}
                  </div>
                  <div style={{fontSize:11, fontWeight:700, color:cfg.c}}>{cat}</div>
                </div>
                <div style={{fontSize:18, fontWeight:900, color:T.ink, marginBottom:4}}>
                  {fu(data.dvb)}
                </div>
                <div style={{display:"flex", justifyContent:"space-between", alignItems:"center"}}>
                  <span style={{fontSize:10, color:T.inkSoft}}>{pct.toFixed(0)}% del total</span>
                  {Math.abs(d) > 0.1 && (
                    <span style={{fontSize:9, fontWeight:700, color:cc(d)}}>
                      {sg(d)}{d.toFixed(1)}%
                    </span>
                  )}
                </div>
                <div style={{marginTop:8}}>
                  <PBar pct={pct} color={cfg.c} h={3} />
                </div>
              </div>
            );
          })}
        </div>
      </div>
      
      {/* ══════ LISTA DE PROGRAMAS ══════ */}
      <div style={{padding:"0 24px 40px"}}>
        {filteredData.map((prog, pi) => {
          const cfg = CATEGORIA_CFG[prog.categoria];
          const tipoCfg = TIPO_CFG[prog.tipo];
          const domCfg = DOMINIO_CFG[prog.dominio];
          
          let progBase = 0, progDvb = 0;
          prog.proyectos.forEach(p => {
            progBase += p.P_base;
            progDvb += calcCap(p, overrides[p.id]);
          });
          const progDelta = dp(progDvb, progBase);
          
          return (
            <div key={`${prog.categoria}-${prog.dominio}-${prog.programa}-${pi}`} 
                 className="fu" style={{animationDelay:`${pi * 0.03}s`}}>
              
              {/* Programa Header */}
              <div style={{
                background:T.card,
                borderRadius:12,
                marginBottom:8,
                border:`1px solid ${T.border}`,
                overflow:"hidden"
              }}>
                {/* Header del programa */}
                <div style={{
                  padding:"12px 16px",
                  borderBottom:`1px solid ${T.borderSm}`,
                  display:"flex", alignItems:"center", justifyContent:"space-between",
                  background:`linear-gradient(135deg, ${cfg.bg}, ${T.card})`
                }}>
                  <div style={{display:"flex", alignItems:"center", gap:12}}>
                    <div style={{
                      width:36, height:36, borderRadius:10, background:cfg.c,
                      display:"flex", alignItems:"center", justifyContent:"center",
                      fontSize:16, color:"#fff"
                    }}>
                      {cfg.icon}
                    </div>
                    <div>
                      <div style={{display:"flex", alignItems:"center", gap:8, marginBottom:2}}>
                        <span style={{fontSize:13, fontWeight:800, color:T.ink}}>{prog.programa}</span>
                        <TipoBadge tipo={prog.tipo} sm />
                      </div>
                      <div style={{fontSize:10, color:T.inkSoft}}>
                        {prog.categoria} → {prog.dominio} · {prog.proyectos.length} proyectos
                      </div>
                    </div>
                  </div>
                  <div style={{textAlign:"right"}}>
                    <div style={{fontSize:16, fontWeight:900, color:T.ink}}>{fu(progDvb)}</div>
                    {Math.abs(progDelta) > 0.1 && (
                      <div style={{fontSize:9, fontWeight:700, color:cc(progDelta)}}>
                        {sg(progDelta)}{progDelta.toFixed(1)}% vs base
                      </div>
                    )}
                  </div>
                </div>
                
                {/* Proyectos */}
                <div style={{padding:"8px"}}>
                  {prog.proyectos.map((proy, i) => {
                    const cap = calcCap(proy, overrides[proy.id]);
                    const d = dp(cap, proy.P_base);
                    const Q = calcQ(proy.dt, overrides[proy.id]?.pt || proy.pt);
                    const P = (overrides[proy.id]?.pf?.pb || proy.pf.pb);
                    
                    return (
                      <div key={proy.id} className="row-hover" style={{
                        padding:"10px 12px",
                        borderRadius:8,
                        marginBottom: i < prog.proyectos.length - 1 ? 4 : 0,
                        display:"grid",
                        gridTemplateColumns:"1fr 120px 100px 100px 80px",
                        alignItems:"center",
                        gap:12
                      }}>
                        {/* Nombre */}
                        <div>
                          <div style={{display:"flex", alignItems:"center", gap:6}}>
                            <span style={{
                              fontSize:9, fontWeight:700, color:T.inkSoft,
                              background:T.surface, padding:"2px 6px", borderRadius:4
                            }}>
                              {proy.id}
                            </span>
                            <span style={{fontSize:12, fontWeight:600, color:T.ink}}>{proy.n}</span>
                          </div>
                          <div style={{fontSize:9, color:T.inkSoft, marginTop:2}}>
                            {proy.componentes?.substring(0, 60)}...
                          </div>
                        </div>
                        
                        {/* Q */}
                        <div style={{textAlign:"center"}}>
                          <div style={{fontSize:11, fontWeight:700, color:T.blue}}>{fn(Q)}</div>
                          <div style={{fontSize:8, color:T.inkSoft}}>{proy.m}</div>
                        </div>
                        
                        {/* P */}
                        <div style={{textAlign:"center"}}>
                          <div style={{fontSize:11, fontWeight:700, color:T.violet}}>{fu(P)}</div>
                          <div style={{fontSize:8, color:T.inkSoft}}>/{proy.m}</div>
                        </div>
                        
                        {/* CAPEX */}
                        <div style={{textAlign:"right"}}>
                          <div style={{fontSize:12, fontWeight:800, color:T.ink}}>{fu(cap)}</div>
                          {Math.abs(d) > 0.1 && (
                            <div style={{fontSize:8, fontWeight:700, color:cc(d)}}>
                              {sg(d)}{d.toFixed(1)}%
                            </div>
                          )}
                        </div>
                        
                        {/* Prioridad */}
                        <div style={{textAlign:"center"}}>
                          <span style={{
                            fontSize:8, fontWeight:700, color:T.inkSoft,
                            background:T.surface, padding:"2px 8px", borderRadius:99,
                            textTransform:"uppercase", letterSpacing:".05em"
                          }}>
                            {proy.prio}
                          </span>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          );
        })}
      </div>
      
      {/* ══════ FOOTER ══════ */}
      <footer style={{
        background:T.card,
        borderTop:`1px solid ${T.border}`,
        padding:"16px 24px",
        textAlign:"center"
      }}>
        <div style={{fontSize:10, color:T.inkSoft}}>
          DVB Command Center v2.0 · Marco CAPEX 2026 · {DATA.reduce((s,p) => s + p.proyectos.length, 0)} proyectos
        </div>
      </footer>
    </div>
  );
}
