import { useState, useMemo, useCallback, useEffect, useRef } from "react";
import "./index.css";
import * as XLSX from "xlsx";
import { loadEstructura } from "./data/loadEstructura";
import { supabase, getProfile, touchLastSeen, fetchScenarios,
         upsertScenario, removeScenario, pushLog, fetchLog,
         authSignOut } from "./supabase";
import AuthLogin from "./AuthLogin";

/* ══════════════════════════════════════════════════════════════════════════
   DVB COMMAND CENTER · Kearney × Claro Colombia · CAPEX 2026
   Design: Outfit · #F7F6F3 bg · Kearney Red #E8182A · cards blancas
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

/* ── Tokens ─────────────────────────────────────────────────────────────── */
const T = {
  red:"#E8182A", redDk:"#B5111F", redBg:"#FFF1F1", redSoft:"#FDD8DA", redXsoft:"#FFF8F8",
  ink:"#111110", inkMid:"#6B6860", inkSoft:"#9C9A95", inkXsoft:"#C8C6C0",
  surface:"#F7F6F3", card:"#FFFFFF", border:"#E8E6E0", borderSm:"#F0EEE9",
  gold:"#C8941C", goldBg:"#FFFBEB", goldBdr:"#FDE68A",
  blue:"#2563EB", blueBg:"#EFF6FF", blueBdr:"#BFDBFE",
  green:"#059669", greenBg:"#ECFDF5", greenBdr:"#A7F3D0",
  violet:"#7C3AED", violetBg:"#F5F3FF", violetBdr:"#DDD6FE",
  orange:"#D97706",
};
const WHITE="#FFFFFF";

const PURPLE_LT="#F5F3FF";

const TIPO_CFG = {
  // Tipos de inversión (nuevos)
  "Crecimiento":     {c:"#059669",bg:"#ECFDF5",bdr:"#A7F3D0",label:"Crecimiento",abv:"CRE"},
  "Modernización":   {c:"#D97706",bg:"#FFFBEB",bdr:"#FDE68A",label:"Modernización",abv:"MOD"},
  "Operación":       {c:"#6B7280",bg:"#F9FAFB",bdr:"#E5E7EB",label:"Operación",abv:"OPE"},
  "Regulatorio":     {c:"#B45309",bg:"#FFFBEB",bdr:"#FCD34D",label:"Regulatorio",abv:"REG"},
  "Transformación":  {c:"#7C3AED",bg:"#F5F3FF",bdr:"#DDD6FE",label:"Transformación",abv:"TRF"},
  "Evolución":       {c:"#2563EB",bg:"#EFF6FF",bdr:"#BFDBFE",label:"Evolución",abv:"EVO"},
  "Comercial":       {c:"#E8182A",bg:"#FFF1F1",bdr:"#FDD8DA",label:"Comercial",abv:"COM"},
  // Categorías
  "Red":                 {c:"#2563EB",bg:"#EFF6FF",bdr:"#BFDBFE",label:"Red",abv:"RED"},
  "Enterprise y B2B":    {c:"#7C3AED",bg:"#F5F3FF",bdr:"#DDD6FE",label:"Enterprise y B2B",abv:"ENT"},
  "Tecnología":          {c:"#0891B2",bg:"#ECFEFF",bdr:"#A5F3FC",label:"Tecnología",abv:"TEC"},
  "Inversión Clientes":  {c:"#E8182A",bg:"#FFF1F1",bdr:"#FDD8DA",label:"Inversión Clientes",abv:"CLI"},
  "Corporativo":         {c:"#6B7280",bg:"#F9FAFB",bdr:"#E5E7EB",label:"Corporativo",abv:"COR"},
  // Mantener compatibilidad con tipos antiguos
  "Network Rollout":        {c:"#2563EB",bg:"#EFF6FF",bdr:"#BFDBFE",label:"Rollout De Red",abv:"NRO"},
  "Network Modernization":  {c:"#D97706",bg:"#FFFBEB",bdr:"#FDE68A",label:"Modernización De Red",abv:"MOD"},
  "Capacity Expansion":     {c:"#059669",bg:"#ECFDF5",bdr:"#A7F3D0",label:"Expansión De Capacidad",abv:"CAP"},
  "Customer Investment":    {c:"#E8182A",bg:"#FFF1F1",bdr:"#FDD8DA",label:"Inversión En Clientes",abv:"CUI"},
  "Enterprise & Wholesale": {c:"#7C3AED",bg:"#F5F3FF",bdr:"#DDD6FE",label:"Empresarial Y Mayorista",abv:"ENT"},
  "Regulatory & Spectrum":  {c:"#B45309",bg:"#FFFBEB",bdr:"#FCD34D",label:"Regulatorio Y Espectro",abv:"REG"},
  "IT & Digital":           {c:"#0891B2",bg:"#ECFEFF",bdr:"#A5F3FC",label:"TI Y Digital",abv:"ITD"},
  "Network Operations":     {c:"#6B7280",bg:"#F9FAFB",bdr:"#E5E7EB",label:"Operaciones De Red",abv:"NOP"},
  "Gestión Administrativa": {c:"#92400E",bg:"#FEF3C7",bdr:"#FDE68A",label:"Gestión Administrativa",abv:"ADM"},
};
const CATEGORIA_CFG = TIPO_CFG;

/* ══════════════════════════════════════════════════════════════════════════
   ÁRBOL PxQ — ESTRUCTURA MULTINIVEL (vacío por ahora - nuevos IDs)
══════════════════════════════════════════════════════════════════════════ */
const PXQ_TREE = {};

// Estructura CAPEX — se reemplaza con /public/capex-estructura.xlsx al iniciar.
// Si el Excel falla, esta definición sirve de fallback.
let DATA = [
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: RED - ACCESO FIJO
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Expansión Cobertura Fijo",categoria:"Red",tipo:"Crecimiento",P_base:45000000,proyectos:[
    {id:"AF.EC.01",n:"FTTH Greenfield",m:"HHP",prio:"CRECIMIENTO",P_base:25000000,dt:"expansion_geo",pt:{objetivo:200000,actual:120000},df:"benchmark_amx",pf:{pa:160,pb:160}},
    {id:"AF.EC.02",n:"FTTH Infill",m:"HHP",prio:"CRECIMIENTO",P_base:12000000,dt:"expansion_geo",pt:{objetivo:150000,actual:100000},df:"benchmark_amx",pf:{pa:80,pb:80}},
    {id:"AF.EC.03",n:"FTTH Red Neutra",m:"HHP",prio:"CRECIMIENTO",P_base:8000000,dt:"expansion_geo",pt:{objetivo:200000,actual:100000},df:"mercado",pf:{pa:40,pb:40}},
  ]},
  
  {macro:"Expansión Capacidad Fijo",categoria:"Red",tipo:"Crecimiento",P_base:8500000,proyectos:[
    {id:"AF.EK.01",n:"Puertos OLT y Agregación",m:"puerto",prio:"CRECIMIENTO",P_base:5000000,dt:"activaciones",pt:{brutas:200000,churn:40000},df:"benchmark_amx",pf:{pa:25,pb:25}},
    {id:"AF.EK.02",n:"CPE Nuevos Clientes",m:"cliente",prio:"CRECIMIENTO",P_base:3500000,dt:"activaciones",pt:{brutas:100000,churn:20000},df:"benchmark_amx",pf:{pa:45,pb:45}},
  ]},
  
  {macro:"Obligaciones Regulatorias Fijo",categoria:"Red",tipo:"Regulatorio",P_base:5000000,proyectos:[
    {id:"AF.OR.01",n:"Servicio Universal Fijo",m:"sitio",prio:"REGULATORIO",P_base:5000000,dt:"cobertura_crc",pt:{comprometido:200,ejecutado:50},df:"benchmark_crc",pf:{pa:33333,pb:33333}},
  ]},
  
  {macro:"Modernización Tecnológica Fijo",categoria:"Red",tipo:"Modernización",P_base:18000000,proyectos:[
    {id:"AF.MT.01",n:"Migración HFC a Fibra",m:"suscriptor",prio:"EBITDA",P_base:12000000,dt:"migraciones",pt:{base_migrable:200000,tasa:0.5},df:"benchmark_amx",pf:{pa:120,pb:120}},
    {id:"AF.MT.02",n:"Upgrade GPON a XGS-PON",m:"OLT",prio:"EBITDA",P_base:4000000,dt:"obsolescencia",pt:{total_activos:500,pct_eol:0.3,meta_pct:0.1},df:"benchmark_amx",pf:{pa:8000,pb:8000}},
    {id:"AF.MT.03",n:"Refresh CPE y ONT",m:"CPE",prio:"PMO",P_base:2000000,dt:"obsolescencia",pt:{total_activos:200000,pct_eol:0.15,meta_pct:0.05},df:"benchmark_amx",pf:{pa:35,pb:35}},
  ]},
  
  {macro:"Operación Acceso Fijo",categoria:"Red",tipo:"Operación",P_base:8000000,proyectos:[
    {id:"AF.OP.01",n:"Mantenimiento Planta Externa",m:"km",prio:"PMO",P_base:5000000,dt:"mantenimiento",pt:{activos:6250,frecuencia:1},df:"historico",pf:{pa:800,pb:800}},
    {id:"AF.OP.02",n:"Inventario y Reposición",m:"activo",prio:"PMO",P_base:3000000,dt:"mantenimiento",pt:{activos:100000,frecuencia:1},df:"historico",pf:{pa:30,pb:30}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: RED - ACCESO MÓVIL
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Expansión Cobertura Móvil",categoria:"Red",tipo:"Crecimiento",P_base:52000000,proyectos:[
    {id:"AM.EC.01",n:"Nuevos Sitios 4G",m:"sitio",prio:"CRECIMIENTO",P_base:19500000,dt:"expansion_geo",pt:{objetivo:500,actual:200},df:"benchmark_amx",pf:{pa:65000,pb:65000}},
    {id:"AM.EC.02",n:"Nuevos Sitios 5G",m:"sitio",prio:"CRECIMIENTO",P_base:28500000,dt:"expansion_geo",pt:{objetivo:400,actual:100},df:"benchmark_amx",pf:{pa:95000,pb:95000}},
    {id:"AM.EC.03",n:"Small Cells y DAS",m:"nodo",prio:"CRECIMIENTO",P_base:4000000,dt:"expansion_geo",pt:{objetivo:200,actual:40},df:"benchmark_amx",pf:{pa:25000,pb:25000}},
  ]},
  
  {macro:"Expansión Capacidad Móvil",categoria:"Red",tipo:"Crecimiento",P_base:21000000,proyectos:[
    {id:"AM.EK.01",n:"Carriers Adicionales 4G",m:"carrier",prio:"CRECIMIENTO",P_base:12000000,dt:"congestion",pt:{total_activos:5000,pct_congestion:0.20,meta_pct:0.05},df:"benchmark_amx",pf:{pa:12000,pb:12000}},
    {id:"AM.EK.02",n:"Carriers Adicionales 5G",m:"carrier",prio:"CRECIMIENTO",P_base:9000000,dt:"congestion",pt:{total_activos:1000,pct_congestion:0.25,meta_pct:0.05},df:"benchmark_amx",pf:{pa:18000,pb:18000}},
  ]},
  
  {macro:"Obligaciones Regulatorias Móvil",categoria:"Red",tipo:"Regulatorio",P_base:58000000,proyectos:[
    {id:"AM.OR.01",n:"Cobertura Mandatoria 4G",m:"sitio",prio:"REGULATORIO",P_base:40000000,dt:"cobertura_crc",pt:{comprometido:100,ejecutado:20},df:"benchmark_crc",pf:{pa:500000,pb:500000}},
    {id:"AM.OR.02",n:"Cobertura Mandatoria 5G",m:"sitio",prio:"REGULATORIO",P_base:18000000,dt:"cobertura_crc",pt:{comprometido:50,ejecutado:10},df:"benchmark_crc",pf:{pa:450000,pb:450000}},
  ]},
  
  {macro:"Modernización Tecnológica Móvil",categoria:"Red",tipo:"Modernización",P_base:85000000,proyectos:[
    {id:"AM.MT.01",n:"Upgrade 4G a 5G",m:"sitio",prio:"EBITDA",P_base:45000000,dt:"migraciones",pt:{base_migrable:2000,tasa:0.5},df:"benchmark_amx",pf:{pa:45000,pb:45000}},
    {id:"AM.MT.02",n:"Refresh Equipos 4G EoL",m:"sitio",prio:"PMO",P_base:15000000,dt:"obsolescencia",pt:{total_activos:3000,pct_eol:0.20,meta_pct:0.05},df:"benchmark_amx",pf:{pa:25000,pb:25000}},
    {id:"AM.MT.03",n:"Refresh Equipos 5G EoL",m:"sitio",prio:"PMO",P_base:7000000,dt:"obsolescencia",pt:{total_activos:500,pct_eol:0.20,meta_pct:0.05},df:"benchmark_amx",pf:{pa:35000,pb:35000}},
    {id:"AM.MT.04",n:"Swap de Vendor RAN",m:"sitio",prio:"EBITDA",P_base:12000000,dt:"obsolescencia",pt:{total_activos:500,pct_eol:0.40,meta_pct:0.05},df:"mercado",pf:{pa:60000,pb:60000}},
    {id:"AM.MT.05",n:"Reubicación de Sitios",m:"sitio",prio:"PMO",P_base:6000000,dt:"mantenimiento",pt:{activos:120,frecuencia:1},df:"mercado",pf:{pa:50000,pb:50000}},
  ]},
  
  {macro:"Operación Acceso Móvil",categoria:"Red",tipo:"Operación",P_base:35000000,proyectos:[
    {id:"AM.OP.01",n:"Mantenimiento Sitios RAN",m:"sitio",prio:"PMO",P_base:21000000,dt:"mantenimiento",pt:{activos:7000,frecuencia:1},df:"historico",pf:{pa:3000,pb:3000}},
    {id:"AM.OP.02",n:"Energía y Climatización",m:"sitio",prio:"PMO",P_base:8000000,dt:"mantenimiento",pt:{activos:4000,frecuencia:1},df:"historico",pf:{pa:2000,pb:2000}},
    {id:"AM.OP.03",n:"Torres e Infraestructura Física",m:"sitio",prio:"PMO",P_base:6000000,dt:"mantenimiento",pt:{activos:3000,frecuencia:1},df:"historico",pf:{pa:2000,pb:2000}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: RED - TRANSPORTE Y CORE
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Expansión Backbone",categoria:"Red",tipo:"Crecimiento",P_base:32000000,proyectos:[
    {id:"TC.EB.01",n:"Fibra Nacional Long-haul",m:"km",prio:"CRECIMIENTO",P_base:15000000,dt:"crecimiento_trafico",pt:{capacidad_actual:1000,crecimiento_pct:0.30},df:"benchmark_amx",pf:{pa:25000,pb:25000}},
    {id:"TC.EB.02",n:"Anillos Metropolitanos",m:"anillo",prio:"CRECIMIENTO",P_base:12000000,dt:"expansion_geo",pt:{objetivo:10,actual:4},df:"benchmark_amx",pf:{pa:2000000,pb:2000000}},
    {id:"TC.EB.03",n:"Capacidad DWDM Lambdas",m:"lambda",prio:"CRECIMIENTO",P_base:5000000,dt:"crecimiento_trafico",pt:{capacidad_actual:200,crecimiento_pct:0.50},df:"benchmark_amx",pf:{pa:50000,pb:50000}},
  ]},
  
  {macro:"Infraestructura 5G Core",categoria:"Red",tipo:"Crecimiento",P_base:35000000,proyectos:[
    {id:"TC.I5.01",n:"Transporte 5G Xhaul",m:"sitio",prio:"CRECIMIENTO",P_base:8000000,dt:"expansion_geo",pt:{objetivo:800,actual:200},df:"benchmark_amx",pf:{pa:12000,pb:12000}},
    {id:"TC.I5.02",n:"Core 5G Standalone",m:"programa",prio:"CRECIMIENTO",P_base:15000000,dt:"crecimiento_trafico",pt:{capacidad_actual:1,crecimiento_pct:1},df:"benchmark_amx",pf:{pa:15000000,pb:15000000}},
    {id:"TC.I5.03",n:"Telco Cloud y vRAN",m:"nodo",prio:"CRECIMIENTO",P_base:12000000,dt:"crecimiento_trafico",pt:{capacidad_actual:10,crecimiento_pct:0.50},df:"benchmark_amx",pf:{pa:500000,pb:500000}},
  ]},
  
  {macro:"Conectividad Internacional",categoria:"Red",tipo:"Crecimiento",P_base:15000000,proyectos:[
    {id:"TC.CI.01",n:"Capacidad IRU Cables Submarinos",m:"Gbps",prio:"CRECIMIENTO",P_base:9000000,dt:"crecimiento_trafico",pt:{capacidad_actual:500,crecimiento_pct:0.40},df:"mercado",pf:{pa:45000,pb:45000}},
    {id:"TC.CI.02",n:"Participación Cables Submarinos",m:"proyecto",prio:"CRECIMIENTO",P_base:6000000,dt:"crecimiento_trafico",pt:{capacidad_actual:1,crecimiento_pct:1},df:"mercado",pf:{pa:6000000,pb:6000000}},
  ]},
  
  {macro:"Modernización Core y TX",categoria:"Red",tipo:"Modernización",P_base:8000000,proyectos:[
    {id:"TC.MC.01",n:"Refresh Routers Core IP/MPLS",m:"nodo",prio:"EBITDA",P_base:8000000,dt:"obsolescencia",pt:{total_activos:40,pct_eol:0.40,meta_pct:0.05},df:"benchmark_amx",pf:{pa:500000,pb:500000}},
  ]},
  
  {macro:"Operación TX y Core",categoria:"Red",tipo:"Operación",P_base:12000000,proyectos:[
    {id:"TC.OP.01",n:"IRUs y Compartición Regulatoria",m:"contrato",prio:"PMO",P_base:7000000,dt:"mantenimiento",pt:{activos:100,frecuencia:1},df:"historico",pf:{pa:70000,pb:70000}},
    {id:"TC.OP.02",n:"Mantenimiento TX y Core",m:"equipo",prio:"PMO",P_base:5000000,dt:"mantenimiento",pt:{activos:500,frecuencia:1},df:"historico",pf:{pa:10000,pb:10000}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: RED - DATA CENTER
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Expansión Capacidad DC",categoria:"Red",tipo:"Crecimiento",P_base:25000000,proyectos:[
    {id:"DC.EC.01",n:"Whitespace y Racks",m:"kW",prio:"CRECIMIENTO",P_base:10000000,dt:"crecimiento_trafico",pt:{capacidad_actual:500,crecimiento_pct:0.30},df:"mercado",pf:{pa:15000,pb:15000}},
    {id:"DC.EC.02",n:"Energía y Refrigeración",m:"MW",prio:"CRECIMIENTO",P_base:12000000,dt:"crecimiento_trafico",pt:{capacidad_actual:2,crecimiento_pct:0.50},df:"mercado",pf:{pa:8000000,pb:8000000}},
    {id:"DC.EC.03",n:"Edge Computing Propio",m:"nodo",prio:"CRECIMIENTO",P_base:3000000,dt:"expansion_geo",pt:{objetivo:20,actual:5},df:"benchmark_amx",pf:{pa:200000,pb:200000}},
  ]},
  
  {macro:"Operación Data Center",categoria:"Red",tipo:"Operación",P_base:4500000,proyectos:[
    {id:"DC.OP.01",n:"Mantenimiento y Refresh Infra DC",m:"m2",prio:"PMO",P_base:4500000,dt:"mantenimiento",pt:{activos:30000,frecuencia:1},df:"historico",pf:{pa:150,pb:150}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: ENTERPRISE Y B2B
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Soluciones Conectividad Enterprise",categoria:"Enterprise y B2B",tipo:"Crecimiento",P_base:18000000,proyectos:[
    {id:"EB.SC.01",n:"Redes Dedicadas y SD-WAN",m:"contrato",prio:"CRECIMIENTO",P_base:10000000,dt:"pipeline_b2b",pt:{pipeline:200,win_rate:0.50},df:"mercado",pf:{pa:100000,pb:100000}},
    {id:"EB.SC.02",n:"Conectividad Multi-Cloud",m:"POP",prio:"CRECIMIENTO",P_base:5000000,dt:"pipeline_b2b",pt:{pipeline:100,win_rate:0.50},df:"mercado",pf:{pa:50000,pb:50000}},
    {id:"EB.SC.03",n:"5G Privado y Network Slicing",m:"despliegue",prio:"CRECIMIENTO",P_base:3000000,dt:"pipeline_b2b",pt:{pipeline:30,win_rate:0.40},df:"mercado",pf:{pa:250000,pb:250000}},
  ]},
  
  {macro:"Servicios Data Center B2B",categoria:"Enterprise y B2B",tipo:"Crecimiento",P_base:15000000,proyectos:[
    {id:"EB.SD.01",n:"Colocation y Housing",m:"rack",prio:"CRECIMIENTO",P_base:6000000,dt:"pipeline_b2b",pt:{pipeline:750,win_rate:0.50},df:"mercado",pf:{pa:8000,pb:8000}},
    {id:"EB.SD.02",n:"Cloud Privado y Managed Services",m:"contrato",prio:"CRECIMIENTO",P_base:5000000,dt:"pipeline_b2b",pt:{pipeline:50,win_rate:0.50},df:"mercado",pf:{pa:100000,pb:100000}},
    {id:"EB.SD.03",n:"Edge Computing Empresarial",m:"nodo",prio:"CRECIMIENTO",P_base:4000000,dt:"pipeline_b2b",pt:{pipeline:40,win_rate:0.40},df:"mercado",pf:{pa:150000,pb:150000}},
  ]},
  
  {macro:"Gobierno y Verticales",categoria:"Enterprise y B2B",tipo:"Crecimiento",P_base:12000000,proyectos:[
    {id:"EB.GV.01",n:"Conectividad Social MINTIC",m:"proyecto",prio:"CRECIMIENTO",P_base:8000000,dt:"pipeline_b2b",pt:{pipeline:10,win_rate:0.50},df:"mercado",pf:{pa:800000,pb:800000}},
    {id:"EB.GV.02",n:"Verticales Estratégicos",m:"proyecto",prio:"CRECIMIENTO",P_base:4000000,dt:"pipeline_b2b",pt:{pipeline:20,win_rate:0.40},df:"mercado",pf:{pa:500000,pb:500000}},
  ]},
  
  {macro:"Capacidad Mayorista",categoria:"Enterprise y B2B",tipo:"Crecimiento",P_base:5000000,proyectos:[
    {id:"EB.WS.01",n:"Servicios Wholesale",m:"Gbps",prio:"CRECIMIENTO",P_base:5000000,dt:"crecimiento_trafico",pt:{capacidad_actual:200,crecimiento_pct:0.50},df:"mercado",pf:{pa:50000,pb:50000}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: TECNOLOGÍA
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Transformación BSS",categoria:"Tecnología",tipo:"Transformación",P_base:45000000,proyectos:[
    {id:"TI.TB.01",n:"Modernización CRM",m:"programa",prio:"TRANSFORMACIÓN",P_base:25000000,dt:"migraciones",pt:{base_migrable:1,tasa:1},df:"mercado",pf:{pa:25000000,pb:25000000}},
    {id:"TI.TB.02",n:"Billing Convergente",m:"programa",prio:"TRANSFORMACIÓN",P_base:20000000,dt:"migraciones",pt:{base_migrable:1,tasa:1},df:"mercado",pf:{pa:20000000,pb:20000000}},
  ]},
  
  {macro:"Transformación OSS",categoria:"Tecnología",tipo:"Transformación",P_base:10000000,proyectos:[
    {id:"TI.TO.01",n:"Orquestación y Automatización Red",m:"programa",prio:"TRANSFORMACIÓN",P_base:10000000,dt:"migraciones",pt:{base_migrable:1,tasa:1},df:"mercado",pf:{pa:10000000,pb:10000000}},
  ]},
  
  {macro:"Plataformas Digitales",categoria:"Tecnología",tipo:"Transformación",P_base:13000000,proyectos:[
    {id:"TI.PD.01",n:"Canales Digitales App y Web",m:"año",prio:"TRANSFORMACIÓN",P_base:5000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"mercado",pf:{pa:5000000,pb:5000000}},
    {id:"TI.PD.02",n:"Data y Analytics",m:"programa",prio:"TRANSFORMACIÓN",P_base:8000000,dt:"migraciones",pt:{base_migrable:1,tasa:1},df:"mercado",pf:{pa:8000000,pb:8000000}},
  ]},
  
  {macro:"Sostenimiento IT",categoria:"Tecnología",tipo:"Operación",P_base:22000000,proyectos:[
    {id:"TI.SI.01",n:"Infraestructura y Licenciamiento",m:"año",prio:"PMO",P_base:12000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"historico",pf:{pa:12000000,pb:12000000}},
    {id:"TI.SI.02",n:"Ciberseguridad",m:"año",prio:"PMO",P_base:5000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"mercado",pf:{pa:5000000,pb:5000000}},
    {id:"TI.SI.03",n:"Mantenimiento Aplicativo",m:"año",prio:"PMO",P_base:5000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"historico",pf:{pa:5000000,pb:5000000}},
  ]},
  
  {macro:"Evolución Plataformas",categoria:"Tecnología",tipo:"Evolución",P_base:6000000,proyectos:[
    {id:"TI.EP.01",n:"Mejoras BSS CRM y Billing",m:"año",prio:"EVOLUCIÓN",P_base:4000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"mercado",pf:{pa:4000000,pb:4000000}},
    {id:"TI.EP.02",n:"Mejoras Canales Digitales",m:"año",prio:"EVOLUCIÓN",P_base:2000000,dt:"mantenimiento",pt:{activos:1,frecuencia:1},df:"mercado",pf:{pa:2000000,pb:2000000}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: INVERSIÓN CLIENTES
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Adquisición Consumidor",categoria:"Inversión Clientes",tipo:"Comercial",P_base:75000000,proyectos:[
    {id:"IC.AC.01",n:"Activaciones Fijo Hogar",m:"activación",prio:"CRECIMIENTO",P_base:32000000,dt:"activaciones",pt:{brutas:400000,churn:100000},df:"benchmark_amx",pf:{pa:80,pb:80}},
    {id:"IC.AC.02",n:"Activación y Subsidio Móvil",m:"alta",prio:"CRECIMIENTO",P_base:43000000,dt:"activaciones",pt:{brutas:500000,churn:70000},df:"benchmark_amx",pf:{pa:100,pb:100}},
  ]},
  
  {macro:"Retención Consumidor",categoria:"Inversión Clientes",tipo:"Comercial",P_base:28000000,proyectos:[
    {id:"IC.RC.01",n:"Programas Retención y Lealtad",m:"save",prio:"RETENCIÓN",P_base:20000000,dt:"migraciones",pt:{base_migrable:400000,tasa:0.25},df:"benchmark_amx",pf:{pa:50,pb:50}},
    {id:"IC.RC.02",n:"Migraciones ARPU-Accretive",m:"upgrade",prio:"CRECIMIENTO",P_base:8000000,dt:"migraciones",pt:{base_migrable:400000,tasa:0.20},df:"benchmark_amx",pf:{pa:30,pb:30}},
  ]},
  
  {macro:"Inversión B2B Clientes",categoria:"Inversión Clientes",tipo:"Comercial",P_base:12000000,proyectos:[
    {id:"IC.IB.01",n:"Adquisición Enterprise y PyME",m:"deal",prio:"CRECIMIENTO",P_base:8000000,dt:"pipeline_b2b",pt:{pipeline:200,win_rate:0.50},df:"mercado",pf:{pa:80000,pb:80000}},
    {id:"IC.IB.02",n:"Renovación Contratos B2B",m:"contrato",prio:"RETENCIÓN",P_base:4000000,dt:"pipeline_b2b",pt:{pipeline:300,win_rate:0.70},df:"mercado",pf:{pa:19000,pb:19000}},
  ]},
  
  // ═══════════════════════════════════════════════════════════════════════════
  // CATEGORÍA: CORPORATIVO
  // ═══════════════════════════════════════════════════════════════════════════
  
  {macro:"Soporte Corporativo",categoria:"Corporativo",tipo:"Operación",P_base:6000000,proyectos:[
    {id:"CO.SC.01",n:"Instalaciones y Sedes",m:"m2",prio:"PMO",P_base:4000000,dt:"mantenimiento",pt:{activos:40000,frecuencia:1},df:"historico",pf:{pa:100,pb:100}},
    {id:"CO.SC.02",n:"Flota Vehicular",m:"vehículo",prio:"PMO",P_base:2000000,dt:"mantenimiento",pt:{activos:72,frecuencia:1},df:"mercado",pf:{pa:28000,pb:28000}},
  ]},
];

/* ── Helpers de formato ─────────────────────────────────────────────────── */
const fu = v=>{
  if(v===undefined||v===null||isNaN(v)) return "—";
  const abs=Math.abs(v);
  if(abs>=1e9)  return `$${(v/1e9).toFixed(2)}B`;
  if(abs>=1e6)  return `$${(v/1e6).toFixed(2)}M`;
  if(abs>=1e3)  return `$${(v/1e3).toFixed(1)}K`;
  return `$${v.toFixed(0)}`;
};
const fn = v=>{
  if(v===undefined||v===null||isNaN(v)) return "—";
  if(Math.abs(v)>=1e6) return `${(v/1e6).toFixed(1)}M`;
  if(Math.abs(v)>=1e3) return `${(v/1e3).toFixed(1)}K`;
  return String(Math.round(v));
};
const dp  = (a,b)=> b>0 ? (a-b)/b*100 : 0;
const sg  = v => v>=0 ? "+" : "";
const cc  = v => v<-0.5 ? T.green : v>0.5 ? T.red : T.inkMid;

/* ══════════════════════════════════════════════════════════════════════════
   CATÁLOGOS DE DRIVERS POR NIVEL
   N1 Macroproyecto → sin driver editable (es agregación)
   N2 Proyecto      → Driver de DEMANDA (¿cuántos necesito?)
   N3 Sub-Proyecto  → Driver de VOLUMEN (¿cuántas unidades físicas?)
   N4 Componente    → Driver de PRECIO  (¿a qué costo unitario?)
══════════════════════════════════════════════════════════════════════════ */

/* N2 — Drivers de Demanda: responden "¿por qué necesito invertir?" */
const DRV_N2 = {
  obsolescencia:       {
    label:"Obsolescencia EoL/EoS", icon:"⚠️", color:"#D97706", tag:"Demanda",
    desc:"Activos fuera de soporte que deben reemplazarse",
    params:[{k:"total_activos",l:"Activos en parque",p:false},{k:"pct_eol",l:"% EoL actual",p:true},{k:"meta_pct",l:"Meta EoL ≤",p:true}],
    calcQ:d=>Math.round((d.total_activos||0)*Math.max(0,(d.pct_eol||0)-(d.meta_pct||0))),
  },
  congestion:          {
    label:"Congestión PRB >80%", icon:"📶", color:"#E8182A", tag:"Demanda",
    desc:"Sitios congestionados que superan el umbral de calidad",
    params:[{k:"total_activos",l:"Total sitios red",p:false},{k:"pct_congestion",l:"% congestionados",p:true},{k:"meta_pct",l:"Meta ≤",p:true}],
    calcQ:d=>Math.round((d.total_activos||0)*Math.max(0,(d.pct_congestion||0)-(d.meta_pct||0))),
  },
  cobertura_crc:       {
    label:"Obligación Regulatoria CRC", icon:"⚖️", color:"#E8182A", tag:"Regulatorio",
    desc:"Compromisos legales de cobertura pendientes de ejecutar",
    params:[{k:"comprometido",l:"Comprometido CRC",p:false},{k:"ejecutado",l:"Ya ejecutado",p:false}],
    calcQ:d=>Math.max(0,(d.comprometido||0)-(d.ejecutado||0)),
  },
  activaciones:        {
    label:"Activaciones Netas Plan", icon:"📈", color:"#059669", tag:"Demanda",
    desc:"Nuevos clientes proyectados menos churn esperado",
    params:[{k:"brutas",l:"Activaciones brutas",p:false},{k:"churn",l:"Churn estimado",p:false}],
    calcQ:d=>Math.max(0,(d.brutas||0)-(d.churn||0)),
  },
  migraciones:         {
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
  expansion_geo:       {
    label:"Expansión de Cobertura", icon:"🗺️", color:"#059669", tag:"Demanda",
    desc:"Gap entre cobertura objetivo y cobertura actual",
    params:[{k:"objetivo",l:"Objetivo (HP/km²)",p:false},{k:"actual",l:"Actual cubierto",p:false}],
    calcQ:d=>Math.max(0,(d.objetivo||0)-(d.actual||0)),
  },
  pipeline_b2b:        {
    label:"Pipeline B2B / Empresas", icon:"🏢", color:"#C8941C", tag:"Demanda",
    desc:"Oportunidades calificadas ponderadas por win-rate",
    params:[{k:"pipeline",l:"Pipeline calificado",p:false},{k:"win_rate",l:"Win-rate histórico",p:true}],
    calcQ:d=>Math.round((d.pipeline||0)*(d.win_rate||0)),
  },
  mantenimiento:       {
    label:"Plan Mantenimiento PMO", icon:"🔧", color:"#6B6860", tag:"Operación",
    desc:"Intervenciones de mantenimiento programadas al año",
    params:[{k:"activos",l:"Activos en parque",p:false},{k:"frecuencia",l:"Intervenciones/año",p:false}],
    calcQ:d=>Math.round((d.activos||0)*(d.frecuencia||1)),
  },
};

/* N3 — Drivers de Volumen: responden "¿cuántas unidades físicas necesito?" */
const DRV_N3 = {
  unidades_fisicas:    {
    label:"Unidades Físicas", icon:"📦", color:"#2563EB", tag:"Volumen",
    desc:"Cantidad de unidades físicas a instalar o reemplazar",
    params:[{k:"cantidad_total",l:"Cantidad requerida",p:false},{k:"ya_instalado",l:"Ya instalado",p:false}],
    calcQ:d=>Math.max(0,(d.cantidad_total||0)-(d.ya_instalado||0)),
  },
  sitios_intervencion: {
    label:"Sitios a Intervenir", icon:"📍", color:"#D97706", tag:"Volumen",
    desc:"Sitios de red que requieren intervención física",
    params:[{k:"sitios_scope",l:"Sitios en scope",p:false},{k:"pct_requiere",l:"% que requiere trabajo",p:true}],
    calcQ:d=>Math.round((d.sitios_scope||0)*(d.pct_requiere||0)),
  },
  capacidad_delta:     {
    label:"Delta de Capacidad", icon:"⚡", color:"#C8941C", tag:"Volumen",
    desc:"Diferencia entre capacidad requerida y disponible",
    params:[{k:"cap_requerida",l:"Capacidad requerida",p:false},{k:"cap_disponible",l:"Disponible actual",p:false}],
    calcQ:d=>Math.max(0,(d.cap_requerida||0)-(d.cap_disponible||0)),
  },
  cobertura_fisica:    {
    label:"Cobertura Física (m²/km)", icon:"📐", color:"#059669", tag:"Volumen",
    desc:"Área o longitud física a cubrir con infraestructura",
    params:[{k:"area_objetivo",l:"Área/longitud objetivo",p:false},{k:"area_actual",l:"Ya cubierta",p:false}],
    calcQ:d=>Math.max(0,(d.area_objetivo||0)-(d.area_actual||0)),
  },
  licencias_sw:        {
    label:"Licencias de Software", icon:"💾", color:"#7C3AED", tag:"Volumen",
    desc:"Número de licencias o suscripciones a adquirir",
    params:[{k:"usuarios_obj",l:"Usuarios objetivo",p:false},{k:"ya_licenciados",l:"Ya licenciados",p:false}],
    calcQ:d=>Math.max(0,(d.usuarios_obj||0)-(d.ya_licenciados||0)),
  },
  hh_trabajo:          {
    label:"Horas-Hombre de Trabajo", icon:"👷", color:"#059669", tag:"Volumen",
    desc:"Esfuerzo total en horas-hombre por actividades del sub-proyecto",
    params:[{k:"actividades",l:"Actividades totales",p:false},{k:"hh_x_actividad",l:"HH por actividad",p:false}],
    calcQ:d=>Math.round((d.actividades||0)*(d.hh_x_actividad||0)),
  },
  contratos_servicio:  {
    label:"Contratos / SLAs", icon:"📋", color:"#6B6860", tag:"Volumen",
    desc:"Número de contratos o SLAs de servicio a suscribir",
    params:[{k:"contratos_total",l:"Contratos en scope",p:false},{k:"pct_renovacion",l:"% a renovar/new",p:true}],
    calcQ:d=>Math.round((d.contratos_total||0)*(d.pct_renovacion||0)),
  },
  sprints_impl:        {
    label:"Sprints de Implementación", icon:"🏃", color:"#7C3AED", tag:"Volumen",
    desc:"Ciclos de desarrollo/implementación ágil requeridos",
    params:[{k:"funcionalidades",l:"Funcionalidades en backlog",p:false},{k:"func_x_sprint",l:"Funcionalidades/sprint",p:false}],
    calcQ:d=>Math.ceil((d.funcionalidades||0)/Math.max(1,(d.func_x_sprint||1))),
  },
};

/* N2 — Drivers Financieros: precio estratégico/contractual del proyecto */
const DRV_FIN_N2 = {
  benchmark_amx:{label:"Benchmark Grupo AMX",        color:"#2563EB", desc:"Precio de referencia del grupo AMX para la misma categoría de inversión en LatAm"},
  historico:    {label:"Histórico Contractual Claro", color:"#6B6860", desc:"Promedio ponderado de contratos ejecutados por Claro en los últimos 3 años"},
  benchmark_crc:{label:"Valor Regulatorio CRC",       color:"#E8182A", desc:"Precio máximo o referencia establecido por resolución regulatoria CRC vigente"},
  ila:          {label:"Indexación IPC + ILA Telco",  color:"#D97706", desc:"Precio del período anterior escalado con IPC + Índice de Laboral y de Activos del sector"},
  mercado:      {label:"Precio de Mercado",           color:"#C8941C", desc:"Mejor precio obtenido en proceso competitivo abierto o cotización spot"},
  cotizacion:   {label:"Cotización Directa",          color:"#059669", desc:"Precio ofertado formalmente por el proveedor en el proceso de selección actual"},
};

/* N3 — Drivers Financieros: precio por sub-proyecto/componente de trabajo */
const DRV_FIN_N3 = {
  cotizacion:   {label:"Cotización Directa",          color:"#059669", desc:"Precio ofertado formalmente por el proveedor en el proceso de selección actual"},
  contrato_marco:{label:"Contrato Marco Vigente",     color:"#2563EB", desc:"Precio pactado en el contrato marco activo con el proveedor principal"},
  benchmark_amx:{label:"Benchmark AMX por Categoría",color:"#7C3AED", desc:"Precio de referencia AMX segmentado por categoría de sub-proyecto (HW, civil, SW...)"},
  mercado:      {label:"Precio de Mercado / Spot",    color:"#C8941C", desc:"Mejor precio obtenido en proceso competitivo abierto o cotización spot"},
};

/* N4 — Drivers Financieros: precio unitario del ítem de costo */
const DRV_FIN_N4 = {
  precio_lista:  {label:"Precio de Lista / Catálogo", color:"#059669", desc:"Precio publicado en lista oficial del proveedor o catálogo de producto"},
  tarifa_marco:  {label:"Tarifa Contrato Marco",      color:"#2563EB", desc:"Tarifa unitaria fija pactada en el contrato marco vigente"},
  precio_hist:   {label:"Precio Histórico Unitario",  color:"#6B6860", desc:"Precio unitario promedio pagado en los últimos 12 meses por el mismo ítem"},
  cotizacion_u:  {label:"Cotización Unitaria",        color:"#059669", desc:"Precio por unidad según la oferta formal más reciente del proveedor"},
  arancel_reg:   {label:"Arancel / Tasa Regulatoria", color:"#E8182A", desc:"Tarifa fija establecida por CRC, ANLA u otro ente regulador"},
  mercado_spot:  {label:"Precio Spot de Mercado",     color:"#C8941C", desc:"Precio unitario vigente en el mercado abierto o subasta puntual"},
};

const DRV_FIN = DRV_FIN_N2; // alias legacy

// Alias para compatibilidad con código legacy

/* Mapeo sugerido: driver N2 → driver N3 más coherente */
const SUGGEST_N3_DT = {
  obsolescencia:"unidades_fisicas", congestion:"sitios_intervencion",
  cobertura_crc:"unidades_fisicas", activaciones:"unidades_fisicas",
  migraciones:"licencias_sw",       crecimiento_trafico:"capacidad_delta",
  expansion_geo:"cobertura_fisica", pipeline_b2b:"contratos_servicio",
  mantenimiento:"sitios_intervencion",
};
const SUGGEST_N3_DF = {
  benchmark_amx:"contrato_marco", historico:"contrato_marco",
  benchmark_crc:"cotizacion",     ila:"cotizacion",
};

/* Defaults de parámetros por driver */
const DEFS_N2 = {
  obsolescencia:{total_activos:500,pct_eol:.30,meta_pct:.20},
  congestion:{total_activos:500,pct_congestion:.12,meta_pct:.08},
  cobertura_crc:{comprometido:100,ejecutado:0},
  activaciones:{brutas:50000,churn:10000},
  migraciones:{base_migrable:5000,tasa:.20},
  crecimiento_trafico:{capacidad_actual:200,crecimiento_pct:.25},
  expansion_geo:{objetivo:10000,actual:5000},
  pipeline_b2b:{pipeline:200,win_rate:.50},
  mantenimiento:{activos:500,frecuencia:1},
};
const DEFS_N3 = {
  unidades_fisicas:{cantidad_total:100,ya_instalado:0},
  sitios_intervencion:{sitios_scope:200,pct_requiere:.80},
  capacidad_delta:{cap_requerida:200,cap_disponible:50},
  cobertura_fisica:{area_objetivo:5000,area_actual:0},
  licencias_sw:{usuarios_obj:10000,ya_licenciados:0},
  hh_trabajo:{actividades:500,hh_x_actividad:8},
  contratos_servicio:{contratos_total:50,pct_renovacion:1.0},
  sprints_impl:{funcionalidades:40,func_x_sprint:5},
};

/* calcQ universal */
const calcQ = (k,d,nivel=2) => {
  const cat = nivel===3 ? DRV_N3 : DRV_N2;
  return cat[k]?.calcQ(d) ?? 0;
};

const calcCap = (p, ov) => {
  // Sin override: CAPEX DVB = Base (punto de partida)
  if(!ov) return p.P_base;
  // Con override guardado explícitamente (_capex viene del botón Guardar)
  if(ov._capex != null) return ov._capex;
  // Con override parcial (solo cambios de driver sin guardar árbol)
  const dk = ov.dt || p.dt;
  const pt = ov.pt ? {...p.pt,...ov.pt} : {...p.pt};
  const pf = ov.pf ? {...p.pf,...ov.pf} : {...p.pf};
  const Q  = calcQ(dk, pt, 2);
  const P  = pf.pb;
  if(Q>0 && P>0) return Q*P;
  return p.P_base;
};

/* back-calc: editar Q → recalcular parámetro principal */
const backCalcPt = (dk,pt,nQ,nivel=2) => {
  const np = {...pt};
  if(dk==="obsolescencia")        np.pct_eol=(pt.meta_pct||0)+nQ/Math.max(1,pt.total_activos);
  else if(dk==="congestion")      np.pct_congestion=(pt.meta_pct||0)+nQ/Math.max(1,pt.total_activos);
  else if(dk==="cobertura_crc")   np.comprometido=(pt.ejecutado||0)+nQ;
  else if(dk==="activaciones")    np.brutas=(pt.churn||0)+nQ;
  else if(dk==="migraciones")     np.tasa=nQ/Math.max(1,pt.base_migrable);
  else if(dk==="crecimiento_trafico") np.crecimiento_pct=nQ/Math.max(1,pt.capacidad_actual);
  else if(dk==="expansion_geo")   np.objetivo=(pt.actual||0)+nQ;
  else if(dk==="pipeline_b2b")    np.win_rate=nQ/Math.max(1,pt.pipeline);
  else if(dk==="mantenimiento")   np.frecuencia=nQ/Math.max(1,pt.activos);
  else if(dk==="unidades_fisicas")np.cantidad_total=(pt.ya_instalado||0)+nQ;
  else if(dk==="sitios_intervencion") np.pct_requiere=nQ/Math.max(1,pt.sitios_scope);
  else if(dk==="capacidad_delta") np.cap_requerida=(pt.cap_disponible||0)+nQ;
  else if(dk==="cobertura_fisica")np.area_objetivo=(pt.area_actual||0)+nQ;
  else if(dk==="licencias_sw")    np.usuarios_obj=(pt.ya_licenciados||0)+nQ;
  else if(dk==="hh_trabajo")      np.hh_x_actividad=nQ/Math.max(1,pt.actividades);
  else if(dk==="contratos_servicio") np.pct_renovacion=nQ/Math.max(1,pt.contratos_total);
  else if(dk==="sprints_impl")    np.funcionalidades=nQ*Math.max(1,pt.func_x_sprint);
  return np;
};

/* Colores y metadatos de cada nivel */
const LVL = {
  1:{c:"#E8182A", bg:"#FFF1F1", bdr:"#FDD8DA", label:"Macroproyecto", short:"M1"},
  2:{c:"#2563EB", bg:"#EFF6FF", bdr:"#BFDBFE", label:"Proyecto",      short:"N2"},
  3:{c:"#7C3AED", bg:"#F5F3FF", bdr:"#DDD6FE", label:"Sub-Proyecto",  short:"N3"},
  4:{c:"#059669", bg:"#ECFDF5", bdr:"#A7F3D0", label:"Componente",    short:"N4"},
};

/* ── Utilidades de UI ───────────────────────────────────────────────────── */
function NoiseSVG(){return(<svg style={{position:"absolute",inset:0,width:"100%",height:"100%",pointerEvents:"none",opacity:.04}} xmlns="http://www.w3.org/2000/svg"><filter id="nz"><feTurbulence type="fractalNoise" baseFrequency=".65" numOctaves="3" stitchTiles="stitch"/></filter><rect width="100%" height="100%" filter="url(#nz)"/></svg>);}
function PBar({pct,color=T.red,h=4}){return(<div style={{height:h,background:T.borderSm,borderRadius:99,overflow:"hidden"}}><div style={{height:"100%",width:`${Math.min(100,Math.max(0,pct))}%`,background:color,borderRadius:99,animation:"barGrow .6s cubic-bezier(.22,1,.36,1) both"}}/></div>);}
function TipoBadge({tipo,sm}){const cfg=TIPO_CFG[tipo]||TIPO_CFG["Network Rollout"];return(<span style={{display:"inline-flex",alignItems:"center",padding:sm?"2px 8px":"3px 10px",borderRadius:99,background:cfg.bg,border:`1px solid ${cfg.bdr}`,fontSize:sm?8.5:10,fontWeight:700,color:cfg.c,letterSpacing:".05em",textTransform:"uppercase",whiteSpace:"nowrap"}}>{cfg.label}</span>);}

function LvlBadge({n}){
  const l=LVL[n];
  return(
    <span style={{display:"inline-flex",alignItems:"center",gap:3,padding:"1px 6px",
      borderRadius:99,background:l.bg,border:`1px solid ${l.bdr}`}}>
      <span style={{fontSize:7.5,fontWeight:900,color:l.c,letterSpacing:".06em"}}>{l.short}</span>
      <span style={{fontSize:7.5,fontWeight:600,color:l.c}}>{l.label}</span>
    </span>
  );
}

/* Conector visual de árbol */
function TreeConnector({isLast,depth=0,open=false}){
  const left = 12 + depth*2;
  return(
    <div style={{position:"relative",width:32,flexShrink:0,alignSelf:"stretch",minHeight:36}}>
      <div style={{position:"absolute",top:0,left:left,width:2,
        background:T.borderSm,
        bottom: isLast&&!open ? undefined : 0,
        height: isLast&&!open ? 20 : undefined,
      }}/>
      <div style={{position:"absolute",top:20,left:left,width:18,height:2,background:T.borderSm}}/>
    </div>
  );
}

function EN({v,onChange,pct=false,size=14,color=T.ink}){
  const[ed,setEd]=useState(false);
  const[val,setVal]=useState(String(pct?+(v*100).toFixed(2):+v.toFixed(2)));
  useEffect(()=>{if(!ed)setVal(String(pct?+(v*100).toFixed(2):+v.toFixed(2)));},[v,ed,pct]);
  if(ed) return(
    <input autoFocus value={val} onChange={e=>setVal(e.target.value)}
      onBlur={()=>{setEd(false);const n=parseFloat(val);if(!isNaN(n))onChange(pct?n/100:n);else setVal(String(pct?+(v*100).toFixed(2):+v.toFixed(2)));}}
      onKeyDown={e=>{if(e.key==="Enter"){setEd(false);const n=parseFloat(val);if(!isNaN(n))onChange(pct?n/100:n);}if(e.key==="Escape")setEd(false);}}
      style={{width:Math.max(48,val.length*8+16),fontSize:size,fontWeight:800,color,background:"transparent",
        border:"none",borderBottom:`2px solid ${color}`,outline:"none",textAlign:"right",
        fontFamily:"'Outfit',system-ui",padding:"0 2px",letterSpacing:"-.02em"}}/>
  );
  const disp = pct ? `${(v*100).toFixed(1)}%` : (Math.abs(v)>=1e6?`${(v/1e6).toFixed(2)}M`:Math.abs(v)>=1e3?`${(v/1e3).toFixed(1)}K`:v.toFixed(v%1===0?0:2));
  return(
    <span onClick={()=>{setEd(true);setVal(String(pct?+(v*100).toFixed(2):+v.toFixed(2)));}}
      style={{fontSize:size,fontWeight:800,color,cursor:"text",letterSpacing:"-.02em",
        borderBottom:`1px dashed ${color}50`,paddingBottom:1,userSelect:"none"}}>
      {disp}
    </span>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   BLOQUE DRIVER TÉCNICO — catálogo varía por nivel (N2 o N3)
══════════════════════════════════════════════════════════════════════════ */
function DrTecBlock({dtKey, pt, nivel=2, Q_unit, onChangeDt, onChangePt, onChangeQ}){
  const cat  = nivel===3 ? DRV_N3 : DRV_N2;
  const defs = nivel===3 ? DEFS_N3 : DEFS_N2;
  const DT   = cat[dtKey] || Object.values(cat)[0];
  const Q    = DT?.calcQ ? DT.calcQ(pt) : 0;
  const lc   = LVL[nivel];

  return(
    <div style={{borderRadius:11,padding:"10px 12px",
      background:T.surface,
      border:`1.5px solid ${DT?.color||lc.c}25`,
      borderTop:`3px solid ${DT?.color||lc.c}`,
      display:"flex",flexDirection:"column",gap:6}}>
      {/* Cabecera */}
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <span style={{fontSize:14}}>{DT?.icon}</span>
        <div>
          <div style={{fontSize:7.5,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",
            letterSpacing:".1em"}}>Cantidad a invertir (Q)</div>
          <div style={{fontSize:9,fontWeight:700,color:DT?.color,lineHeight:1.2}}>
            {DT?.tag} — {DT?.label}
          </div>
        </div>
      </div>
      {/* Selector */}
      <select value={dtKey}
        onChange={e=>{ const nk=e.target.value; onChangeDt(nk, defs[nk]||{}); }}
        style={{width:"100%",padding:"5px 8px",borderRadius:7,
          border:`1.5px solid ${T.borderSm}`,background:T.card,
          color:T.ink,fontSize:9.5,fontWeight:600,outline:"none",
          appearance:"none",cursor:"pointer",fontFamily:"'Outfit',system-ui"}}>
        {Object.entries(cat).map(([k,v])=>(
          <option key={k} value={k}>{v.icon} {v.label}</option>
        ))}
      </select>
      {/* Descripción */}
      <div style={{fontSize:8.5,color:T.inkSoft,fontStyle:"italic",lineHeight:1.5,
        padding:"4px 8px",background:T.card,borderRadius:5,
        borderLeft:`3px solid ${DT?.color||lc.c}`}}>
        {DT?.desc}
      </div>
      {/* Parámetros */}
      <div style={{borderRadius:7,overflow:"hidden",border:`1px solid ${T.borderSm}`}}>
        {(DT?.params||[]).map(({k,l,p},i)=>(
          <div key={k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",
            padding:"5px 9px",
            background:i%2===0?T.card:"#FAFAF8",
            borderBottom:i<(DT.params.length-1)?`1px solid ${T.borderSm}`:"none"}}>
            <span style={{fontSize:9,color:T.inkMid,maxWidth:"55%",lineHeight:1.3}}>{l}</span>
            <EN v={pt[k]??0} pct={p} color={DT?.color||lc.c} onChange={val=>onChangePt(k,val)}/>
          </div>
        ))}
      </div>
      {/* Q resultante */}
      <div style={{background:`${DT?.color||lc.c}14`,borderRadius:7,padding:"6px 10px",
        display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div style={{fontSize:7.5,fontWeight:700,color:DT?.color||lc.c,
          textTransform:"uppercase",letterSpacing:".1em"}}>Q calculado</div>
        <div style={{display:"flex",alignItems:"baseline",gap:4}}>
          <EN v={Q} onChange={nQ=>onChangeQ(backCalcPt(dtKey,pt,nQ,nivel))}
            size={16} color={DT?.color||lc.c}/>
          <span style={{fontSize:8,color:T.inkMid}}>{Q_unit}</span>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   BLOQUE DRIVER FINANCIERO — catálogo varía por nivel (N2, N3 o N4)
══════════════════════════════════════════════════════════════════════════ */
function DrFinBlock({dfKey, pa, pb, Q_unit, nivel=2, color, onChangeDf, onChangePa, onChangePb}){
  const cat = nivel===4 ? DRV_FIN_N4 : nivel===3 ? DRV_FIN_N3 : DRV_FIN_N2;
  const DF  = cat[dfKey] || cat[Object.keys(cat)[0]];
  const lc  = LVL[nivel];
  const c   = color || DF?.color || lc.c;
  const gap = dp(pb??0, pa??0);

  return(
    <div style={{borderRadius:11,padding:"10px 12px",
      background:T.surface,
      border:`1.5px solid ${c}25`,
      borderTop:`3px solid ${c}`,
      display:"flex",flexDirection:"column",gap:6}}>
      {/* Cabecera */}
      <div style={{display:"flex",alignItems:"center",gap:6}}>
        <span style={{fontSize:14}}>💰</span>
        <div>
          <div style={{fontSize:7.5,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",
            letterSpacing:".1em"}}>Precio / Costo unitario (P)</div>
          <div style={{fontSize:9,fontWeight:700,color:c,lineHeight:1.2}}>
            Precio / Costo unitario
          </div>
        </div>
      </div>
      {/* Selector */}
      <select value={dfKey} onChange={e=>onChangeDf(e.target.value)}
        style={{width:"100%",padding:"5px 8px",borderRadius:7,
          border:`1.5px solid ${T.borderSm}`,background:T.card,
          color:T.ink,fontSize:9.5,fontWeight:600,outline:"none",
          appearance:"none",cursor:"pointer",fontFamily:"'Outfit',system-ui"}}>
        {Object.entries(cat).map(([k,v])=>(
          <option key={k} value={k}>{v.label}</option>
        ))}
      </select>
      {/* Descripción */}
      <div style={{fontSize:8.5,color:T.inkSoft,fontStyle:"italic",lineHeight:1.5,
        padding:"4px 8px",background:T.card,borderRadius:5,
        borderLeft:`3px solid ${DF?.color||c}`}}>
        {DF?.desc}
      </div>
      {/* P actual vs P benchmark */}
      <div style={{borderRadius:7,overflow:"hidden",border:`1px solid ${T.borderSm}`}}>
        {[{l:"P actual / contratado",v:pa,set:onChangePa,faded:true},
          {l:"P benchmark (driver)",  v:pb,set:onChangePb,faded:false}].map(({l,v,set,faded},i)=>(
          <div key={i} style={{display:"flex",justifyContent:"space-between",alignItems:"center",
            padding:"5px 9px",
            background:i===0?T.card:"#FAFAF8",
            borderBottom:i===0?`1px solid ${T.borderSm}`:"none"}}>
            <span style={{fontSize:9,color:T.inkMid}}>{l}</span>
            <div style={{display:"flex",gap:3,alignItems:"baseline"}}>
              <EN v={v??0} color={faded?T.inkSoft:c} onChange={set}/>
              <span style={{fontSize:7.5,color:T.inkSoft}}>/{Q_unit}</span>
            </div>
          </div>
        ))}
      </div>
      {/* P driver resumen */}
      <div style={{background:`${c}12`,borderRadius:7,padding:"6px 10px",
        display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:7.5,fontWeight:700,color:c,textTransform:"uppercase",letterSpacing:".1em"}}>
            P driver
          </div>
          {Math.abs(gap)>.1&&(
            <div style={{fontSize:8.5,fontWeight:700,color:gap<0?T.green:T.red}}>
              {sg(gap)}{gap.toFixed(1)}% vs actual
            </div>
          )}
        </div>
        <div style={{display:"flex",alignItems:"baseline",gap:4}}>
          <EN v={pb??0} onChange={onChangePb} size={16} color={c}/>
          <span style={{fontSize:8,color:T.inkMid}}>/{Q_unit}</span>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   PxQPANEL — 4 niveles: Macro(N1) → Proyecto(N2) → Sub-Proy(N3) → Comp(N4)
   · N2: driver de Demanda (¿por qué?)  ×  driver estratégico de Precio
   · N3: driver de Volumen (¿cuánto?)   ×  driver de precio de componente
   · N4: solo driver de Precio unitario (¿a qué costo?)
   · Cálculos bottom-up: N4 → N3 → N2 → N1
══════════════════════════════════════════════════════════════════════════ */
function PxQPanel({proy, macroData, macroTipo, ov, onChange, onSave}){
  const[openN3, setOpenN3] = useState(null);
  const[openN4, setOpenN4] = useState(null);
  const[saved,  setSaved]  = useState(false); // feedback visual del botón

  /* ── N2: estado del proyecto ── */
  const dt2 = ov?.dt || proy.dt;
  const df2 = ov?.df || proy.df;
  const pt2 = useMemo(()=> ov?.pt ? {...proy.pt,...ov.pt} : {...proy.pt}, [proy.pt,ov]);
  const pf2 = useMemo(()=> ov?.pf ? {...proy.pf,...ov.pf} : {...proy.pf}, [proy.pf,ov]);
  const Q2  = calcQ(dt2, pt2, 2);
  const P2  = pf2.pb;
  const cfg = TIPO_CFG[macroTipo]||TIPO_CFG["Network Rollout"];

  /* ── N3+N4: estado del árbol ── */
  const baseTree = PXQ_TREE[proy.id];

  const buildTree = useCallback(()=>{
    if(!baseTree) return {};
    const n3dt_sug = SUGGEST_N3_DT[proy.dt]||"unidades_fisicas";
    const n3df_sug = SUGGEST_N3_DF[proy.df]||"cotizacion";
    return Object.fromEntries(baseTree.componentes.map(sp=>([sp.id,{
      dt: n3dt_sug,
      pt: {...(DEFS_N3[n3dt_sug]||{})},
      df: n3df_sug,
      pa: sp.P, pb: sp.P,
      // N4: componentes individuales — solo precio
      comps: Object.fromEntries((sp.sub||[]).map((c,ci)=>[ci,{
        df:"precio_lista", pa:c.p, pb:c.p, q:c.q,
      }])),
    }])));
  },[proy.id, proy.dt, proy.df]);

  const[treeState, setTreeState] = useState(()=> ov?._tree || buildTree());
  useEffect(()=>{ setTreeState(ov?._tree || buildTree()); setOpenN3(null); setOpenN4(null); },[proy.id]);

  const updN3 = (spId, patch) => setTreeState(p=>({...p,[spId]:{...p[spId],...patch}}));
  const updN4 = (spId,ci,patch) => setTreeState(p=>({
    ...p,[spId]:{...p[spId],comps:{...p[spId].comps,[ci]:{...p[spId].comps?.[ci],...patch}}}
  }));

  /* ── Cálculos bottom-up ── */
  const n4Total = useCallback((sp,ci)=>{
    const base = sp.sub[ci];
    const cv   = treeState[sp.id]?.comps?.[ci];
    return (cv?.q??base.q) * (cv?.pb??base.p);
  },[treeState]);

  const n3Total = useCallback((sp)=>
    (sp.sub||[]).reduce((s,_,ci)=>s+n4Total(sp,ci),0)
  ,[n4Total]);

  const Q2_base = useMemo(()=>calcQ(proy.dt, proy.pt, 2), [proy.dt, proy.pt]);

  const treeTotal = useMemo(()=>{
    if(!baseTree) return (Q2>0 && P2>0) ? Q2*P2 : proy.P_base;
    // Árbol base: suma fija de componentes N4
    const baseSum = baseTree.componentes.reduce((s,sp)=>s+n3Total(sp),0);
    // Escalar por cambio de Q2 respecto al Q original del proyecto
    const ratio = Q2_base>0 ? Q2/Q2_base : 1;
    return baseSum * ratio;
  },[baseTree,n3Total,Q2,Q2_base,P2,proy.P_base]);

  // Detectar si hay cambios sin guardar
  const isDirty = useMemo(()=>{
    if(ov?._capex != null && Math.abs(ov._capex - treeTotal) > 1) return true;
    if(ov?._capex == null && Math.abs(treeTotal - proy.P_base) > 1) return true;
    return false;
  },[treeTotal, ov, proy.P_base]);

  const handleSave = () => {
    // Guardar _capex + estado N2 en overrides para que el dashboard lo refleje
    onChange({
      _capex: treeTotal,
      dt: dt2, df: df2,
      pt: pt2, pf: pf2,
      _tree: treeState, // guardar estado del árbol
    });
    setSaved(true);
    setTimeout(()=>setSaved(false), 2000);
  };

  const dProy  = dp(treeTotal, proy.P_base);
  const dMacro = macroData ? dp(treeTotal, macroData.P_base) : 0;

  /* ── Fondo del header: color por tipo ── */
  return(
    <div style={{display:"flex",flexDirection:"column",height:"100%"}}>

      {/* ════ HEADER ════ */}
      <div style={{background:`linear-gradient(135deg,${T.redDk},${T.red})`,
        padding:"14px 18px",position:"relative",overflow:"hidden",flexShrink:0}}>
        <NoiseSVG/>
        <div style={{position:"relative",zIndex:1}}>
          {/* Breadcrumb jerárquico */}
          <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:8,flexWrap:"wrap"}}>
            {macroData&&<>
              <span style={{fontSize:8,fontWeight:700,color:"rgba(255,255,255,.4)",
                background:"rgba(255,255,255,.08)",borderRadius:99,padding:"1px 7px"}}>
                N1 · {macroData.macro}
              </span>
              <span style={{color:"rgba(255,255,255,.25)",fontSize:9}}>›</span>
            </>}
            <span style={{fontSize:8,fontWeight:700,color:"rgba(255,255,255,.8)",
              background:"rgba(255,255,255,.15)",borderRadius:99,padding:"1px 7px"}}>
              N2 · {proy.id} {proy.n}
            </span>
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-end"}}>
            <div>
              <div style={{fontSize:15,fontWeight:800,color:"#fff",
                letterSpacing:"-.015em",marginBottom:5}}>{proy.n}</div>
              <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                <TipoBadge tipo={macroTipo} sm/>
                {proy.prio&&proy.prio!=="—"&&
                  <span style={{background:"rgba(255,255,255,.12)",border:"1px solid rgba(255,255,255,.2)",
                    borderRadius:99,padding:"2px 9px",fontSize:8.5,fontWeight:700,
                    color:"rgba(255,255,255,.8)",textTransform:"uppercase"}}>{proy.prio}</span>}
              </div>
            </div>
            <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:6,flexShrink:0,marginLeft:12}}>
              <div style={{textAlign:"right"}}>
                <div style={{fontSize:7.5,color:"rgba(255,255,255,.45)",textTransform:"uppercase",
                  letterSpacing:".1em",marginBottom:2}}>CAPEX DVB</div>
                <EN v={treeTotal} onChange={nC=>{if(Q2>0)onChange({pf:{...pf2,pb:nC/Q2}});}}
                  size={19} color="#fff"/>
                {Math.abs(dProy)>.1&&<div style={{fontSize:9,fontWeight:700,
                  color:dProy<0?"#86EFAC":"#FCA5A5",marginTop:1}}>
                  {sg(dProy)}{dProy.toFixed(1)}% vs base
                </div>}
              </div>
              {/* Botón Guardar */}
              <button onClick={handleSave}
                style={{display:"flex",alignItems:"center",gap:5,padding:"6px 14px",
                  borderRadius:99,border:"none",cursor:"pointer",fontFamily:"'Outfit',system-ui",
                  fontWeight:700,fontSize:11,letterSpacing:"-.01em",
                  background: saved ? "rgba(45,200,120,0.9)" : isDirty ? "#fff" : "rgba(255,255,255,.25)",
                  color: saved ? "#fff" : isDirty ? T.red : "rgba(255,255,255,.7)",
                  boxShadow: isDirty&&!saved ? "0 4px 16px rgba(0,0,0,0.2)" : "none",
                  transition:"all .2s",transform: isDirty&&!saved ? "scale(1.04)" : "scale(1)"}}>
                <span style={{fontSize:13}}>{saved?"✅":"💾"}</span>
                {saved ? "Guardado" : isDirty ? "Guardar cambios" : "Sin cambios"}
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* ════ BODY ════ */}
      <div style={{flex:1,overflow:"auto",padding:"14px 16px 28px"}}>

        {/* ══ N1: MACROPROYECTO — contexto ══ */}
        {macroData&&(
          <div style={{marginBottom:12,padding:"10px 14px",borderRadius:12,
            background:LVL[1].bg, border:`1.5px solid ${LVL[1].bdr}`}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <div style={{width:28,height:28,borderRadius:8,background:LVL[1].c,
                  display:"flex",alignItems:"center",justifyContent:"center",
                  fontSize:13,flexShrink:0}}>🏗</div>
                <div>
                  <div style={{fontSize:8,fontWeight:700,color:LVL[1].c,textTransform:"uppercase",
                    letterSpacing:".1em"}}>N1 · Macroproyecto</div>
                  <div style={{fontSize:11,fontWeight:800,color:T.ink}}>{macroData.macro}</div>
                </div>
              </div>
              <div style={{textAlign:"right",flexShrink:0}}>
                <div style={{fontSize:7.5,color:T.inkSoft,textTransform:"uppercase"}}>Base macro</div>
                <div style={{fontSize:12,fontWeight:900,color:T.inkMid}}>{fu(macroData.P_base)}</div>
                {Math.abs(dMacro)>.1&&<div style={{fontSize:8,fontWeight:700,color:cc(dMacro)}}>
                  {sg(dMacro)}{dMacro.toFixed(1)}% este proyecto
                </div>}
              </div>
            </div>
            {/* Pills de proyectos hermanos */}
            <div style={{marginTop:8,display:"flex",gap:4,flexWrap:"wrap"}}>
              {macroData.proyectos.map(p=>(
                <span key={p.id} style={{fontSize:8,padding:"2px 8px",borderRadius:99,
                  fontWeight:p.id===proy.id?700:400,
                  background:p.id===proy.id?"rgba(232,24,42,.1)":T.surface,
                  border:`1px solid ${p.id===proy.id?T.red:T.borderSm}`,
                  color:p.id===proy.id?T.red:T.inkSoft}}>
                  {p.id} · {p.n}
                </span>
              ))}
            </div>
          </div>
        )}

        {/* ══ N2: PROYECTO — Drivers de Demanda × Precio estratégico ══ */}
        <div style={{marginBottom:14}}>
          {/* Nodo N2 */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
            <div style={{width:28,height:28,borderRadius:8,background:LVL[2].c,
              display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>
              📁
            </div>
            <div style={{flex:1}}>
              <div style={{fontSize:8,fontWeight:700,color:LVL[2].c,textTransform:"uppercase",
                letterSpacing:".1em"}}>N2 · Proyecto — Drivers de Demanda</div>
              <div style={{fontSize:9.5,color:T.inkSoft}}>
                Q×P = <b style={{color:LVL[2].c}}>{fn(Q2)}</b> {proy.m}
                <span style={{margin:"0 4px",color:T.inkXsoft}}>×</span>
                <b style={{color:LVL[2].c}}>{fu(P2)}/{proy.m}</b>
                <span style={{margin:"0 4px",color:T.inkXsoft}}>=</span>
                <b style={{color:T.red}}>{fu(Q2>0&&P2>0?Q2*P2:proy.P_base)}</b>
              </div>
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,
            marginLeft:36,paddingLeft:12,borderLeft:`2px solid ${LVL[2].c}25`}}>
            <DrTecBlock
              dtKey={dt2} pt={pt2} nivel={2} Q_unit={proy.m}
              onChangeDt={(nk,npt)=>onChange({dt:nk,pt:npt})}
              onChangePt={(k,v)=>onChange({pt:{...pt2,[k]:v}})}
              onChangeQ={np=>onChange({pt:np})}
            />
            <DrFinBlock
              dfKey={df2} pa={pf2.pa} pb={pf2.pb} Q_unit={proy.m} nivel={2}
              onChangeDf={nk=>onChange({df:nk})}
              onChangePa={v=>onChange({pf:{...pf2,pa:v}})}
              onChangePb={v=>onChange({pf:{...pf2,pb:v}})}
            />
          </div>
        </div>

        {/* ══ N3 + N4: SUB-PROYECTOS y COMPONENTES ══ */}
        {baseTree ? (<>
          {/* Cabecera N3 */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <div style={{width:28,height:28,borderRadius:8,background:LVL[3].c,
              display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,flexShrink:0}}>
              📂
            </div>
            <div style={{flex:1}}>
              <div style={{fontSize:8,fontWeight:700,color:LVL[3].c,textTransform:"uppercase",
                letterSpacing:".1em"}}>N3 · Sub-Proyectos — Drivers de Volumen físico</div>
              <div style={{fontSize:9,color:T.inkSoft}}>
                {baseTree.componentes.length} sub-proyectos · Total árbol:
                <b style={{color:T.red,marginLeft:4}}>{fu(treeTotal)}</b>
              </div>
            </div>
          </div>

          {/* Barra composición */}
          <div style={{display:"flex",height:5,borderRadius:99,overflow:"hidden",
            marginBottom:7,marginLeft:36,gap:2}}>
            {baseTree.componentes.map(sp=>(
              <div key={sp.id} title={sp.label}
                style={{flex:n3Total(sp)||sp.pct,background:sp.color,opacity:.85}}/>
            ))}
          </div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10,marginLeft:36}}>
            {baseTree.componentes.map(sp=>{
              const t=n3Total(sp);
              return(
                <div key={sp.id} style={{display:"flex",alignItems:"center",gap:3}}>
                  <div style={{width:6,height:6,borderRadius:2,background:sp.color,flexShrink:0}}/>
                  <span style={{fontSize:8,color:T.inkMid,fontWeight:600}}>
                    {sp.label} <span style={{color:sp.color,fontWeight:700}}>{fu(t)}</span>
                    <span style={{color:T.inkXsoft}}> {treeTotal>0?(t/treeTotal*100).toFixed(0):0}%</span>
                  </span>
                </div>
              );
            })}
          </div>

          {/* Sub-Proyectos N3 */}
          <div style={{display:"flex",flexDirection:"column"}}>
            {baseTree.componentes.map((sp,spi)=>{
              const isLast  = spi===baseTree.componentes.length-1;
              const isOpenN = openN3===sp.id;
              const spv     = treeState[sp.id]||{};
              const spTotal = n3Total(sp);
              const spDT    = DRV_N3[spv.dt]||Object.values(DRV_N3)[0];
              const spDF    = DRV_FIN_N3[spv.df]||Object.values(DRV_FIN_N3)[0];
              const spQ     = spDT?.calcQ ? spDT.calcQ(spv.pt||{}) : sp.Q;

              return(
                <div key={sp.id} style={{display:"flex",alignItems:"flex-start"}}>
                  <TreeConnector isLast={isLast} open={isOpenN}/>
                  <div style={{flex:1,minWidth:0,marginBottom:isOpenN?10:5}}>

                    {/* Cabecera N3 */}
                    <div onClick={()=>{setOpenN3(isOpenN?null:sp.id); setOpenN4(null);}}
                      style={{display:"flex",alignItems:"center",gap:7,padding:"8px 10px",
                        borderRadius:11,cursor:"pointer",
                        background:isOpenN?`${sp.color}10`:T.card,
                        border:`1.5px solid ${isOpenN?sp.color:T.borderSm}`,
                        boxShadow:isOpenN?`0 2px 12px ${sp.color}18`:"0 1px 3px rgba(0,0,0,0.04)",
                        transition:"all .18s"}}>
                      <div style={{width:30,height:30,borderRadius:8,flexShrink:0,
                        background:isOpenN?sp.color:`${sp.color}15`,fontSize:15,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        transition:"background .18s"}}>{sp.icon}</div>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                          <span style={{fontSize:11,fontWeight:700,
                            color:isOpenN?sp.color:T.ink,transition:"color .15s",
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                            {sp.label}
                          </span>
                          <LvlBadge n={3}/>
                        </div>
                        <div style={{fontSize:8.5,color:T.inkSoft,display:"flex",gap:5}}>
                          <span style={{color:spDT?.color,fontWeight:600}}>
                            {spDT?.icon} Q:{fn(spQ)} {sp.Q_unit}
                          </span>
                          <span style={{color:T.inkXsoft}}>·</span>
                          <span style={{color:spDF?.color||T.blue,fontWeight:600}}>
                            💰 P:{fu(spv.pb??sp.P)}
                          </span>
                        </div>
                      </div>
                      <div style={{textAlign:"right",flexShrink:0}}>
                        <div style={{fontSize:14,fontWeight:900,color:sp.color}}>{fu(spTotal)}</div>
                        <div style={{fontSize:8,color:T.inkSoft}}>
                          {treeTotal>0?(spTotal/treeTotal*100).toFixed(0):0}%
                        </div>
                      </div>
                      <div style={{width:18,height:18,borderRadius:"50%",flexShrink:0,
                        background:isOpenN?`${sp.color}20`:T.borderSm,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        transition:"all .18s"}}>
                        <span style={{fontSize:10,color:isOpenN?sp.color:T.inkSoft,fontWeight:700,
                          transform:isOpenN?"rotate(90deg)":"none",
                          display:"block",transition:"transform .2s",lineHeight:1}}>›</span>
                      </div>
                    </div>

                    {/* N3 EXPANDIDO */}
                    {isOpenN&&(
                      <div className="fi" style={{marginTop:5,marginLeft:8,
                        paddingLeft:12,borderLeft:`2px solid ${sp.color}35`}}>

                        {/* Drivers N3 */}
                        <div style={{fontSize:8,fontWeight:700,color:LVL[3].c,
                          textTransform:"uppercase",letterSpacing:".1em",
                          marginBottom:7,display:"flex",alignItems:"center",gap:4}}>
                          <span style={{width:14,height:14,borderRadius:4,background:LVL[3].c,
                            display:"inline-flex",alignItems:"center",justifyContent:"center",
                            fontSize:7,fontWeight:900,color:"#fff",flexShrink:0}}>3</span>
                          Drivers de Volumen · {sp.label}
                        </div>
                        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:7,marginBottom:10}}>
                          <DrTecBlock
                            dtKey={spv.dt||(SUGGEST_N3_DT[proy.dt]||"unidades_fisicas")}
                            pt={spv.pt||(DEFS_N3[spv.dt||"unidades_fisicas"]||{})}
                            nivel={3} Q_unit={sp.Q_unit}
                            onChangeDt={(nk,npt)=>updN3(sp.id,{dt:nk,pt:npt})}
                            onChangePt={(k,v)=>updN3(sp.id,{pt:{...(spv.pt||{}),[k]:v}})}
                            onChangeQ={np=>updN3(sp.id,{pt:np})}
                          />
                          <DrFinBlock
                            dfKey={spv.df||(SUGGEST_N3_DF[proy.df]||"cotizacion")}
                            pa={spv.pa??sp.P} pb={spv.pb??sp.P}
                            Q_unit={sp.Q_unit} nivel={3} color={sp.color}
                            onChangeDf={nk=>updN3(sp.id,{df:nk})}
                            onChangePa={v=>updN3(sp.id,{pa:v})}
                            onChangePb={v=>updN3(sp.id,{pb:v})}
                          />
                        </div>

                        {/* Componentes N4 */}
                        <div style={{fontSize:8,fontWeight:700,color:LVL[4].c,
                          textTransform:"uppercase",letterSpacing:".1em",
                          marginBottom:6,display:"flex",alignItems:"center",gap:4}}>
                          <span style={{width:14,height:14,borderRadius:4,background:LVL[4].c,
                            display:"inline-flex",alignItems:"center",justifyContent:"center",
                            fontSize:7,fontWeight:900,color:"#fff",flexShrink:0}}>4</span>
                          Componentes de Costo · Driver de Precio unitario
                        </div>

                        <div style={{display:"flex",flexDirection:"column"}}>
                          {(sp.sub||[]).map((comp,ci)=>{
                            const ck   = `${sp.id}-${ci}`;
                            const isOC = openN4===ck;
                            const isLC = ci===sp.sub.length-1;
                            const cv   = treeState[sp.id]?.comps?.[ci]||{};
                            const cQ   = cv.q??comp.q;
                            const cPb  = cv.pb??comp.p;
                            const cPa  = cv.pa??comp.p;
                            const cTot = cQ*cPb;
                            const cDF  = DRV_FIN_N4[cv.df||"precio_lista"]||Object.values(DRV_FIN_N4)[0];
                            const dCP  = dp(cPb,cPa);

                            return(
                              <div key={ci} style={{display:"flex",alignItems:"flex-start"}}>
                                {/* Conector N3→N4 */}
                                <div style={{position:"relative",width:22,flexShrink:0,
                                  alignSelf:"stretch",minHeight:32}}>
                                  <div style={{position:"absolute",top:0,left:9,width:1.5,
                                    background:T.borderSm,
                                    bottom: isLC&&!isOC ? undefined : 0,
                                    height: isLC&&!isOC ? 16 : undefined}}/>
                                  <div style={{position:"absolute",top:16,left:9,
                                    width:12,height:1.5,background:T.borderSm}}/>
                                </div>

                                <div style={{flex:1,minWidth:0,marginBottom:isOC?6:3}}>
                                  {/* Fila N4 */}
                                  <div onClick={()=>setOpenN4(isOC?null:ck)}
                                    style={{display:"flex",alignItems:"center",gap:6,
                                      padding:"5px 8px",borderRadius:9,cursor:"pointer",
                                      background:isOC?`${sp.color}08`:T.surface,
                                      border:`1.5px solid ${isOC?sp.color:T.borderSm}`,
                                      transition:"all .15s"}}>
                                    {/* Dot */}
                                    <div style={{width:8,height:8,borderRadius:"50%",flexShrink:0,
                                      background:isOC?sp.color:`${sp.color}50`,
                                      border:`1.5px solid ${sp.color}60`,transition:"all .15s"}}/>
                                    {/* Label + badge driver precio */}
                                    <div style={{flex:1,minWidth:0}}>
                                      <div style={{display:"flex",alignItems:"center",gap:4,marginBottom:1}}>
                                        <span style={{fontSize:10,fontWeight:isOC?700:500,
                                          color:isOC?sp.color:T.ink,transition:"color .15s",
                                          overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                                          {comp.label}
                                        </span>
                                        <LvlBadge n={4}/>
                                      </div>
                                      <span style={{fontSize:7.5,fontWeight:700,
                                        color:cDF?.color||T.green,
                                        background:`${cDF?.color||T.green}12`,
                                        borderRadius:99,padding:"1px 6px",
                                        border:`1px solid ${cDF?.color||T.green}22`}}>
                                        💰 {cDF?.label}
                                      </span>
                                    </div>
                                    {/* Q inline */}
                                    <div style={{textAlign:"right",flexShrink:0,minWidth:48}}>
                                      <div style={{fontSize:7.5,color:T.inkSoft,marginBottom:1}}>Q</div>
                                      <EN v={cQ} onChange={v=>updN4(sp.id,ci,{q:v})} size={11} color={T.ink}/>
                                      <div style={{fontSize:7,color:T.inkSoft}}>{sp.Q_unit}</div>
                                    </div>
                                    {/* P inline */}
                                    <div style={{textAlign:"right",flexShrink:0,minWidth:56}}>
                                      <div style={{fontSize:7.5,color:T.inkSoft,marginBottom:1}}>P unit.</div>
                                      <EN v={cPb} onChange={v=>updN4(sp.id,ci,{pb:v})} size={11} color={LVL[4].c}/>
                                      {Math.abs(dCP)>.3&&<div style={{fontSize:7,fontWeight:700,color:cc(dCP)}}>
                                        {sg(dCP)}{dCP.toFixed(0)}%
                                      </div>}
                                    </div>
                                    {/* Subtotal — live */}
                                    <div style={{textAlign:"right",flexShrink:0,minWidth:60}}>
                                      <div style={{fontSize:7.5,color:T.inkSoft,marginBottom:1}}>Subtotal</div>
                                      <div style={{fontSize:12,fontWeight:800,color:sp.color}}>{fu(cTot)}</div>
                                      {Math.abs(dp(cTot,comp.q*comp.p))>.3&&
                                        <div style={{fontSize:7,fontWeight:700,color:cc(dp(cTot,comp.q*comp.p))}}>
                                          {sg(dp(cTot,comp.q*comp.p))}{dp(cTot,comp.q*comp.p).toFixed(0)}%
                                        </div>}
                                    </div>
                                    {/* Chevron */}
                                    <div style={{width:14,height:14,borderRadius:"50%",flexShrink:0,
                                      background:isOC?`${sp.color}18`:T.borderSm,
                                      display:"flex",alignItems:"center",justifyContent:"center",
                                      transition:"all .15s"}}>
                                      <span style={{fontSize:8,color:isOC?sp.color:T.inkSoft,fontWeight:700,
                                        transform:isOC?"rotate(90deg)":"none",
                                        display:"block",transition:"transform .18s",lineHeight:1}}>›</span>
                                    </div>
                                  </div>

                                  {/* N4 EXPANDIDO — solo driver de precio */}
                                  {isOC&&(
                                    <div className="fi" style={{marginTop:4,marginLeft:6,
                                      paddingLeft:10,borderLeft:`2px solid ${sp.color}28`,
                                      paddingBottom:4}}>
                                      <div style={{fontSize:8,fontWeight:700,color:LVL[4].c,
                                        textTransform:"uppercase",letterSpacing:".1em",
                                        marginBottom:6,display:"flex",alignItems:"center",gap:4}}>
                                        <span style={{width:14,height:14,borderRadius:4,background:LVL[4].c,
                                          display:"inline-flex",alignItems:"center",justifyContent:"center",
                                          fontSize:7,fontWeight:900,color:"#fff",flexShrink:0}}>4</span>
                                        Driver de Precio · {comp.label}
                                      </div>
                                      <DrFinBlock
                                        dfKey={cv.df||"precio_lista"}
                                        pa={cPa} pb={cPb}
                                        Q_unit={sp.Q_unit} nivel={4} color={sp.color}
                                        onChangeDf={nk=>updN4(sp.id,ci,{df:nk})}
                                        onChangePa={v=>updN4(sp.id,ci,{pa:v})}
                                        onChangePb={v=>updN4(sp.id,ci,{pb:v})}
                                      />
                                      {/* Mini total N4 */}
                                      <div style={{marginTop:7,background:T.card,borderRadius:8,
                                        padding:"7px 10px",border:`1.5px solid ${sp.color}22`,
                                        display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                                        <div style={{fontSize:9,color:T.inkMid}}>
                                          <b style={{color:T.ink}}>{fn(cQ)}</b>
                                          <span style={{color:T.inkXsoft,margin:"0 4px"}}>×</span>
                                          <b style={{color:LVL[4].c}}>{fu(cPb)}</b>
                                          <span style={{color:T.inkXsoft,margin:"0 4px"}}>=</span>
                                        </div>
                                        <div style={{textAlign:"right"}}>
                                          <div style={{fontSize:15,fontWeight:900,color:sp.color}}>{fu(cTot)}</div>
                                          {Math.abs(dp(cTot,comp.q*comp.p))>.3&&
                                            <div style={{fontSize:8,fontWeight:700,color:cc(dp(cTot,comp.q*comp.p))}}>
                                              {sg(dp(cTot,comp.q*comp.p))}{dp(cTot,comp.q*comp.p).toFixed(1)}% vs base {fu(comp.q*comp.p)}
                                            </div>}
                                        </div>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>

                        {/* Total N3 */}
                        <div style={{marginTop:7,background:`${sp.color}10`,borderRadius:10,
                          padding:"8px 12px",border:`1.5px solid ${sp.color}22`,
                          display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <div>
                            <div style={{fontSize:7.5,fontWeight:700,color:sp.color,
                              textTransform:"uppercase",letterSpacing:".1em"}}>
                              Total N3 · {sp.label}
                            </div>
                            <div style={{fontSize:8.5,color:T.inkMid,marginTop:1}}>
                              {(sp.sub||[]).length} componentes (N4)
                            </div>
                          </div>
                          <div style={{fontSize:16,fontWeight:900,color:sp.color}}>{fu(spTotal)}</div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>

          {/* ── TOTAL ÁRBOL ── */}
          <div style={{marginTop:10,marginLeft:36,background:T.redXsoft,borderRadius:13,
            border:`1.5px solid ${T.redSoft}`,padding:"11px 15px",
            display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div>
              <div style={{fontSize:8,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",
                letterSpacing:".12em",marginBottom:3}}>CAPEX DVB = Σ N4 (Q×P)</div>
              <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                {baseTree.componentes.map(sp=>(
                  <span key={sp.id} style={{fontSize:8,color:sp.color,fontWeight:700}}>
                    {sp.label.split(" ")[0]}: {fu(n3Total(sp))}
                  </span>
                ))}
              </div>
            </div>
            <div style={{textAlign:"right",flexShrink:0,marginLeft:10}}>
              <div style={{fontSize:22,fontWeight:900,color:T.red,letterSpacing:"-.025em"}}>
                {fu(treeTotal)}
              </div>
              <div style={{fontSize:9,fontWeight:700,color:cc(dProy),marginTop:2}}>
                {sg(dProy)}{dProy.toFixed(1)}% vs base {fu(proy.P_base)}
              </div>
            </div>
          </div>
        </>):(
          <div style={{marginLeft:36,background:T.surface,borderRadius:12,padding:"16px",
            border:`1px solid ${T.borderSm}`,textAlign:"center"}}>
            <div style={{fontSize:22,marginBottom:6}}>🌱</div>
            <div style={{fontSize:12,fontWeight:700,color:T.ink,marginBottom:3}}>
              Sub-Proyectos (N3) en construcción
            </div>
            <div style={{fontSize:10.5,color:T.inkSoft,lineHeight:1.7,maxWidth:280,margin:"0 auto"}}>
              Los N3 y N4 serán levantados en talleres con los directores de área responsables.
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function Drawer({proy,macroData,macroTipo,ov,onChange,onClose}){
  const [width, setWidth] = useState(680);
  const dragging = useRef(false);
  const startX   = useRef(0);
  const startW   = useRef(0);
  const onMouseDown = (e) => {
    dragging.current=true; startX.current=e.clientX; startW.current=width;
    document.body.style.cursor="ew-resize"; document.body.style.userSelect="none";
  };
  useEffect(() => {
    const onMove=(e)=>{ if(!dragging.current) return;
      const delta=startX.current-e.clientX;
      setWidth(Math.round(Math.min(Math.max(startW.current+delta,420),window.innerWidth*0.92)));
    };
    const onUp=()=>{ dragging.current=false; document.body.style.cursor=""; document.body.style.userSelect=""; };
    window.addEventListener("mousemove",onMove); window.addEventListener("mouseup",onUp);
    return()=>{ window.removeEventListener("mousemove",onMove); window.removeEventListener("mouseup",onUp); };
  },[]);
  return(
    <div style={{position:"fixed",inset:0,zIndex:400,display:"flex"}}>
      <div style={{flex:1,background:"rgba(0,0,0,.44)",backdropFilter:"blur(5px)"}} onClick={onClose}/>
      <div className="sr" style={{width,minWidth:420,background:T.card,borderLeft:`1px solid ${T.borderSm}`,display:"flex",flexDirection:"column",overflow:"hidden",boxShadow:"-20px 0 60px rgba(0,0,0,0.12)",position:"relative"}}>
        <div onMouseDown={onMouseDown} style={{position:"absolute",left:0,top:0,bottom:0,width:6,cursor:"ew-resize",zIndex:10,background:"transparent",transition:"background .15s"}}
          onMouseEnter={e=>e.currentTarget.style.background="rgba(232,24,42,0.18)"}
          onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <div style={{position:"absolute",top:"50%",left:"50%",transform:"translate(-50%,-50%)",display:"flex",flexDirection:"column",gap:3}}>
            {[0,1,2].map(i=>(<div key={i} style={{width:2,height:16,borderRadius:99,background:"rgba(156,154,149,0.5)"}}/>))}
          </div>
        </div>
        <div style={{flexShrink:0,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 16px",background:T.surface,borderBottom:`1px solid ${T.borderSm}`}}>
          <span style={{fontSize:10,color:T.inkSoft,fontWeight:600}}>{width}px</span>
          <button onClick={onClose} style={{width:30,height:30,borderRadius:9,border:`1px solid ${T.borderSm}`,background:T.card,color:T.inkMid,cursor:"pointer",fontSize:16,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
        </div>
        <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
          <PxQPanel proy={proy} macroData={macroData} macroTipo={macroTipo} ov={ov} onChange={onChange} onSave={onChange}/>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   MODAL AJUSTE INVERSO
══════════════════════════════════════════════════════════════════════════ */
function InverseModal({overrides,onApply,onClose}){
  const tDVB=useMemo(()=>DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+calcCap(p,overrides[p.id]),0),0),[overrides]);
  const[targetM,setTargetM]=useState((tDVB/1e6).toFixed(1));
  const[mode,setMode]=useState("q");
  const[sel,setSel]=useState(()=>Object.fromEntries(DATA.flatMap(m=>m.proyectos.map(p=>[p.id,true]))));
  const tgt=(parseFloat(targetM)||0)*1e6;
  const dTarget=dp(tgt,tDVB);
  const tSel=DATA.reduce((s,m)=>s+m.proyectos.filter(p=>sel[p.id]).reduce((sp,p)=>sp+calcCap(p,overrides[p.id]),0),0);
  const factor=tSel>0?(tgt-(tDVB-tSel))/tSel:1;
  const selN=Object.values(sel).filter(Boolean).length;
  const apply=()=>{
    const upd={};
    DATA.forEach(m=>m.proyectos.forEach(p=>{
      if(!sel[p.id])return;
      const ov=overrides[p.id]||{};const pt=ov.pt?{...p.pt,...ov.pt}:{...p.pt};const pf=ov.pf?{...p.pf,...ov.pf}:{...p.pf};
      const Q=calcQ(ov.dt||p.dt,pt),np={...pt},dk=ov.dt||p.dt;
      if(mode!=="p"){if(dk==="obsolescencia")np.pct_eol=(pt.meta_pct||0)+Q*factor/(pt.total_activos||1);else if(dk==="congestion")np.pct_congestion=(pt.meta_pct||0)+Q*factor/(pt.total_activos||1);else if(dk==="activaciones")np.brutas=(pt.churn||0)+Q*factor;else if(dk==="migraciones")np.tasa=Q*factor/(pt.base_migrable||1);else if(dk==="expansion_geo")np.objetivo=(pt.actual||0)+Q*factor;else if(dk==="pipeline_b2b")np.win_rate=Q*factor/(pt.pipeline||1);else if(dk==="mantenimiento")np.frecuencia=Q*factor/(pt.activos||1);else if(dk==="crecimiento_trafico")np.crecimiento_pct=Q*factor/(pt.capacidad_actual||1);else if(dk==="cobertura_crc")np.comprometido=(pt.ejecutado||0)+Q*factor;}
      upd[p.id]={dt:dk,df:ov.df||p.df,pt:np,pf:{...pf,pb:mode==="q"?pf.pb:pf.pb*factor}};
    }));
    onApply(upd);onClose();
  };
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:500,backdropFilter:"blur(6px)",padding:20}}>
      <div className="si" onClick={e=>e.stopPropagation()} style={{width:"100%",maxWidth:520,background:T.card,borderRadius:24,border:`1px solid ${T.borderSm}`,boxShadow:"0 40px 80px rgba(0,0,0,0.18)",overflow:"hidden"}}>
        <div style={{background:`linear-gradient(135deg,${T.redDk},${T.red})`,padding:"22px 26px",position:"relative",overflow:"hidden"}}>
          <NoiseSVG/>
          <div style={{position:"relative",zIndex:1}}>
            <div style={{fontSize:17,fontWeight:900,color:"#fff",letterSpacing:"-.02em",marginBottom:3}}>Ajuste Inverso de CAPEX</div>
            <div style={{fontSize:11,color:"rgba(255,255,255,.55)"}}>Define el target financiero → propaga a drivers P y/o Q</div>
          </div>
        </div>
        <div style={{padding:"22px 26px"}}>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,padding:"14px 16px",background:T.surface,borderRadius:14,border:`1px solid ${T.borderSm}`,marginBottom:18}}>
            {[{l:"DVB Actual",v:fu(tDVB)},{l:"Nuevo Target ($M)",v:null},{l:"Variación",v:`${sg(dTarget)}${dTarget.toFixed(1)}%`,c:cc(dTarget)}].map((k,i)=>(
              <div key={i}>
                <div style={{fontSize:9,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",letterSpacing:".1em",marginBottom:5}}>{k.l}</div>
                {i===1?(<div style={{display:"flex",alignItems:"baseline",gap:4}}><span style={{fontSize:14,color:T.inkMid,fontWeight:700}}>$</span><input value={targetM} onChange={e=>setTargetM(e.target.value)} style={{width:72,background:"transparent",border:"none",borderBottom:`2.5px solid ${T.red}`,color:T.red,fontSize:22,fontWeight:900,outline:"none",textAlign:"right",fontFamily:"'Outfit',system-ui"}}/><span style={{fontSize:12,color:T.inkMid,fontWeight:700}}>M</span></div>)
                :<div style={{fontSize:20,fontWeight:900,color:k.c||T.ink,letterSpacing:"-.02em"}}>{k.v}</div>}
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:8,marginBottom:16}}>
            {[["q","Solo Q"],["p","Solo P"],["ambos","Q y P"]].map(([id,lbl])=>(
              <button key={id} onClick={()=>setMode(id)} style={{flex:1,padding:"9px",borderRadius:11,border:`1.5px solid ${mode===id?T.red:T.borderSm}`,background:mode===id?T.redXsoft:T.surface,color:mode===id?T.red:T.inkMid,fontSize:11,fontWeight:700,cursor:"pointer",transition:"all .15s",textTransform:"uppercase",letterSpacing:".05em"}}>{lbl}</button>
            ))}
          </div>
          <div style={{border:`1px solid ${T.borderSm}`,borderRadius:12,overflow:"hidden",maxHeight:200,overflowY:"auto",marginBottom:16}}>
            {DATA.slice(0,14).map(m=>{const allS=m.proyectos.every(p=>sel[p.id]);const cfg=TIPO_CFG[m.tipo]||TIPO_CFG["Network Rollout"];return(<div key={m.macro} onClick={()=>{const v=!allS;setSel(s=>({...s,...Object.fromEntries(m.proyectos.map(p=>[p.id,v]))}));}} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 14px",background:T.card,cursor:"pointer",borderBottom:`1px solid ${T.borderSm}`}}>
              <div style={{width:13,height:13,borderRadius:4,border:`1.5px solid ${allS?T.red:T.inkXsoft}`,background:allS?T.red:"transparent",flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>{allS&&<span style={{color:"#fff",fontSize:8,fontWeight:900,lineHeight:1}}>✓</span>}</div>
              <span style={{fontSize:11,fontWeight:600,color:T.ink,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.macro}</span>
              <span style={{display:"inline-flex",padding:"1px 8px",borderRadius:99,background:cfg.bg,fontSize:8.5,fontWeight:700,color:cfg.c,border:`1px solid ${cfg.bdr}`,whiteSpace:"nowrap"}}>{cfg.label}</span>
            </div>);})}
          </div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"12px 16px",background:T.surface,borderRadius:12,border:`1px solid ${T.borderSm}`,marginBottom:18}}>
            <span style={{fontSize:12,color:T.inkMid}}>Factor sobre {selN} proyectos:</span>
            <span style={{fontSize:24,fontWeight:900,color:factor<1?T.green:factor>1?T.red:T.ink,letterSpacing:"-.02em"}}>{(factor*100).toFixed(1)}%</span>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button onClick={onClose} style={{padding:"10px 22px",borderRadius:12,border:`1.5px solid ${T.borderSm}`,background:T.surface,color:T.inkMid,fontWeight:600,fontSize:12,cursor:"pointer"}}>Cancelar</button>
            <button className="btn-primary" onClick={apply} style={{padding:"10px 26px",borderRadius:12,border:"none",background:`linear-gradient(135deg,${T.red},${T.redDk})`,color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer",boxShadow:`0 4px 16px rgba(232,24,42,.3)`}}>Aplicar ajuste</button>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   VISTA TABLERO — DASHBOARD CON CARDS
══════════════════════════════════════════════════════════════════════════ */
/* ══════════════════════════════════════════════════════════════════════════
   RESET MODAL — Devolver ajustes a línea base (selectivo)
══════════════════════════════════════════════════════════════════════════ */
function ResetModal({overrides, onApply, onClose}){
  // Construir árbol: solo macros/proyectos que tienen overrides
  const tree = DATA.map(m=>{
    const proyAjustados = m.proyectos.filter(p=>overrides[p.id]);
    if(!proyAjustados.length) return null;
    return {...m, proyAjustados};
  }).filter(Boolean);

  // Estado de selección: {macroIdx: true/false/null(indeterminate), "pid": true/false}
  const initSel = ()=>{
    const s={};
    tree.forEach(m=>{
      m.proyAjustados.forEach(p=>{ s[p.id]=true; });
    });
    return s;
  };
  const[sel,setSel]=useState(initSel);

  // Helpers de selección
  const macroState = m => {
    const ids = m.proyAjustados.map(p=>p.id);
    const checked = ids.filter(id=>sel[id]).length;
    if(checked===0)        return "none";
    if(checked===ids.length) return "all";
    return "partial";
  };

  const toggleMacro = m => {
    const ids = m.proyAjustados.map(p=>p.id);
    const st  = macroState(m);
    setSel(prev=>{
      const n={...prev};
      ids.forEach(id=>{ n[id] = st!=="all"; });
      return n;
    });
  };

  const toggleProj = id => setSel(prev=>({...prev,[id]:!prev[id]}));

  const selectedIds = Object.entries(sel).filter(([,v])=>v).map(([k])=>k);
  const totalSel    = selectedIds.length;
  const totalAll    = Object.keys(overrides).length;

  // Calcular delta de lo que se va a resetear
  const deltaReset  = selectedIds.reduce((s,id)=>{
    const p = DATA.flatMap(m=>m.proyectos).find(p=>p.id===id);
    if(!p) return s;
    return s + (calcCap(p, overrides[id]) - p.P_base);
  }, 0);

  const[expandedMacros, setExpandedMacros]=useState(()=>new Set(tree.map(m=>m.macro)));
  const toggleExpand = macro => setExpandedMacros(prev=>{
    const n=new Set(prev);
    n.has(macro)?n.delete(macro):n.add(macro);
    return n;
  });

  return(
    <div style={{position:"fixed",inset:0,zIndex:500,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,.5)",backdropFilter:"blur(6px)"}}>
      <div className="si" style={{width:560,maxHeight:"82vh",background:T.card,borderRadius:20,
        boxShadow:"0 32px 80px rgba(0,0,0,0.22)",display:"flex",flexDirection:"column",overflow:"hidden"}}>

        {/* Header */}
        <div style={{padding:"18px 22px 14px",borderBottom:`1px solid ${T.borderSm}`,flexShrink:0}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
            <div style={{display:"flex",alignItems:"center",gap:10}}>
              <div style={{width:36,height:36,borderRadius:10,background:T.redBg,
                border:`1.5px solid ${T.redSoft}`,display:"flex",alignItems:"center",
                justifyContent:"center",fontSize:18}}>↺</div>
              <div>
                <div style={{fontSize:15,fontWeight:800,color:T.ink,letterSpacing:"-.01em"}}>
                  Devolver a línea base
                </div>
                <div style={{fontSize:10.5,color:T.inkSoft}}>
                  Selecciona qué ajustes quieres revertir
                </div>
              </div>
            </div>
            <button onClick={onClose} style={{width:28,height:28,borderRadius:8,border:`1px solid ${T.borderSm}`,
              background:T.surface,cursor:"pointer",fontSize:15,color:T.inkSoft,
              display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
          </div>
          {/* Barra resumen */}
          <div style={{display:"flex",gap:8,marginTop:10,flexWrap:"wrap"}}>
            <div style={{padding:"4px 12px",borderRadius:99,background:T.surface,
              border:`1px solid ${T.borderSm}`,fontSize:9.5,color:T.inkMid,fontWeight:600}}>
              {totalSel} de {totalAll} proyectos seleccionados
            </div>
            {totalSel>0&&<div style={{padding:"4px 12px",borderRadius:99,
              background:deltaReset<0?T.greenBg:T.redBg,
              border:`1px solid ${deltaReset<0?T.greenBdr:T.redSoft}`,
              fontSize:9.5,fontWeight:700,color:deltaReset<0?T.green:T.red}}>
              {deltaReset<0?"Recupera ahorro":"Revierte incremento"}: {fu(Math.abs(deltaReset))}
            </div>}
            {/* Seleccionar todo / ninguno */}
            <div style={{marginLeft:"auto",display:"flex",gap:6}}>
              <button onClick={()=>setSel(initSel())}
                style={{padding:"3px 10px",borderRadius:99,border:`1px solid ${T.borderSm}`,
                  background:T.surface,fontSize:9,fontWeight:600,color:T.inkMid,cursor:"pointer"}}>
                Todo
              </button>
              <button onClick={()=>setSel({})}
                style={{padding:"3px 10px",borderRadius:99,border:`1px solid ${T.borderSm}`,
                  background:T.surface,fontSize:9,fontWeight:600,color:T.inkMid,cursor:"pointer"}}>
                Ninguno
              </button>
            </div>
          </div>
        </div>

        {/* Lista scrollable */}
        <div style={{flex:1,overflow:"auto",padding:"8px 0"}}>
          {tree.length===0?(
            <div style={{textAlign:"center",padding:"40px 20px",color:T.inkSoft,fontSize:12}}>
              No hay ajustes guardados
            </div>
          ):tree.map(m=>{
            const cfg    = TIPO_CFG[m.tipo]||TIPO_CFG["Network Rollout"];
            const mState = macroState(m);
            const isExp  = expandedMacros.has(m.macro);
            const mDelta = m.proyAjustados.reduce((s,p)=>s+(sel[p.id]?calcCap(p,overrides[p.id])-p.P_base:0),0);

            return(
              <div key={m.macro} style={{borderBottom:`1px solid ${T.borderSm}`}}>
                {/* Fila Macroproyecto */}
                <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 22px",
                  cursor:"pointer",background:mState==="all"?"#FFF8F7":T.card,
                  transition:"background .12s"}}
                  className="row-hover">
                  {/* Checkbox macro */}
                  <div onClick={()=>toggleMacro(m)}
                    style={{width:18,height:18,borderRadius:5,flexShrink:0,cursor:"pointer",
                      border:`2px solid ${mState==="none"?T.borderSm:T.red}`,
                      background:mState==="all"?T.red:T.card,
                      display:"flex",alignItems:"center",justifyContent:"center",
                      transition:"all .15s"}}>
                    {mState==="all"&&<span style={{color:"#fff",fontSize:11,lineHeight:1,fontWeight:900}}>✓</span>}
                    {mState==="partial"&&<span style={{color:T.red,fontSize:13,lineHeight:.8,fontWeight:900}}>−</span>}
                  </div>
                  {/* Label macro */}
                  <div style={{flex:1,minWidth:0}} onClick={()=>toggleExpand(m.macro)}>
                    <div style={{display:"flex",alignItems:"center",gap:7}}>
                      <TipoBadge tipo={m.tipo} sm/>
                      <span style={{fontSize:11,fontWeight:700,color:T.ink,
                        overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                        {m.macro}
                      </span>
                    </div>
                    <div style={{fontSize:9,color:T.inkSoft,marginTop:2}}>
                      {m.proyAjustados.length} proyecto{m.proyAjustados.length>1?"s":""} ajustado{m.proyAjustados.length>1?"s":""}
                    </div>
                  </div>
                  {/* Delta macro */}
                  {mState!=="none"&&mDelta!==0&&(
                    <div style={{fontSize:10,fontWeight:700,flexShrink:0,
                      color:mDelta<0?T.green:T.red}}>
                      {mDelta<0?"-":"+"}{fu(Math.abs(mDelta))}
                    </div>
                  )}
                  {/* Expand toggle */}
                  <div onClick={()=>toggleExpand(m.macro)}
                    style={{width:20,height:20,borderRadius:6,background:T.surface,
                      border:`1px solid ${T.borderSm}`,display:"flex",alignItems:"center",
                      justifyContent:"center",flexShrink:0,cursor:"pointer",transition:"all .15s"}}>
                    <span style={{fontSize:10,color:T.inkSoft,fontWeight:700,
                      display:"block",transition:"transform .18s",lineHeight:1,
                      transform:isExp?"rotate(90deg)":"none"}}>›</span>
                  </div>
                </div>

                {/* Proyectos hijos */}
                {isExp&&m.proyAjustados.map(p=>{
                  const ov  = overrides[p.id];
                  const dvb = calcCap(p,ov);
                  const d   = dvb - p.P_base;
                  const pct = dp(dvb, p.P_base);
                  const cfg2= TIPO_CFG[m.tipo]||TIPO_CFG["Network Rollout"];

                  // Qué tiene este override
                  const tags=[];
                  if(ov?._capex!=null&&ov?._tree)  tags.push({l:"Árbol N3/N4",  c:T.violet, icon:"🌳"});
                  if(ov?.dt&&ov.dt!==p.dt)          tags.push({l:"Driver Q",     c:T.blue,   icon:"📐"});
                  if(ov?.pf?.pb&&ov.pf.pb!==p.pf?.pb) tags.push({l:"Driver P",  c:T.green,  icon:"💰"});

                  return(
                    <div key={p.id} onClick={()=>toggleProj(p.id)}
                      style={{display:"flex",alignItems:"center",gap:10,
                        padding:"8px 22px 8px 52px",
                        background:sel[p.id]?"#FFF8F7":T.surface,
                        borderTop:`1px solid ${T.borderSm}`,
                        cursor:"pointer",transition:"background .1s"}}
                      className="row-hover">
                      {/* Checkbox proyecto */}
                      <div style={{width:16,height:16,borderRadius:4,flexShrink:0,
                        border:`2px solid ${sel[p.id]?T.red:T.borderSm}`,
                        background:sel[p.id]?T.red:T.card,
                        display:"flex",alignItems:"center",justifyContent:"center",
                        transition:"all .15s"}}>
                        {sel[p.id]&&<span style={{color:"#fff",fontSize:10,lineHeight:1,fontWeight:900}}>✓</span>}
                      </div>
                      {/* Info proyecto */}
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:5,marginBottom:2}}>
                          <span style={{fontSize:8,fontWeight:700,color:T.inkSoft,
                            background:T.surface,border:`1px solid ${T.borderSm}`,
                            borderRadius:4,padding:"1px 5px"}}>{p.id}</span>
                          <span style={{fontSize:10.5,fontWeight:600,color:T.ink,
                            overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.n}</span>
                        </div>
                        <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                          {tags.map((t,ti)=>(
                            <span key={ti} style={{fontSize:8,fontWeight:700,
                              color:t.c,background:t.c+"15",
                              borderRadius:99,padding:"1px 6px",
                              border:`1px solid ${t.c}25`}}>
                              {t.icon} {t.l}
                            </span>
                          ))}
                        </div>
                      </div>
                      {/* Valores base → dvb */}
                      <div style={{textAlign:"right",flexShrink:0}}>
                        <div style={{display:"flex",alignItems:"center",gap:5,justifyContent:"flex-end"}}>
                          <span style={{fontSize:9.5,color:T.inkSoft}}>{fu(p.P_base)}</span>
                          <span style={{fontSize:9,color:T.inkXsoft}}>→</span>
                          <span style={{fontSize:10,fontWeight:700,color:d<0?T.green:T.red}}>{fu(dvb)}</span>
                        </div>
                        <div style={{fontSize:9,fontWeight:700,color:d<0?T.green:T.red,textAlign:"right"}}>
                          {d<0?"-":"+"}{fu(Math.abs(d))} ({sg(-pct)}{(-pct).toFixed(1)}%)
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>

        {/* Footer acciones */}
        <div style={{padding:"14px 22px",borderTop:`1px solid ${T.borderSm}`,flexShrink:0,
          background:T.surface,display:"flex",alignItems:"center",justifyContent:"space-between",gap:10}}>
          <div style={{fontSize:10,color:T.inkSoft}}>
            {totalSel===0?"Selecciona al menos un proyecto para revertir":`Revertir ${totalSel} proyecto${totalSel>1?"s":""} a sus valores Apollo`}
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={onClose}
              style={{padding:"9px 18px",borderRadius:10,border:`1.5px solid ${T.borderSm}`,
                background:T.card,color:T.inkMid,fontSize:11,fontWeight:600,cursor:"pointer"}}>
              Cancelar
            </button>
            <button onClick={()=>totalSel>0&&onApply(selectedIds)}
              disabled={totalSel===0}
              style={{padding:"9px 20px",borderRadius:10,border:"none",
                background:totalSel>0?`linear-gradient(135deg,${T.red},${T.redDk})`:"#E0E0E0",
                color:totalSel>0?"#fff":"#aaa",fontSize:11,fontWeight:700,
                cursor:totalSel>0?"pointer":"not-allowed",
                boxShadow:totalSel>0?"0 4px 14px rgba(232,24,42,.3)":"none",
                transition:"all .18s",display:"flex",alignItems:"center",gap:6}}>
              <span style={{fontSize:13}}>↺</span>
              Devolver a base{totalSel>0?` (${totalSel})`:""}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}


function ViewTablero({overrides,setOverrides,onChangeLog}){
  const[drawer,setDrawer]=useState(null);
  const[fTipo,setFTipo]=useState(null);
  const[inverso,setInverso]=useState(false);
  const[search,setSearch]=useState("");
  const[resetModal,setResetModal]=useState(false);

  const capexP=useCallback(p=>calcCap(p,overrides[p.id]),[overrides]);
  const mDVB=useCallback(m=>m.proyectos.reduce((s,p)=>s+capexP(p),0),[capexP]);
  const mBase=m=>m.proyectos.reduce((s,p)=>s+p.P_base,0);
  const tBase=DATA.reduce((s,m)=>s+mBase(m),0);
  const tDVB=DATA.reduce((s,m)=>s+mDVB(m),0);
  const dT=dp(tDVB,tBase);
  const nOv=Object.keys(overrides).length;
  const tProj=DATA.reduce((s,m)=>s+m.proyectos.length,0);

  const CAT_KEYS=["Crecimiento","Modernización","Regulatorio","Operación","Transformación","Evolución","Comercial"];
  const porTipo=useMemo(()=>CAT_KEYS.map(k=>{
    const cfg=TIPO_CFG[k];
    const ms=DATA.filter(m=>m.tipo===k);
    const base=ms.reduce((s,m)=>s+mBase(m),0);
    const dvb=ms.reduce((s,m)=>s+mDVB(m),0);
    return{tipo:k,cfg,base,dvb,d:dp(dvb,base),count:ms.reduce((s,m)=>s+m.proyectos.length,0),pct:base/tBase};
  }),[overrides]);

  const filtered=useMemo(()=>DATA.filter(m=>(fTipo?m.tipo===fTipo:true)&&(!search||m.macro.toLowerCase().includes(search.toLowerCase()))),[fTipo,search]);
  const handleChange=useCallback((id,patch)=>{setOverrides(p=>({...p,[id]:{...p[id],...patch,pt:patch.pt?{...(p[id]?.pt||{}),...patch.pt}:p[id]?.pt,pf:patch.pf?{...(p[id]?.pf||{}),...patch.pf}:p[id]?.pf}}));},[setOverrides]);

  return(
    <div style={{maxWidth:1200,margin:"0 auto",padding:"24px 28px 60px"}}>

      {/* ── HERO BANNER ── */}
      <div className="fu hover-lift" style={{borderRadius:24,marginBottom:20,overflow:"hidden",position:"relative",
        background:`linear-gradient(150deg,${T.redDk} 0%,${T.red} 55%,#E85A20 100%)`,
        boxShadow:"0 24px 64px rgba(232,24,42,0.22)"}}>
        <NoiseSVG/>
        <div style={{position:"absolute",top:"-20%",right:"5%",width:450,height:450,borderRadius:"50%",background:"radial-gradient(circle,rgba(255,255,255,.1),transparent 65%)",pointerEvents:"none",zIndex:2}}/>
        <div style={{position:"absolute",inset:0,zIndex:2,backgroundImage:"radial-gradient(circle,rgba(255,255,255,.05) 1px,transparent 1px)",backgroundSize:"28px 28px",pointerEvents:"none"}}/>
        <div style={{position:"relative",zIndex:3,display:"grid",gridTemplateColumns:"1fr 1px 1fr 1px 1fr 1px 1fr"}}>
          {[
            {label:"CAPEX Total 2026",       val:fu(tBase),           sub:"Línea base Apollo v04 · "+DATA.reduce((s,m)=>s+m.proyectos.length,0)+" proyectos"},
            {label:"CAPEX DVB Propuesto", val:fu(tDVB), sub:`${sg(dT)}${dT.toFixed(1)}% vs base · ${nOv} proy. ajustados`, hot:nOv>0},
            {label:"Proyectos configurados", val:String(nOv),         sub:`de ${tProj} proyectos totales`},
            {label:"Eficiencia", val:nOv===0?"—":fu(Math.abs(tBase-tDVB)), hot:nOv>0&&tDVB<tBase, sub:nOv===0?"Ajusta drivers para calcular":"Base "+fu(tBase)+" − DVB "+fu(tDVB)},
          ].map((k,i)=>(
            <>
              {i>0&&<div key={`s${i}`} style={{background:"rgba(255,255,255,.12)",width:1,margin:"24px 0"}}/>}
              <div key={i} style={{padding:"28px 30px"}}>
                <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.45)",textTransform:"uppercase",letterSpacing:".14em",marginBottom:10}}>{k.label}</div>
                <div style={{fontSize:30,fontWeight:900,color:k.hot?"#86EFAC":"#fff",letterSpacing:"-.025em",lineHeight:1,marginBottom:6}}>{k.val}</div>
                <div style={{fontSize:i===3?10:11,color:i===3&&k.hot?"rgba(134,239,172,.9)":"rgba(255,255,255,.5)",fontWeight:i===3?700:400,lineHeight:1.5}}>{k.sub}</div>
              </div>
            </>
          ))}
        </div>
      </div>

      {/* ── MINI-KPIs POR TIPO ── */}
      <div className="fu1" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12,marginBottom:22}}>
        {porTipo.map(({tipo,cfg,base,dvb,d,count,pct})=>{
          const isActive=fTipo===tipo;
          const isDimmed=fTipo&&!isActive;
          return(
          <div key={tipo} onClick={()=>setFTipo(fTipo===tipo?null:tipo)}
            style={{background:isDimmed?"#F5F5F3":T.card,borderRadius:18,padding:"18px 20px",
              border:`1.5px solid ${isActive?cfg.c:isDimmed?"#E5E3DF":T.borderSm}`,
              boxShadow:isActive?`0 8px 28px ${cfg.c}28`:"0 2px 8px rgba(0,0,0,0.04)",
              cursor:"pointer",position:"relative",overflow:"hidden",
              opacity:isDimmed?0.42:1,
              transform:isActive?"translateY(-2px) scale(1.015)":"translateY(0) scale(1)",
              transition:"all .22s cubic-bezier(.22,1,.36,1)"}}>
            <div style={{position:"absolute",top:-10,right:-10,width:56,height:56,borderRadius:"50%",
              background:isDimmed?"transparent":`${cfg.c}10`,pointerEvents:"none"}}/>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:10}}>
              <TipoBadge tipo={tipo}/>
              {isActive&&<div style={{width:8,height:8,borderRadius:"50%",background:cfg.c,marginTop:3,flexShrink:0}}/>}
            </div>
            <div style={{fontSize:24,fontWeight:900,color:isDimmed?T.inkSoft:T.ink,letterSpacing:"-.02em",marginBottom:4}}>{fu(dvb)}</div>
            <div style={{fontSize:10,color:T.inkSoft,marginBottom:10}}>{fu(base)} base · {count} proy. · {(pct*100).toFixed(1)}%</div>
            <PBar pct={pct*100*3.5} color={isDimmed?"#D4D0CA":cfg.c} h={3}/>
            <div style={{fontSize:10,fontWeight:700,color:isDimmed?T.inkSoft:cc(d),marginTop:6}}>{sg(d)}{d.toFixed(1)}% vs base</div>
          </div>
        );})}
      </div>

      {/* ── TOOLBAR ── */}
      <div className="fu2" style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:16,fontWeight:800,color:T.ink,letterSpacing:"-.02em"}}>Macroproyectos · CAPEX 2026</div>
          <div style={{fontSize:11,color:T.inkSoft,marginTop:2}}>{filtered.length} macros · {filtered.reduce((s,m)=>s+m.proyectos.length,0)} proyectos</div>
        </div>
        <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
          <div style={{position:"relative"}}>
            <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:T.inkSoft,fontSize:14}}>⌕</span>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar macroproyecto…"
              style={{padding:"8px 12px 8px 30px",borderRadius:11,border:`1.5px solid ${T.borderSm}`,background:T.card,fontSize:11,color:T.ink,outline:"none",width:200,fontFamily:"'Outfit',system-ui",transition:"border-color .15s"}}
              onFocus={e=>e.target.style.borderColor=T.red} onBlur={e=>e.target.style.borderColor=T.borderSm}/>
          </div>
          <button onClick={()=>setInverso(true)} style={{display:"flex",alignItems:"center",gap:7,padding:"8px 16px",borderRadius:11,border:`1.5px solid ${T.borderSm}`,background:T.card,color:T.inkMid,fontSize:11,fontWeight:600,cursor:"pointer",transition:"all .15s"}}
            onMouseEnter={e=>e.currentTarget.style.borderColor=T.red} onMouseLeave={e=>e.currentTarget.style.borderColor=T.borderSm}>
            ⇄ Ajuste Inverso
            {nOv>0&&<span style={{background:T.red,color:"#fff",borderRadius:99,padding:"1px 8px",fontSize:9,fontWeight:900}}>{nOv}</span>}
          </button>
          {nOv>0&&<button onClick={()=>setResetModal(true)} style={{display:"flex",alignItems:"center",gap:6,padding:"8px 14px",borderRadius:10,border:`1.5px solid ${T.borderSm}`,background:T.surface,color:T.inkMid,fontSize:11,fontWeight:600,cursor:"pointer",transition:"all .15s"}} onMouseEnter={e=>{e.currentTarget.style.borderColor=T.red;e.currentTarget.style.color=T.red}} onMouseLeave={e=>{e.currentTarget.style.borderColor=T.borderSm;e.currentTarget.style.color=T.inkMid}}><span style={{fontSize:14}}>↺</span> Devolver a base</button>}
        </div>
      </div>

      {/* ── GRID DE CARDS (3 columnas) ── */}
      <div className="fu3" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:14}}>
        {filtered.map(m=>{
          const mb=mBase(m),md=mDVB(m),dM=dp(md,mb);
          const cfg=TIPO_CFG[m.tipo]||TIPO_CFG["Network Rollout"];
          const ovCount=m.proyectos.filter(p=>overrides[p.id]).length;
          const anyO=ovCount>0;
          const pctTotal=mb/tBase*100;

          return(
            <div key={m.macro} className="macro-card" style={{background:T.card,borderRadius:18,
              border:`1.5px solid ${anyO?T.red+"50":T.borderSm}`,
              boxShadow:anyO?"0 4px 20px rgba(232,24,42,.1)":"0 2px 8px rgba(0,0,0,0.04)",
              overflow:"hidden"}}>
              {/* Card header */}
              <div style={{padding:"16px 18px 12px",borderBottom:`1px solid ${T.borderSm}`,position:"relative"}}>
                {anyO&&<div style={{position:"absolute",top:0,left:0,right:0,height:3,background:`linear-gradient(90deg,${T.red},${T.redDk})`}}/>}
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:8}}>
                  <div style={{flex:1,minWidth:0,paddingRight:8}}>
                    <div style={{fontSize:11,fontWeight:700,color:T.ink,lineHeight:1.35,marginBottom:6,overflow:"hidden",display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical"}}>{m.macro}</div>
                    <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                      <TipoBadge tipo={m.tipo} sm/>
                      {anyO&&<span style={{display:"inline-flex",padding:"2px 8px",borderRadius:99,background:T.redXsoft,border:`1px solid ${T.redSoft}`,fontSize:8,fontWeight:700,color:T.red,textTransform:"uppercase",letterSpacing:".06em"}}>DVB {ovCount}/{m.proyectos.length}</span>}
                    </div>
                  </div>
                  <div style={{textAlign:"right",flexShrink:0}}>
                    <div style={{fontSize:18,fontWeight:900,color:cc(dM),letterSpacing:"-.02em",lineHeight:1}}>{fu(md)}</div>
                    {Math.abs(dM)>.3&&<div style={{fontSize:10,fontWeight:700,color:cc(dM),marginTop:2}}>{sg(dM)}{dM.toFixed(1)}%</div>}
                  </div>
                </div>
                {/* Barra de peso en el portfolio */}
                <div>
                  <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                    <span style={{fontSize:8.5,color:T.inkSoft}}>Base: {fu(mb)}</span>
                    <span style={{fontSize:8.5,color:T.inkSoft}}>{pctTotal.toFixed(1)}% del CAPEX</span>
                  </div>
                  <PBar pct={pctTotal*4.2} color={anyO?T.red:cfg.c} h={3}/>
                </div>
              </div>

              {/* Proyectos mini-lista */}
              <div>
                {m.proyectos.map((p,pi)=>{
                  const cap=capexP(p);const dP=dp(cap,p.P_base);
                  const hasO=!!overrides[p.id];const hasTree=!!PXQ_TREE[p.id];
                  const DT=DRV_N2[p.dt];
                  return(
                    <div key={p.id} className="row-hover" style={{display:"flex",alignItems:"center",gap:8,padding:"8px 16px",borderBottom:pi<m.proyectos.length-1?`1px solid ${T.borderSm}`:"none",background:pi%2===0?T.card:"#FAFAFA"}}>
                      <div style={{flex:1,minWidth:0}}>
                        <div style={{fontSize:11,fontWeight:hasO?700:400,color:hasO?T.ink:T.inkMid,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.n}</div>
                        <div style={{fontSize:8.5,color:DT?.color||T.inkSoft,marginTop:1,fontWeight:600,display:"flex",gap:5,alignItems:"center"}}>
                          <span>{DT?.icon}</span>
                          <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:120}}>{DT?.label}</span>
                        </div>
                      </div>
                      <div style={{textAlign:"right",flexShrink:0,minWidth:58}}>
                        <div style={{fontSize:11,fontWeight:700,color:cc(dP)}}>{fu(cap)}</div>
                        {Math.abs(dP)>.3&&<div style={{fontSize:8.5,fontWeight:700,color:cc(dP)}}>{sg(dP)}{dP.toFixed(1)}%</div>}
                      </div>
                      <button onClick={()=>setDrawer({proy:{...p},macroData:m,macroTipo:m.tipo})}
                        className="btn-primary"
                        style={{padding:"5px 10px",borderRadius:9,border:hasO?"none":`1.5px solid ${T.borderSm}`,
                          background:hasO?`linear-gradient(135deg,${T.red},${T.redDk})`:T.surface,
                          color:hasO?"#fff":T.inkMid,fontSize:9,fontWeight:700,cursor:"pointer",flexShrink:0,
                          boxShadow:hasO?`0 3px 10px rgba(232,24,42,.26)`:"none",transition:"all .15s",
                          display:"flex",alignItems:"center",gap:4,textTransform:"uppercase",letterSpacing:".06em",whiteSpace:"nowrap"}}>
                        {hasTree&&<span>🌳</span>}
                        {hasO?"DVB ✓":"DVB"}
                      </button>
                    </div>
                  );
                })}
              </div>

              {/* Card footer total macro */}
              <div style={{padding:"9px 16px",background:T.surface,borderTop:`1px solid ${T.borderSm}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span style={{fontSize:9,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",letterSpacing:".1em"}}>{m.proyectos.length} proyectos</span>
                <span style={{fontSize:11,fontWeight:700,color:cc(dM)}}>{sg(dM)}{dM.toFixed(1)}% DVB vs base</span>
              </div>
            </div>
          );
        })}
      </div>

      {drawer&&<Drawer proy={drawer.proy} macroData={drawer.macroData} macroTipo={drawer.macroTipo} ov={overrides[drawer.proy.id]}
        onChange={patch=>(onChangeLog||handleChange)(drawer.proy.id,patch,drawer.proy,drawer.macroData,drawer.macroTipo)} onClose={()=>setDrawer(null)}/>}
      {resetModal&&<ResetModal overrides={overrides} onApply={ids=>{setOverrides(p=>{const n={...p};ids.forEach(id=>delete n[id]);return n;});setResetModal(false);}} onClose={()=>setResetModal(false)}/>}
      {inverso&&<InverseModal overrides={overrides}
        onApply={upd=>setOverrides(p=>({...p,...Object.fromEntries(Object.entries(upd).map(([id,ov])=>[id,{...p[id],...ov}]))}))}
        onClose={()=>setInverso(false)}/>}
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   VISTA METODOLOGÍA
══════════════════════════════════════════════════════════════════════════ */
function ViewMetodologia(){
  const[open,setOpen]=useState(null);
  const ETAPAS=[
    {n:"01",title:"Línea Base",             statusC:T.green,   status:"COMPLETADO", output:"Mapa de inversiones",      body:"Se categoriza el CAPEX en paquetes, sub-paquetes y actividades. Se mapea el presupuesto, su pareto de gasto y las metas estratégicas del negocio.",bullets:["36 Macroproyectos · CAPEX total Apollo v04","101 proyectos con P actual y Q implícito","Categorización por tipo: RED, CLIENTE, IT, ADMIN","Jerarquía: Macro → Proyecto → Componente → Sub"]},
    {n:"02",title:"Drivers Técnicos → Q",   statusC:"#D97706", status:"EN CURSO",   output:"Q aspiracional por nivel",  body:"Se definen drivers técnicos para cada nivel del árbol y se dimensiona la necesidad Q de cada proyecto y componente con parámetros reales.",bullets:["9 tipos de driver: EoL, congestión, CRC, activaciones, migraciones…","Q calculado automáticamente desde parámetros IBP / plan","Bidireccional: editar Q → ajusta parámetro del driver","Árbol PxQ: Q propio por componente (no solo proyecto)"]},
    {n:"03",title:"Árboles P×Q por Niveles",statusC:"#D97706", status:"EN CURSO",   output:"Inductores por componente", body:"Se construyen árboles a 3 niveles. Cada nivel tiene Q (driver técnico) y P (driver financiero) propios, permitiendo optimizar en el nivel correcto.",bullets:["Nivel 1 — Proyecto: Q_total × P_promedio = CAPEX","Nivel 2 — Componente: HW, Civiles, M.O., O&M, SW (% asignado)","Nivel 3 — Sub-componente: línea de costo específica con qd + pd","CAPEX = Σ (Q_comp × P_comp) acumulado desde nivel 3"]},
    {n:"04",title:"Drivers Financieros → P", statusC:"#D97706", status:"EN CURSO",   output:"P benchmark por nivel",    body:"Se comparan los árboles de costo contra benchmarks financieros en cada nivel y se define el P objetivo por componente y sub-componente.",bullets:["6 drivers: AMX LatAm, cotización, histórico, CRC, IPC+ILA, spot","Gap visible por nivel: P_actual vs P_driver","Edición directa de P → recalcula CAPEX en cascada","Drill-down: oportunidad real en HW vs Civiles vs M.O."]},
    {n:"05",title:"Calibrar y Desplegar",    statusC:T.inkSoft, status:"PENDIENTE",  output:"Modelo PxQ operacional",   body:"Despliegue holístico con un único modelo de seguimiento por meta estratégica, driver y nivel PxQ. Ajuste inverso de target financiero.",bullets:["CAPEX DVB = Σ componentes (Q×P) ajustados","Ajuste inverso: target $M → redistribuir en P, Q o ambos","SPI = Q ejecutado / Q planeado (cierre mensual)","Monitoreo de gap driver en cada nivel del árbol"]},
  ];
  return(
    <div style={{maxWidth:1060,margin:"0 auto",padding:"28px 28px 52px"}}>
      <div className="fu hover-lift" style={{borderRadius:24,marginBottom:18,overflow:"hidden",position:"relative",background:`linear-gradient(150deg,${T.redDk},${T.red})`,boxShadow:"0 32px 80px rgba(0,0,0,0.22)"}}>
        <NoiseSVG/>
        <div style={{position:"absolute",top:"-20%",right:"5%",width:500,height:500,borderRadius:"50%",background:"radial-gradient(circle,rgba(255,255,255,.1),transparent 65%)",pointerEvents:"none",zIndex:2}}/>
        <div style={{position:"absolute",inset:0,zIndex:2,backgroundImage:"radial-gradient(circle,rgba(255,255,255,.05) 1px,transparent 1px)",backgroundSize:"28px 28px",pointerEvents:"none"}}/>
        <div style={{position:"relative",zIndex:3,display:"grid",gridTemplateColumns:"1fr 200px"}}>
          <div style={{padding:"48px 52px"}}>
            <div style={{fontSize:9,fontWeight:700,color:"rgba(255,255,255,.4)",letterSpacing:".2em",textTransform:"uppercase",marginBottom:10}}>Kearney · Claro Colombia · S&OP / IBP 2026</div>
            <div style={{fontSize:36,fontWeight:900,color:"#fff",letterSpacing:"-.025em",lineHeight:1.18,marginBottom:12}}>Drivers Value<br/><em style={{fontStyle:"italic",color:"rgba(255,255,255,.9)"}}>Budgeting</em></div>
            <div style={{fontSize:13,color:"rgba(255,255,255,.55)",lineHeight:1.8,maxWidth:500}}>Árbol PxQ a 3 niveles con drivers técnicos (Q) y financieros (P) propios en cada nivel — desde el proyecto hasta el sub-componente de costo.</div>
          </div>
          <div style={{borderLeft:"1px solid rgba(255,255,255,.08)",background:"rgba(255,255,255,.04)",display:"flex",flexDirection:"column",justifyContent:"center",padding:"44px 28px"}}>
            {[{n:fu(DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+p.P_base,0),0)),l:"CAPEX 2026"},{n:"36",l:"Macros"},{n:"101",l:"Proyectos"},{n:"3",l:"Niveles PxQ"}].map((s,i)=>(
              <div key={s.n} style={{padding:"14px 0",borderBottom:i<3?"1px solid rgba(255,255,255,.06)":"none"}}>
                <div style={{fontSize:26,fontWeight:900,color:"#fff",letterSpacing:"-.03em",lineHeight:1,marginBottom:3}}>{s.n}</div>
                <div style={{fontSize:10.5,color:"rgba(255,255,255,.38)",fontWeight:500}}>{s.l}</div>
              </div>
            ))}
          </div>
        </div>
      </div>

      <div className="fu1" style={{background:T.card,borderRadius:18,border:`1px solid ${T.borderSm}`,padding:"22px 24px",marginBottom:14,boxShadow:"0 2px 8px rgba(0,0,0,0.04)"}}>
        <div style={{fontSize:9,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",letterSpacing:".16em",marginBottom:16}}>Las 5 Etapas DVB · Clic para expandir</div>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {ETAPAS.map(e=>{const isO=open===e.n;return(
            <div key={e.n} style={{borderRadius:14,overflow:"hidden",border:`1.5px solid ${isO?T.red:T.borderSm}`,transition:"all .2s",boxShadow:isO?"0 4px 20px rgba(232,24,42,.1)":"0 1px 3px rgba(0,0,0,0.03)"}}>
              <div onClick={()=>setOpen(isO?null:e.n)} style={{display:"grid",gridTemplateColumns:"auto 1fr auto",alignItems:"center",background:isO?T.redXsoft:"#FAFAF8",cursor:"pointer",transition:"background .15s"}}>
                <div style={{padding:"14px 16px",display:"flex",alignItems:"center",gap:12}}>
                  <div style={{width:44,height:44,borderRadius:12,background:isO?T.redSoft:T.borderSm,border:`1.5px solid ${isO?T.red+"50":T.borderSm}`,display:"flex",alignItems:"center",justifyContent:"center",transition:"all .2s",flexShrink:0}}>
                    <span style={{fontSize:17,fontWeight:900,color:isO?T.red:T.inkMid}}>{e.n}</span>
                  </div>
                  <div>
                    <div style={{fontSize:13,fontWeight:700,color:isO?T.red:T.ink,transition:"color .15s"}}>{e.title}</div>
                    <span style={{fontSize:9,fontWeight:700,color:e.statusC,background:`${e.statusC}18`,borderRadius:99,padding:"2px 8px",border:`1px solid ${e.statusC}30`,textTransform:"uppercase",letterSpacing:".08em",marginTop:4,display:"inline-block"}}>{e.status}</span>
                  </div>
                </div>
                <div/>
                <div style={{padding:"0 20px",display:"flex",alignItems:"center",gap:10}}>
                  <div style={{background:T.red,borderRadius:10,padding:"8px 14px"}}>
                    <div style={{fontSize:9,color:"rgba(255,255,255,.6)",textTransform:"uppercase",letterSpacing:".1em",marginBottom:2}}>Output</div>
                    <div style={{fontSize:10.5,fontWeight:700,color:"#fff",maxWidth:150,lineHeight:1.3}}>{e.output}</div>
                  </div>
                  <div style={{width:26,height:26,borderRadius:"50%",background:isO?T.redSoft:T.borderSm,display:"flex",alignItems:"center",justifyContent:"center",transition:"all .2s"}}>
                    <span style={{fontSize:14,color:isO?T.red:T.inkSoft,transform:isO?"rotate(90deg)":"none",display:"block",transition:"transform .22s",fontWeight:700}}>›</span>
                  </div>
                </div>
              </div>
              {isO&&(<div className="fi" style={{padding:"20px 24px",borderTop:`1px solid ${T.red}22`,background:T.card}}>
                <p style={{fontSize:12.5,color:T.inkMid,lineHeight:1.8,margin:"0 0 16px"}}>{e.body}</p>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
                  {e.bullets.map((b,j)=>(
                    <div key={j} style={{display:"flex",gap:10,alignItems:"flex-start",background:T.surface,borderRadius:10,padding:"10px 12px",border:`1px solid ${T.borderSm}`}}>
                      <span style={{color:T.red,flexShrink:0,fontSize:12,fontWeight:700,marginTop:1}}>—</span>
                      <span style={{fontSize:11.5,color:T.inkMid,lineHeight:1.6}}>{b}</span>
                    </div>
                  ))}
                </div>
              </div>)}
            </div>
          );})}
        </div>
      </div>

      {/* Catálogo drivers */}
      <div className="fu2" style={{background:T.card,borderRadius:18,border:`1px solid ${T.borderSm}`,padding:"22px 24px",boxShadow:"0 2px 8px rgba(0,0,0,0.04)"}}>
        <div style={{fontSize:9,fontWeight:700,color:T.inkSoft,textTransform:"uppercase",letterSpacing:".16em",marginBottom:16}}>Catálogo de Drivers Técnicos · 9 tipos</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:10}}>
          {Object.entries(DRV_N2).map(([k,v])=>(
            <div key={k} className="hover-lift" style={{borderRadius:14,border:`1.5px solid ${v.color}30`,background:`${v.color}08`,padding:"14px 16px"}}>
              <div style={{display:"flex",gap:10,alignItems:"flex-start",marginBottom:8}}>
                <span style={{fontSize:20,flexShrink:0}}>{v.icon}</span>
                <div style={{fontSize:12,fontWeight:700,color:T.ink,lineHeight:1.3}}>{v.label}</div>
              </div>
              <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                {v.params.map(pr=>(<span key={pr.k} style={{fontSize:9,fontWeight:600,color:v.color,background:`${v.color}15`,borderRadius:99,padding:"2px 8px",border:`1px solid ${v.color}30`}}>{pr.l}</span>))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}


/* ══════════════════════════════════════════════════════════════════════════
   VIEW EFICIENCIAS DVB
   Muestra por proyecto: qué cambió, cuánto vale el ahorro/incremento,
   y qué driver lo generó.
══════════════════════════════════════════════════════════════════════════ */
function ViewEficiencias({overrides, tBase, tDVB}){
  const eficiencia = tBase - tDVB; // positivo = ahorro, negativo = incremento
  const nAjustes   = Object.keys(overrides).filter(id=>overrides[id]?._capex!=null||overrides[id]?.pf||overrides[id]?.pt).length;

  // Construir lista de proyectos ajustados con su delta
  const ajustados = DATA.flatMap(m=>
    m.proyectos.map(p=>{
      const ov  = overrides[p.id];
      const dvb = calcCap(p, ov);
      const delta = p.P_base - dvb; // positivo = ahorro
      const pct   = dp(dvb, p.P_base);
      if(!ov || Math.abs(delta) < 1) return null;

      // Determinar qué cambió
      const cambios = [];
      if(ov._capex != null && ov._tree) cambios.push({tipo:"árbol",label:"Árbol N3/N4 ajustado",icon:"🌳"});
      if(ov.dt && ov.dt !== p.dt) cambios.push({tipo:"driver_q",label:`Driver Q: ${DRV_N2[ov.dt]?.label||ov.dt}`,icon:"📐"});
      if(ov.df && ov.df !== p.df) cambios.push({tipo:"driver_p",label:`Driver P: ${DRV_FIN_N2[ov.df]?.label||ov.df}`,icon:"💰"});
      if(ov.pt) {
        const qBase = calcQ(p.dt, p.pt, 2);
        const qNew  = calcQ(ov.dt||p.dt, {...p.pt,...ov.pt}, 2);
        if(Math.abs(qNew-qBase)>0.5) cambios.push({tipo:"q",label:`Q: ${fn(qBase)} → ${fn(qNew)} ${p.m}`,icon:"📊"});
      }
      if(ov.pf) {
        const pBase = p.pf.pb;
        const pNew  = ov.pf.pb ?? pBase;
        if(Math.abs(pNew-pBase)>0.01) cambios.push({tipo:"p",label:`P: ${fu(pBase)} → ${fu(pNew)}/${p.m}`,icon:"🏷"});
      }
      if(cambios.length===0) cambios.push({tipo:"general",label:"Ajuste manual de CAPEX",icon:"✏️"});

      return {p, m: DATA.find(m=>m.proyectos.some(pp=>pp.id===p.id)), dvb, delta, pct, cambios, tipo:DATA.find(mm=>mm.proyectos.some(pp=>pp.id===p.id))?.tipo};
    }).filter(Boolean)
  ).sort((a,b)=>Math.abs(b.delta)-Math.abs(a.delta));

  const ahorros     = ajustados.filter(a=>a.delta>0);
  const incrementos = ajustados.filter(a=>a.delta<0);

  if(nAjustes===0) return(
    <div style={{maxWidth:900,margin:"0 auto",padding:"40px 28px",textAlign:"center"}}>
      <div style={{fontSize:56,marginBottom:16}}>💡</div>
      <div style={{fontSize:22,fontWeight:800,color:T.ink,marginBottom:8,letterSpacing:"-.02em"}}>
        Sin ajustes registrados
      </div>
      <div style={{fontSize:14,color:T.inkSoft,lineHeight:1.7,maxWidth:460,margin:"0 auto"}}>
        Abre cualquier proyecto en el Tablero DVB, ajusta los drivers Q y P, y presiona <b>💾 Guardar cambios</b>. Las eficiencias aparecerán aquí automáticamente.
      </div>
      <div style={{marginTop:28,display:"inline-flex",alignItems:"center",gap:10,
        padding:"12px 24px",borderRadius:14,background:T.card,
        border:`1.5px solid ${T.borderSm}`,boxShadow:"0 2px 12px rgba(0,0,0,0.05)"}}>
        <span style={{fontSize:24}}>🎯</span>
        <div style={{textAlign:"left"}}>
          <div style={{fontSize:12,fontWeight:700,color:T.ink}}>CAPEX DVB = CAPEX Base</div>
          <div style={{fontSize:11,color:T.inkSoft}}>CAPEX base = punto de partida DVB</div>
        </div>
      </div>
    </div>
  );

  return(
    <div style={{maxWidth:1100,margin:"0 auto",padding:"24px 28px 60px"}}>

      {/* ── HERO EFICIENCIA ── */}
      <div className="fu" style={{borderRadius:20,marginBottom:20,overflow:"hidden",position:"relative",
        background:`linear-gradient(135deg,${eficiencia>=0?"#064E3B":"#7F1D1D"} 0%,${eficiencia>=0?"#059669":"#E8182A"} 100%)`,
        boxShadow:`0 20px 50px ${eficiencia>=0?"rgba(5,150,105,.25)":"rgba(232,24,42,.25)"}`}}>
        <NoiseSVG/>
        <div style={{position:"relative",zIndex:1,display:"grid",gridTemplateColumns:"1fr 1px 1fr 1px 1fr 1px 1fr",padding:0}}>
          {[
            {label:"CAPEX Base 2026",   val:fu(tBase),        sub:"Línea base Apollo — punto de partida"},
            {label:"CAPEX DVB",         val:fu(tDVB),         sub:`${sg(dp(tDVB,tBase))}${dp(tDVB,tBase).toFixed(1)}% vs base`},
            {label:eficiencia>=0?"Ahorro identificado":"Incremento proyectado",
                                        val:fu(Math.abs(eficiencia)),
                                        sub:`${fu(tBase)} − ${fu(tDVB)} = ${eficiencia>=0?"ahorro":"sobrecosto"}`},
            {label:"Proyectos ajustados",val:String(nAjustes), sub:`${ahorros.length} ahorros · ${incrementos.length} incrementos`},
          ].map((k,i)=>(
            <>
              {i>0&&<div style={{background:"rgba(255,255,255,.15)",width:1,margin:"20px 0"}}/>}
              <div style={{padding:"24px 28px"}}>
                <div style={{fontSize:8.5,fontWeight:700,color:"rgba(255,255,255,.45)",
                  textTransform:"uppercase",letterSpacing:".14em",marginBottom:8}}>{k.label}</div>
                <div style={{fontSize:28,fontWeight:900,color:"#fff",
                  letterSpacing:"-.025em",lineHeight:1,marginBottom:5}}>{k.val}</div>
                <div style={{fontSize:10.5,color:"rgba(255,255,255,.55)",lineHeight:1.5}}>{k.sub}</div>
              </div>
            </>
          ))}
        </div>
      </div>

      {/* ── BARRA COMPOSICIÓN ── */}
      <div className="fu1" style={{background:T.card,borderRadius:14,padding:"16px 20px",
        marginBottom:18,border:`1px solid ${T.borderSm}`}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
          <div style={{fontSize:12,fontWeight:700,color:T.ink}}>Composición del ajuste por proyecto</div>
          <div style={{display:"flex",gap:12}}>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              <div style={{width:10,height:10,borderRadius:3,background:T.green}}/><span style={{fontSize:9,color:T.inkMid}}>Ahorro</span>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:5}}>
              <div style={{width:10,height:10,borderRadius:3,background:T.red}}/><span style={{fontSize:9,color:T.inkMid}}>Incremento</span>
            </div>
          </div>
        </div>
        {/* Barra waterfall */}
        <div style={{display:"flex",gap:2,height:32,borderRadius:8,overflow:"hidden",marginBottom:8}}>
          {ajustados.map((a,i)=>{
            const w = Math.abs(a.delta)/Math.max(...ajustados.map(x=>Math.abs(x.delta)))*100;
            const cfg = TIPO_CFG[a.tipo]||TIPO_CFG["Network Rollout"];
            return(
              <div key={i} title={`${a.p.n}: ${a.delta>=0?"+":"-"}${fu(Math.abs(a.delta))}`}
                style={{flex:Math.max(w,1),background:a.delta>=0?T.green:T.red,opacity:.85,
                  display:"flex",alignItems:"center",justifyContent:"center",minWidth:4,
                  borderRadius:4,cursor:"default",transition:"opacity .15s"}}/>
            );
          })}
          {/* Barra base ajustada sin cambios */}
          <div style={{flex:20,background:T.borderSm,borderRadius:4}}/>
        </div>
        <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
          {ajustados.slice(0,8).map((a,i)=>(
            <div key={i} style={{display:"flex",alignItems:"center",gap:4}}>
              <div style={{width:7,height:7,borderRadius:2,flexShrink:0,
                background:a.delta>=0?T.green:T.red}}/>
              <span style={{fontSize:8.5,color:T.inkMid,fontWeight:600}}>
                {a.p.id} <span style={{color:a.delta>=0?T.green:T.red,fontWeight:700}}>{a.delta>=0?"+":""}{fu(-a.delta*-1)}</span>
              </span>
            </div>
          ))}
          {ajustados.length>8&&<span style={{fontSize:8.5,color:T.inkSoft}}>+{ajustados.length-8} más</span>}
        </div>
      </div>

      {/* ── TABLA DE EFICIENCIAS ── */}
      <div className="fu2" style={{background:T.card,borderRadius:16,border:`1px solid ${T.borderSm}`,overflow:"hidden"}}>
        {/* Cabecera tabla */}
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 3fr",
          padding:"10px 18px",background:T.surface,
          borderBottom:`1px solid ${T.borderSm}`}}>
          {["Proyecto","Base","DVB","Δ CAPEX","Drivers ajustados"].map(h=>(
            <div key={h} style={{fontSize:8.5,fontWeight:700,color:T.inkSoft,
              textTransform:"uppercase",letterSpacing:".1em"}}>{h}</div>
          ))}
        </div>

        {ajustados.map((a,i)=>{
          const cfg = TIPO_CFG[a.tipo]||TIPO_CFG["Network Rollout"];
          const isAhorro = a.delta>=0;
          return(
            <div key={a.p.id} style={{
              display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 3fr",
              padding:"12px 18px",
              borderBottom:i<ajustados.length-1?`1px solid ${T.borderSm}`:"none",
              background:i%2===0?T.card:"#FAFAF8",
              transition:"background .12s"}} className="row-hover">

              {/* Proyecto */}
              <div>
                <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:3}}>
                  <div style={{width:24,height:24,borderRadius:7,background:cfg.bg,
                    border:`1px solid ${cfg.bdr}`,display:"flex",alignItems:"center",
                    justifyContent:"center",flexShrink:0}}>
                    <span style={{fontSize:11,color:cfg.c,fontWeight:800}}>{a.p.id.split(".")[0].slice(-2)}</span>
                  </div>
                  <div>
                    <div style={{fontSize:11,fontWeight:700,color:T.ink,lineHeight:1.2}}>{a.p.n}</div>
                    <div style={{fontSize:8.5,color:T.inkSoft}}>{a.m?.macro}</div>
                  </div>
                </div>
                <TipoBadge tipo={a.tipo} sm/>
              </div>

              {/* Base */}
              <div style={{display:"flex",alignItems:"center"}}>
                <div style={{fontSize:12,fontWeight:700,color:T.inkMid}}>{fu(a.p.P_base)}</div>
              </div>

              {/* DVB */}
              <div style={{display:"flex",alignItems:"center"}}>
                <div style={{fontSize:12,fontWeight:700,color:isAhorro?T.green:T.red}}>{fu(a.dvb)}</div>
              </div>

              {/* Delta */}
              <div style={{display:"flex",alignItems:"center"}}>
                <div style={{padding:"4px 10px",borderRadius:99,
                  background:isAhorro?T.greenBg:T.redBg,
                  border:`1px solid ${isAhorro?T.greenBdr:T.redSoft}`}}>
                  <div style={{fontSize:11,fontWeight:800,color:isAhorro?T.green:T.red}}>
                    {isAhorro?"-":"+"}${Math.abs(a.delta/1e6)<1?
                      (Math.abs(a.delta/1e3)).toFixed(0)+"K":
                      (Math.abs(a.delta/1e6)).toFixed(2)+"M"}
                  </div>
                  <div style={{fontSize:8,color:isAhorro?T.green:T.red,fontWeight:600}}>
                    {sg(-a.pct)}{(-a.pct).toFixed(1)}%
                  </div>
                </div>
              </div>

              {/* Drivers ajustados */}
              <div style={{display:"flex",flexWrap:"wrap",gap:4,alignItems:"center"}}>
                {a.cambios.map((c,ci)=>(
                  <div key={ci} style={{display:"flex",alignItems:"center",gap:4,
                    padding:"3px 8px",borderRadius:99,
                    background:T.surface,border:`1px solid ${T.borderSm}`}}>
                    <span style={{fontSize:11}}>{c.icon}</span>
                    <span style={{fontSize:8.5,fontWeight:600,color:T.inkMid}}>{c.label}</span>
                  </div>
                ))}
              </div>
            </div>
          );
        })}

        {/* Footer resumen */}
        <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 3fr",
          padding:"12px 18px",background:T.surface,
          borderTop:`2px solid ${T.border}`}}>
          <div style={{fontSize:11,fontWeight:800,color:T.ink}}>
            TOTAL ({ajustados.length} proyectos)
          </div>
          <div style={{fontSize:12,fontWeight:900,color:T.inkMid}}>{fu(tBase)}</div>
          <div style={{fontSize:12,fontWeight:900,color:eficiencia>=0?T.green:T.red}}>{fu(tDVB)}</div>
          <div>
            <div style={{fontSize:13,fontWeight:900,color:eficiencia>=0?T.green:T.red}}>
              {eficiencia>=0?"-":"+"}{fu(Math.abs(eficiencia))}
            </div>
            <div style={{fontSize:8.5,fontWeight:700,color:eficiencia>=0?T.green:T.red}}>
              {sg(-dp(tDVB,tBase))}{(-dp(tDVB,tBase)).toFixed(1)}% total
            </div>
          </div>
          <div style={{fontSize:10,color:T.inkSoft,display:"flex",alignItems:"center"}}>
            {ahorros.length} proyectos con ahorro · {incrementos.length} con incremento
          </div>
        </div>
      </div>

      {/* ── POR TIPO ── */}
      {Object.entries(TIPO_CFG).map(([tipo,cfg])=>{
        const items = ajustados.filter(a=>a.tipo===tipo);
        if(!items.length) return null;
        const dTipo = items.reduce((s,a)=>s+a.delta,0);
        return(
          <div key={tipo} className="fu3" style={{marginTop:12,background:T.card,borderRadius:14,
            padding:"12px 16px",border:`1.5px solid ${cfg.bdr}`}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
              <div style={{display:"flex",alignItems:"center",gap:8}}>
                <TipoBadge tipo={tipo}/>
                <span style={{fontSize:11,color:T.inkSoft}}>{items.length} proyectos ajustados</span>
              </div>
              <div style={{fontSize:14,fontWeight:900,color:dTipo>=0?T.green:T.red}}>
                {dTipo>=0?"Ahorro":"Incremento"}: {fu(Math.abs(dTipo))}
              </div>
            </div>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {items.map(a=>(
                <div key={a.p.id} style={{padding:"5px 10px",borderRadius:9,
                  background:a.delta>=0?T.greenBg:T.redBg,
                  border:`1px solid ${a.delta>=0?T.greenBdr:T.redSoft}`}}>
                  <div style={{fontSize:9.5,fontWeight:700,color:T.ink}}>{a.p.id} · {a.p.n}</div>
                  <div style={{fontSize:10,fontWeight:800,color:a.delta>=0?T.green:T.red}}>
                    {a.delta>=0?"-":"+"}{fu(Math.abs(a.delta))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}


/* ══════════════════════════════════════════════════════════════════════════
   APP ROOT
══════════════════════════════════════════════════════════════════════════ */

/* ══════════════════════════════════════════════════════════════════════════
   XLSX UTILS — leer/escribir Excel en browser via SheetJS
══════════════════════════════════════════════════════════════════════════ */
const Q_PARAMS_ORDER=['brutas','churn','base_migrable','tasa','total_activos','pct_eol',
  'meta_pct','comprometido','ejecutado','activos','frecuencia',
  'capacidad_actual','crecimiento_pct','objetivo','actual',
  'pipeline','win_rate','pct_congestion'];

// Detectar versión del schema y delegar al parser correcto
function parseXLSX(arrayBuffer){
  
  const wb=XLSX.read(arrayBuffer,{type:'arraybuffer',cellFormula:false,cellHTML:false});
  const schemaWs=wb.Sheets['DVB_Schema'];
  const schemaId=schemaWs ? (XLSX.utils.sheet_to_json(schemaWs,{header:1,defval:''})[0]||[])[0]||'' : '';
  if(schemaId==='DVB_SCHEMA_V2' && wb.Sheets['DVB_Granular']){
    return parseXLSX_V2(wb);
  }
  return parseXLSX_V1(wb);
}

// V1: hoja DVB_Input — nivel N2 proyecto
function parseXLSX_V1(wb){
  
  const ws=wb.Sheets['DVB_Input'];
  if(!ws) throw new Error('No se encontró la hoja DVB_Input');
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  const dataRows=rows.slice(3).filter(r=>r[2]&&String(r[2]).match(/^\d+\.\d+$/));
  return dataRows.map(r=>({
    id:    String(r[2]||'').trim(),
    macro: String(r[0]||'').trim(),
    tipo:  String(r[1]||'').trim(),
    n:     String(r[3]||'').trim(),
    m:     String(r[4]||'').trim(),
    prio:  String(r[5]||'').trim(),
    P_base:Number(r[6])||0,
    dt:    String(r[7]||'').trim(),
    df:    String(r[8]||'').trim(),
    pa:    Number(r[9])||0,
    pb:    Number(r[10])||0,
    pt:    Q_PARAMS_ORDER.reduce((acc,k,qi)=>{
      const v=r[11+qi];
      if(v!==''&&v!=='—'&&v!=null) acc[k]=Number(v)||0;
      return acc;
    },{}),
    _schema:'V1',
    _n3_overrides:{}, // sin datos granulares
  }));
}

// V2: hoja DVB_Granular — N3/N4
// Lee filas N4 (col A="N4"), extrae q_dvb (K) y p_dvb (L) editados
// Devuelve el mismo formato que V1 pero con _n3_overrides y _n4_overrides poblados
function parseXLSX_V2(wb){
  
  const ws=wb.Sheets['DVB_Granular'];
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
  // Cabeceras en fila 4 (índice 3), datos desde fila 5 (índice 4)
  const dataRows=rows.slice(4);

  // Agrupar por proyecto
  const byProj={};
  let curProj=null, curN3=null;

  dataRows.forEach(r=>{
    const nivel=String(r[0]||'').trim();
    const tipo =String(r[1]||'').trim();
    const code =String(r[2]||'').trim();
    const name =String(r[3]||'').trim();

    if(nivel==='N2'){
      curProj={
        id:code, macro:String(r[2]||''), tipo,
        n:name.replace(/\s*\[.*\]/,'').trim(),
        P_base:Number(r[12])||0,
        dt:String(r[5]||'').trim(),
        df:String(r[6]||'').trim(),
        _n3_overrides:{},_n4_overrides:{},_schema:'V2',
        pt:{},pf:{pb:Number(r[11])||0},pa:Number(r[9])||0,pb:Number(r[11])||0,
      };
      byProj[code]=curProj;
      curN3=null;
    }
    else if(nivel==='N3' && curProj){
      curN3={
        id:code.split('.').pop(),  // último segmento = n3_id
        label:name.replace(/^\s+/,'').replace(/^[\S]+\s+/,'').trim(),
        q_dvb: Number(r[10])||0,  // col K
        p_dvb: Number(r[11])||0,  // col L
        q_as:  Number(r[7])||0,   // col H
        p_as:  Number(r[9])||0,   // col J
      };
      // Guardar override N3 si Q fue editado
      const n3key=curN3.id;
      curProj._n3_overrides[n3key]={
        q_dvb:curN3.q_dvb, p_dvb:curN3.p_dvb,
        q_as:curN3.q_as,   p_as:curN3.p_as,
      };
    }
    else if(nivel==='N4' && curProj && curN3){
      const n4key=name.replace(/^\s+/,'').trim();
      if(!curProj._n4_overrides[curN3.id]) curProj._n4_overrides[curN3.id]={};
      curProj._n4_overrides[curN3.id][n4key]={
        q_dvb: Number(r[10])||0,  // col K
        p_dvb: Number(r[11])||0,  // col L
        q_as:  Number(r[7])||0,
        p_as:  Number(r[9])||0,
      };
    }
  });

  return Object.values(byProj);
}

// Generar xlsx con overrides aplicados usando SheetJS
function generateXLSX(proyRows, overrides){
  
  const wb=XLSX.utils.book_new();

  const Q_LABELS=['Activ. Brutas','Churn Estimado','Base Migrable','Tasa Migración',
    'Total Activos','% EoL Actual','Meta %','Comprometido CRC','Ya Ejecutado',
    'Activos Parque','Intervenc./Año','Capacidad Actual','Crec. % Proy.',
    'Objetivo HP/km','Actual Cubierto','Pipeline Calif.','Win-Rate','% Congestion'];

  const header1=['Macro','Tipo','ID','Nombre Proyecto','Unidad','Prioridad',
    'CAPEX Base USD','Driver Q','Driver P','P Unit. AS-IS USD','P Unit. DVB USD',
    ...Q_LABELS];

  const rows=[header1];
  let totalBase=0, totalDVB=0;

  proyRows.forEach(p=>{
    const ov=overrides[p.id]||{};
    const pt_merged={...p.pt,...(ov.pt||{})};
    const pb_dvb = ov.pf?.pb ?? ov._capex!=null
      ? (ov._capex / Math.max(1, calcQ(ov.dt||p.dt, pt_merged, 2)))
      : p.pb;
    const dvb_capex=calcCap(
      {id:p.id,P_base:p.P_base,dt:p.dt,df:p.df,pt:p.pt,pf:{pa:p.pa,pb:p.pb}},
      ov
    );
    totalBase+=p.P_base; totalDVB+=dvb_capex;
    const qvals=Q_PARAMS_ORDER.map(k=>{
      const v=pt_merged[k];
      return v!=null?v:'';
    });
    rows.push([p.macro,p.tipo,p.id,p.n,p.m,p.prio,
      p.P_base, p.dt, p.df, p.pa, pb_dvb, ...qvals]);
  });

  // Fila totales
  rows.push(['TOTAL CAPEX 2026','','','','','',totalBase,'','',
    '','', ...Q_PARAMS_ORDER.map(()=>'')]);

  // Fila delta
  const delta=totalDVB-totalBase;
  rows.push(['AJUSTE DVB vs BASE','','','','','',delta,'','',
    '','', ...Q_PARAMS_ORDER.map(()=>'')]);

  const ws=XLSX.utils.aoa_to_sheet(rows);

  // Anchos aprox
  ws['!cols']=[{wch:30},{wch:10},{wch:8},{wch:22},{wch:9},{wch:11},{wch:16},
    {wch:15},{wch:15},{wch:16},{wch:16},...Q_PARAMS_ORDER.map(()=>({wch:14}))];

  XLSX.utils.book_append_sheet(wb,'DVB_Output',ws);

  // Hoja resumen ajustes
  const adjRows=[['ID','Proyecto','Macro','CAPEX Base','CAPEX DVB','Δ USD','Δ %','Drivers ajustados']];
  proyRows.forEach(p=>{
    const ov=overrides[p.id];
    if(!ov) return;
    const dvb=calcCap({id:p.id,P_base:p.P_base,dt:p.dt,df:p.df,pt:p.pt,pf:{pa:p.pa,pb:p.pb}},ov);
    const delta=dvb-p.P_base;
    const tags=[];
    if(ov._tree) tags.push('Árbol N3/N4');
    if(ov.pt) tags.push('Driver Q');
    if(ov.pf) tags.push('Driver P');
    adjRows.push([p.id,p.n,p.macro,p.P_base,dvb,delta,(delta/p.P_base*100).toFixed(1)+'%',tags.join(', ')]);
  });
  const ws2=XLSX.utils.aoa_to_sheet(adjRows);
  ws2['!cols']=[{wch:8},{wch:22},{wch:30},{wch:16},{wch:16},{wch:14},{wch:8},{wch:24}];
  XLSX.utils.book_append_sheet(wb,'DVB_Ajustes',ws2);

  const buf=XLSX.write(wb,{bookType:'xlsx',type:'array'});
  const blob=new Blob([buf],{type:'application/octet-stream'});
  const url=URL.createObjectURL(blob);
  const a=document.createElement('a'); a.href=url;
  a.download=`DVB_Claro_${new Date().toISOString().slice(0,10)}.xlsx`;
  a.click(); URL.revokeObjectURL(url);
}

/* ══════════════════════════════════════════════════════════════════════════
   TAB CONTROL — Carga / Descarga plantilla + parámetros globales
══════════════════════════════════════════════════════════════════════════ */
function ViewControl({overrides, setOverrides, loadedFile, setLoadedFile, globalParams, setGlobalParams, auditLog=[], userEmail=''}){
  const fileRef=useRef();
  const [parsing,setParsing]=useState(false);
  const [parseResult,setParseResult]=useState(null); // {ok, rows, errors}
  const [exporting,setExporting]=useState(false);
  const [gpEdit,setGpEdit]=useState(false);

  // Proyectos planos para exportar
  const allProjs=DATA.flatMap(m=>m.proyectos.map(p=>({
    ...p, macro:m.macro, tipo:m.tipo,
    pa:p.pf?.pa??0, pb:p.pf?.pb??0
  })));

  const nOv=Object.keys(overrides).length;
  const tBase=DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+p.P_base,0),0);
  const tDVB=DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+calcCap(p,overrides[p.id]),0),0);
  const delta=tDVB-tBase;

  // Cargar archivo
  const handleFile=useCallback(e=>{
    const file=e.target.files[0]; if(!file) return;
    setParsing(true); setParseResult(null);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const rows=parseXLSX(ev.target.result);
        const errors=[];
        rows.forEach(r=>{
          if(!r.id) errors.push(`Fila sin ID`);
          if(!r.dt) errors.push(`${r.id}: sin driver Q`);
        });
        const found=rows.filter(r=>allProjs.find(p=>p.id===r.id));
        setParseResult({ok:errors.length===0,rows,errors,found:found.length,total:rows.length});
        setLoadedFile({name:file.name,rows,date:new Date().toLocaleString('es-CO')});
        setParsing(false);
      } catch(err){
        setParseResult({ok:false,errors:[err.message],rows:[],found:0,total:0});
        setParsing(false);
      }
    };
    reader.readAsArrayBuffer(file);
  },[allProjs]);

  // Aplicar plantilla al modelo (V1 y V2)
  const applyToModel=()=>{
    if(!parseResult?.ok||!parseResult.rows.length) return;
    const newOv={...overrides};
    parseResult.rows.forEach(r=>{
      const orig=allProjs.find(p=>p.id===r.id);
      if(!orig) return;

      if(r._schema==='V2'){
        // V2: calcular _capex desde N4 overrides
        // Reconstruir el árbol con los q_dvb/p_dvb del xlsx
        const n4Ov=r._n4_overrides||{};
        const n3Ov=r._n3_overrides||{};
        // Guardar como _tree serializado simple para que PxQPanel lo restaure
        const hasChanges=Object.keys(n4Ov).length>0||Object.keys(n3Ov).length>0;
        if(hasChanges){
          newOv[r.id]={
            ...(newOv[r.id]||{}),
            _n3_overrides:n3Ov,
            _n4_overrides:n4Ov,
          };
        }
        // También aplicar P unitario N2 si cambió
        if(r.pb && r.pb!==orig.pf?.pb){
          newOv[r.id]={...(newOv[r.id]||{}),pf:{...(newOv[r.id]?.pf||{}),pb:r.pb}};
        }
      } else {
        // V1: lógica original
        const ptDiff={};
        Q_PARAMS_ORDER.forEach(k=>{
          if(r.pt[k]!=null && r.pt[k]!==orig.pt[k]) ptDiff[k]=r.pt[k];
        });
        const pbChanged=r.pb!==orig.pb;
        if(Object.keys(ptDiff).length||pbChanged){
          newOv[r.id]={
            ...(newOv[r.id]||{}),
            ...(Object.keys(ptDiff).length?{pt:{...(newOv[r.id]?.pt||{}), ...ptDiff}}:{}),
            ...(pbChanged?{pf:{...(newOv[r.id]?.pf||{}), pb:r.pb}}:{})
          };
        }
      }
    });
    setOverrides(newOv);
    setParseResult(pr=>({...pr,applied:true}));
  };

  // Exportar
  const handleExport=()=>{
    if(!XLSX){
      alert('SheetJS no disponible. Asegúrate de que el artifact tiene acceso a CDN.');
      return;
    }
    setExporting(true);
    setTimeout(()=>{
      generateXLSX(allProjs,overrides);
      setExporting(false);
    },100);
  };

  const S={
    card:{background:T.card,borderRadius:18,padding:'22px 24px',
          boxShadow:'0 2px 8px rgba(0,0,0,.04)',border:`1px solid ${T.borderSm}`},
    label:{fontSize:9,fontWeight:700,color:T.inkSoft,textTransform:'uppercase',
           letterSpacing:'.1em',marginBottom:4},
    val:{fontSize:22,fontWeight:900,color:T.ink,letterSpacing:'-.02em'},
    btn:(c='#E8182A',bg='#FFF1F1')=>({
      display:'flex',alignItems:'center',gap:8,padding:'10px 20px',
      background:bg,border:`1.5px solid ${c}`,borderRadius:10,
      cursor:'pointer',fontSize:12,fontWeight:700,color:c,
    }),
    section:{marginTop:28},
    sectionTitle:{fontSize:11,fontWeight:800,color:T.ink,letterSpacing:'-.01em',
                  marginBottom:14,display:'flex',alignItems:'center',gap:8},
  };

  // Parámetros globales del modelo
  const GP_FIELDS=[
    {k:'ipc',   l:'IPC Colombia %',    u:'%',  tip:'Inflación anual Colombia'},
    {k:'ila',   l:'ILA Telecomunicaciones %', u:'%', tip:'Índice laboral sector TIC'},
    {k:'trm',   l:'TRM USD/COP',       u:'COP',tip:'Tasa representativa del mercado'},
    {k:'delta_amx', l:'Delta AMX Benchmark %', u:'%', tip:'Ajuste sobre precios de referencia AMX'},
    {k:'contingencia', l:'Factor Contingencia %', u:'%', tip:'% sobre CAPEX por imprevistos'},
  ];
  const GP_DEFAULT={ipc:5.2,ila:3.8,trm:4180,delta_amx:0,contingencia:3.0};
  const gp={...GP_DEFAULT,...globalParams};

  return(
    <div style={{padding:'24px 28px',maxWidth:1200,margin:'0 auto'}}>

      {/* KPIs superiores */}
      <div style={{display:'grid',gridTemplateColumns:'repeat(4,1fr)',gap:14,marginBottom:28}}>
        {[
          {l:'CAPEX Base',v:`$${(tBase/1e6).toFixed(1)}M`,c:T.inkMid,bg:T.card},
          {l:'CAPEX DVB',v:`$${(tDVB/1e6).toFixed(1)}M`,c:T.red,bg:'#FFF1F1'},
          {l:'Δ vs Base',v:`${delta>=0?'+':''}$${(delta/1e6).toFixed(1)}M`,
           c:delta<0?T.green:delta>0?T.red:T.inkMid,
           bg:delta<0?'#ECFDF5':delta>0?'#FFF1F1':T.card},
          {l:'Proyectos Ajustados',v:`${nOv} / ${DATA.reduce((s,m)=>s+m.proyectos.length,0)}`,c:T.violet,bg:PURPLE_LT??'#F5F3FF'},
        ].map(k=>(
          <div key={k.l} style={{...S.card,background:k.bg}}>
            <div style={S.label}>{k.l}</div>
            <div style={{...S.val,color:k.c}}>{k.v}</div>
          </div>
        ))}
      </div>

      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:22}}>

        {/* ── COLUMNA IZQUIERDA: Carga ── */}

        {/* ══ HISTORIAL DE AJUSTES ══ */}
        <div style={{marginBottom:24,background:T.card,borderRadius:14,
          border:`1px solid ${T.borderSm}`,overflow:"hidden"}}>
          <div style={{padding:"12px 18px",borderBottom:`1px solid ${T.borderSm}`,
            display:"flex",alignItems:"center",justifyContent:"space-between"}}>
            <div style={{fontSize:13,fontWeight:700,color:T.ink}}>Historial de ajustes</div>
            <span style={{fontSize:10,color:T.inkSoft,background:T.surface,
              padding:"2px 8px",borderRadius:99,border:`1px solid ${T.borderSm}`}}>
              {auditLog.length} registros
            </span>
          </div>
          {auditLog.length===0 ? (
            <div style={{padding:"28px",textAlign:"center"}}>
              <div style={{fontSize:22,marginBottom:8}}>📋</div>
              <div style={{fontSize:12,fontWeight:600,color:T.ink,marginBottom:4}}>Sin ajustes todavía</div>
              <div style={{fontSize:11,color:T.inkSoft}}>Modifica un Q o P en el tablero y aparecerá aquí</div>
            </div>
          ):(
            <div style={{overflowX:"auto"}}>
              <div style={{display:"grid",gridTemplateColumns:"110px 1fr 130px 90px 120px",
                padding:"7px 18px",background:T.surface,
                borderBottom:`1px solid ${T.borderSm}`,minWidth:600}}>
                {["Cuándo","Proyecto","Campo","Valor nuevo","Quién"].map(h=>(
                  <span key={h} style={{fontSize:8.5,fontWeight:700,color:T.inkSoft,
                    textTransform:"uppercase",letterSpacing:".07em"}}>{h}</span>
                ))}
              </div>
              {auditLog.slice(0,80).map((a,i)=>{
                const who  = a.user_name||a.user_email?.split("@")[0]||"";
                const when = new Date(a.created_at);
                const diff = Math.round((Date.now()-when)/60000);
                const t    = diff<1?"ahora":diff<60?`${diff}m`:diff<1440?`${Math.round(diff/60)}h`:when.toLocaleDateString("es-CO",{day:"2-digit",month:"short"});
                const val  = a.value_after?.v!=null
                  ? typeof a.value_after.v==="number"
                    ? a.value_after.v>9999?`$${(a.value_after.v/1e3).toFixed(0)}K`:`${a.value_after.v}`
                    : String(a.value_after.v)
                  : "—";
                const isMe = a.user_email===userEmail;
                return(
                  <div key={a.id} style={{display:"grid",
                    gridTemplateColumns:"110px 1fr 130px 90px 120px",
                    padding:"8px 18px",minWidth:600,
                    background:i%2===0?T.card:"#FAFAF9",
                    borderBottom:i<auditLog.slice(0,80).length-1?`1px solid ${T.borderSm}`:"none",
                    alignItems:"center"}}>
                    <span style={{fontSize:9.5,color:T.inkSoft}}>{t}</span>
                    <div style={{overflow:"hidden"}}>
                      <div style={{fontSize:11,fontWeight:600,color:T.ink,
                        whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>
                        {a.project_name||a.project_id}
                      </div>
                      <div style={{fontSize:9,color:T.inkSoft}}>{a.macro_name||""}</div>
                    </div>
                    <span style={{fontSize:10,color:T.inkMid}}>{a.field}</span>
                    <span style={{fontSize:11,fontWeight:800,color:T.green}}>{val}</span>
                    <div style={{display:"flex",alignItems:"center",gap:5}}>
                      <div style={{width:20,height:20,borderRadius:"50%",flexShrink:0,
                        background:isMe?T.red:"#E5E7EB",
                        display:"flex",alignItems:"center",justifyContent:"center",
                        fontSize:9,fontWeight:800,color:isMe?"#fff":T.inkSoft}}>
                        {who.charAt(0).toUpperCase()}
                      </div>
                      <span style={{fontSize:9.5,color:isMe?T.ink:T.inkMid,fontWeight:isMe?700:400}}>
                        {who}{isMe&&" (tú)"}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
        <div>
          <div style={S.card}>
            <div style={S.sectionTitle}>
              <span style={{fontSize:18}}>📥</span>
              <span>Cargar Plantilla DVB</span>
            </div>
            <p style={{fontSize:12,color:T.inkMid,lineHeight:1.6,marginBottom:16}}>
              Carga la plantilla Excel con los parámetros AS-IS o con ajustes DVB ya realizados.
              Los valores de <b>P Unitario DVB</b> y <b>Parámetros Q</b> se aplicarán al modelo.
            </p>

            {/* Drop zone */}
            <div
              onClick={()=>fileRef.current?.click()}
              style={{border:`2px dashed ${T.borderSm}`,borderRadius:12,padding:'28px 20px',
                      textAlign:'center',cursor:'pointer',background:T.surface,
                      transition:'all .2s'}}
              onMouseEnter={e=>e.currentTarget.style.borderColor=T.red}
              onMouseLeave={e=>e.currentTarget.style.borderColor=T.borderSm}
            >
              <div style={{fontSize:32,marginBottom:8}}>📊</div>
              <div style={{fontSize:13,fontWeight:700,color:T.ink}}>
                {loadedFile ? loadedFile.name : 'Haz clic para seleccionar'}
              </div>
              <div style={{fontSize:11,color:T.inkSoft,marginTop:4}}>
                {loadedFile
                  ? `Cargado el ${loadedFile.date}`
                  : 'DVB_Plantilla_Claro_2026.xlsx'}
              </div>
            </div>
            <input ref={fileRef} type="file" accept=".xlsx"
                   style={{display:'none'}} onChange={handleFile}/>

            {/* Resultado parseo */}
            {parsing&&(
              <div style={{marginTop:14,padding:'12px 16px',background:T.surface,
                           borderRadius:10,fontSize:12,color:T.inkMid}}>
                ⏳ Leyendo archivo...
              </div>
            )}
            {parseResult&&(
              <div style={{marginTop:14,padding:'14px 16px',
                           background:parseResult.ok?'#ECFDF5':'#FFF1F1',
                           borderRadius:10,border:`1px solid ${parseResult.ok?T.green:T.red}`}}>
                <div style={{fontSize:12,fontWeight:700,
                             color:parseResult.ok?T.green:T.red,marginBottom:6}}>
                  {parseResult.ok?'✅ Plantilla válida':'❌ Errores en la plantilla'}
                </div>
                <div style={{fontSize:11,color:T.inkMid}}>
                  {parseResult.total} filas leídas · {parseResult.found} proyectos coinciden con el modelo
                </div>
                {parseResult.errors.slice(0,3).map((e,i)=>(
                  <div key={i} style={{fontSize:10,color:T.red,marginTop:3}}>• {e}</div>
                ))}
                {parseResult.ok&&!parseResult.applied&&(
                  <button onClick={applyToModel}
                    style={{...S.btn(T.green,'#ECFDF5'),marginTop:12,width:'100%',
                            justifyContent:'center'}}>
                    ⚡ Aplicar al Modelo DVB
                  </button>
                )}
                {parseResult.applied&&(
                  <div style={{marginTop:10,fontSize:12,fontWeight:700,color:T.green}}>
                    ✅ Aplicado al modelo
                  </div>
                )}
              </div>
            )}
          </div>

          {/* Descargar plantilla vacía */}
          <div style={{...S.card,marginTop:16}}>
            <div style={S.sectionTitle}>
              <span style={{fontSize:18}}>📋</span>
              <span>Plantilla Base</span>
            </div>
            <p style={{fontSize:12,color:T.inkMid,lineHeight:1.6,marginBottom:14}}>
              Descarga la plantilla Excel con los 76 proyectos pre-cargados con datos Apollo AS-IS.
              Edita las columnas azules y verdes y vuelve a cargar.
            </p>
            <a href="https://raw.githubusercontent.com/placeholder/dvb/main/DVB_Plantilla.xlsx"
               style={{display:'none'}}>_</a>
            <div style={{fontSize:11,color:T.inkSoft,background:T.surface,
                         borderRadius:8,padding:'10px 14px',lineHeight:1.7}}>
              <b>Plantilla incluye:</b><br/>
              • Hoja <code>DVB_Input</code> — 76 proyectos con datos Apollo<br/>
              • Columnas editables: P Unitario DVB + 18 Parámetros Q<br/>
              • Hoja <code>DVB_Resumen</code> — totales por tipo/macro<br/>
              • Hoja <code>DVB_Schema</code> — mapa de columnas para validación
            </div>
            <div style={{marginTop:12,fontSize:11,color:T.inkSoft}}>
              💡 El archivo <code>DVB_Plantilla_Claro_2026.xlsx</code> se descarga junto con este artifact.
            </div>
          </div>
        </div>

        {/* ── COLUMNA DERECHA: Exportar + Parámetros Globales ── */}
        <div>
          <div style={S.card}>
            <div style={S.sectionTitle}>
              <span style={{fontSize:18}}>📤</span>
              <span>Exportar Modelo DVB Ajustado</span>
            </div>
            <p style={{fontSize:12,color:T.inkMid,lineHeight:1.6,marginBottom:14}}>
              Descarga la plantilla con los ajustes DVB aplicados en la sesión actual.
              Incluye CAPEX ajustado, Δ vs base y drivers modificados por proyecto.
            </p>
            {nOv>0?(
              <div style={{background:'#ECFDF5',borderRadius:10,padding:'12px 16px',
                           marginBottom:14,border:`1px solid ${T.green}`}}>
                <div style={{fontSize:12,fontWeight:700,color:T.green}}>
                  {nOv} proyectos ajustados en esta sesión
                </div>
                <div style={{fontSize:11,color:T.inkMid,marginTop:3}}>
                  Δ Total: {delta>=0?'+':''}${(delta/1e6).toFixed(2)}M vs base
                </div>
              </div>
            ):(
              <div style={{background:T.surface,borderRadius:10,padding:'12px 16px',
                           marginBottom:14,fontSize:12,color:T.inkSoft}}>
                Sin ajustes DVB en esta sesión. El export reflejará datos AS-IS.
              </div>
            )}
            <button onClick={handleExport} disabled={exporting}
              style={{...S.btn(),width:'100%',justifyContent:'center',
                      opacity:exporting?.6:1}}>
              {exporting?'⏳ Generando...':'⬇️ Descargar DVB_Claro_[fecha].xlsx'}
            </button>
            <div style={{marginTop:10,fontSize:10,color:T.inkSoft}}>
              Genera hojas: <code>DVB_Output</code> (modelo completo) + <code>DVB_Ajustes</code> (solo cambios)
            </div>
          </div>

          {/* Parámetros Globales */}
          <div style={{...S.card,marginTop:16}}>
            <div style={{...S.sectionTitle,justifyContent:'space-between'}}>
              <div style={{display:'flex',alignItems:'center',gap:8}}>
                <span style={{fontSize:18}}>⚙️</span>
                <span>Parámetros Globales del Modelo</span>
              </div>
              <button onClick={()=>setGpEdit(v=>!v)}
                style={{fontSize:10,fontWeight:700,color:T.red,background:'#FFF1F1',
                        border:`1px solid ${T.red}`,borderRadius:8,padding:'4px 10px',
                        cursor:'pointer'}}>
                {gpEdit?'✅ Guardar':'✏️ Editar'}
              </button>
            </div>
            <div style={{fontSize:11,color:T.inkSoft,marginBottom:14,lineHeight:1.6}}>
              Estos parámetros aplican a todo el modelo. Afectan drivers de precio P
              basados en ILA, IPC y benchmarks de mercado.
            </div>
            <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
              {GP_FIELDS.map(f=>(
                <div key={f.k} style={{background:T.surface,borderRadius:10,
                                       padding:'10px 14px',border:`1px solid ${T.borderSm}`}}>
                  <div style={{fontSize:9,fontWeight:700,color:T.inkSoft,
                               textTransform:'uppercase',letterSpacing:'.08em',marginBottom:4}}>
                    {f.l}
                  </div>
                  {gpEdit?(
                    <input type="number" step="0.1"
                      value={gp[f.k]}
                      onChange={e=>setGlobalParams(prev=>({...prev,[f.k]:parseFloat(e.target.value)||0}))}
                      style={{width:'100%',border:`1.5px solid ${T.red}`,borderRadius:6,
                              padding:'4px 8px',fontSize:13,fontWeight:700,
                              color:T.ink,background:'#FFF9F9',outline:'none'}}/>
                  ):(
                    <div style={{fontSize:16,fontWeight:800,color:T.ink}}>
                      {f.u==='%'?`${gp[f.k].toFixed(1)}%`:
                       f.u==='COP'?`$${gp[f.k].toLocaleString('es-CO')}`:
                       gp[f.k]}
                      <span style={{fontSize:9,color:T.inkSoft,marginLeft:4}}>{f.u}</span>
                    </div>
                  )}
                  <div style={{fontSize:9,color:T.inkSoft,marginTop:3}}>{f.tip}</div>
                </div>
              ))}
            </div>
            {gpEdit&&(
              <div style={{marginTop:12,padding:'10px 14px',background:'#FFFBEB',
                           borderRadius:8,border:'1px solid #D97706',fontSize:11,
                           color:'#92400E'}}>
                ⚠️ Cambiar IPC/ILA requiere re-calcular precios P en el Tablero.
                Los proyectos con driver P = <code>ila</code> o <code>historico</code> se recalcularán.
              </div>
            )}
          </div>

          {/* Log de cambios */}
          {nOv>0&&(
            <div style={{...S.card,marginTop:16}}>
              <div style={S.sectionTitle}>
                <span style={{fontSize:16}}>📝</span>
                <span>Registro de Ajustes DVB</span>
              </div>
              <div style={{maxHeight:200,overflowY:'auto',fontSize:11}}>
                {DATA.flatMap(m=>m.proyectos.map(p=>{
                  const ov=overrides[p.id]; if(!ov) return null;
                  const dvb=calcCap(p,ov);
                  const d=dvb-p.P_base;
                  return(
                    <div key={p.id} style={{display:'flex',alignItems:'center',
                      justifyContent:'space-between',padding:'6px 10px',
                      borderBottom:`1px solid ${T.borderSm}`,
                      background:d<0?'#ECFDF5':d>0?'#FFF1F1':T.surface}}>
                      <div>
                        <span style={{fontWeight:700,color:T.ink}}>{p.id}</span>
                        <span style={{color:T.inkMid,marginLeft:6}}>{p.n}</span>
                      </div>
                      <div style={{fontWeight:700,
                                   color:d<0?T.green:d>0?T.red:T.inkMid}}>
                        {d>=0?'+':''}{(d/1e6).toFixed(2)}M
                      </div>
                    </div>
                  );
                }).filter(Boolean))}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ══════════════════════════════════════════════════════════════════════════
   TAB ESCENARIOS — Multi-escenario con comparación
══════════════════════════════════════════════════════════════════════════ */
function ViewEscenarios({escenarios, setEscenarios, activeScen, setActiveScen, setOverrides, onSaveDB, onDeleteDB}){
  const [compareA,setCompareA]=useState(null);
  const [compareB,setCompareB]=useState(null);
  const [newName,setNewName]=useState('');
  const [creating,setCreating]=useState(false);

  const tBase=DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+p.P_base,0),0);

  const scenTotal=(ov)=>DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+calcCap(p,ov[p.id]),0),0);

  const createScenario=(name,fromOv={})=>{
    const id='scen_'+Date.now();
    const ns={id,name:name||`Escenario ${escenarios.length+1}`,
              overrides:{...fromOv},
              createdAt:new Date().toLocaleString('es-CO'),
              color:['#E8182A','#2563EB','#059669','#7C3AED','#D97706'][escenarios.length%5]};
    setEscenarios(prev=>[...prev,ns]);
    if(onSaveDB) onSaveDB(ns);
    return id;
  };

  const deleteScenario=(id)=>{
    const scen=escenarios.find(s=>s.id===id);
    setEscenarios(prev=>prev.filter(s=>s.id!==id));
    if(activeScen===id) setActiveScen(null);
    if(onDeleteDB&&scen) onDeleteDB(scen);
    if(compareA===id) setCompareA(null);
    if(compareB===id) setCompareB(null);
  };

  const activateScenario=(id)=>{
    const s=escenarios.find(e=>e.id===id);
    if(s){ setActiveScen(id); setOverrides({...s.overrides}); }
  };

  const S={
    card:{background:T.card,borderRadius:16,padding:'18px 20px',
          boxShadow:'0 2px 8px rgba(0,0,0,.04)',border:`1px solid ${T.borderSm}`},
    scenCard:(active,color)=>({
      background:T.card,borderRadius:14,padding:'16px 18px',
      border:`2px solid ${active?color:T.borderSm}`,cursor:'pointer',
      boxShadow:active?`0 0 0 3px ${color}22`:'none',
      transition:'all .15s',
    }),
    badge:(c)=>({
      display:'inline-flex',alignItems:'center',padding:'2px 8px',
      borderRadius:99,background:c+'22',
      fontSize:9,fontWeight:700,color:c,letterSpacing:'.06em'
    }),
  };

  // Datos de comparación
  const scenA=escenarios.find(e=>e.id===compareA);
  const scenB=escenarios.find(e=>e.id===compareB);

  const compData=useMemo(()=>{
    if(!scenA&&!scenB) return [];
    return DATA.flatMap(m=>m.proyectos.map(p=>{
      const vA=scenA?calcCap(p,scenA.overrides[p.id]):p.P_base;
      const vB=scenB?calcCap(p,scenB.overrides[p.id]):p.P_base;
      return {id:p.id,n:p.n,macro:m.macro,tipo:m.tipo,
              P_base:p.P_base,vA,vB,dA:vA-p.P_base,dB:vB-p.P_base,diff:vB-vA};
    }));
  },[scenA,scenB,escenarios]);

  const totA=compData.reduce((s,r)=>s+r.vA,0);
  const totB=compData.reduce((s,r)=>s+r.vB,0);

  return(
    <div style={{padding:'24px 28px',maxWidth:1300,margin:'0 auto'}}>

      {/* Header + crear */}
      <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:22}}>
        <div>
          <div style={{fontSize:18,fontWeight:900,color:T.ink}}>Escenarios DVB</div>
          <div style={{fontSize:12,color:T.inkMid,marginTop:2}}>
            Crea múltiples versiones del presupuesto y compáralas lado a lado
          </div>
        </div>
        <div style={{display:'flex',gap:10}}>
          {creating?(
            <div style={{display:'flex',gap:8,alignItems:'center'}}>
              <input value={newName} onChange={e=>setNewName(e.target.value)}
                placeholder="Nombre del escenario..."
                style={{padding:'8px 12px',border:`1.5px solid ${T.red}`,borderRadius:8,
                        fontSize:12,outline:'none',width:200}}/>
              <button onClick={()=>{
                if(newName.trim()){
                  const id=createScenario(newName.trim());
                  activateScenario(id); setNewName(''); setCreating(false);
                }
              }} style={{padding:'8px 14px',background:T.red,color:'#fff',
                          border:'none',borderRadius:8,fontWeight:700,cursor:'pointer',fontSize:12}}>
                Crear
              </button>
              <button onClick={()=>setCreating(false)}
                style={{padding:'8px 10px',background:T.surface,border:`1px solid ${T.borderSm}`,
                        borderRadius:8,cursor:'pointer',fontSize:12,color:T.inkMid}}>
                ✕
              </button>
            </div>
          ):(
            <>
              <button onClick={()=>setCreating(true)}
                style={{padding:'9px 18px',background:'#FFF1F1',border:`1.5px solid ${T.red}`,
                        borderRadius:10,cursor:'pointer',fontSize:12,fontWeight:700,color:T.red}}>
                + Nuevo Escenario
              </button>
              {escenarios.length>0&&(
                <button onClick={()=>{
                  const active=escenarios.find(e=>e.id===activeScen);
                  if(active){
                    const id=createScenario(`${active.name} (copia)`,active.overrides);
                    activateScenario(id);
                  }
                }} style={{padding:'9px 18px',background:T.surface,border:`1px solid ${T.borderSm}`,
                           borderRadius:10,cursor:'pointer',fontSize:12,fontWeight:700,color:T.inkMid}}>
                  📋 Duplicar activo
                </button>
              )}
            </>
          )}
        </div>
      </div>

      {/* Sin escenarios */}
      {escenarios.length===0&&(
        <div style={{...S.card,textAlign:'center',padding:'48px 24px'}}>
          <div style={{fontSize:48,marginBottom:12}}>🎭</div>
          <div style={{fontSize:15,fontWeight:700,color:T.ink,marginBottom:8}}>
            Sin escenarios todavía
          </div>
          <div style={{fontSize:12,color:T.inkMid,maxWidth:400,margin:'0 auto',lineHeight:1.7}}>
            Crea un escenario para guardar el estado actual del modelo DVB.
            Cada escenario captura todos los overrides y permite comparación.
          </div>
          <button onClick={()=>setCreating(true)}
            style={{marginTop:20,padding:'10px 24px',background:T.red,color:'#fff',
                    border:'none',borderRadius:10,fontWeight:700,cursor:'pointer',fontSize:13}}>
            + Crear primer escenario
          </button>
        </div>
      )}

      {/* Grid de escenarios */}
      {escenarios.length>0&&(
        <div style={{display:'grid',gridTemplateColumns:'repeat(auto-fill,minmax(260px,1fr))',
                     gap:14,marginBottom:28}}>
          {escenarios.map(s=>{
            const tot=scenTotal(s.overrides);
            const d=tot-tBase;
            const nAdj=Object.keys(s.overrides).length;
            const isActive=activeScen===s.id;
            return(
              <div key={s.id} style={S.scenCard(isActive,s.color)}
                   onClick={()=>activateScenario(s.id)}>
                <div style={{display:'flex',alignItems:'flex-start',justifyContent:'space-between',
                             marginBottom:10}}>
                  <div style={{display:'flex',alignItems:'center',gap:8}}>
                    <div style={{width:10,height:10,borderRadius:'50%',background:s.color}}/>
                    <div style={{fontSize:13,fontWeight:800,color:T.ink}}>{s.name}</div>
                  </div>
                  <div style={{display:'flex',gap:6}}>
                    {isActive&&<span style={S.badge(s.color)}>ACTIVO</span>}
                    <button onClick={e=>{e.stopPropagation();deleteScenario(s.id);}}
                      style={{background:'none',border:'none',cursor:'pointer',
                              fontSize:13,color:T.inkSoft,padding:'0 2px'}}>✕</button>
                  </div>
                </div>

                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:8,marginBottom:10}}>
                  <div style={{background:T.surface,borderRadius:8,padding:'8px 10px'}}>
                    <div style={{fontSize:8,fontWeight:700,color:T.inkSoft,
                                 textTransform:'uppercase',letterSpacing:'.08em'}}>DVB Total</div>
                    <div style={{fontSize:16,fontWeight:900,color:s.color}}>
                      ${(tot/1e6).toFixed(1)}M
                    </div>
                  </div>
                  <div style={{background:d<0?'#ECFDF5':d>0?'#FFF1F1':T.surface,
                               borderRadius:8,padding:'8px 10px'}}>
                    <div style={{fontSize:8,fontWeight:700,color:T.inkSoft,
                                 textTransform:'uppercase',letterSpacing:'.08em'}}>Δ vs Base</div>
                    <div style={{fontSize:16,fontWeight:900,
                                 color:d<0?T.green:d>0?T.red:T.inkMid}}>
                      {d>=0?'+':''}{(d/1e6).toFixed(1)}M
                    </div>
                  </div>
                </div>

                <div style={{display:'flex',justifyContent:'space-between',
                             fontSize:10,color:T.inkSoft}}>
                  <span>{nAdj} proyectos ajustados</span>
                  <span>{s.createdAt}</span>
                </div>

                {/* Barra de ahorro */}
                <div style={{marginTop:10,height:4,background:T.borderSm,borderRadius:99,overflow:'hidden'}}>
                  <div style={{width:`${Math.min(100,Math.abs(d)/tBase*100*10)}%`,
                               height:'100%',background:d<0?T.green:T.red,
                               borderRadius:99,transition:'width .3s'}}/>
                </div>

                {/* Botones */}
                <div style={{display:'flex',gap:8,marginTop:12}} onClick={e=>e.stopPropagation()}>
                  <button
                    onClick={()=>setCompareA(prev=>prev===s.id?null:s.id)}
                    style={{flex:1,padding:'5px',border:`1px solid ${compareA===s.id?scenA?.color||T.red:T.borderSm}`,
                            borderRadius:7,background:compareA===s.id?scenA?.color+'22':'none',
                            fontSize:10,fontWeight:700,cursor:'pointer',
                            color:compareA===s.id?scenA?.color||T.red:T.inkMid}}>
                    {compareA===s.id?'✓ A':'A'}
                  </button>
                  <button
                    onClick={()=>setCompareB(prev=>prev===s.id?null:s.id)}
                    style={{flex:1,padding:'5px',border:`1px solid ${compareB===s.id?scenB?.color||T.blue:T.borderSm}`,
                            borderRadius:7,background:compareB===s.id?scenB?.color+'22':'none',
                            fontSize:10,fontWeight:700,cursor:'pointer',
                            color:compareB===s.id?scenB?.color||T.blue:T.inkMid}}>
                    {compareB===s.id?'✓ B':'B'}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Comparador A vs B */}
      {(compareA||compareB)&&(
        <div style={S.card}>
          <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:18}}>
            <div style={{fontSize:14,fontWeight:800,color:T.ink}}>
              Comparación de Escenarios
            </div>
            <div style={{display:'flex',gap:16,fontSize:12}}>
              <span style={{display:'flex',alignItems:'center',gap:6}}>
                <span style={{width:10,height:10,borderRadius:'50%',
                              background:scenA?.color||T.red,display:'inline-block'}}/>
                <b>A:</b> {scenA?.name||'Base'} — ${(totA/1e6).toFixed(1)}M
              </span>
              <span style={{display:'flex',alignItems:'center',gap:6}}>
                <span style={{width:10,height:10,borderRadius:'50%',
                              background:scenB?.color||T.blue,display:'inline-block'}}/>
                <b>B:</b> {scenB?.name||'Base'} — ${(totB/1e6).toFixed(1)}M
              </span>
              <span style={{fontWeight:700,
                            color:totB-totA<0?T.green:totB-totA>0?T.red:T.inkMid}}>
                B−A: {totB-totA>=0?'+':''}{((totB-totA)/1e6).toFixed(2)}M
              </span>
            </div>
          </div>

          {/* Tabla comparación — solo proyectos con diferencia */}
          <div style={{overflowX:'auto'}}>
            <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
              <thead>
                <tr style={{background:T.surface}}>
                  {['ID','Proyecto','Tipo','Base',
                    `${scenA?.name||'A'}`,`Δ A`,
                    `${scenB?.name||'B'}`,`Δ B`,
                    'B − A'].map(h=>(
                    <th key={h} style={{padding:'8px 10px',textAlign:h==='ID'||h==='Tipo'?'center':'right',
                                        fontSize:9,fontWeight:700,color:T.inkSoft,
                                        textTransform:'uppercase',letterSpacing:'.06em',
                                        borderBottom:`2px solid ${T.borderSm}`}}>
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {compData
                  .filter(r=>Math.abs(r.dA)>100||Math.abs(r.dB)>100)
                  .sort((a,b)=>Math.abs(b.diff)-Math.abs(a.diff))
                  .map((r,i)=>{
                    const TIPO_C={'Network Rollout':'#2563EB','Network Modernization':'#D97706','Capacity Expansion':'#059669','Customer Investment':'#E8182A','Enterprise & Wholesale':'#7C3AED','Regulatory & Spectrum':'#B45309','IT & Digital':'#0891B2','Network Operations':'#6B7280','Gestión Administrativa':'#92400E'};
                    return(
                      <tr key={r.id} style={{background:i%2===0?T.surface:T.card,
                                             borderBottom:`1px solid ${T.borderSm}`}}>
                        <td style={{padding:'7px 10px',textAlign:'center',
                                    fontSize:10,fontWeight:700,color:T.inkMid}}>{r.id}</td>
                        <td style={{padding:'7px 10px',fontWeight:600,color:T.ink,maxWidth:180,
                                    overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>
                          {r.n}
                        </td>
                        <td style={{padding:'7px 10px',textAlign:'center'}}>
                          <span style={{background:(TIPO_C[r.tipo]||T.inkMid)+'20',
                                        color:TIPO_C[r.tipo]||T.inkMid,
                                        borderRadius:99,padding:'2px 7px',
                                        fontSize:9,fontWeight:700}}>
                            {r.tipo}
                          </span>
                        </td>
                        {[r.P_base,r.vA,r.dA,r.vB,r.dB,r.diff].map((v,vi)=>{
                          const isD=[2,4,5].includes(vi);
                          const col=isD?(v<0?T.green:v>0?T.red:T.inkMid):T.ink;
                          return(
                            <td key={vi} style={{padding:'7px 10px',textAlign:'right',
                                                 fontWeight:isD?700:400,color:col}}>
                              {isD&&v>=0?'+':''}{(v/1e6).toFixed(2)}M
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                {/* Total */}
                <tr style={{background:T.ink,borderTop:`2px solid ${T.borderSm}`}}>
                  <td colSpan={3} style={{padding:'9px 10px',fontSize:11,fontWeight:700,
                                          color:WHITE}}>TOTAL</td>
                  {[tBase,totA,totA-tBase,totB,totB-tBase,totB-totA].map((v,vi)=>{
                    const isD=[2,4,5].includes(vi);
                    const col=isD?(v<0?'#6EE7B7':v>0?'#FCA5A5':WHITE):WHITE;
                    return(
                      <td key={vi} style={{padding:'9px 10px',textAlign:'right',
                                           fontWeight:800,color:col,fontSize:12}}>
                        {isD&&v>=0?'+':''}{(v/1e6).toFixed(2)}M
                      </td>
                    );
                  })}
                </tr>
              </tbody>
            </table>
          </div>

          {compData.filter(r=>Math.abs(r.dA)<100&&Math.abs(r.dB)<100).length>0&&(
            <div style={{marginTop:10,fontSize:10,color:T.inkSoft,textAlign:'right'}}>
              {compData.filter(r=>Math.abs(r.dA)<100&&Math.abs(r.dB)<100).length} proyectos sin diferencia ocultos
            </div>
          )}
        </div>
      )}
    </div>
  );
}


export default function App(){
  // ══ TODOS LOS HOOKS PRIMERO ════════════════════════════════════

  // Estructura CAPEX desde Excel
  const [dataReady, setDataReady] = useState(false);

  // Auth
  const [session,   setSession]   = useState(null);
  const [profile,   setProfile]   = useState(null);
  const [authReady, setAuthReady] = useState(false);

  // App
  const [overrides,    setOverrides]    = useState({});
  const [auditLog,     setAuditLog]     = useState([]);
  const [tab,          setTab]          = useState("tablero");
  const [escenarios,   setEscenarios]   = useState([]);
  const [activeScen,   setActiveScen]   = useState(null);
  const [loadedFile,   setLoadedFile]   = useState(null);
  const [globalParams, setGlobalParams] = useState({ipc:5.2,ila:3.8,trm:4180,delta_amx:0,contingencia:3.0});

  // ══ initUser como useCallback — DEBE estar antes de los useEffect que lo usan ══
  const initUser = useCallback(async (s) => {
    touchLastSeen(s.user.id);
    const {data:p}    = await getProfile(s.user.id);
    setProfile(p);
    const {data:log}  = await fetchLog(s.user.id);
    setAuditLog(log||[]);
    const {data:scens}= await fetchScenarios(s.user.id);
    if(scens?.length){
      setEscenarios(scens.map((sc,i)=>({
        id:'db_'+sc.id, dbId:sc.id,
        name:sc.name, description:sc.description||'',
        overrides:sc.overrides||{},
        createdAt:new Date(sc.created_at).toLocaleString('es-CO'),
        color:['#E8182A','#2563EB','#059669','#7C3AED','#D97706'][i%5]
      })));
    }
  }, []);

  // Carga estructura CAPEX desde Excel (fallback: DATA hardcodeado)
  useEffect(()=>{
    loadEstructura()
      .then(d=>{ if(d.length>0) DATA=d; })
      .catch(()=>{})
      .finally(()=>setDataReady(true));
  },[]);

  // Estilos globales
  useEffect(()=>{
    const s=document.createElement("style");s.textContent=GS;document.head.appendChild(s);
    return()=>s.remove();
  },[]);

  // Auth listener
  useEffect(()=>{
    supabase.auth.getSession().then(({data:{session:s}})=>{
      setSession(s);
      if(s) initUser(s);
      setAuthReady(true);
    });
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_,s)=>{
      setSession(s);
      if(s) initUser(s);
      else { setProfile(null); setAuditLog([]); setEscenarios([]); }
    });
    return ()=>subscription.unsubscribe();
  },[initUser]);

  // ══ GUARDS — después de todos los hooks ════════════════════════
  if(!authReady||!dataReady) return(
    <div style={{minHeight:"100vh",background:"#F7F6F3",display:"flex",
      alignItems:"center",justifyContent:"center",
      fontFamily:"'Outfit',system-ui,sans-serif"}}>
      <div style={{fontSize:13,color:"#9CA3AF"}}>Cargando...</div>
    </div>
  );
  if(!session) return <AuthLogin onAuth={(s)=>{ setSession(s); initUser(s); }}/>;

  // ══ FUNCIONES — no son hooks, van después de los guards ════════

  const handleSignOut = async ()=>{
    await authSignOut();
    setSession(null); setProfile(null); setAuditLog([]); setEscenarios([]);
  };

  const handleChange = (id, patch) => {
    setOverrides(p=>{
      const prev=p[id]||{};
      return {...p,[id]:{...prev,...patch,
        pt:patch.pt?{...(prev.pt||{}),...patch.pt}:prev.pt,
        pf:patch.pf?{...(prev.pf||{}),...patch.pf}:prev.pf}};
    });
  };

  const handleChangeWithLog = (id, patch, proy, macro, categoria) => {
    handleChange(id, patch);
    const field = patch.pf?.pb!==undefined?'P unitario':
                  patch.pf?.pa!==undefined?'P actual':
                  patch.pt?'Parámetros Q':patch.dt?'Driver Q':'ajuste';
    pushLog({
      user_id:session.user.id,
      user_email:session.user.email,
      user_name:profile?.full_name||session.user.email.split('@')[0],
      project_id:id, project_name:proy?.n||id,
      macro_name:macro?.macro||'', categoria:categoria||'', field,
      value_before:null,
      value_after:patch.pf?.pb!=null?{v:patch.pf.pb}:patch.pf?.pa!=null?{v:patch.pf.pa}:null,
    }).then(()=>fetchLog(session.user.id).then(({data})=>setAuditLog(data||[])));
  };

  const saveScenToDB = async (scen, totalCapex, deltaPct) => {
    const {data} = await upsertScenario(session.user.id, {...scen,totalCapex,deltaPct});
    if(data) setEscenarios(prev=>prev.map(s=>s.id===scen.id?{...s,dbId:data.id}:s));
  };

  const deleteScenFromDB = async (scen) => {
    if(scen.dbId) await removeScenario(session.user.id, scen.dbId);
  };

  const tBase=DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+p.P_base,0),0);
  const tDVB=DATA.reduce((s,m)=>s+m.proyectos.reduce((sp,p)=>sp+calcCap(p,overrides[p.id]),0),0);
  const dT=dp(tDVB,tBase);
  const nOv=Object.keys(overrides).length;

  const TICK=DATA.flatMap(m=>m.proyectos.slice(0,2).map(p=>{const c=calcCap(p,overrides[p.id]),d=dp(c,p.P_base);return`${p.id} · ${p.n} · ${fu(c)} ${d>=0?"▲":"▼"}${Math.abs(d).toFixed(1)}%`;}));
const TABS=[{id:"tablero",label:"Tablero DVB",icon:"📊"},{id:"eficiencias",label:"Eficiencias DVB",icon:"💡"},{id:"escenarios",label:"Escenarios",icon:"🎭"},{id:"control",label:"Control",icon:"⚙️"},{id:"metodologia",label:"Metodología",icon:"🗂"}];

  return(
    <div style={{minHeight:"100vh",background:T.surface,display:"flex",flexDirection:"column"}}>
      {/* TICKER */}
      <div style={{background:T.card,borderBottom:`1px solid ${T.borderSm}`,padding:"4px 0",overflow:"hidden",flexShrink:0}}>
        <div className="tick-wrap">
          <div className="tick-inner">
            {[...TICK,...TICK].map((t,i)=>(
              <span key={i} style={{fontSize:9.5,fontWeight:600,color:T.inkSoft,padding:"0 20px",borderRight:`1px solid ${T.borderSm}`}}>{t}</span>
            ))}
          </div>
        </div>
      </div>

      {/* HEADER */}
      <header style={{background:"rgba(255,255,255,0.9)",backdropFilter:"blur(16px)",WebkitBackdropFilter:"blur(16px)",borderBottom:`1px solid ${T.borderSm}`,padding:"0 28px",height:58,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0,position:"sticky",top:0,zIndex:200,boxShadow:"0 1px 0 rgba(0,0,0,0.04),0 4px 12px rgba(0,0,0,0.03)"}}>
        <div style={{display:"flex",alignItems:"center",gap:16}}>
          <div style={{fontSize:13,fontWeight:900,color:T.ink,letterSpacing:".14em",textTransform:"uppercase"}}>KEARNEY</div>
          <div style={{width:1,height:26,background:T.borderSm}}/>
          <div>
            <div style={{fontSize:12,fontWeight:700,color:T.ink,letterSpacing:"-.01em"}}>Claro Colombia · DVB Command Center</div>
            <div style={{fontSize:9.5,color:T.inkSoft}}>Drivers Value Budgeting · CAPEX 2026</div>
          </div>
          <div style={{width:1,height:26,background:T.borderSm}}/>
          {/* Usuario */}
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:28,height:28,borderRadius:"50%",background:T.red,
              display:"flex",alignItems:"center",justifyContent:"center",
              fontSize:11,fontWeight:800,color:"#fff",flexShrink:0}}>
              {(profile?.full_name||session.user.email).charAt(0).toUpperCase()}
            </div>
            <div style={{lineHeight:1.2}}>
              <div style={{fontSize:11,fontWeight:700,color:T.ink}}>
                {profile?.full_name||session.user.email.split("@")[0]}
              </div>
              <div style={{fontSize:9,color:T.inkSoft,textTransform:"capitalize"}}>
                {profile?.role||"editor"}
              </div>
            </div>
            <button onClick={handleSignOut}
              style={{padding:"4px 10px",borderRadius:7,border:`1px solid ${T.borderSm}`,
                background:"transparent",color:T.inkSoft,fontSize:10,fontWeight:600,
                cursor:"pointer",marginLeft:4,fontFamily:"'Outfit',system-ui"}}>
              Salir
            </button>
          </div>
          <div style={{width:1,height:26,background:T.borderSm}}/>
          {/* Live totals en header */}
          <div style={{display:"flex",gap:6}}>
            {[{l:"Base",v:fu(tBase),c:T.inkMid},{l:"DVB",v:fu(tDVB),c:cc(dT)},...(Math.abs(dT)>.1?[{l:"Δ",v:`${dT>=0?"+":""}${dT.toFixed(1)}%`,c:cc(dT)}]:[])].map(k=>(
              <div key={k.l} style={{display:"flex",alignItems:"center",gap:6,padding:"4px 12px",background:T.surface,borderRadius:99,border:`1px solid ${T.borderSm}`}}>
                <span style={{fontSize:9,fontWeight:600,color:T.inkSoft}}>{k.l}</span>
                <span style={{fontSize:13,fontWeight:900,color:k.c,letterSpacing:"-.01em"}}>{k.v}</span>
              </div>
            ))}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{display:"flex",gap:2,background:T.surface,borderRadius:12,padding:"3px",border:`1px solid ${T.borderSm}`}}>
            {TABS.map(t=>(
              <button key={t.id} className="tab-pill" onClick={()=>setTab(t.id)} style={{padding:"6px 16px",borderRadius:10,border:"none",cursor:"pointer",background:tab===t.id?T.card:"transparent",color:tab===t.id?T.red:T.inkMid,fontWeight:tab===t.id?700:500,fontSize:11,borderBottom:tab===t.id?`2px solid ${T.red}`:"2px solid transparent",boxShadow:tab===t.id?"0 2px 6px rgba(0,0,0,0.08)":"none",whiteSpace:"nowrap",fontFamily:"'Outfit',system-ui"}}>{t.icon} {t.label}</button>
            ))}
          </div>
          <div style={{display:"flex",alignItems:"center",gap:6,padding:"5px 12px",background:T.surface,borderRadius:99,border:`1px solid ${T.borderSm}`}}>
            <span className="ldot"/>
            <span style={{fontSize:9,fontWeight:700,color:T.green,letterSpacing:".08em",textTransform:"uppercase"}}>Live</span>
            {nOv>0&&<span style={{fontSize:9,fontWeight:700,color:T.red}}>{nOv} dvb</span>}
          </div>
        </div>
      </header>

      {/* CONTENT */}
      <div style={{flex:1,overflow:"auto"}}>
        {tab==="tablero"    &&<ViewTablero overrides={overrides} setOverrides={setOverrides} onChangeLog={handleChangeWithLog}/>}
        {tab==="eficiencias"&&<ViewEficiencias overrides={overrides} tBase={tBase} tDVB={tDVB}/>}
        {tab==="escenarios" &&<ViewEscenarios escenarios={escenarios} setEscenarios={setEscenarios} activeScen={activeScen} setActiveScen={setActiveScen} setOverrides={setOverrides} onSaveDB={saveScenToDB} onDeleteDB={deleteScenFromDB}/>}
        {tab==="control"    &&<ViewControl overrides={overrides} setOverrides={setOverrides} loadedFile={loadedFile} setLoadedFile={setLoadedFile} globalParams={globalParams} setGlobalParams={setGlobalParams} auditLog={auditLog} userEmail={session?.user?.email||''}/>}
        {tab==="metodologia"&&<ViewMetodologia/>}
      </div>

      {/* FOOTER */}
      <footer style={{background:T.card,borderTop:`1px solid ${T.borderSm}`,height:50,display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 28px",flexShrink:0}}>
        {TABS.map((t,i)=>i===0?(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 14px",borderRadius:10,border:`1px solid ${T.borderSm}`,background:T.surface,cursor:"pointer",fontFamily:"'Outfit',system-ui",transition:"all .15s"}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor=T.red;e.currentTarget.style.color=T.red;}} onMouseLeave={e=>{e.currentTarget.style.borderColor=T.borderSm;e.currentTarget.style.color=T.ink;}}>
            <span style={{fontSize:13}}>←</span>
            <div style={{textAlign:"left"}}>
              <div style={{fontSize:8.5,color:T.inkSoft,fontWeight:700,textTransform:"uppercase",letterSpacing:".1em",lineHeight:1}}>Vista</div>
              <div style={{fontSize:11,fontWeight:700,lineHeight:1.4}}>{t.icon} {t.label}</div>
            </div>
          </button>
        ):null)}
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          {TABS.map(t=>(<div key={t.id} onClick={()=>setTab(t.id)} style={{width:tab===t.id?20:6,height:6,borderRadius:99,background:tab===t.id?T.red:T.borderSm,cursor:"pointer",transition:"all .25s cubic-bezier(.22,1,.36,1)"}}/>))}
        </div>
        {TABS.map((t,i)=>i===1?(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 14px",borderRadius:10,border:`1px solid ${T.borderSm}`,background:T.surface,cursor:"pointer",fontFamily:"'Outfit',system-ui",transition:"all .15s"}}
            onMouseEnter={e=>{e.currentTarget.style.borderColor=T.red;e.currentTarget.style.color=T.red;}} onMouseLeave={e=>{e.currentTarget.style.borderColor=T.borderSm;e.currentTarget.style.color=T.ink;}}>
            <div style={{textAlign:"right"}}>
              <div style={{fontSize:8.5,color:T.inkSoft,fontWeight:700,textTransform:"uppercase",letterSpacing:".1em",lineHeight:1}}>Vista</div>
              <div style={{fontSize:11,fontWeight:700,lineHeight:1.4}}>{t.icon} {t.label}</div>
            </div>
            <span style={{fontSize:13}}>→</span>
          </button>
        ):null)}
      </footer>
    </div>
  );
}
