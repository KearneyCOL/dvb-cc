/**
 * generar-estructura.cjs
 * Genera public/capex-estructura.xlsx desde el DATA hardcodeado.
 * Uso: node scripts/generar-estructura.cjs
 */

const XLSX = require("xlsx");
const path = require("path");

// ── DATA (copiado de App.jsx) ────────────────────────────────────────────────
const DATA = [
  // CATEGORÍA: RED - ACCESO FIJO
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

  // CATEGORÍA: RED - ACCESO MÓVIL
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

  // CATEGORÍA: RED - TRANSPORTE Y CORE
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

  // CATEGORÍA: RED - DATA CENTER
  {macro:"Expansión Capacidad DC",categoria:"Red",tipo:"Crecimiento",P_base:25000000,proyectos:[
    {id:"DC.EC.01",n:"Whitespace y Racks",m:"kW",prio:"CRECIMIENTO",P_base:10000000,dt:"crecimiento_trafico",pt:{capacidad_actual:500,crecimiento_pct:0.30},df:"mercado",pf:{pa:15000,pb:15000}},
    {id:"DC.EC.02",n:"Energía y Refrigeración",m:"MW",prio:"CRECIMIENTO",P_base:12000000,dt:"crecimiento_trafico",pt:{capacidad_actual:2,crecimiento_pct:0.50},df:"mercado",pf:{pa:8000000,pb:8000000}},
    {id:"DC.EC.03",n:"Edge Computing Propio",m:"nodo",prio:"CRECIMIENTO",P_base:3000000,dt:"expansion_geo",pt:{objetivo:20,actual:5},df:"benchmark_amx",pf:{pa:200000,pb:200000}},
  ]},
  {macro:"Operación Data Center",categoria:"Red",tipo:"Operación",P_base:4500000,proyectos:[
    {id:"DC.OP.01",n:"Mantenimiento y Refresh Infra DC",m:"m2",prio:"PMO",P_base:4500000,dt:"mantenimiento",pt:{activos:30000,frecuencia:1},df:"historico",pf:{pa:150,pb:150}},
  ]},

  // CATEGORÍA: ENTERPRISE Y B2B
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

  // CATEGORÍA: TECNOLOGÍA
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

  // CATEGORÍA: INVERSIÓN CLIENTES
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

  // CATEGORÍA: CORPORATIVO
  {macro:"Soporte Corporativo",categoria:"Corporativo",tipo:"Operación",P_base:6000000,proyectos:[
    {id:"CO.SC.01",n:"Instalaciones y Sedes",m:"m2",prio:"PMO",P_base:4000000,dt:"mantenimiento",pt:{activos:40000,frecuencia:1},df:"historico",pf:{pa:100,pb:100}},
    {id:"CO.SC.02",n:"Flota Vehicular",m:"vehículo",prio:"PMO",P_base:2000000,dt:"mantenimiento",pt:{activos:72,frecuencia:1},df:"mercado",pf:{pa:28000,pb:28000}},
  ]},
];

// ── Aplanar DATA en filas ────────────────────────────────────────────────────
function aplanar(data) {
  const filas = [];
  for (const m of data) {
    for (const p of m.proyectos) {
      const ptKeys = Object.keys(p.pt || {});
      filas.push({
        categoria:       m.categoria,
        macro:           m.macro,
        tipo:            m.tipo,
        P_base_macro:    m.P_base,
        id:              p.id,
        nombre:          p.n,
        medida:          p.m,
        prioridad:       p.prio,
        P_base:          p.P_base,
        driver_demanda:  p.dt,
        driver_financiero: p.df,
        pa:              p.pf?.pa ?? "",
        pb:              p.pf?.pb ?? "",
        pt_p1:           ptKeys[0] ?? "",
        pt_v1:           ptKeys[0] !== undefined ? p.pt[ptKeys[0]] : "",
        pt_p2:           ptKeys[1] ?? "",
        pt_v2:           ptKeys[1] !== undefined ? p.pt[ptKeys[1]] : "",
        pt_p3:           ptKeys[2] ?? "",
        pt_v3:           ptKeys[2] !== undefined ? p.pt[ptKeys[2]] : "",
      });
    }
  }
  return filas;
}

// ── Escribir Excel ───────────────────────────────────────────────────────────
const filas = aplanar(DATA);
const ws = XLSX.utils.json_to_sheet(filas);

// Anchos de columna
ws["!cols"] = [
  {wch:20},{wch:32},{wch:16},{wch:14},   // categoria, macro, tipo, P_base_macro
  {wch:12},{wch:36},{wch:12},{wch:14},   // id, nombre, medida, prioridad
  {wch:14},{wch:20},{wch:20},            // P_base, driver_demanda, driver_financiero
  {wch:12},{wch:12},                     // pa, pb
  {wch:16},{wch:12},{wch:16},{wch:12},{wch:16},{wch:12}, // pt_p1..pt_v3
];

const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, "Estructura");

const outPath = path.join(__dirname, "..", "public", "capex-estructura.xlsx");
XLSX.writeFile(wb, outPath);
console.log(`✅ Generado: ${outPath}`);
console.log(`   ${filas.length} proyectos exportados.`);
