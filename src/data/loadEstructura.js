/**
 * loadEstructura.js
 * Carga el árbol CAPEX desde /public/capex-estructura.xlsx
 *
 * Soporta dos formatos:
 *   1. "Árbol PxQ" — formato jerárquico del usuario (N0/N1/N2/N3)
 *   2. "Estructura"  — formato plano interno (generado por generar-estructura.cjs)
 */
import * as XLSX from "xlsx";

// ── Metadatos por nombre de proyecto (id, driver demanda, driver financiero, prioridad) ──
const META = {
  "FTTH Greenfield":                   { id:"AF.EC.01", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "FTTH Infill":                       { id:"AF.EC.02", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "FTTH Red Neutra":                   { id:"AF.EC.03", dt:"expansion_geo",       df:"mercado",       prio:"CRECIMIENTO"  },
  "Puertos OLT y Agregación":          { id:"AF.EK.01", dt:"activaciones",        df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "CPE Nuevos Clientes":               { id:"AF.EK.02", dt:"activaciones",        df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Servicio Universal Fijo":           { id:"AF.OR.01", dt:"cobertura_crc",       df:"benchmark_crc", prio:"REGULATORIO"  },
  "Migración HFC a Fibra":             { id:"AF.MT.01", dt:"migraciones",         df:"benchmark_amx", prio:"EBITDA"       },
  "Upgrade GPON a XGS-PON":           { id:"AF.MT.02", dt:"obsolescencia",       df:"benchmark_amx", prio:"EBITDA"       },
  "Refresh CPE y ONT":                 { id:"AF.MT.03", dt:"obsolescencia",       df:"benchmark_amx", prio:"PMO"          },
  "Mantenimiento Planta Externa":      { id:"AF.OP.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Inventario y Reposición":           { id:"AF.OP.02", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Nuevos Sitios 4G":                  { id:"AM.EC.01", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Nuevos Sitios 5G":                  { id:"AM.EC.02", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Small Cells y DAS":                 { id:"AM.EC.03", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Carriers Adicionales 4G":           { id:"AM.EK.01", dt:"congestion",          df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Carriers Adicionales 5G":           { id:"AM.EK.02", dt:"congestion",          df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Cobertura Mandatoria 4G":           { id:"AM.OR.01", dt:"cobertura_crc",       df:"benchmark_crc", prio:"REGULATORIO"  },
  "Cobertura Mandatoria 5G":           { id:"AM.OR.02", dt:"cobertura_crc",       df:"benchmark_crc", prio:"REGULATORIO"  },
  "Upgrade 4G a 5G":                   { id:"AM.MT.01", dt:"migraciones",         df:"benchmark_amx", prio:"EBITDA"       },
  "Refresh Equipos 4G EoL":            { id:"AM.MT.02", dt:"obsolescencia",       df:"benchmark_amx", prio:"PMO"          },
  "Refresh Equipos 5G EoL":            { id:"AM.MT.03", dt:"obsolescencia",       df:"benchmark_amx", prio:"PMO"          },
  "Swap de Vendor RAN":                { id:"AM.MT.04", dt:"obsolescencia",       df:"mercado",       prio:"EBITDA"       },
  "Reubicación de Sitios":             { id:"AM.MT.05", dt:"mantenimiento",       df:"mercado",       prio:"PMO"          },
  "Mantenimiento Sitios RAN":          { id:"AM.OP.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Energía y Climatización":           { id:"AM.OP.02", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Torres e Infraestructura Física":   { id:"AM.OP.03", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Fibra Nacional Long-haul":          { id:"TC.EB.01", dt:"crecimiento_trafico", df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Anillos Metropolitanos":            { id:"TC.EB.02", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Capacidad DWDM Lambdas":            { id:"TC.EB.03", dt:"crecimiento_trafico", df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Transporte 5G Xhaul":               { id:"TC.I5.01", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Core 5G Standalone":                { id:"TC.I5.02", dt:"crecimiento_trafico", df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Telco Cloud y vRAN":                { id:"TC.I5.03", dt:"crecimiento_trafico", df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Capacidad IRU Cables Submarinos":   { id:"TC.CI.01", dt:"crecimiento_trafico", df:"mercado",       prio:"CRECIMIENTO"  },
  "Participación Cables Submarinos":   { id:"TC.CI.02", dt:"crecimiento_trafico", df:"mercado",       prio:"CRECIMIENTO"  },
  "Refresh Routers Core IP/MPLS":      { id:"TC.MC.01", dt:"obsolescencia",       df:"benchmark_amx", prio:"EBITDA"       },
  "IRUs y Compartición Regulatoria":   { id:"TC.OP.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Mantenimiento TX y Core":           { id:"TC.OP.02", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Whitespace y Racks":                { id:"DC.EC.01", dt:"crecimiento_trafico", df:"mercado",       prio:"CRECIMIENTO"  },
  "Energía y Refrigeración":           { id:"DC.EC.02", dt:"crecimiento_trafico", df:"mercado",       prio:"CRECIMIENTO"  },
  "Edge Computing Propio":             { id:"DC.EC.03", dt:"expansion_geo",       df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Mantenimiento y Refresh Infra DC":  { id:"DC.OP.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Redes Dedicadas y SD-WAN":          { id:"EB.SC.01", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Conectividad Multi-Cloud":          { id:"EB.SC.02", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "5G Privado y Network Slicing":      { id:"EB.SC.03", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Colocation y Housing":              { id:"EB.SD.01", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Cloud Privado y Managed Services":  { id:"EB.SD.02", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Edge Computing Empresarial":        { id:"EB.SD.03", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Conectividad Social MINTIC":        { id:"EB.GV.01", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Verticales Estratégicos":           { id:"EB.GV.02", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Servicios Wholesale":               { id:"EB.WS.01", dt:"crecimiento_trafico", df:"mercado",       prio:"CRECIMIENTO"  },
  "Modernización CRM":                 { id:"TI.TB.01", dt:"migraciones",         df:"mercado",       prio:"TRANSFORMACIÓN"},
  "Billing Convergente":               { id:"TI.TB.02", dt:"migraciones",         df:"mercado",       prio:"TRANSFORMACIÓN"},
  "Orquestación y Automatización Red": { id:"TI.TO.01", dt:"migraciones",         df:"mercado",       prio:"TRANSFORMACIÓN"},
  "Canales Digitales App y Web":       { id:"TI.PD.01", dt:"mantenimiento",       df:"mercado",       prio:"TRANSFORMACIÓN"},
  "Data y Analytics":                  { id:"TI.PD.02", dt:"migraciones",         df:"mercado",       prio:"TRANSFORMACIÓN"},
  "Infraestructura y Licenciamiento":  { id:"TI.SI.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Ciberseguridad":                    { id:"TI.SI.02", dt:"mantenimiento",       df:"mercado",       prio:"PMO"          },
  "Mantenimiento Aplicativo":          { id:"TI.SI.03", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Mejoras BSS CRM y Billing":         { id:"TI.EP.01", dt:"mantenimiento",       df:"mercado",       prio:"EVOLUCIÓN"    },
  "Mejoras Canales Digitales":         { id:"TI.EP.02", dt:"mantenimiento",       df:"mercado",       prio:"EVOLUCIÓN"    },
  "Activaciones Fijo Hogar":           { id:"IC.AC.01", dt:"activaciones",        df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Activación y Subsidio Móvil":       { id:"IC.AC.02", dt:"activaciones",        df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Programas Retención y Lealtad":     { id:"IC.RC.01", dt:"migraciones",         df:"benchmark_amx", prio:"RETENCIÓN"    },
  "Migraciones ARPU-Accretive":        { id:"IC.RC.02", dt:"migraciones",         df:"benchmark_amx", prio:"CRECIMIENTO"  },
  "Adquisición Enterprise y PyME":     { id:"IC.IB.01", dt:"pipeline_b2b",        df:"mercado",       prio:"CRECIMIENTO"  },
  "Renovación Contratos B2B":          { id:"IC.IB.02", dt:"pipeline_b2b",        df:"mercado",       prio:"RETENCIÓN"    },
  "Instalaciones y Sedes":             { id:"CO.SC.01", dt:"mantenimiento",       df:"historico",     prio:"PMO"          },
  "Flota Vehicular":                   { id:"CO.SC.02", dt:"mantenimiento",       df:"mercado",       prio:"PMO"          },
};

// N0 del Excel → categoria en la app
const N0_CAT = {
  "Admin":     "Corporativo",
  "Comercial": "Inversión Clientes",
  "IT":        "Tecnología",
  "Red":       "Red",
};

// N2 (programa) → tipo
function inferTipo(n2) {
  const n = (n2 || "").toLowerCase();
  if (/transformación/.test(n))                                   return "Transformación";
  if (/evolución/.test(n))                                        return "Evolución";
  if (/obligacion|regulatori/.test(n))                            return "Regulatorio";
  if (/modernización|swap|refresh/.test(n))                       return "Modernización";
  if (/operación|sostenimiento|soporte corporativo/.test(n))      return "Operación";
  if (/retención|inversión b2b/.test(n))                          return "Comercial";
  return "Crecimiento";
}

// N2 (programa) → driver de demanda (dt)
function inferDt(n2) {
  const n = (n2 || "").toLowerCase();
  if (/obligacion|regulatori/.test(n))                            return "cobertura_crc";
  if (/modernización|refresh|swap/.test(n))                       return "obsolescencia";
  if (/operación|sostenimiento|soporte|mantenimiento/.test(n))    return "mantenimiento";
  if (/retención/.test(n))                                        return "migraciones";
  if (/transformación|evolución|plataforma|digital/.test(n))      return "migraciones";
  if (/adquisición consumidor/.test(n))                           return "activaciones";
  if (/inversión b2b|gobierno|vertical|mayorista|soluciones conectividad|servicios data center|pipeline/.test(n)) return "pipeline_b2b";
  if (/expansión capacidad|conectividad int|backbone|data center|infraestructura 5g/.test(n)) return "crecimiento_trafico";
  return "expansion_geo";  // default para expansión cobertura, etc.
}

// N0 y N2 → driver financiero (df)
function inferDf(n0, n2) {
  const n = (n2 || "").toLowerCase();
  if (/obligacion|regulatori/.test(n)) return "benchmark_crc";
  if (/operación|sostenimiento|soporte|mantenimiento|iru/.test(n)) return "historico";
  if (n0 === "Red")        return "benchmark_amx";
  return "mercado";
}

// tipo → prioridad
function inferPrio(tipo) {
  const t = (tipo || "").toLowerCase();
  if (t === "regulatorio")    return "REGULATORIO";
  if (t === "modernización")  return "EBITDA";
  if (t === "operación")      return "PMO";
  if (t === "transformación") return "TRANSFORMACIÓN";
  if (t === "evolución")      return "EVOLUCIÓN";
  if (t === "comercial")      return "RETENCIÓN";
  return "CRECIMIENTO";
}

// dt + Q → pt (parámetros del driver de demanda)
function inferPt(dt, Q) {
  Q = Number(Q) || 0;
  switch (dt) {
    case "expansion_geo":       return { objetivo: Q, actual: 0 };
    case "activaciones":        return { brutas: Q, churn: 0 };
    case "migraciones":         return { base_migrable: Q || 1, tasa: 1 };
    case "crecimiento_trafico": return { capacidad_actual: Q || 1, crecimiento_pct: 1 };
    case "obsolescencia":       return { total_activos: Q || 100, pct_eol: 1, meta_pct: 0 };
    case "cobertura_crc":       return { comprometido: Q, ejecutado: 0 };
    case "mantenimiento":       return { activos: Q || 1, frecuencia: 1 };
    case "pipeline_b2b":        return { pipeline: Q, win_rate: 1 };
    case "congestion":          return { total_activos: Q || 100, pct_congestion: 0.2, meta_pct: 0.05 };
    default:                    return { activos: Q || 1, frecuencia: 1 };
  }
}

// ── Sufijos por N1 para desambiguar N2 con nombre genérico ───────────────────
// (p.ej. "Operación" aparece bajo Acceso Fijo, Acceso Móvil, Data Center, etc.)
const N1_SUFFIX = {
  "Acceso Fijo":       "Fijo",
  "Acceso Móvil":      "Móvil",
  "Data Center":       "DC",
  "Transporte y Core": "Core",   // "Core" ya aparece en algunos N2 → se omite automáticamente
};

function macroNombre(n1, n2) {
  const sfx = N1_SUFFIX[n1];
  // Solo agregar sufijo si el N2 no lo contiene ya
  return sfx && !n2.includes(sfx) ? `${n2} ${sfx}` : n2;
}

// ── Generador de IDs para proyectos nuevos (no en META) ─────────────────────
const _idCounters = {};
function genId(n0, n1, n2) {
  const prefix = ((n0[0]||"") + (n1[0]||"") + n2.replace(/\s/g,"").slice(0,2)).toUpperCase();
  _idCounters[prefix] = (_idCounters[prefix] || 0) + 1;
  return `${prefix}.${String(_idCounters[prefix]).padStart(2,"0")}`;
}

// ── Parser principal: formato "Árbol PxQ" ───────────────────────────────────
function parseArbolPxQ(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  let curN0 = "", curN1 = "", curN2 = "", curMacroKey = "", curN2Capex = 0;
  const macroMap = new Map();

  for (const row of rows) {
    const nivel = String(row[0] || "").trim();
    if (!nivel || nivel === "Nivel") continue;

    const n0    = String(row[1] || "").trim();
    const n1    = String(row[2] || "").trim();
    const n2    = String(row[3] || "").trim();
    const n3    = String(row[4] || "").trim();
    const Q     = Number(row[7]) || 0;
    const unid  = String(row[8] || "").trim();
    const p     = Number(row[9]) || 0;
    const capex = Number(row[10]) || 0;

    if (nivel === "N0") { curN0 = n0; continue; }
    if (nivel === "N1 · Dominio") { curN0 = n0; curN1 = n1; continue; }

    if (nivel === "N2 · Programa") {
      curN0 = n0; curN1 = n1; curN2 = n2; curN2Capex = capex;
      curMacroKey = `${curN1}/${curN2}`;                    // clave única N1+N2
      const macroName = macroNombre(curN1, curN2);          // nombre de display
      const cat  = N0_CAT[curN0] || curN0;
      const tipo = inferTipo(curN2);
      if (!macroMap.has(curMacroKey)) {
        macroMap.set(curMacroKey, {
          macro: macroName,
          categoria: cat,
          tipo,
          P_base: curN2Capex,
          proyectos: [],
        });
      }
      continue;
    }

    if (nivel === "N3 · Proyecto" && n3) {
      curN0 = n0; curN1 = n1; curN2 = n2;
      curMacroKey = `${curN1}/${curN2}`;
      const macroName = macroNombre(curN1, curN2);
      const cat  = N0_CAT[curN0] || curN0;
      const tipo = inferTipo(curN2);

      if (!macroMap.has(curMacroKey)) {
        macroMap.set(curMacroKey, { macro: macroName, categoria: cat, tipo, P_base: curN2Capex, proyectos: [] });
      }

      const meta = META[n3];
      const dt   = meta?.dt   || inferDt(curN2);
      const df   = meta?.df   || inferDf(curN0, curN2);
      const prio = meta?.prio || inferPrio(tipo);
      const id   = meta?.id   || genId(curN0, curN1, curN2);
      const pt   = inferPt(dt, Q);

      macroMap.get(curMacroKey).proyectos.push({
        id,
        n:    n3,
        m:    unid || "unidad",
        prio,
        P_base: capex,
        dt,
        df,
        pt,
        pf: { pa: p, pb: p },
      });
    }
  }

  return Array.from(macroMap.values()).filter(m => m.proyectos.length > 0);
}

// ── Parser secundario: formato plano interno "Estructura" ────────────────────
function parseEstructura(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  const macroMap = new Map();
  for (const r of rows) {
    if (!macroMap.has(r.macro)) {
      macroMap.set(r.macro, {
        macro: String(r.macro || ""), categoria: String(r.categoria || ""),
        tipo: String(r.tipo || ""), P_base: Number(r.P_base_macro) || 0,
        proyectos: [],
      });
    }
    const pt = {};
    if (r.pt_p1 !== "" && r.pt_v1 !== "") pt[r.pt_p1] = Number(r.pt_v1);
    if (r.pt_p2 !== "" && r.pt_v2 !== "") pt[r.pt_p2] = Number(r.pt_v2);
    if (r.pt_p3 !== "" && r.pt_v3 !== "") pt[r.pt_p3] = Number(r.pt_v3);
    macroMap.get(r.macro).proyectos.push({
      id: String(r.id || ""), n: String(r.nombre || ""), m: String(r.medida || ""),
      prio: String(r.prioridad || ""), P_base: Number(r.P_base) || 0,
      dt: String(r.driver_demanda || ""), df: String(r.driver_financiero || ""),
      pt, pf: { pa: Number(r.pa) || 0, pb: Number(r.pb) || 0 },
    });
  }
  return Array.from(macroMap.values());
}

// ── Punto de entrada ─────────────────────────────────────────────────────────
// Busca los archivos en orden de prioridad
const EXCEL_FILES = ["/marco-capex.xlsx", "/capex-estructura.xlsx"];

export async function loadEstructura() {
  for (const path of EXCEL_FILES) {
    try {
      const res = await fetch(path);
      if (!res.ok) continue;
      const ab = await res.arrayBuffer();
      const wb = XLSX.read(ab, { type: "array" });
      if (wb.Sheets["Árbol PxQ"])  return parseArbolPxQ(wb.Sheets["Árbol PxQ"]);
      if (wb.Sheets["Estructura"]) return parseEstructura(wb.Sheets["Estructura"]);
    } catch { /* intentar siguiente */ }
  }
  throw new Error("No se encontró ningún archivo de estructura CAPEX en /public");
}
