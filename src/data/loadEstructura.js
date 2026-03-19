/**
 * loadEstructura.js
 * Carga el árbol CAPEX desde /public/capex-estructura.xlsx
 * y devuelve el array DATA con el mismo formato que tenía hardcodeado.
 *
 * Columnas esperadas en la hoja "Estructura":
 *   categoria, macro, tipo, P_base_macro,
 *   id, nombre, medida, prioridad, P_base,
 *   driver_demanda, driver_financiero, pa, pb,
 *   pt_p1, pt_v1, pt_p2, pt_v2, pt_p3, pt_v3
 */
import * as XLSX from "xlsx";

export async function loadEstructura() {
  const res = await fetch("/capex-estructura.xlsx");
  if (!res.ok) throw new Error(`No se pudo cargar capex-estructura.xlsx (${res.status})`);
  const ab = await res.arrayBuffer();
  const wb = XLSX.read(ab, { type: "array" });
  const ws = wb.Sheets["Estructura"];
  if (!ws) throw new Error('El archivo no contiene la hoja "Estructura"');
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return rowsToData(rows);
}

function rowsToData(rows) {
  const macroMap = new Map();
  for (const r of rows) {
    const macroKey = r.macro;
    if (!macroMap.has(macroKey)) {
      macroMap.set(macroKey, {
        macro:     String(r.macro     || ""),
        categoria: String(r.categoria || ""),
        tipo:      String(r.tipo      || ""),
        P_base:    Number(r.P_base_macro) || 0,
        proyectos: [],
      });
    }
    const pt = {};
    if (r.pt_p1 !== "" && r.pt_v1 !== "") pt[r.pt_p1] = Number(r.pt_v1);
    if (r.pt_p2 !== "" && r.pt_v2 !== "") pt[r.pt_p2] = Number(r.pt_v2);
    if (r.pt_p3 !== "" && r.pt_v3 !== "") pt[r.pt_p3] = Number(r.pt_v3);

    macroMap.get(macroKey).proyectos.push({
      id:   String(r.id      || ""),
      n:    String(r.nombre  || ""),
      m:    String(r.medida  || ""),
      prio: String(r.prioridad || ""),
      P_base: Number(r.P_base) || 0,
      dt:   String(r.driver_demanda    || ""),
      df:   String(r.driver_financiero || ""),
      pt,
      pf: {
        pa: Number(r.pa) || 0,
        pb: Number(r.pb) || 0,
      },
    });
  }
  return Array.from(macroMap.values());
}
