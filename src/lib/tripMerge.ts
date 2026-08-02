import * as XLSX from "xlsx";

export type Row = Record<string, any>;

export const CITIES = [
  "Bengaluru",
  "Mumbai",
  "Trichy",
  "Coimbatore",
  "Nashik",
  "Chennai",
  "Hyderabad",
  "Mum Watsapp",
] as const;

/** Some "cities" in the dashboard are really a different view of an existing city's sheet. */
const CITY_SHEET_ALIAS: Record<string, string> = {
  "Mum Watsapp": "Mumbai",
};
function sheetNameForCity(city: string): string {
  return CITY_SHEET_ALIAS[city] ?? city;
}

export const MERGE_TYPES = [
  { id: "fnv", label: "FNV Merge", needsBase: true, needsFg: false, needsLoading: false, needsGroups: true, needsBaseKind: true, needsGroundFile: false, needsRemarksFile: false },
  { id: "gro", label: "GRO Merge", needsBase: false, needsFg: true, needsLoading: true, needsGroups: false, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: false },
  { id: "fnv_gro_bread", label: "FNV + GRO Merge (Bread)", needsBase: false, needsFg: true, needsLoading: false, needsGroups: false, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: false },
  { id: "fnv_gro_milk", label: "FNV + GRO Merge (Milk)", needsBase: false, needsFg: true, needsLoading: true, needsGroups: false, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: false },
  
  { id: "gro_bread_milk", label: "GRO Merge (Bread + Milk)", needsBase: true, needsFg: false, needsLoading: false, needsGroups: true, needsBaseKind: true, needsGroundFile: false, needsRemarksFile: false },
  { id: "fnv_gro_cbe_trichy", label: "FNV + GRO Merge (Coimbatore / Trichy)", needsBase: false, needsFg: true, needsLoading: false, needsGroups: false, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: false },
  { id: "fnv_gro_chennai", label: "FNV + GRO Merge (Chennai drops)", needsBase: false, needsFg: true, needsLoading: false, needsGroups: true, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: false },
  { id: "watsapp", label: "WhatsApp Merge", needsBase: true, needsFg: false, needsLoading: false, needsGroups: true, needsBaseKind: true, needsGroundFile: false, needsRemarksFile: false },
  { id: "mum_watsapp_fnv", label: "FNV", needsBase: false, needsFg: true, needsLoading: true, needsGroups: true, needsBaseKind: true, needsGroundFile: true, needsRemarksFile: false },
  { id: "mum_egg_vehicle", label: "Egg Vehicle (Egg & Bread)", needsBase: false, needsFg: true, needsLoading: false, needsGroups: true, needsBaseKind: true, needsGroundFile: true, needsRemarksFile: false },
  { id: "mum_milk_vehicle_egg_bread", label: "Milk Vehicle (Egg & Bread)", needsBase: false, needsFg: true, needsLoading: false, needsGroups: false, needsBaseKind: false, needsGroundFile: false, needsRemarksFile: true },
] as const;

export type MergeTypeId = (typeof MERGE_TYPES)[number]["id"];

export const CITY_MERGE_TYPES: Record<string, MergeTypeId[]> = {
  Hyderabad: ["fnv"],
  Nashik: ["fnv"],
  Coimbatore: ["fnv", "fnv_gro_cbe_trichy"],
  Trichy: ["fnv", "fnv_gro_cbe_trichy"],
  Chennai: ["fnv", "fnv_gro_chennai"],
  Bengaluru: ["fnv", "gro", "fnv_gro_bread", "fnv_gro_milk", "gro_bread_milk", "watsapp"],
  Mumbai: ["fnv", "gro", "fnv_gro_bread", "fnv_gro_milk", "gro_bread_milk", "watsapp"],
  "Mum Watsapp": ["mum_watsapp_fnv", "mum_egg_vehicle", "mum_milk_vehicle_egg_bread"],
};

export function readWorkbook(file: ArrayBuffer): XLSX.WorkBook {
  return XLSX.read(file, { type: "array" });
}

export function readSheet(wb: XLSX.WorkBook, sheetName: string): Row[] {
  const s = wb.Sheets[sheetName];
  if (!s) return [];
  return XLSX.utils.sheet_to_json(s, { defval: "" });
}

/** Loose string compare */
const norm = (v: any) => String(v ?? "").trim().toLowerCase();

/** Find column key case-insensitively */
function pickCol(row: Row, ...names: string[]): string | undefined {
  const keys = Object.keys(row);
  for (const n of names) {
    const k = keys.find((k) => k.toLowerCase().replace(/[_\s]/g, "") === n.toLowerCase().replace(/[_\s]/g, ""));
    if (k) return k;
  }
  return undefined;
}

interface OutRow {
  SoId: string | number;
  TrmId: string | number;
}

/** For a set of items (customer names OR trip ids), collect SoIds from rows matching filter */
function collectSoIds(rows: Row[], keyCol: string, item: string, extraFilter?: (r: Row) => boolean): { soIds: any[]; trmIds: any[] } {
  const matched = rows.filter((r) => norm(r[keyCol]) === norm(item) && (extraFilter ? extraFilter(r) : true));
  return {
    soIds: matched.map((r) => r[pickCol(r, "SoId", "So Id") || "SoId"]).filter((v) => v !== "" && v != null),
    trmIds: matched.map((r) => r[pickCol(r, "TrmId", "Trmid", "Trm Id") || "TrmId"]).filter((v) => v !== "" && v != null),
  };
}

/**
 * Generic group processor:
 * For each group -> first item's TRMID (single) + every other item's SoIds -> rows.
 * If only 1 item in group, skipped (per spec).
 * `includeAnchorSoId`: when true, the first (anchor) item's own SoIds are ALSO
 * included in the output (mapped to its own TrmId) instead of being skipped.
 */
function processGroups(
  rows: Row[],
  keyCol: string,
  groups: string[][],
  trmFilter: (r: Row) => boolean,
  soIdFilter: (r: Row) => boolean,
  includeAnchorSoId = false,
): OutRow[] {
  const out: OutRow[] = [];
  // When the anchor's own SoId is included, a single-item "group" is valid —
  // it's a self-merge (same store's own TrmId + its own SoId), e.g. a single
  // point store in the Chennai drops merge. Otherwise, at least 2 items are
  // required (an anchor plus something to merge into it).
  const minItems = includeAnchorSoId ? 1 : 2;
  for (const group of groups) {
    const items = group.filter((x) => x && x.trim());
    if (items.length < minItems) continue;
    const firstTrmInfo = collectSoIds(rows, keyCol, items[0], trmFilter);
    const trmId = firstTrmInfo.trmIds[0];
    if (trmId == null) continue;
    const startIdx = includeAnchorSoId ? 0 : 1;
    for (let i = startIdx; i < items.length; i++) {
      const info = collectSoIds(rows, keyCol, items[i], soIdFilter);
      for (const s of info.soIds) out.push({ TrmId: trmId, SoId: s });
    }
  }
  return out;
}

/**
 * WhatsApp merge: for a group of items (trip ids or customer names),
 * find the item with the MOST SoIds (rows) matched in the source file —
 * its TrmId becomes the anchor. Every OTHER item's SoIds are mapped to
 * that anchor TrmId. The anchor item's own SoIds are skipped (not output).
 *
 * `itemFilter(item)` returns a row-filter to apply for that specific item —
 * lets different items in the same group use different type filters
 * (e.g. an exception-list store using Milk while everyone else uses FNV).
 * Defaults to no filter (every row counts).
 */
function processWatsappGroups(
  rows: Row[],
  keyCol: string,
  groups: string[][],
  itemFilter: (item: string) => (r: Row) => boolean = () => () => true,
): OutRow[] {
  const out: OutRow[] = [];
  for (const group of groups) {
    const items = group.filter((x) => x && x.trim());
    if (items.length < 2) continue;
    const infos = items.map((it) => collectSoIds(rows, keyCol, it, itemFilter(it)));
    let maxIdx = 0;
    for (let i = 1; i < infos.length; i++) {
      if (infos[i].soIds.length > infos[maxIdx].soIds.length) maxIdx = i;
    }
    const anchorTrm = infos[maxIdx].trmIds[0];
    if (anchorTrm == null) continue;
    for (let i = 0; i < infos.length; i++) {
      if (i === maxIdx) continue; // skip the anchor's own SoIds
      for (const s of infos[i].soIds) out.push({ TrmId: anchorTrm, SoId: s });
    }
  }
  return out;
}

/** Per-store mapping: for each row, TrmId from typeA and SoId from typeB (matched by CustomerName) */
function perStoreMap(fg: Row[], trmFilter: (r: Row) => boolean, soIdFilter: (r: Row) => boolean): OutRow[] {
  const out: OutRow[] = [];
  const byCust = new Map<string, { trms: any[]; sos: any[] }>();
  for (const r of fg) {
    const custKey = pickCol(r, "CustomerName") || "CustomerName";
    const cust = norm(r[custKey]);
    if (!cust) continue;
    if (!byCust.has(cust)) byCust.set(cust, { trms: [], sos: [] });
    const entry = byCust.get(cust)!;
    if (trmFilter(r)) {
      const t = r[pickCol(r, "TrmId") || "TrmId"];
      if (t != null && t !== "") entry.trms.push(t);
    }
    if (soIdFilter(r)) {
      const s = r[pickCol(r, "SoId") || "SoId"];
      if (s != null && s !== "") entry.sos.push(s);
    }
  }
  for (const { trms, sos } of byCust.values()) {
    if (!trms.length || !sos.length) continue;
    const trmId = trms[0];
    for (const s of sos) out.push({ TrmId: trmId, SoId: s });
  }
  return out;
}

// -- Loading sheet parsing (BLR-style) --
// The sheet has repeated triplets (Trip No, Loading Priority, Store Site ID).
// We scan raw cells, find header row containing "Trip No" and "Store Site ID",
// then read rows below. Trip No may be blank for continuation rows.
/**
 * Ground file (route/loading plan): reads Store Name + Vehicle type columns.
 * A merged Vehicle-type cell shows its value only on the first row of the
 * block and blank on the rows below it — those blank-vehicle rows are the
 * continuation of the merged block above. Each block's Store Names form
 * one merge group.
 */
export function parseGroundFile(wb: XLSX.WorkBook): string[][] {
  const sheetName = wb.SheetNames[0];
  const ws = sheetName ? wb.Sheets[sheetName] : undefined;
  if (!ws) return [];
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (!rows.length) return [];
  const storeCol = pickCol(rows[0], "StoreName", "Store Name") || "Store Name";
  const vehicleCol = pickCol(rows[0], "VehicleType", "Vehicle Type", "Vehicle type") || "Vehicle Type";
  const groups: string[][] = [];
  let current: string[] | null = null;
  for (const r of rows) {
    const store = String(r[storeCol] ?? "").trim();
    if (!store) continue;
    const vehicle = String(r[vehicleCol] ?? "").trim();
    if (vehicle || !current) {
      current = [store];
      groups.push(current);
    } else {
      current.push(store);
    }
  }
  return groups;
}

/**
 * Reads a store list with a trailing "Remark's"/"Remarks" column
 * (e.g. "Milk, Bread & Egg"). Only stores with a non-empty remark are
 * eligible for the Milk Vehicle (Egg & Bread) merge.
 */
export function parseEligibleStores(wb: XLSX.WorkBook): Set<string> {
  const sheetName = wb.SheetNames[0];
  const ws = sheetName ? wb.Sheets[sheetName] : undefined;
  if (!ws) return new Set();
  const rows: Row[] = XLSX.utils.sheet_to_json(ws, { defval: "" });
  if (!rows.length) return new Set();
  const storeCol = pickCol(rows[0], "StoreName", "Store Name") || "Store Name";
  const remarksCol = pickCol(rows[0], "Remarks", "Remark's", "Remark") || "Remark's";
  const set = new Set<string>();
  for (const r of rows) {
    const store = String(r[storeCol] ?? "").trim();
    const remark = String(r[remarksCol] ?? "").trim();
    if (store && remark) set.add(norm(store));
  }
  return set;
}

/**
 * Alternative to the Loading sheet for GRO / FNV+GRO(Milk): a vehicle/route
 * trip file with columns like "TRIP NO.", "Vehicle Type", "Store Name",
 * "Driver Name", "Drop Point Number", "Remark's". Trip No is repeated on
 * every row (not blank-continuation), so no forward-fill quirks — but we
 * forward-fill anyway to be safe. Header matching is loose (trailing dots /
 * "Store Name" vs "Store Site ID") so the same function works for either
 * layout.
 */
export function parseTripGroupFile(wb: XLSX.WorkBook): { tripNo: string; store: string }[] {
  const results: { tripNo: string; store: string }[] = [];
  const looseNorm = (v: any) => String(v ?? "").trim().toLowerCase().replace(/\.+$/, "");
  for (const sn of wb.SheetNames) {
    const sheet = wb.Sheets[sn];
    const aoa: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    let headerRowIdx = -1;
    let tCol = -1;
    let sCol = -1;
    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i];
      if (!row) continue;
      const tIdx = row.findIndex((c) => looseNorm(c) === "trip no");
      const sIdx = row.findIndex((c) => looseNorm(c) === "store name" || looseNorm(c) === "store site id");
      if (tIdx !== -1 && sIdx !== -1) {
        headerRowIdx = i;
        tCol = tIdx;
        sCol = sIdx;
        break;
      }
    }
    if (headerRowIdx === -1) continue;
    let currentTrip = "";
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const t = String(row[tCol] ?? "").trim();
      const store = String(row[sCol] ?? "").trim();
      if (t) currentTrip = t;
      if (store && currentTrip) results.push({ tripNo: currentTrip, store });
    }
  }
  return results;
}

export function parseLoadingSheet(wb: XLSX.WorkBook): { tripNo: string; store: string }[] {
  const results: { tripNo: string; store: string }[] = [];
  for (const sn of wb.SheetNames) {
    const sheet = wb.Sheets[sn];
    const aoa: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    // Find header row
    let headerRowIdx = -1;
    for (let i = 0; i < aoa.length; i++) {
      const row = aoa[i];
      if (row && row.some((c) => norm(c) === "trip no") && row.some((c) => norm(c) === "store site id")) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;
    const header = aoa[headerRowIdx];
    // Find all column pairs (Trip No idx, Store Site ID idx). We iterate columns and pair each Trip No with the NEXT Store Site ID column after it (before next Trip No).
    const tripIdxs: number[] = [];
    const storeIdxs: number[] = [];
    header.forEach((c, idx) => {
      if (norm(c) === "trip no") tripIdxs.push(idx);
      if (norm(c) === "store site id") storeIdxs.push(idx);
    });
    // pair
    const pairs: [number, number][] = [];
    for (const t of tripIdxs) {
      const s = storeIdxs.find((x) => x > t);
      if (s !== undefined) pairs.push([t, s]);
    }
    // Use only the FIRST pair to avoid duplicate reads (all pivots reflect same base)
    if (!pairs.length) continue;
    const [tCol, sCol] = pairs[0];
    let currentTrip = "";
    for (let r = headerRowIdx + 1; r < aoa.length; r++) {
      const row = aoa[r] || [];
      const t = String(row[tCol] ?? "").trim();
      const store = String(row[sCol] ?? "").trim();
      if (t) currentTrip = t;
      if (store && currentTrip) results.push({ tripNo: currentTrip, store });
    }
  }
  return results;
}

/** Group loading sheet stores by trip number */
export function groupsFromLoading(loading: { tripNo: string; store: string }[]): string[][] {
  const map = new Map<string, string[]>();
  for (const { tripNo, store } of loading) {
    if (!map.has(tripNo)) map.set(tripNo, []);
    map.get(tripNo)!.push(store);
  }
  return [...map.values()];
}

export interface RunInput {
  mergeType: MergeTypeId;
  city: string;
  baseKind?: "TripID" | "CustomerName";
  baseWb?: XLSX.WorkBook | null;
  fgWb?: XLSX.WorkBook | null;
  loadingWb?: XLSX.WorkBook | null;
  /** Alternative to loadingWb for "gro" / "fnv_gro_milk": a vehicle/route trip
   * file (TRIP NO. + Store Name columns) giving real trip->store groupings
   * directly, without needing the FNV/numeric-trip heuristics the Loading
   * sheet requires. When provided, it takes priority over loadingWb. */
  tripGroupsWb?: XLSX.WorkBook | null;
  groundWb?: XLSX.WorkBook | null;
  remarksWb?: XLSX.WorkBook | null;
  groups?: string[][]; // rows of items
}

const typeIs = (r: Row, val: string) => norm(r[pickCol(r, "Type") || "Type"]) === norm(val);
const sotypeIs = (r: Row, val: string) => norm(r[pickCol(r, "sotype_fnv_gro", "sotype") || "sotype_fnv_gro"]) === norm(val);

export function runMerge(input: RunInput): OutRow[] {
  const { mergeType, city, baseKind = "CustomerName" } = input;
  const baseRows = input.baseWb ? readSheet(input.baseWb, sheetNameForCity(city)) : [];
  const fgRows = input.fgWb ? readSheet(input.fgWb, sheetNameForCity(city)) : [];

  const keyOf = (r: Row) => {
    if (baseKind === "TripID") return pickCol(r, "TripID", "Trip ID", "TripId") || "TripID";
    return pickCol(r, "CustomerName", "Customer Name") || "CustomerName";
  };

  switch (mergeType) {
    case "fnv": {
      if (!baseRows.length || !input.groups) return [];
      const kc = keyOf(baseRows[0]);
      return processGroups(baseRows, kc, input.groups, (r) => sotypeIs(r, "FNV"), (r) => sotypeIs(r, "FNV"));
    }
    case "gro": {
      if (!fgRows.length) return [];
      let groups: string[][];
      if (input.tripGroupsWb) {
        // Vehicle/route trip file: same file can carry an "FNV" trip block
        // (Trip No literally = "FNV") alongside real numeric trips — those
        // FNV rows must NOT be treated as GRO trips, so exclude them here,
        // same as the numeric-only filtering done for the Loading sheet below.
        const parsed = parseTripGroupFile(input.tripGroupsWb);
        const numericOnly = parsed.filter((l) => !/fnv/i.test(String(l.tripNo)));
        groups = groupsFromLoading(numericOnly);
      } else {
        if (!input.loadingWb) return [];
        const loading = parseLoadingSheet(input.loadingWb);
        // Only trips whose Trip No is purely numeric (exclude FNV trips)
        const numericOnly = loading.filter((l) => /^\d+$/.test(String(l.tripNo).trim()));
        groups = groupsFromLoading(numericOnly);
      }
      const kc = pickCol(fgRows[0], "CustomerName") || "CustomerName";
      return processGroups(fgRows, kc, groups, (r) => typeIs(r, "Milk"), (r) => typeIs(r, "Milk"));
    }
    case "fnv_gro_bread": {
      if (!fgRows.length) return [];
      return perStoreMap(fgRows, (r) => typeIs(r, "FNV"), (r) => typeIs(r, "Bakery_and_Egg"));
    }
    case "fnv_gro_milk": {
      if (!fgRows.length) return [];
      if (input.tripGroupsWb) {
        // Same store-under-same-trip merge as "gro", but the anchor is
        // whichever store in the trip actually has an FNV row (not simply
        // the first row in the vehicle file — the FNV store can be listed
        // in any position), and every store in the trip (including that
        // anchor) contributes its Milk SoIds.
        const kc = pickCol(fgRows[0], "CustomerName") || "CustomerName";
        // Only the "FNV" trip block is relevant here — numeric (plain GRO)
        // trips in this same file should be ignored for the Milk merge.
        const parsed = parseTripGroupFile(input.tripGroupsWb);
        const fnvOnly = parsed.filter((l) => /fnv/i.test(String(l.tripNo)));
        const rawGroups = groupsFromLoading(fnvOnly);
        const groups = rawGroups.map((group) => {
          const fnvIdx = group.findIndex(
            (store) => collectSoIds(fgRows, kc, store, (r) => typeIs(r, "FNV")).trmIds[0] != null,
          );
          if (fnvIdx <= 0) return group; // already first, or no FNV store found (group is skipped downstream)
          // Move the actual FNV store to the front so it becomes the anchor.
          return [group[fnvIdx], ...group.slice(0, fnvIdx), ...group.slice(fnvIdx + 1)];
        });
        return processGroups(
          fgRows,
          kc,
          groups,
          (r) => typeIs(r, "FNV"),
          (r) => typeIs(r, "Milk"),
          true, // includeAnchorSoId: the anchor store's own Milk SoIds count too
        );
      }
      if (!input.loadingWb) return [];
      const loading = parseLoadingSheet(input.loadingWb);
      const custKey = pickCol(fgRows[0], "CustomerName") || "CustomerName";
      // Only trips whose Trip No contains "FNV"
      const fnvTrips = new Set(
        loading.filter((l) => /fnv/i.test(String(l.tripNo))).map((l) => l.tripNo),
      );
      const relevantStores = new Set(
        loading.filter((l) => fnvTrips.has(l.tripNo)).map((l) => norm(l.store)),
      );
      const out: OutRow[] = [];
      for (const store of relevantStores) {
        const rowsForStore = fgRows.filter((r) => norm(r[custKey]) === store);
        if (!rowsForStore.length) continue;
        const fnvRow = rowsForStore.find((r) => typeIs(r, "FNV"));
        if (!fnvRow) continue;
        const trm = fnvRow[pickCol(fnvRow, "TrmId") || "TrmId"];
        if (trm == null || trm === "") continue;
        for (const r of rowsForStore) {
          if (typeIs(r, "Milk")) {
            const s = r[pickCol(r, "SoId") || "SoId"];
            if (s != null && s !== "") out.push({ TrmId: trm, SoId: s });
          }
        }
      }
      return out;
    }
    case "gro_bread_milk": {
      if (!baseRows.length || !input.groups) return [];
      const out: OutRow[] = [];
      const kc = keyOf(baseRows[0]);
      const tripKey = pickCol(baseRows[0], "TripID", "Trip ID", "TripId") || "TripID";
      const trmKey = pickCol(baseRows[0], "TrmId", "Trmid", "Trm Id") || "TrmId";
      const soKey = pickCol(baseRows[0], "SoId", "So Id") || "SoId";
      for (const grp of input.groups) {
        const items = grp.filter((x) => x && x.trim());
        if (items.length < 2) continue;
        const rowsInGroup = baseRows.filter((r) => items.some((it) => norm(r[kc]) === norm(it)));
        if (!rowsInGroup.length) continue;
        const tripToSoIds = new Map<string, any[]>();
        const tripToTrm = new Map<string, any>();
        for (const r of rowsInGroup) {
          const tid = String(r[tripKey] ?? "").trim();
          if (!tid) continue;
          if (!tripToSoIds.has(tid)) tripToSoIds.set(tid, []);
          const s = r[soKey];
          if (s != null && s !== "") tripToSoIds.get(tid)!.push(s);
          const t = r[trmKey];
          if (t != null && t !== "" && !tripToTrm.has(tid)) tripToTrm.set(tid, t);
        }
        let anchorTrm: any = null;
        for (const [tid, sos] of tripToSoIds) {
          if (sos.length > 1 && tripToTrm.has(tid)) {
            anchorTrm = tripToTrm.get(tid);
            break;
          }
        }
        if (anchorTrm == null) {
          for (const t of tripToTrm.values()) {
            if (t != null && t !== "") { anchorTrm = t; break; }
          }
        }
        if (anchorTrm == null) continue;
        for (const r of rowsInGroup) {
          const s = r[soKey];
          if (s != null && s !== "") out.push({ TrmId: anchorTrm, SoId: s });
        }
      }
      return out;
    }
    case "fnv_gro_cbe_trichy": {
      if (!fgRows.length) return [];
      return perStoreMap(fgRows, (r) => sotypeIs(r, "FNV"), (r) => sotypeIs(r, "GRO"));
    }
    case "watsapp": {
      if (!baseRows.length || !input.groups) return [];
      const kc = keyOf(baseRows[0]);
      return processWatsappGroups(baseRows, kc, input.groups);
    }
    case "fnv_gro_chennai": {
      if (!fgRows.length || !input.groups) return [];
      const kc = pickCol(fgRows[0], "CustomerName") || "CustomerName";
      // Customer 1's FNV TrmId is the anchor. Customer 1's own GRO SoId is
      // included (not skipped), plus every other customer's GRO SoId — all
      // mapped to customer 1's FNV TrmId.
      return processGroups(
        fgRows,
        kc,
        input.groups,
        (r) => typeIs(r, "FNV"),
        (r) => typeIs(r, "Milk") || typeIs(r, "Bakery_and_Egg"),
        true,
      );
    }
    case "mum_watsapp_fnv": {
      if (!fgRows.length || !input.groups) return [];
      const kc = keyOf(fgRows[0]);
      // Normally only FNV + Bakery_and_Egg count (Milk is excluded). Exception:
      // stores that show up as FNV trips (Trip No containing "FNV") — read
      // from whichever of these is uploaded: the Loading sheet, the vehicle
      // trip file, or the Ground file (it often carries the same TRIP NO.
      // column). For those stores alone, Milk-type rows are also counted.
      let exceptionStores = new Set<string>();
      const loadingSource = input.tripGroupsWb
        ? parseTripGroupFile(input.tripGroupsWb)
        : input.loadingWb
          ? parseLoadingSheet(input.loadingWb)
          : input.groundWb
            ? parseTripGroupFile(input.groundWb)
            : null;
      if (loadingSource) {
        const loading = loadingSource;
        exceptionStores = new Set(
          loading.filter((l) => /fnv/i.test(String(l.tripNo))).map((l) => norm(l.store)),
        );
      }
      const filterFor = (item: string) => (r: Row) => {
        if (typeIs(r, "FNV") || typeIs(r, "Bakery_and_Egg")) return true;
        if (typeIs(r, "Milk") && exceptionStores.has(norm(item))) return true;
        return false;
      };
      return processWatsappGroups(fgRows, kc, input.groups, filterFor);
    }
    case "mum_egg_vehicle": {
      if (!fgRows.length || !input.groups) return [];
      const kc = keyOf(fgRows[0]);
      // Any one customer's TrmId is the anchor; rest customers contribute
      // SoIds. Only Bakery_and_Egg type counts — FNV and Milk are ignored.
      return processGroups(
        fgRows,
        kc,
        input.groups,
        (r) => typeIs(r, "Bakery_and_Egg"),
        (r) => typeIs(r, "Bakery_and_Egg"),
      );
    }
    case "mum_milk_vehicle_egg_bread": {
      if (!fgRows.length || !input.remarksWb) return [];
      const eligible = parseEligibleStores(input.remarksWb);
      if (!eligible.size) return [];
      const custKey = pickCol(fgRows[0], "CustomerName") || "CustomerName";
      // Only stores flagged in the remarks file ("Milk, Bread & Egg") are considered.
      const eligibleRows = fgRows.filter((r) => eligible.has(norm(r[custKey])));
      // That store's Milk-type TrmId is matched with its own Bakery_and_Egg SoIds.
      return perStoreMap(eligibleRows, (r) => typeIs(r, "Milk"), (r) => typeIs(r, "Bakery_and_Egg"));
    }
  }
  return [];
}

export function downloadOutput(rows: OutRow[], filename: string) {
  // Sample format: saleOrderId,trmId,vehicleType,split,updateVehicleType
  const fmt = (v: any) => {
    if (v == null) return "";
    if (typeof v === "number") {
      // Avoid scientific notation for large IDs
      return Number.isInteger(v) ? v.toFixed(0) : String(v);
    }
    return String(v).trim();
  };
  const escape = (v: any) => {
    const s = fmt(v);
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  const header = "saleOrderId,trmId,vehicleType,split,updateVehicleType";
  const lines = [
    header,
    ...rows.map((r) => `${escape(r.SoId)},${escape(r.TrmId)},,,`),
  ];
  const csv = lines.join("\r\n");
  const name = filename.replace(/\.(xlsx|xls)$/i, ".csv");
  const finalName = /\.csv$/i.test(name) ? name : `${name}.csv`;
  const blob = new Blob(["\ufeff" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = finalName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}



/** Parse textarea input: each line = one group, items separated by comma/tab/whitespace.
 * If a line contains 2+ standalone long numeric tokens (5+ digits) mixed in with other
 * text — e.g. a pasted spreadsheet row with serial numbers, store names, and trip IDs —
 * only those numeric tokens are used as the group (the trip IDs), the rest is ignored.
 */
export function parseGroupsText(text: string): string[][] {
  return text
    .split(/\r?\n/)
    .map((line) => {
      const longNumbers = line.match(/\d{5,}/g);
      if (longNumbers && longNumbers.length >= 2) return longNumbers;
      return line.split(/[,\t]+/).map((s) => s.trim()).filter(Boolean);
    })
    .filter((arr) => arr.length > 0);
}

/** Parse groups from uploaded excel: each row = one group, all non-empty cells are items */
export function parseGroupsFromWorkbook(wb: XLSX.WorkBook): string[][] {
  const groups: string[][] = [];
  for (const sn of wb.SheetNames) {
    const aoa: any[][] = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: "" });
    // If first row looks like headers (D1,D2,Drop 1 etc), skip it
    let start = 0;
    if (aoa.length && aoa[0].every((c) => /^(d\d+|drop\s*\d+|s\d+|store\s*\d+)$/i.test(String(c ?? "").trim()))) {
      start = 1;
    }
    for (let i = start; i < aoa.length; i++) {
      const items = (aoa[i] || []).map((c) => String(c ?? "").trim()).filter(Boolean);
      if (items.length) groups.push(items);
    }
  }
  return groups;
}
