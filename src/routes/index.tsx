import { createFileRoute } from "@tanstack/react-router";
import type * as React from "react";
import { useMemo, useRef, useState } from "react";
import {
  CITIES,
  MERGE_TYPES,
  CITY_MERGE_TYPES,
  type MergeTypeId,
  readWorkbook,
  runMerge,
  downloadOutput,
  parseGroupsText,
  parseGroupsFromWorkbook,
  parseGroundFile,
  parseTripGroupFile,
} from "@/lib/tripMerge";
import type * as XLSX from "xlsx";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "Trip Merge Dashboard" },
      { name: "description", content: "Merge FNV, GRO and combined trips across cities into a single loading file." },
      { property: "og:title", content: "Trip Merge Dashboard" },
      { property: "og:description", content: "Merge FNV, GRO and combined trips across cities into a single loading file." },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary_large_image" },
    ],
  }),
  component: Dashboard,
});

type WbState = { wb: XLSX.WorkBook; name: string } | null;

function FileInput({
  label,
  hint,
  value,
  onChange,
  accept = ".xlsx,.xls",
}: {
  label: string;
  hint?: string;
  value: WbState;
  onChange: (v: WbState) => void;
  accept?: string;
}) {
  const ref = useRef<HTMLInputElement>(null);
  return (
    <div className="rounded-lg border border-border bg-card p-4">
      <div className="flex items-start justify-between gap-3">
        <div>
          <div className="font-medium text-foreground">{label}</div>
          {hint && <div className="mt-0.5 text-xs text-muted-foreground">{hint}</div>}
          {value && <div className="mt-1 text-xs text-primary">✓ {value.name}</div>}
        </div>
        <div className="flex gap-2">
          <button
            onClick={() => ref.current?.click()}
            className="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90"
          >
            {value ? "Replace" : "Upload"}
          </button>
          {value && (
            <button
              onClick={() => onChange(null)}
              className="rounded-md border border-border px-3 py-1.5 text-xs text-foreground hover:bg-accent"
            >
              Clear
            </button>
          )}
        </div>
      </div>
      <input
        ref={ref}
        type="file"
        accept={accept}
        className="hidden"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (!f) return;
          const buf = await f.arrayBuffer();
          onChange({ wb: readWorkbook(buf), name: f.name });
          if (ref.current) ref.current.value = "";
        }}
      />
    </div>
  );
}

function Dashboard() {
  const [city, setCity] = useState<string>("Bengaluru");
  const availableMergeTypes = useMemo(
    () => MERGE_TYPES.filter((m) => (CITY_MERGE_TYPES[city] ?? []).includes(m.id)),
    [city],
  );
  const [mergeType, setMergeType] = useState<MergeTypeId>("fnv");
  // Ensure selected merge type is valid for the current city
  if (availableMergeTypes.length && !availableMergeTypes.some((m) => m.id === mergeType)) {
    setMergeType(availableMergeTypes[0].id);
  }
  const [baseKind, setBaseKind] = useState<"TripID" | "CustomerName">("CustomerName");
  const [baseWb, setBaseWb] = useState<WbState>(null);
  const [fgWb, setFgWb] = useState<WbState>(null);
  const [loadingWb, setLoadingWb] = useState<WbState>(null);
  const [tripGroupsWb, setTripGroupsWb] = useState<WbState>(null);
  const [groupsText, setGroupsText] = useState("");
  const [groupsFile, setGroupsFile] = useState<WbState>(null);
  const [groundFile, setGroundFile] = useState<WbState>(null);
  const [remarksFile, setRemarksFile] = useState<WbState>(null);
  const [status, setStatus] = useState<string>("");
  const [ocrBusy, setOcrBusy] = useState(false);
  const [ocrPreview, setOcrPreview] = useState<string | null>(null);
  const ocrFileRef = useRef<HTMLInputElement>(null);

  const mt = useMemo(() => MERGE_TYPES.find((m) => m.id === mergeType)!, [mergeType]);
  // Merge types where the vehicle trip file can be uploaded instead of the Loading sheet.
  const canUseTripGroups = mergeType === "gro" || mergeType === "fnv_gro_milk" || mergeType === "mum_watsapp_fnv";

  const groups = useMemo(() => {
    if (!mt.needsGroups) return [];
    if (mt.needsGroundFile && groundFile) return parseGroundFile(groundFile.wb);
    if (groupsFile) return parseGroupsFromWorkbook(groupsFile.wb);
    return parseGroupsText(groupsText);
  }, [mt.needsGroups, mt.needsGroundFile, groundFile, groupsFile, groupsText]);

  const runOcr = async (imageFile: File): Promise<string> => {
    const Tesseract = await import("tesseract.js");
    const { data } = await Tesseract.recognize(imageFile, "eng");
    return data.text
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)
      .join(", ");
  };

  const handlePasteImage = async (e: React.ClipboardEvent<HTMLTextAreaElement>) => {
    const items = e.clipboardData?.items;
    if (!items) return;
    let imageFile: File | null = null;
    for (const item of items) {
      if (item.type.startsWith("image/")) {
        const blob = item.getAsFile();
        if (blob) imageFile = blob;
        break;
      }
    }
    if (!imageFile) return; // normal text paste — let it through as-is
    e.preventDefault();
    setOcrBusy(true);
    setStatus("Reading text from pasted image…");
    try {
      const extracted = await runOcr(imageFile);
      setOcrPreview(extracted);
      setStatus(`✓ Extracted text from image — review below, then click "Insert" to add it`);
    } catch (err: any) {
      setStatus(`OCR failed: ${err?.message ?? err}`);
    } finally {
      setOcrBusy(false);
    }
  };

  const handleUploadImage = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (ocrFileRef.current) ocrFileRef.current.value = "";
    if (!file) return;
    setOcrBusy(true);
    setStatus("Reading text from uploaded image…");
    try {
      const extracted = await runOcr(file);
      setOcrPreview(extracted);
      setStatus(`✓ Extracted text from image — review below, then click "Insert" to add it`);
    } catch (err: any) {
      setStatus(`OCR failed: ${err?.message ?? err}`);
    } finally {
      setOcrBusy(false);
    }
  };

  const canRun = () => {
    if (mt.needsBase && !baseWb) return "Upload 6620 Base file";
    if (mt.needsFg && !fgWb) return mergeType === "fnv_gro_cbe_trichy" ? "Upload 6620 Base file" : "Upload FNV/GRO merge file";
    if (mt.needsLoading && !loadingWb && !(canUseTripGroups && tripGroupsWb)) {
      return "Upload loading sheet (or the vehicle trip file)";
    }
    if (mt.needsRemarksFile && !remarksFile) return "Upload the store remarks file";
    if (mt.needsGroups && groups.length === 0) return "Enter or upload merge groups";
    return null;
  };

  const handleRun = () => {
    const err = canRun();
    if (err) {
      setStatus(err);
      return;
    }
    try {
      const rows = runMerge({
        mergeType,
        city,
        baseKind,
        baseWb: baseWb?.wb ?? null,
        fgWb: fgWb?.wb ?? null,
        loadingWb: loadingWb?.wb ?? null,
        tripGroupsWb: tripGroupsWb?.wb ?? null,
        groundWb: groundFile?.wb ?? null,
        remarksWb: remarksFile?.wb ?? null,
        groups,
      });
      if (!rows.length) {
        setStatus("No matching rows produced. Check inputs, city sheet, and item names/IDs.");
        return;
      }
      const stamp = new Date().toISOString().slice(0, 16).replace(/[:T]/g, "-");
      downloadOutput(rows, `merged_${mergeType}_${city}_${stamp}.xlsx`);
      setStatus(`✓ Generated ${rows.length} rows`);
    } catch (e: any) {
      setStatus(`Error: ${e?.message ?? e}`);
    }
  };

  return (
    <div className="min-h-screen bg-background">
      <header className="border-b border-border bg-card">
        <div className="mx-auto max-w-5xl px-6 py-5">
          <h1 className="text-2xl font-semibold tracking-tight text-foreground">Trip Merge Dashboard</h1>
          <p className="mt-1 text-sm text-muted-foreground">
            Merge trips across customers into a single loading file (SoId + TrmId).
          </p>
        </div>
      </header>

      <main className="mx-auto max-w-5xl px-6 py-8 space-y-6">
        {/* Step 1: City & Merge type */}
        <section className="grid grid-cols-1 gap-4 md:grid-cols-3">
          <div className="rounded-lg border border-border bg-card p-4">
            <label className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
              1. City
            </label>
            <select
              value={city}
              onChange={(e) => setCity(e.target.value)}
              className="mt-2 w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
            >
              {CITIES.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </div>

          <div className="rounded-lg border border-border bg-card p-4 md:col-span-2">
            <label className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
              2. Merge Type
            </label>
            <select
              value={mergeType}
              onChange={(e) => setMergeType(e.target.value as MergeTypeId)}
              className="mt-2 w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
            >
              {availableMergeTypes.map((m) => (
                <option key={m.id} value={m.id}>
                  {m.label}
                </option>
              ))}
            </select>
          </div>
        </section>

        {/* Step 2: Base kind */}
        {mt.needsBaseKind && (
          <section className="rounded-lg border border-border bg-card p-4">
            <label className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
              3. Group items are
            </label>
            <div className="mt-2 flex gap-3">
              {(["TripID", "CustomerName"] as const).map((k) => (
                <label key={k} className="flex items-center gap-2 text-sm text-foreground">
                  <input
                    type="radio"
                    name="baseKind"
                    checked={baseKind === k}
                    onChange={() => setBaseKind(k)}
                  />
                  {k === "TripID" ? "Trip ID" : "Customer Name"}
                </label>
              ))}
            </div>
          </section>
        )}

        {/* Step 3: File uploads */}
        <section className="space-y-3">
          <h2 className="text-sm font-medium uppercase tracking-wide text-muted-foreground">
            Files
          </h2>
          {mt.needsBase && (
            <FileInput
              label="6620 Base file"
              hint="Sheets per city with CustomerName, TripID, SoId, sotype_fnv_gro, Trmid"
              value={baseWb}
              onChange={setBaseWb}
            />
          )}
          {mt.needsFg && (
            <FileInput
              label={mergeType === "fnv_gro_cbe_trichy" ? "6620 Base file" : "FNV / GRO merge file"}
              hint={
                mergeType === "fnv_gro_cbe_trichy"
                  ? "Upload your 6620 Base file here (Coimbatore / Trichy sheet) — same file, this dashboard just reads it directly for this merge type."
                  : "Sheets per city with Type column (FNV, Milk, Bakery_and_Egg)"
              }
              value={fgWb}
              onChange={setFgWb}
            />
          )}
          {mt.needsLoading && (
            <FileInput
              label={mergeType === "mum_watsapp_fnv" ? "Loading sheet (for the Milk exception stores)" : "Loading sheet"}
              hint={
                mergeType === "mum_watsapp_fnv"
                  ? "Trip No + Store Site ID columns. Stores that show up as FNV trips here are the exception — Milk-type SoIds are also counted for them alone."
                  : "Trip No + Store Site ID columns; used to derive trip groups / FNV trips"
              }
              value={loadingWb}
              onChange={(v) => {
                setLoadingWb(v);
                if (v && canUseTripGroups) setTripGroupsWb(null);
              }}
            />
          )}
          {mt.needsLoading && canUseTripGroups && (
            <>
              <div className="text-xs text-muted-foreground">— or —</div>
              <FileInput
                label="Vehicle trip file (alternative to Loading sheet)"
                hint={
                  mergeType === "mum_watsapp_fnv"
                    ? 'TRIP NO. + Store Name columns. Stores in trips whose TRIP NO. contains "FNV" are the exception — Milk-type SoIds are also counted for them alone (same rule as the Loading sheet, just read from this file instead).'
                    : 'TRIP NO. + Store Name columns (e.g. a route plan with Vehicle Type, Driver Name, Drop Point Number). Stores sharing the same TRIP NO. are merged together directly — no FNV/numeric filtering needed.'
                }
                value={tripGroupsWb}
                onChange={(v) => {
                  setTripGroupsWb(v);
                  if (v) setLoadingWb(null);
                }}
              />
              {tripGroupsWb && (
                <div className="text-xs text-muted-foreground">
                  {(() => {
                    const g = new Map<string, number>();
                    for (const { tripNo } of parseTripGroupFile(tripGroupsWb.wb)) {
                      g.set(tripNo, (g.get(tripNo) ?? 0) + 1);
                    }
                    const multi = [...g.values()].filter((n) => n > 1).length;
                    return `${g.size} trips parsed, ${multi} with multiple stores to merge.`;
                  })()}
                </div>
              )}
            </>
          )}
          {mt.needsRemarksFile && (
            <FileInput
              label="Store remarks file"
              hint='Store Name + a trailing Remarks column (e.g. "Milk, Bread & Egg"). Only stores with a remark are used.'
              value={remarksFile}
              accept=".xlsx,.xls,.csv"
              onChange={setRemarksFile}
            />
          )}
        </section>

        {/* Step 4: Groups */}
        {mt.needsGroups && (
          <section className="rounded-lg border border-border bg-card p-4 space-y-3">
            <div>
              <label className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                Merge groups
              </label>
              <p className="mt-1 text-xs text-muted-foreground">
                One group per line. Separate items with commas. First item's TRMID is used; remaining items contribute SoIds.
                {mergeType === "gro_bread_milk" && " (Bangalore: uses 6620 base file; anchor TRMID is the trip with multiple SoIds; all drops' SoIds map to it.)"}
                {mergeType === "watsapp" &&
                  " (WhatsApp Merge: uses the 6620 Base file only. Each line is a trip/customer group to be merged. Whichever item has the MOST SoIds in the base file becomes the anchor — its TRMID is used, and its own SoIds are skipped. Every other item's SoIds are matched to that anchor TRMID.)"}
                {mergeType === "fnv_gro_chennai" &&
                  " (Chennai: customer 1's FNV TRMID is the anchor. Customer 1's own GRO SoId is included, plus every other customer's GRO SoId — all mapped to customer 1's TRMID.)"}
                {mergeType === "mum_watsapp_fnv" &&
                  " (Mum Watsapp: uses the FNV/GRO merge file, Mumbai sheet. Only FNV and Bakery_and_Egg count (Milk is excluded) — except stores flagged as FNV trips (Trip No containing \"FNV\") in the Loading sheet, the vehicle trip file, or the Ground file, where Milk counts too. Whichever item has the MOST matching SoIds becomes the anchor — its TRMID is used, its own SoIds are skipped. You can also upload a Ground file below to auto-build groups from merged Store Name rows.)"}
                {mergeType === "mum_egg_vehicle" &&
                  " (Egg Vehicle: uses the FNV/GRO merge file, Mumbai sheet. Only Bakery_and_Egg counts — FNV and Milk are ignored. Any one item's TRMID is the anchor; the rest contribute SoIds. Upload the Ground file below to auto-build groups from merged Vehicle-type rows.)"}
                {" "}You can also paste a screenshot directly into the box below, or upload one, and it'll be read via OCR.
              </p>
            </div>
            <textarea
              value={groupsText}
              onChange={(e) => {
                setGroupsText(e.target.value);
                setGroupsFile(null);
                setGroundFile(null);
              }}
              onPaste={handlePasteImage}
              placeholder={
                "Ben_102_RR Nagar, Ben_104_Gottigere, Ben_105_Sanjay Nagar\n9745241, 9745250, 9745212\n\n(or paste/upload a screenshot)"
              }
              rows={6}
              disabled={!!groupsFile || !!groundFile || ocrBusy}
              className="w-full rounded-md border border-input bg-background px-3 py-2 font-mono text-xs"
            />

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => ocrFileRef.current?.click()}
                disabled={ocrBusy}
                className="rounded-md border border-border px-3 py-1.5 text-xs font-medium text-foreground hover:bg-accent disabled:opacity-50"
              >
                Upload image for OCR
              </button>
              {ocrBusy && <span className="text-xs text-primary">Reading text from image…</span>}
              <input
                ref={ocrFileRef}
                type="file"
                accept="image/*"
                className="hidden"
                onChange={handleUploadImage}
              />
            </div>

            {ocrPreview !== null && (
              <div className="rounded-md border border-primary/40 bg-accent/40 p-3 space-y-2">
                <label className="text-xs font-medium uppercase tracking-wide text-muted-foreground">
                  Detected text — review/edit, then click Insert
                </label>
                <textarea
                  value={ocrPreview}
                  onChange={(e) => setOcrPreview(e.target.value)}
                  rows={4}
                  className="w-full rounded-md border border-input bg-background px-3 py-2 font-mono text-xs"
                />
                <div className="flex gap-2">
                  <button
                    type="button"
                    onClick={() => {
                      setGroupsFile(null);
                      setGroupsText((prev) => (prev ? `${prev}\n${ocrPreview}` : ocrPreview));
                      setOcrPreview(null);
                      setStatus("✓ Inserted OCR text into merge groups");
                    }}
                    className="rounded-md bg-primary px-3 py-1.5 text-xs font-medium text-primary-foreground hover:opacity-90"
                  >
                    Insert
                  </button>
                  <button
                    type="button"
                    onClick={() => setOcrPreview(null)}
                    className="rounded-md border border-border px-3 py-1.5 text-xs text-foreground hover:bg-accent"
                  >
                    Discard
                  </button>
                </div>
              </div>
            )}

            <div className="text-xs text-muted-foreground">— or —</div>
            <FileInput
              label="Upload groups Excel"
              hint="Each row = one merge group. Column headers (D1, D2, Drop 1…) are skipped."
              value={groupsFile}
              onChange={(v) => {
                setGroupsFile(v);
                if (v) setGroundFile(null);
                if (v) setGroupsText("");
              }}
            />
            {mt.needsGroundFile && (
              <>
                <div className="text-xs text-muted-foreground">— or —</div>
                <FileInput
                  label="Upload Ground file"
                  hint="Route plan / loading sheet with Store Name + Vehicle type columns. Rows sharing a merged Vehicle-type cell are grouped together automatically."
                  value={groundFile}
                  accept=".xlsx,.xls,.csv"
                  onChange={(v) => {
                    setGroundFile(v);
                    if (v) {
                      setGroupsFile(null);
                      setGroupsText("");
                    }
                  }}
                />
              </>
            )}
            {groups.length > 0 && (
              <div className="text-xs text-muted-foreground">
                {groups.length} group{groups.length !== 1 ? "s" : ""} parsed.
              </div>
            )}
          </section>
        )}

        {/* Action */}
        <section className="flex items-center justify-between gap-4">
          <div className="text-sm text-muted-foreground">{status}</div>
          <button
            onClick={handleRun}
            className="rounded-md bg-primary px-5 py-2.5 text-sm font-medium text-primary-foreground shadow hover:opacity-90"
          >
            Generate & Download
          </button>
        </section>

        <footer className="pt-4 text-xs text-muted-foreground">
          Output: column A = SoId, column B = TrmId. Columns C–E (vehicleType, split, updateVehicleType) are kept blank, matching the standard upload template.
        </footer>
      </main>
    </div>
  );
}
