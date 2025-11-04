// src/ExamPlannerCSV.tsx
import { useMemo, useState, useEffect } from "react";
import { DndContext, useDraggable, useDroppable } from "@dnd-kit/core";
import { restrictToWindowEdges } from "@dnd-kit/modifiers";
import {
  format,
  addDays,
  subDays,
  isBefore,
  isAfter,
  startOfDay,
  parseISO,
} from "date-fns";
import * as Papa from "papaparse";
import * as XLSX from "xlsx-js-style";
import {
  compressToEncodedURIComponent,
  decompressFromEncodedURIComponent,
} from "lz-string";

/* ---------- Helpers ---------- */
function mondayOfWeek(d: Date) {
  const day = d.getDay();                // 0=dg … 6=ds
  const diff = (day + 6) % 7;            // 0 si dilluns
  return startOfDay(subDays(d, diff));
}
function fridayOfWeek(d: Date) {
  const mon = mondayOfWeek(d);
  return startOfDay(addDays(mon, 4));
}
function* eachWeek(mondayStart: Date, fridayEnd: Date) {
  let cur = new Date(mondayStart);
  while (!isAfter(cur, fridayEnd)) {
    const mon = new Date(cur);
    const fri = addDays(mon, 4);
    yield { mon, fri };
    cur = addDays(mon, 7);
  }
}
function fmtDM(d: Date) { return format(d, "dd/MM"); }
function iso(d: Date) { return format(d, "yyyy-MM-dd"); }

/* ---------- Tipus i models ---------- */
type TipusPeriode = "PARCIAL" | "FINAL" | "REAVALUACIÓ";

interface Period {
  id: number;
  label: string;
  tipus: TipusPeriode;
  startStr: string;   // "yyyy-MM-dd"
  endStr: string;     // "yyyy-MM-dd"
  curs?: number;      // any acadèmic (inici), ex. 2025
  quad?: 1 | 2;       // quadrimestre del període
  blackouts?: string[];
}

interface Subject {
  id: string;
  codi: string;
  sigles: string;
  nivell?: string;
  curs?: string;
  quadrimestre?: 1 | 2;
  MET?: string;
  MATT?: string;
  MEE?: string;
  MCYBERS?: string;
}

interface TimeSlot { start: string; end: string; } // "HH:mm"

type AssignedMap = Record<string, string[]>;     // "YYYY-MM-DD|slotIndex" → [subjectId,...]
type AssignedPerPeriod = Record<number, AssignedMap>;
type SlotsPerPeriod = Record<number, TimeSlot[]>;

/** Informació d’aules i matriculats per cel·la i assignatura */
type RoomsEnroll = {
  rooms: string[];
  students?: number;
};
type RoomsMapPerCell = Record<string, RoomsEnroll>; // subjectId → info
type RoomsDataPerPeriod = Record<number, Record<string, RoomsMapPerCell>>; // pid → (dateIso|slotIdx) → map

/* ---------- Subcomponents ---------- */
function MastersLines({ s }: { s: Subject }) {
  const lines = [s.MET, s.MATT, s.MEE, s.MCYBERS]
    .filter((v) => v && String(v).trim() !== "")
    .map((v) => String(v).trim());
  if (!lines.length) return null;
  return (
    <div className="text-xs opacity-80 leading-4 whitespace-pre-line">
      {lines.join("\n")}
    </div>
  );
}

function Chip({ id, s }: { id: string; s: Subject }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } =
    useDraggable({ id });
  return (
    <div
      ref={setNodeRef}
      {...listeners}
      {...attributes}
      className={`inline-flex flex-col px-3 py-2 rounded-2xl shadow-sm border text-sm cursor-grab active:cursor-grabbing select-none bg-white ${
        isDragging ? "opacity-70" : ""
      }`}
      style={{
        transform: transform
          ? `translate3d(${transform.x}px, ${transform.y}px, 0)`
          : undefined,
        maxWidth: 280,
      }}
      title={`${s.sigles} · ${s.codi}`}
    >
      <span className="font-medium truncate">{s.sigles} · {s.codi}</span>
      {s.nivell ? (
        <span className="text-xs opacity-80 leading-4">Nivell: {s.nivell}</span>
      ) : (
        <MastersLines s={s} />
      )}
    </div>
  );
}

function DropCell({
  id,
  disabled,
  assignedList,
  extrasForSubjects,
  onRemoveOne,
}: {
  id: string;
  disabled?: boolean;
  assignedList?: Subject[];
  extrasForSubjects?: Record<string, RoomsEnroll>; // subjectId -> info
  onRemoveOne?: (subjectId: string) => void;
}) {
  const { setNodeRef, isOver } = useDroppable({ id, disabled });
  return (
    <td
      ref={setNodeRef}
      className={`align-top min-w-[170px] h-20 p-2 border ${
        disabled ? "bg-gray-100 text-gray-400"
                 : isOver ? "ring-2 ring-indigo-400"
                          : "bg-white"
      }`}
    >
      {assignedList && assignedList.length ? (
        <div className="space-y-2">
          {assignedList.map((s) => {
            const extra = extrasForSubjects?.[s.id];
            const hasRooms = extra && extra.rooms && extra.rooms.length > 0;
            const hasStud = extra && typeof extra.students === "number" && !Number.isNaN(extra.students);
            return (
              <div
                key={s.id}
                className={`relative p-2 rounded-xl border shadow-sm ${disabled ? "opacity-60" : "bg-gray-50"}`}
              >
                <div className="text-sm font-semibold leading-tight">
                  {s.sigles} · {s.codi}
                </div>
                {s.nivell ? (
                  <div className="text-xs opacity-80">Nivell: {s.nivell}</div>
                ) : (
                  <MastersLines s={s} />
                )}

                {(hasRooms || hasStud) && (
                  <div className="mt-1 space-y-0.5 text-xs">
                    {hasRooms && (
                      <div>
                        <span className="font-medium">Aules/Rooms:</span>{" "}
                        {extra!.rooms.join(", ")}
                      </div>
                    )}
                    {hasStud && (
                      <div>
                        <span className="font-medium">Estudiants/Students:</span>{" "}
                        {extra!.students}
                      </div>
                    )}
                  </div>
                )}

                {!disabled && onRemoveOne && (
                  <button
                    onClick={() => onRemoveOne(s.id)}
                    className="absolute -top-2 -right-2 w-6 h-6 rounded-full border bg-white shadow text-xs"
                    aria-label="Eliminar"
                    title="Eliminar d’aquesta cel·la"
                  >
                    ×
                  </button>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <div className="text-xs text-gray-400 italic">
          {disabled ? "No disponible" : "Arrossega aquí"}
        </div>
      )}
    </td>
  );
}

/* ---------- Component principal ---------- */
export default function ExamPlannerCSV() {
  /* Assignatures (demo inicial – es sobreescriuran amb CSV/JSON) */
  const [subjects, setSubjects] = useState<Subject[]>([
    { id: "mat101", codi: "MAT101", sigles: "CALC I", nivell: "GRAU", curs: "2025", quadrimestre: 1 },
    { id: "fis201", codi: "FIS201", sigles: "FIS II", nivell: "GRAU", curs: "2025", quadrimestre: 1 },
    { id: "prg150", codi: "PRG150", sigles: "PRG", nivell: "GRAU", curs: "2025", quadrimestre: 1 },
    { id: "tic500", codi: "TIC500", sigles: "CIBER", curs: "2025", quadrimestre: 2, MCYBERS: "Sí", MET: "Optativa" },
  ]);

  /* Períodes (amb curs/quad propis) */
  const [periods, setPeriods] = useState<Period[]>([{
    id: 1,
    label: "Període 1",
    tipus: "PARCIAL",
    startStr: format(mondayOfWeek(new Date()), "yyyy-MM-dd"),
    endStr: format(fridayOfWeek(new Date()), "yyyy-MM-dd"),
    curs: undefined,
    quad: undefined,
    blackouts: [],
  }]);
  const [activePid, setActivePid] = useState<number>(1);

  /* Franges per període */
  const [slotsPerPeriod, setSlotsPerPeriod] = useState<SlotsPerPeriod>({
    1: [
      { start: "08:00", end: "10:00" },
      { start: "10:30", end: "12:30" },
      { start: "15:00", end: "17:00" },
    ],
  });

  /* Assignacions per període */
  const [assignedPerPeriod, setAssignedPerPeriod] = useState<AssignedPerPeriod>({});

  /* Aules/Matriculats per període/cel·la/assignatura */
  const [roomsData, setRoomsData] = useState<RoomsDataPerPeriod>({});

  /* Períodes on cada assignatura està permesa (derivat del CSV d’import) */
  const [allowedPeriodsBySubject, setAllowedPeriodsBySubject] = useState<Record<string, number[]>>({});

  const activePeriod = periods.find((p) => p.id === activePid)!;

  function isDisabledDay(d: Date, p: Period) {
    const sd = parseISO(p.startStr);
    const ed = parseISO(p.endStr);
    const outside = isBefore(d, sd) || isAfter(d, ed);
    if (outside) return true;
    const bl = p.blackouts ?? [];
    return bl.includes(iso(d));
  }
  function cellKey(dateIso: string, slotIndex: number) {
    return `${dateIso}|${slotIndex}`;
  }

  /* Assignatures ja utilitzades — NOMÉS en el període actiu (important per “un cop per període”) */
  const usedIds = useMemo(() => {
    const amap = assignedPerPeriod[activePid] ?? {};
    const s = new Set<string>();
    for (const list of Object.values(amap)) for (const id of list) s.add(id);
    return s;
  }, [assignedPerPeriod, activePid]);

  /* Filtrat automàtic segons curs/quad del període actiu + pertinença explícita al període */
  const availableSubjects = useMemo(() => {
    const pcurs = activePeriod?.curs != null ? String(activePeriod.curs) : undefined;
    const pquad = activePeriod?.quad;
    const pid = activePid;

    return subjects
      .filter((s) => !usedIds.has(s.id))
      .filter((s) => (pcurs ? s.curs === pcurs : true))
      .filter((s) => (pquad ? s.quadrimestre === pquad : true))
      .filter((s) => {
        const allowed = allowedPeriodsBySubject[s.id];
        return Array.isArray(allowed) ? allowed.includes(pid) : true;
      });
  }, [subjects, usedIds, activePeriod?.curs, activePeriod?.quad, activePid, allowedPeriodsBySubject]);

  /* Guardar/Carregar estat a URL (hash) */
  function saveStateToUrl() {
    const payload = { subjects, periods, slotsPerPeriod, assignedPerPeriod, activePid, roomsData, allowedPeriodsBySubject };
    const packed = compressToEncodedURIComponent(JSON.stringify(payload));
    const url = new URL(window.location.href);
    url.hash = `state=${packed}`;
    history.replaceState(null, "", url.toString());
    alert("Estat guardat a l’enllaç!");
  }
  function loadStateFromUrl(): boolean {
    const m = (window.location.hash || "").match(/[#&]state=([^&]+)/);
    if (!m) return false;
    try {
      const json = decompressFromEncodedURIComponent(m[1]);
      if (!json) return false;
      const data = JSON.parse(json);
      if (Array.isArray(data.subjects)) setSubjects(data.subjects);
      if (Array.isArray(data.periods)) setPeriods(data.periods);
      if (data.slotsPerPeriod) setSlotsPerPeriod(data.slotsPerPeriod);
      if (data.assignedPerPeriod) setAssignedPerPeriod(data.assignedPerPeriod);
      if (data.roomsData) setRoomsData(data.roomsData);
      if (data.allowedPeriodsBySubject) setAllowedPeriodsBySubject(data.allowedPeriodsBySubject);
      if (typeof data.activePid === "number") setActivePid(data.activePid);
      return true;
    } catch { return false; }
  }
  function copyLinkToClipboard() {
    if (!window.location.hash.includes("state=")) {
      saveStateToUrl();
      return;
    }
    navigator.clipboard.writeText(window.location.href)
      .then(() => alert("Enllaç copiat!"))
      .catch(() => alert("No s’ha pogut copiar l’enllaç."));
  }
  useEffect(() => { loadStateFromUrl(); /* un sol cop */ }, []);

  /* DnD */
  function onDragEnd(e: any) {
    const subjectId = e.active?.id as string;
    const dropId = e.over?.id as string | undefined;
    if (!dropId || !dropId.startsWith("cell:")) return;
    // id = cell:periodId:YYYY-MM-DD:slotIndex
    const [, pidStr, dateIso, slotIndexStr] = dropId.split(":");
    const pid = Number(pidStr);
    const period = periods.find((p) => p.id === pid);
    if (!period) return;

    const dayDate = parseISO(dateIso);
    if (isDisabledDay(dayDate, period)) return;

    if (usedIds.has(subjectId)) {
      alert("Aquesta assignatura ja està programada al període actiu.");
      return;
    }
    const key = cellKey(dateIso, Number(slotIndexStr));
    setAssignedPerPeriod((prev) => {
      const prevMap = prev[pid] ?? {};
      const nextList = (prevMap[key] ?? []).includes(subjectId)
        ? prevMap[key]!
        : [...(prevMap[key] ?? []), subjectId];
      return { ...prev, [pid]: { ...prevMap, [key]: nextList } };
    });
  }
  function removeOneFromCell(pid: number, dateIso: string, slotIndex: number, subjectId: string) {
    const key = cellKey(dateIso, slotIndex);
    setAssignedPerPeriod((prev) => {
      const prevMap = prev[pid] ?? {};
      const next = (prevMap[key] ?? []).filter((id) => id !== subjectId);
      const copy: AssignedMap = { ...prevMap };
      if (next.length) copy[key] = next; else delete copy[key];
      return { ...prev, [pid]: copy };
    });
  }

  /* Gestió períodes */
  function addPeriod() {
    if (periods.length >= 5) { alert("Pots tenir com a màxim 5 períodes."); return; }
    const newId = Math.max(0, ...periods.map((p) => p.id)) + 1;
    const today = new Date();
    const newPeriod: Period = {
      id: newId,
      label: `Període ${newId}`,
      tipus: "PARCIAL",
      startStr: format(mondayOfWeek(today), "yyyy-MM-dd"),
      endStr: format(fridayOfWeek(today), "yyyy-MM-dd"),
      curs: undefined,
      quad: undefined,
      blackouts: [],
    };
    setPeriods([...periods, newPeriod]);
    setSlotsPerPeriod((sp) => ({ ...sp, [newId]: [{ start: "08:00", end: "10:00" }] }));
    setActivePid(newId);
  }
  function removePeriod(id: number) {
    if (!confirm("Segur que vols eliminar aquest període?")) return;
    setPeriods(periods.filter((p) => p.id !== id));
    setAssignedPerPeriod((ap) => { const c = { ...ap }; delete c[id]; return c; });
    setSlotsPerPeriod((sp) => { const c = { ...sp }; delete c[id]; return c; });
    setRoomsData((rd) => { const c = { ...rd }; delete c[id]; return c; });
    // allowedPeriodsBySubject no cal tocar-ho
    if (activePid === id) {
      const rest = periods.filter((p) => p.id !== id);
      if (rest.length) setActivePid(rest[0].id);
    }
  }

  /* Exportacions */
  function exportJSON() {
    const data = { periods, slotsPerPeriod, assignedPerPeriod, subjects, roomsData, allowedPeriodsBySubject };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "planificador-examens.json";
    a.click();
    URL.revokeObjectURL(a.href);
  }
  function importJSON(ev: React.ChangeEvent<HTMLInputElement>) {
    const f = ev.target.files?.[0];
    if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(String(reader.result));
        if (Array.isArray(data.periods)) setPeriods(data.periods);
        if (data.slotsPerPeriod) setSlotsPerPeriod(data.slotsPerPeriod);
        if (data.assignedPerPeriod) setAssignedPerPeriod(data.assignedPerPeriod);
        if (Array.isArray(data.subjects)) setSubjects(data.subjects);
        if (data.roomsData) setRoomsData(data.roomsData);
        if (data.allowedPeriodsBySubject) setAllowedPeriodsBySubject(data.allowedPeriodsBySubject);
        if (Array.isArray(data.periods) && data.periods.length) setActivePid(data.periods[0].id);
      } catch { alert("JSON no vàlid"); }
    };
    reader.readAsText(f);
    ev.currentTarget.value = "";
  }

  // Inferència (per si falta curs/quad a alguna assignatura puntual)
  function inferCursFromDate(d: Date): string {
    const y = d.getFullYear(), m = d.getMonth() + 1;
    return (m >= 9 ? y : y - 1).toString();  // curs = any d'inici
  }
  function inferQuadFromDate(d: Date): 1 | 2 {
    const m = d.getMonth() + 1;
    return (m >= 9 || m === 1) ? 1 : 2;       // set–gen: Q1; feb–jul: Q2
  }

  function exportCSV() {
    const lines: string[] = [];
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      for (const { mon } of eachWeek(mondayOfWeek(parseISO(p.startStr)), fridayOfWeek(parseISO(p.endStr)))) {
        for (let i = 0; i < 5; i++) {
          const day = addDays(mon, i);
          if (isDisabledDay(day, p)) continue;
          const dateIso = iso(day);
          for (let si = 0; si < slots.length; si++) {
            const key = `${dateIso}|${si}`;
            const ids = amap[key] ?? [];
            if (!ids.length) continue;
            for (const id of ids) {
              const s = subjects.find((x) => x.id === id);
              if (!s) continue;
              const CENTRE = "230";
              const CURS = s.curs?.toString() ?? String(p.curs ?? inferCursFromDate(day));
              const QUADRIMESTRE = String(s.quadrimestre ?? p.quad ?? inferQuadFromDate(day));
              const TIPUS_EXAMEN = p.tipus === "REAVALUACIÓ" ? "REAVALUACIO" : p.tipus;
              const DIA = format(day, "dd-MM-yyyy");
              const HORA_INICI = slots[si].start;
              const HORA_FI = slots[si].end;
              const UNITAT_DOCENT = s.codi;
              const GRUPS = "";
              lines.push([CENTRE, CURS, QUADRIMESTRE, TIPUS_EXAMEN, DIA, HORA_INICI, HORA_FI, UNITAT_DOCENT, GRUPS].join(","));
            }
          }
        }
      }
    }
    const csv = lines.join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "examens_export.csv";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function formatSubjectForCell(s: Subject) {
    const header = `${s.sigles} · ${s.codi}`;
    const masters = [s.MET, s.MATT, s.MEE, s.MCYBERS]
      .filter((v) => v && String(v).trim() !== "")
      .map((v) => String(v).trim())
      .join("\n");
    const details = s.nivell ? `Nivell: ${s.nivell}` : masters;
    return details ? `${header}\n${details}` : header;
  }

  function exportExcel() {
    const wb = XLSX.utils.book_new();
    const slotColors = ["#E3F2FD", "#E8F5E9", "#FFF3E0", "#F3E5F5", "#E0F7FA", "#FBE9E7"];
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      const allRows: any[][] = [];
      for (const { mon } of eachWeek(mondayOfWeek(parseISO(p.startStr)), fridayOfWeek(parseISO(p.endStr)))) {
        const dayHeaders = Array.from({ length: 5 }).map((_, i) => {
          const d = addDays(mon, i);
          return `${["Dl/Mon", "Dt/Tu", "Dc/Wed", "Dj/Thu", "Dv/Fri"][i]} ${format(d, "dd/MM")}`;
        });
        allRows.push(["franja horària/Time slot", ...dayHeaders]);
        for (let si = 0; si < slots.length; si++) {
          const slot = slots[si];
          const row: any[] = [`${slot.start}-${slot.end}`];
          for (let i = 0; i < 5; i++) {
            const day = addDays(mon, i);
            if (isDisabledDay(day, p)) { row.push(""); continue; }
            const key = `${iso(day)}|${si}`;
            const ids = (amap[key] ?? []);
            const list = ids.map(id => subjects.find(x => x.id === id)).filter(Boolean) as Subject[];
            const text = list.map(s => {
              const base = formatSubjectForCell(s);
              const extra = roomsData?.[p.id]?.[key]?.[s.id];
              if (!extra) return base;
              const lines: string[] = [base];
              if (extra.rooms?.length) lines.push(`Aules/Rooms: ${extra.rooms.join(", ")}`);
              if (typeof extra.students === "number") lines.push(`Estudiants/Students: ${extra.students}`);
              return lines.join("\n");
            }).join("\n\n");
            row.push(text);
          }
          allRows.push(row);
        }
        allRows.push([]);
      }
      const ws = XLSX.utils.aoa_to_sheet(allRows);
      if (ws["!ref"]) {
        const range = XLSX.utils.decode_range(ws["!ref"] as string);
        for (let R = range.s.r; R <= range.e.r; R++) {
          const firstCell = ws[XLSX.utils.encode_cell({ r: R, c: 0 })];
          const v = (firstCell?.v as string || "").toLowerCase();
          const isHeader = v.includes("franja horària") || v.includes("franja horaria") || v.includes("time slot");
          if (isHeader) {
            for (let C = 0; C <= range.e.c; C++) {
              const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
              if (!cell) continue;
              (cell as any).s = { font: { bold: true }, alignment: { horizontal: C === 0 ? "left" : "center" } };
            }
            continue;
          }
          const isSlotRow = typeof firstCell?.v === "string" && /^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$/.test(firstCell.v);
          if (isSlotRow) {
            const idx = (slots.findIndex(s => `${s.start}-${s.end}` === (firstCell?.v as string)) + slotColors.length) % slotColors.length;
            const rgb = slotColors[idx].replace("#", "");
            if (firstCell) (firstCell as any).s = { font: { bold: true } };
            for (let C = 1; C <= range.e.c; C++) {
              const cell = ws[XLSX.utils.encode_cell({ r: R, c: C })];
              if (!cell) continue;
              (cell as any).s = { alignment: { vertical: "top", wrapText: true }, fill: { fgColor: { rgb } } };
            }
          }
        }
        ws["!cols"] = [{ wch: 16 }];
        const totalCols = (XLSX.utils.decode_range(ws["!ref"] as string).e.c + 1);
        while ((ws["!cols"] as any[]).length < totalCols) (ws["!cols"] as any[]).push({ wch: 36 });
        ws["!rows"] = allRows.map(row => ({ hpt: row.length ? 42 : 10 }));
      }
      XLSX.utils.book_append_sheet(wb, ws, `${p.tipus}`);
    }
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "calendari_examens.xlsx";
    a.click();
    URL.revokeObjectURL(a.href);
  }

  function exportTXT() {
    const LEN = { CODI: 10, CURS: 4, QUAD: 1, NOM: 120, DIA: 10, HORA: 5, DESC: 2000 } as const;
    const padText = (v: string, len: number) => {
      const s = (v ?? "").toString();
      return s.length >= len ? s.slice(0, len) : s + " ".repeat(len - s.length);
    };
    const padNum = (v: number | string | undefined, len: number) => {
      const s = (v ?? "").toString();
      return s.length >= len ? s.slice(0, len) : " ".repeat(len - s.length) + s;
    };
    const lines: string[] = [];
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      for (const { mon } of eachWeek(mondayOfWeek(parseISO(p.startStr)), fridayOfWeek(parseISO(p.endStr)))) {
        for (let i = 0; i < 5; i++) {
          const day = addDays(mon, i);
          if (isDisabledDay(day, p)) continue;
          for (let si = 0; si < slots.length; si++) {
            const ids = (amap[`${iso(day)}|${si}`] ?? []);
            if (!ids.length) continue;
            for (const id of ids) {
              const s = subjects.find(x => x.id === id);
              if (!s) continue;
              const CODI = padText(s.codi, LEN.CODI);
              const CURS = padNum(s.curs ?? String(p.curs ?? inferCursFromDate(day)), LEN.CURS);
              const QUAD = padNum(s.quadrimestre ?? p.quad ?? inferQuadFromDate(day), LEN.QUAD);
              const NOM  = padText(s.sigles, LEN.NOM);
              const DIA  = padText(format(day, "dd-MM-yyyy"), LEN.DIA);
              const HORA = padText((slots[si].start || "").replace(":", "-"), LEN.HORA);
              const DESC = padText(p.tipus === "REAVALUACIÓ" ? "REAVALUACIO" : p.tipus, LEN.DESC);
              lines.push([CODI, CURS, QUAD, NOM, DIA, HORA, DESC].join(" "));
            }
          }
        }
      }
    }
    const txt = lines.join("\n");
    const blob = new Blob([txt], { type: "text/plain;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "examens_export.txt";
    a.click();
    URL.revokeObjectURL(a.href);
  }

/* ---------- Import CSV (assignatures + períodes) — SENSE duplicar assignatures ---------- */
const handleImportCSV: React.ChangeEventHandler<HTMLInputElement> = (e) => {
  const f = e.target.files?.[0];
  if (!f) return;

  // Auxiliars
  const parseDate = (raw: any): string | undefined => {
    if (!raw) return undefined;
    const s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    const m2 = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
    return undefined;
  };
  const parseSlots = (raw: any): TimeSlot[] => {
    if (!raw) return [];
    return String(raw)
      .split(/[;,|]/)
      .map((p) => p.trim())
      .filter(Boolean)
      .map((pair) => {
        const mm = pair.match(/^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$/);
        if (!mm) return null;
        const [_, a, b] = mm;
        const pad = (h: string) => h.split(":").map((x) => x.padStart(2, "0")).join(":");
        return { start: pad(a), end: pad(b) };
      })
      .filter(Boolean) as TimeSlot[];
  };
  const parseBlackouts = (raw: any): string[] => {
    if (!raw) return [];
    const toks = String(raw).split(/[;,|]/).map((s) => s.trim()).filter(Boolean);
    const out: string[] = [];
    for (const t of toks) {
      const d = parseDate(t);
      if (d) out.push(d);
    }
    return Array.from(new Set(out)).sort();
  };
  const normQuad = (raw: any): 1 | 2 | undefined => {
    if (raw == null || raw === "") return undefined;
    const n = Number(String(raw).replace(/\D/g, ""));
    return n === 1 || n === 2 ? (n as 1 | 2) : undefined;
  };
  const normCursAny = (raw: any): string | undefined => {
    if (!raw && raw !== 0) return undefined;
    const s = String(raw).trim();
    const m = s.match(/^(\d{4})\s*[-/]/);
    if (m) return m[1];
    const y = s.match(/^\d{4}$/);
    if (y) return y[0];
    const first4 = s.match(/(\d{4})/);
    return first4 ? first4[1] : undefined;
  };

  Papa.parse(f, {
    header: true,
    skipEmptyLines: true,
    complete: (res: Papa.ParseResult<any>) => {
      try {
        const rows = (res.data as any[]).filter(Boolean);

        // --- Col·lectors ---
        // 1) Subjects sense duplicar (clau: codi||sigles)
        const subjByKey = new Map<string, Subject>();
        // 2) Períodes on apareix cada clau
        const periodsByKey = new Map<string, Set<number>>();
        // 3) Períodes definits i slots
        const periodMap = new Map<number, Period>();
        const slotsMap: SlotsPerPeriod = {};

        // Pistes per deduccions
        const quadSeenPerPid = new Map<number, 1|2>();
        const cursSeenPerPid = new Map<number, number>();

        const keyOf = (codi: any, sigles: any) =>
          `${String(codi||"").trim().toLowerCase()}||${String(sigles||"").trim().toLowerCase()}`;

        for (const r of rows) {
          // --- Subject (sense duplicar) ---
          const codi = r.codi ?? r.codigo ?? r.CODI ?? r.CODIGO ?? r.code;
          const sigles = r.sigles ?? r.SIGLES ?? r.siglas ?? r.SIGLAS;
          if (codi || sigles) {
            const k = keyOf(codi, sigles);
            const existed = subjByKey.get(k);

            const nivell = (r.nivell ?? r.NIVELL ?? r.nivel ?? r.NIVEL)?.toString();
            const curs = normCursAny(r.curs ?? r.CURS ?? r.curso ?? r.CURSO);
            const quadrimestre = normQuad(r.quadrimestre ?? r.QUADRIMESTRE ?? r.quad ?? r.QUAD);
            const MET = r.MET ?? r.met;
            const MATT = r.MATT ?? r.matt;
            const MEE = r.MEE ?? r.mee;
            const MCYBERS = r.MCYBERS ?? r.mcybers;

            if (!existed) {
              subjByKey.set(k, {
                id: String(codi || sigles),
                codi: String(codi || ""),
                sigles: String(sigles || ""),
                nivell: nivell || undefined,
                curs: curs || undefined,
                quadrimestre: quadrimestre,
                MET: MET ? String(MET) : undefined,
                MATT: MATT ? String(MATT) : undefined,
                MEE: MEE ? String(MEE) : undefined,
                MCYBERS: MCYBERS ? String(MCYBERS) : undefined,
              });
            } else {
              // Omple camps buits si en aquesta fila venen informats
              if (!existed.nivell && nivell) existed.nivell = nivell;
              if (!existed.curs && curs) existed.curs = curs;
              if (!existed.quadrimestre && quadrimestre) existed.quadrimestre = quadrimestre;
              if (!existed.MET && MET) existed.MET = String(MET);
              if (!existed.MATT && MATT) existed.MATT = String(MATT);
              if (!existed.MEE && MEE) existed.MEE = String(MEE);
              if (!existed.MCYBERS && MCYBERS) existed.MCYBERS = String(MCYBERS);
            }
          }

          // --- Període associat a la fila ---
          const pidRaw = r.period_id ?? r.PERIOD_ID ?? r.PeriodId ?? r.periode ?? r.PERIODO ?? r.PERIOD;
          const pid = pidRaw ? Number(pidRaw) : NaN;
          if (Number.isFinite(pid) && pid >= 1 && pid <= 5) {
            // Registra que aquesta assignatura (clau) pertany a aquest període
            const k = keyOf(r.codi ?? r.CODI ?? r.code, r.sigles ?? r.SIGLES);
            if (k.trim() !== "||") {
              if (!periodsByKey.has(k)) periodsByKey.set(k, new Set<number>());
              periodsByKey.get(k)!.add(pid);
            }
          }

          // Tipus de període + dates + slots
          const tipusRaw = (r.period_tipus ?? r.PERIOD_TIPUS ?? r.tipo ?? r.TIPO ?? "")
            .toString().toUpperCase();
          const tipusNorm: TipusPeriode =
            tipusRaw === "FINAL" ? "FINAL"
            : (tipusRaw === "REAVALUACIO" || tipusRaw === "REAVALUACIÓ" || tipusRaw === "REAVALUACION") ? "REAVALUACIÓ"
            : "PARCIAL";

          const startStr =
            parseDate(r.period_inici ?? r.PERIOD_INICI ?? r.start) ||
            format(mondayOfWeek(new Date()), "yyyy-MM-dd");
          const endStr =
            parseDate(r.period_fi ?? r.PERIOD_FI ?? r.end) ||
            format(fridayOfWeek(new Date()), "yyyy-MM-dd");

          const slots = parseSlots(r.period_slots ?? r.PERIOD_SLOTS ?? r.slots) || [{ start: "08:00", end: "10:00" }];
          const blackouts = parseBlackouts(r.period_blackouts ?? r.PERIOD_BLACKOUTS ?? r.blackouts ?? r.BLOCKED_DATES);

          // Curs/quad declarats com a columnes de període
          const periodCurs = normCursAny(r.period_curs ?? r.PERIOD_CURS);
          const periodQuad = normQuad(r.period_quad ?? r.PERIOD_QUAD);

          // Pistes
          const filaCurs = normCursAny(r.curs ?? r.CURS ?? r.curso ?? r.CURSO);
          const filaQuad = normQuad(r.quadrimestre ?? r.QUADRIMESTRE ?? r.quad ?? r.QUAD);
          if (Number.isFinite(pid)) {
            if (filaQuad) quadSeenPerPid.set(pid, filaQuad);
            if (filaCurs) cursSeenPerPid.set(pid, Number(filaCurs));
          }

          if (Number.isFinite(pid) && pid >= 1 && pid <= 5) {
            if (!periodMap.has(pid)) {
              periodMap.set(pid, {
                id: pid,
                label: `Període ${pid}`,
                tipus: tipusNorm,
                startStr,
                endStr,
                curs: periodCurs ? Number(periodCurs) : undefined,
                quad: periodQuad,
                blackouts,
              });
              slotsMap[pid] = slots;
            } else {
              const p = periodMap.get(pid)!;
              if (!p.curs && periodCurs) p.curs = Number(periodCurs);
              if (!p.quad && periodQuad) p.quad = periodQuad;
            }
          }
        }

        // Deducció final
        for (const [pid, p] of periodMap) {
          if (p.quad == null && quadSeenPerPid.has(pid)) p.quad = quadSeenPerPid.get(pid)!;
          if (p.curs == null && cursSeenPerPid.has(pid)) p.curs = cursSeenPerPid.get(pid)!;
        }

        //  Subjects únics (sense sufixos aleatoris)
        const uniqueSubjects = Array.from(subjByKey.values());

        // allowedPeriodsBySubject des de periodsByKey
        const nextAllowed: Record<string, number[]> = {};
        for (const s of uniqueSubjects) {
          const key = `${s.codi.trim().toLowerCase()}||${s.sigles.trim().toLowerCase()}`;
          const set = periodsByKey.get(key);
          if (set && set.size) nextAllowed[s.id] = Array.from(set).sort((a,b)=>a-b);
        }

        setSubjects(uniqueSubjects);

        if (periodMap.size > 0) {
          const ordered = Array.from(periodMap.keys()).sort((a, b) => a - b);
          const list = ordered.map((k) => periodMap.get(k)!);
          setPeriods(list);
          setSlotsPerPeriod(slotsMap);
          setAssignedPerPeriod({});
          setRoomsData({});
          setAllowedPeriodsBySubject(nextAllowed);
          setActivePid(list[0].id);
          alert(`Importades ${uniqueSubjects.length} assignatures i ${list.length} períodes del CSV.`);
        } else {
          setAllowedPeriodsBySubject(nextAllowed);
          alert(`Importades ${uniqueSubjects.length} assignatures del CSV.`);
        }
      } catch (err) {
        console.error(err);
        alert("Error processant el CSV");
      }
    },
    error: () => alert("No s'ha pogut llegir el fitxer CSV"),
  });

  e.currentTarget.value = "";
};

/* ---------- Import CSV (Aules + Matriculats) — igual ---------- */
const handleImportRoomsCSV: React.ChangeEventHandler<HTMLInputElement> = (e) => {
  const f = e.target.files?.[0];
  if (!f) return;

  const parseDate = (raw: any): string | undefined => {
    if (!raw) return undefined;
    const s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    const m2 = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m2) return `${m2[3]}-${m2[2]}-${m2[1]}`;
    return undefined;
  };
  const normTime = (t: any): string | undefined => {
    if (!t && t !== 0) return undefined;
    const s = String(t).trim();
    const m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (!m) return undefined;
    return `${m[1].padStart(2,"0")}:${m[2]}`;
  };

  Papa.parse(f, {
    header: true,
    skipEmptyLines: true,
    complete: (res: Papa.ParseResult<any>) => {
      try {
        const rows = (res.data as any[]).filter(Boolean);
        if (!rows.length) { alert("CSV buit."); return; }

        let attached = 0, skipped = 0;

        // Índex ràpid: per període → "HH:mm|HH:mm" → slotIndex
        const slotIdxByPid: Record<number, Record<string, number>> = {};
        for (const p of periods) {
          const slots = slotsPerPeriod[p.id] ?? [];
          slotIdxByPid[p.id] = {};
          slots.forEach((sl, i) => {
            slotIdxByPid[p.id][`${sl.start}|${sl.end}`] = i;
          });
        }

        const nextRoomsData: RoomsDataPerPeriod = JSON.parse(JSON.stringify(roomsData || {}));

        for (const r of rows) {
          const codi = r.codi ?? r.CODI ?? r.codigo ?? r.CODIGO ?? r.code;
          const sigles = r.sigles ?? r.SIGLES ?? r.siglas ?? r.SIGLAS;

          const periode =
            r.periode ?? r.PERIode ?? r.PERIODO ?? r.periode_id ?? r.period_id ?? r.PERIOD ?? r.Period ?? r.Període;
          const pid = Number(periode);
          if (!Number.isFinite(pid) || !(periods.some(p => p.id === pid))) { skipped++; continue; }

          const dayIso = parseDate(r["dia d'examen"] ?? r.dia ?? r.DIA ?? r.fecha ?? r.FECHA ?? r.day);
          const start = normTime(r["hora d'inici de l'examen"] ?? r.hora_inici ?? r.inici ?? r.start ?? r.HORA_INICI);
          const end = normTime(r["hora de fi de l'examen"] ?? r.hora_fi ?? r.fi ?? r.end ?? r.HORA_FI);
          const aula = (r.aula ?? r.AULA ?? r.room ?? r.ROOM ?? r.classroom ?? "").toString().trim();
          const nStudRaw = r["número d'estudiants matriculats"] ?? r.matriculats ?? r.matriculados ?? r.students ?? r.STUDENTS ?? r.num_estudiants ?? r.NUM_ESTUDIANTS;
          const nStudents = nStudRaw != null && nStudRaw !== "" ? Number(String(nStudRaw).replace(/[^\d]/g,"")) : undefined;

          if (!dayIso || !start || !end || !aula) { skipped++; continue; }
          const idx = slotIdxByPid[pid]?.[`${start}|${end}`];
          if (idx == null) { skipped++; continue; }

          const key = `${dayIso}|${idx}`;

          const assignedIdsInCell = (assignedPerPeriod[pid]?.[key] ?? []);
          const matchId =
            assignedIdsInCell.find(id => subjects.find(s => s.id === id)?.codi === String(codi || "")) ??
            assignedIdsInCell.find(id => subjects.find(s => s.id === id)?.sigles === String(sigles || ""));

          if (!matchId) { skipped++; continue; }

          if (!nextRoomsData[pid]) nextRoomsData[pid] = {};
          if (!nextRoomsData[pid][key]) nextRoomsData[pid][key] = {};
          if (!nextRoomsData[pid][key][matchId]) nextRoomsData[pid][key][matchId] = { rooms: [] };

          const entry = nextRoomsData[pid][key][matchId];

          if (aula && !entry.rooms.includes(aula)) entry.rooms.push(aula);
          if (typeof nStudents === "number" && Number.isFinite(nStudents) && entry.students == null) {
            entry.students = nStudents;
          }

          attached++;
        }

        setRoomsData(nextRoomsData);
        alert(`Aules/Matrículats processats. Afegits: ${attached}. Omesos: ${skipped}.`);
      } catch (err) {
        console.error(err);
        alert("Error processant el CSV d’aules/matriculats");
      }
    },
    error: () => alert("No s'ha pogut llegir el fitxer CSV d’aules/matriculats"),
  });

  e.currentTarget.value = "";
};

  /* ---------- Render ---------- */
  return (
    <div className="p-6 max-w-[1200px] mx-auto">
      <h1 className="text-2xl font-bold mb-2">Planificador d'exàmens — períodes amb curs/quadrimestre</h1>
      <p className="text-sm mb-6">
        CSV esperat (assignatures/períodes): <code>codi,sigles,nivell,curs,quadrimestre,period_id,period_tipus,period_inici,period_fi,period_slots,period_blackouts</code>.
        Opcional: <code>MET,MATT,MEE,MCYBERS</code>. També admet <code>period_curs,period_quad</code>.
      </p>

      {/* Intercanvi de dades */}
      <div className="p-4 rounded-2xl border shadow-sm bg-white mb-6">
        <h2 className="font-semibold mb-3">Dades i intercanvi</h2>
        <div className="flex flex-wrap gap-3 items-center">
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar CSV
            <input type="file" accept=".csv,text/csv" className="hidden" onChange={handleImportCSV} />
          </label>

          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar Aules/Matriculats (CSV)
            <input type="file" accept=".csv,text/csv" className="hidden" onChange={handleImportRoomsCSV} />
          </label>

          <button onClick={exportCSV} className="px-3 py-2 border rounded-xl shadow-sm">Exportar CSV</button>
          <button onClick={exportTXT} className="px-3 py-2 border rounded-xl shadow-sm">Exportar TXT</button>
          <button onClick={exportExcel} className="px-3 py-2 border rounded-xl shadow-sm">Exportar Excel</button>
          <button onClick={exportJSON} className="px-3 py-2 border rounded-xl shadow-sm">Exportar JSON</button>
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar JSON
            <input type="file" accept="application/json" className="hidden" onChange={importJSON} />
          </label>

          <button onClick={saveStateToUrl} className="px-3 py-2 border rounded-xl shadow-sm">Guardar a l’enllaç</button>
          <button
            onClick={() => { if (!loadStateFromUrl()) alert("No s’ha trobat cap estat a l’enllaç."); }}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Carregar de l’enllaç
          </button>
          <button onClick={copyLinkToClipboard} className="px-3 py-2 border rounded-xl shadow-sm">Copiar enllaç</button>
          <span className="text-xs text-gray-500 ml-auto">
            Disponibles: {availableSubjects.length}/{subjects.length}
          </span>
        </div>
      </div>

      {/* Pestanyes (sense mostrar id) */}
      <div className="mb-4 flex items-center justify-between gap-2">
        <div className="flex flex-wrap gap-2">
          {periods.map((p) => (
            <button
              key={p.id}
              onClick={() => setActivePid(p.id)}
              className={`px-3 py-2 rounded-xl border shadow-sm ${p.id === activePid ? "bg-indigo-50 border-indigo-300" : "bg-white"}`}
              title="Canviar de període"
            >
              {p.tipus}
            </button>
          ))}
        </div>
        <div className="flex gap-2">
          <button onClick={addPeriod} className="px-3 py-2 border rounded-xl shadow-sm">Afegir període</button>
          {periods.length > 1 && (
            <button onClick={() => removePeriod(activePid)} className="px-3 py-2 border rounded-xl shadow-sm">
              Eliminar període actiu
            </button>
          )}
        </div>
      </div>

      {/* Configuració període actiu */}
      {activePeriod && (
        <div className="grid md:grid-cols-3 gap-4 mb-6">
          <div className="p-4 rounded-2xl border shadow-sm bg-white">
            <h2 className="font-semibold mb-3">Configuració del període</h2>

            <label className="block text-sm mb-1">Tipus</label>
            <select
              value={activePeriod.tipus}
              onChange={(e) => {
                const v = e.target.value as TipusPeriode;
                setPeriods(arr => arr.map(p => p.id === activePid ? { ...p, tipus: v } : p));
              }}
              className="w-full border rounded-xl p-2"
            >
              <option>PARCIAL</option>
              <option>FINAL</option>
              <option>REAVALUACIÓ</option>
            </select>

            <label className="block text-sm mt-3 mb-1">Curs (any d’inici)</label>
            <input
              type="number"
              placeholder="Ex. 2025"
              value={activePeriod.curs ?? ""}
              onChange={(e) => {
                const n = e.target.value ? Number(e.target.value) : undefined;
                setPeriods(arr => arr.map(p => p.id === activePid ? { ...p, curs: n } : p));
              }}
              className="w-full border rounded-xl p-2"
            />

            <label className="block text-sm mt-3 mb-1">Quadrimestre del període</label>
            <select
              value={activePeriod.quad ?? 0}
              onChange={(e) => {
                const v = Number(e.target.value) as 0 | 1 | 2;
                setPeriods(arr =>
                  arr.map(p =>
                    p.id === activePid ? { ...p, quad: v === 1 || v === 2 ? (v as 1 | 2) : undefined } : p
                  )
                );
              }}
              className="w-full border rounded-xl p-2"
            >
              <option value={0}>(Sense)</option>
              <option value={1}>1</option>
              <option value={2}>2</option>
            </select>

            <label className="block text-sm mt-3 mb-1">Inici</label>
            <input
              type="date"
              value={activePeriod.startStr}
              onChange={(e)=> setPeriods(arr => arr.map(p => p.id===activePid? {...p, startStr: e.target.value}: p))}
              className="w-full border rounded-xl p-2"
            />
            <label className="block text-sm mt-3 mb-1">Fi</label>
            <input
              type="date"
              value={activePeriod.endStr}
              onChange={(e)=> setPeriods(arr => arr.map(p => p.id===activePid? {...p, endStr: e.target.value}: p))}
              className="w-full border rounded-xl p-2"
            />

            {/* Blackouts */}
            <div className="mt-4">
              <h3 className="font-semibold mb-2 text-sm">Dies no disponibles</h3>
              <div className="flex gap-2 items-center">
                <input
                  type="date"
                  min={activePeriod.startStr}
                  max={activePeriod.endStr}
                  onChange={(e) => {
                    const v = e.target.value;
                    if (!v) return;
                    setPeriods(arr => arr.map(p=>{
                      if (p.id !== activePid) return p;
                      const set = new Set(p.blackouts ?? []);
                      set.add(v);
                      return {...p, blackouts: Array.from(set).sort()};
                    }));
                    e.currentTarget.value = "";
                  }}
                  className="border rounded-xl p-2"
                />
                <span className="text-xs text-gray-500">Afegeix un dia del rang per bloquejar-lo</span>
              </div>
              <div className="mt-2 flex flex-wrap gap-2">
                {(activePeriod.blackouts ?? []).map(d => (
                  <span key={d} className="inline-flex items-center gap-2 px-2 py-1 rounded-lg border text-xs bg-gray-50">
                    {format(parseISO(d), "dd/MM/yyyy")}
                    <button
                      className="w-5 h-5 rounded-full border bg-white text-[10px]"
                      title="Eliminar"
                      onClick={()=>{
                        setPeriods(arr => arr.map(p=>{
                          if (p.id !== activePid) return p;
                          const next = (p.blackouts ?? []).filter(x=>x!==d);
                          return {...p, blackouts: next};
                        }));
                      }}
                    >×</button>
                  </span>
                ))}
                {(activePeriod.blackouts ?? []).length === 0 && (
                  <span className="text-xs text-gray-500 italic">No hi ha dies bloquejats</span>
                )}
              </div>
            </div>
          </div>

          {/* Franges horàries */}
          <div className="p-4 rounded-2xl border shadow-sm bg-white md:col-span-2">
            <h2 className="font-semibold mb-3">Franges horàries (per a aquest període)</h2>
            <div className="space-y-2">
              {(slotsPerPeriod[activePid] ?? []).map((s, i) => (
                <div key={i} className="flex items-center gap-2">
                  <span className="text-sm w-6">{i+1}.</span>
                  <input
                    value={s.start}
                    onChange={(e)=>{
                      const v=e.target.value;
                      setSlotsPerPeriod(sp => {
                        const arr = [...(sp[activePid] ?? [])];
                        arr[i] = {...arr[i], start: v};
                        return {...sp, [activePid]: arr};
                      });
                    }}
                    className="border rounded-xl p-2 w-28" placeholder="HH:mm"
                  />
                  <span>–</span>
                  <input
                    value={s.end}
                    onChange={(e)=>{
                      const v=e.target.value;
                      setSlotsPerPeriod(sp => {
                        const arr = [...(sp[activePid] ?? [])];
                        arr[i] = {...arr[i], end: v};
                        return {...sp, [activePid]: arr};
                      });
                    }}
                    className="border rounded-xl p-2 w-28" placeholder="HH:mm"
                  />
                  <button
                    onClick={()=>{
                      setSlotsPerPeriod(sp=>{
                        const arr=[...(sp[activePid]??[])].filter((_,idx)=> idx!==i);
                        return {...sp, [activePid]: arr};
                      });
                      setAssignedPerPeriod(ap=>{
                        const amap = {...(ap[activePid] ?? {})};
                        for (const k of Object.keys(amap)) {
                          const slotIdx = Number(k.split("|")[1]);
                          if (slotIdx === i) delete amap[k];
                        }
                        return {...ap, [activePid]: amap};
                      });
                      setRoomsData(rd=>{
                        const per = {...(rd[activePid] ?? {})};
                        for (const k of Object.keys(per)) {
                          const slotIdx = Number(k.split("|")[1]);
                          if (slotIdx === i) delete per[k];
                        }
                        return {...rd, [activePid]: per};
                      });
                    }}
                    className="ml-2 text-xs px-2 py-1 border rounded-lg"
                  >
                    Eliminar
                  </button>
                </div>
              ))}
            </div>
            <button
              onClick={()=>{
                setSlotsPerPeriod(sp=>{
                  const cur = sp[activePid] ?? [];
                  const last = cur[cur.length-1];
                  const nextStart = last? last.end : "08:00";
                  const [h,m] = nextStart.split(":").map(Number);
                  const endH = (h+2).toString().padStart(2,"0");
                  const next = { start: nextStart, end: `${endH}:${(m||0).toString().padStart(2,"0")}` };
                  return {...sp, [activePid]: [...cur, next]};
                });
              }}
              className="mt-3 px-3 py-2 border rounded-xl shadow-sm"
            >
              Afegir franja
            </button>
          </div>
        </div>
      )}

      {/* Calaix + Calendari */}
      <DndContext onDragEnd={onDragEnd} modifiers={[restrictToWindowEdges]}>
        {/* Safata d'assignatures */}
        <div className="p-4 rounded-2xl border shadow-sm bg-white mb-6">
          <h2 className="font-semibold mb-3">Assignatures (arrossega)</h2>
          <div className="flex flex-wrap gap-2">
            {availableSubjects.map((s) => (
              <Chip key={s.id} id={s.id} s={s} />
            ))}
            {!availableSubjects.length && (
              <div className="text-xs text-gray-500 italic">
                No hi ha assignatures per al curs/quadrimestre i període d’aquest calendari, o ja estan totes programades.
              </div>
            )}
          </div>
        </div>

        {/* Calendari del període actiu */}
        {activePeriod && (
          <div className="mb-8">
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-lg font-semibold">
                {activePeriod.tipus} — {format(parseISO(activePeriod.startStr), "dd/MM")} a {format(parseISO(activePeriod.endStr), "dd/MM")}
              </h3>
              <span className="text-sm text-gray-500">(dl–dv)</span>
            </div>

            {[...eachWeek(mondayOfWeek(parseISO(activePeriod.startStr)), fridayOfWeek(parseISO(activePeriod.endStr)))].map(({mon, fri}, wIdx) => (
              <div key={wIdx} className="mt-6">
                <div className="flex items-center gap-3 mb-2">
                  <h4 className="font-semibold">Setmana {format(mon,"dd/MM")} — {format(fri,"dd/MM")}</h4>
                  <span className="text-xs text-gray-500">(dl–dv)</span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-sm">
                    <thead>
                      <tr>
                        <th className="border p-2 w-[160px] text-left">franja horària/Time slot</th>
                        {Array.from({length:5}).map((_,i)=>(
                          <th key={i} className="border p-2 min-w-[170px] text-left">
                            <div className="font-semibold">{["Dl/Mon","Dt/Tu","Dc/Wed","Dj/Thu","Dv/Fri"][i]}</div>
                            <div className="text-xs text-gray-500">{fmtDM(addDays(mon, i))}</div>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(slotsPerPeriod[activePid] ?? []).map((s, slotIndex) => (
                        <tr key={slotIndex}>
                          <td className="border p-2 align-top font-medium whitespace-nowrap">{s.start}-{s.end}</td>
                          {Array.from({length:5}).map((_,i)=>{
                            const day = addDays(mon, i);
                            const dateIso = iso(day);
                            const disabled = isDisabledDay(day, activePeriod);
                            const amap = assignedPerPeriod[activePid] ?? {};
                            const subjIds = (amap[cellKey(dateIso, slotIndex)] ?? []);
                            const assignedList = subjIds.map(id => subjects.find(x => x.id === id)).filter(Boolean) as Subject[];

                            const extrasForSubjects: Record<string, RoomsEnroll> = {};
                            const extrasCell = roomsData?.[activePid]?.[cellKey(dateIso, slotIndex)] ?? {};
                            for (const sid of subjIds) {
                              if (extrasCell[sid]) extrasForSubjects[sid] = extrasCell[sid];
                            }

                            return (
                              <DropCell
                                key={i}
                                id={`cell:${activePid}:${dateIso}:${slotIndex}`}
                                disabled={disabled}
                                assignedList={assignedList}
                                extrasForSubjects={extrasForSubjects}
                                onRemoveOne={(subjectId)=> removeOneFromCell(activePid, dateIso, slotIndex, subjectId)}
                              />
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}
          </div>
        )}
      </DndContext>

      <div className="mt-8 text-xs text-gray-500">
        <ul className="list-disc ml-5 space-y-1">
          <li>La safata mostra només les assignatures del <em>quadrimestre</em> i el <em>període</em> actiu segons el CSV.</li>
          <li>Una assignatura es pot programar una vegada per cada període on apareix al CSV.</li>
          <li>Exportacions (CSV/TXT/Excel) inclouen totes les pestanyes; l’Excel afegeix Aules/Estudiants si s’han importat.</li>
        </ul>
      </div>
    </div>
  );
}
