import React, { useMemo, useState, useEffect } from "react";
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

/* ---------- Helpers ---------- */
function mondayOfWeek(d: Date) {
  const day = d.getDay();
  const diff = (day + 6) % 7;
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
function fmtDM(d: Date) {
  return format(d, "dd/MM");
}

/* ---------- Types ---------- */
type TipusPeriode = "PARCIAL" | "FINAL" | "REAVALUACIÓ";
interface Subject {
  id: string;
  codigo: string;
  siglas: string;
  nivel: string;
  curs?: string;
  quad?: 1 | 2;
}
interface TimeSlot { start: string; end: string; }
interface PeriodMeta {
  id: number;
  tipus: TipusPeriode;
  any: number;       // 2025..2090
  quad: 1 | 2;
  startStr: string;  // yyyy-MM-dd
  endStr: string;    // yyyy-MM-dd
}
type AssignedMap = Record<string, string[]>;     // "YYYY-MM-DD|slotIndex" → [subjectId,...]
type AssignedPerPeriod = Record<number, AssignedMap>;
type SlotsPerPeriod = Record<number, TimeSlot[]>;

/* ---------- Draggable chip ---------- */
function Chip({ id, label }: { id: string; label: string }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } =
    useDraggable({ id });
  return (
    <div
      ref={setNodeRef}
      {...listeners}
      {...attributes}
      className={`inline-flex items-center gap-2 px-3 py-2 rounded-2xl shadow-sm border text-sm cursor-grab active:cursor-grabbing select-none bg-white ${
        isDragging ? "opacity-70" : ""
      }`}
      style={{
        transform: transform
          ? `translate3d(${transform.x}px, ${transform.y}px, 0)`
          : undefined,
      }}
      title={label}
    >
      <span className="font-medium truncate max-w-[20ch]">{label}</span>
    </div>
  );
}

/* ---------- Droppable cell ---------- */
function DropCell({
  id,
  disabled,
  assignedList,
  onRemoveOne,
}: {
  id: string;
  disabled?: boolean;
  assignedList?: Subject[];
  onRemoveOne?: (subjectId: string) => void;
}) {
  const { setNodeRef, isOver } = useDroppable({ id, disabled });
  return (
    <td
      ref={setNodeRef}
      className={`align-top min-w-[170px] h-20 p-2 border ${
        disabled
          ? "bg-gray-100 text-gray-400"
          : isOver
          ? "ring-2 ring-indigo-400"
          : "bg-white"
      }`}
    >
      {assignedList && assignedList.length > 0 ? (
        <div className="space-y-2">
          {assignedList.map((s) => (
            <div
              key={s.id}
              className={`relative p-2 rounded-xl border shadow-sm ${
                disabled ? "opacity-60" : "bg-gray-50"
              }`}
            >
              <div className="text-sm font-semibold leading-tight">
                {s.siglas} · {s.codigo}
              </div>
              <div className="text-xs opacity-80">Nivell: {s.nivel}</div>
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
          ))}
        </div>
      ) : (
        <div className="text-xs text-gray-400 italic">
          {disabled ? "No disponible" : "Arrossega aquí"}
        </div>
      )}
    </td>
  );
}

/* ---------- Main ---------- */
export default function ExamPlannerCSV() {
  /* Subjects */
  const [subjects, setSubjects] = useState<Subject[]>([
    { id: "mat101", codigo: "MAT101", siglas: "CALC I", nivel: "GRAU" },
    { id: "fis201", codigo: "FIS201", siglas: "FIS II", nivel: "GRAU" },
    { id: "prg150", codigo: "PRG150", siglas: "PRG", nivel: "GRAU" },
    { id: "alg300", codigo: "ALG300", siglas: "ALG", nivel: "MÀSTER" },
  ]);

  /* Períodes */
  const [periods, setPeriods] = useState<PeriodMeta[]>([
    {
      id: 1,
      tipus: "PARCIAL",
      any: new Date().getFullYear(),
      quad: 1,
      startStr: format(mondayOfWeek(new Date()), "yyyy-MM-dd"),
      endStr: format(fridayOfWeek(new Date()), "yyyy-MM-dd"),
    },
  ]);
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
  const [assignedPerPeriod, setAssignedPerPeriod] = useState<AssignedPerPeriod>(
    {}
  );

  /* Filtres calaix */
  const [filterCurs, setFilterCurs] = useState<string | "">("");
  const [filterQuad, setFilterQuad] = useState<0 | 1 | 2>(0);

  const activePeriod = periods.find((p) => p.id === activePid)!;

  function isDisabledDay(d: Date, p: PeriodMeta) {
    const sd = parseISO(p.startStr);
    const ed = parseISO(p.endStr);
    return isBefore(d, sd) || isAfter(d, ed);
  }
  function cellKey(dateIso: string, slotIndex: number) {
    return `${dateIso}|${slotIndex}`;
  }

  const allCursos = useMemo(
    () => Array.from(new Set(subjects.map((s) => s.curs).filter(Boolean))) as string[],
    [subjects]
  );

  useEffect(() => {
    const ap = periods.find((p) => p.id === activePid);
    if (ap) setFilterQuad(ap.quad);
  }, [activePid, periods]);

  const usedIds = useMemo(() => {
    const s = new Set<string>();
    for (const amap of Object.values(assignedPerPeriod)) {
      for (const list of Object.values(amap)) {
        for (const id of list) s.add(id);
      }
    }
    return s;
  }, [assignedPerPeriod]);

  const availableSubjects = useMemo(() => {
    return subjects
      .filter((s) => !usedIds.has(s.id))
      .filter((s) => (filterCurs ? s.curs === filterCurs : true))
      .filter((s) => (filterQuad ? s.quad === filterQuad : true));
  }, [subjects, usedIds, filterCurs, filterQuad]);

  /* Drag & drop */
  function onDragEnd(e: any) {
    const subjectId = e.active?.id as string;
    const dropId = e.over?.id as string | undefined;
    if (!dropId) return;
    if (!dropId.startsWith("cell:")) return;

    // id = cell:periodId:YYYY-MM-DD:slotIndex
    const [, pidStr, dateIso, slotIndexStr] = dropId.split(":");
    const pid = Number(pidStr);
    const period = periods.find((p) => p.id === pid);
    if (!period) return;

    const dayDate = parseISO(dateIso);
    if (isDisabledDay(dayDate, period)) return;

    if (usedIds.has(subjectId)) {
      alert("Aquesta assignatura ja està programada al calendari.");
      return;
    }

    const key = cellKey(dateIso, Number(slotIndexStr));
    setAssignedPerPeriod((prev) => {
      const prevMap = prev[pid] ?? {};
      const list = prevMap[key] ?? [];
      if (list.includes(subjectId)) return prev;
      const nextMap: AssignedMap = { ...prevMap, [key]: [...list, subjectId] };
      return { ...prev, [pid]: nextMap };
    });
  }

  function removeOneFromCell(
    pid: number,
    dateIso: string,
    slotIndex: number,
    subjectId: string
  ) {
    const key = cellKey(dateIso, slotIndex);
    setAssignedPerPeriod((prev) => {
      const prevMap = prev[pid] ?? {};
      const list = prevMap[key] ?? [];
      const next = list.filter((id) => id !== subjectId);
      const copy: AssignedMap = { ...prevMap };
      if (next.length === 0) delete copy[key];
      else copy[key] = next;
      return { ...prev, [pid]: copy };
    });
  }

  /* Gestió períodes */
  function addPeriod() {
    if (periods.length >= 5) {
      alert("Pots tenir com a màxim 5 períodes.");
      return;
    }
    const newId = Math.max(0, ...periods.map((p) => p.id)) + 1;
    const today = new Date();
    const meta: PeriodMeta = {
      id: newId,
      tipus: "PARCIAL",
      any: today.getFullYear(),
      quad: 1,
      startStr: format(mondayOfWeek(today), "yyyy-MM-dd"),
      endStr: format(fridayOfWeek(today), "yyyy-MM-dd"),
    };
    setPeriods([...periods, meta]);
    setSlotsPerPeriod((sp) => ({ ...sp, [newId]: [{ start: "08:00", end: "10:00" }] }));
    setActivePid(newId);
  }
  function removePeriod(id: number) {
    if (!confirm("Segur que vols eliminar aquest període?")) return;
    setPeriods(periods.filter((p) => p.id !== id));
    setAssignedPerPeriod((ap) => {
      const c = { ...ap };
      delete c[id];
      return c;
    });
    setSlotsPerPeriod((sp) => {
      const c = { ...sp };
      delete c[id];
      return c;
    });
    if (activePid === id && periods.length > 1) {
      const rest = periods.filter((p) => p.id !== id);
      setActivePid(rest[0].id);
    }
  }

  /* Exportacions bàsiques (CSV/TXT/JSON) */
  function exportJSON() {
    const data = { periods, slotsPerPeriod, assignedPerPeriod, subjects };
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "planificador-examens.json"; a.click();
    URL.revokeObjectURL(url);
  }
  function importJSON(ev: React.ChangeEvent<HTMLInputElement>) {
    const f = ev.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const data = JSON.parse(String(reader.result));
        if (Array.isArray(data.periods)) setPeriods(data.periods);
        if (data.slotsPerPeriod) setSlotsPerPeriod(data.slotsPerPeriod);
        if (data.assignedPerPeriod) setAssignedPerPeriod(data.assignedPerPeriod);
        if (Array.isArray(data.subjects)) setSubjects(data.subjects);
        if (Array.isArray(data.periods) && data.periods.length) setActivePid(data.periods[0].id);
      } catch { alert("JSON no vàlid"); }
    };
    reader.readAsText(f);
    ev.currentTarget.value = "";
  }
  function exportCSV() {
    const rows: string[] = [];
    rows.push("Periode,Data,Slot,HoraInici,HoraFi,Codigo,Siglas,Nivel,Curs,Quadrimestre");
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      for (const { mon } of eachWeek(mondayOfWeek(parseISO(p.startStr)), fridayOfWeek(parseISO(p.endStr)))) {
        for (let si = 0; si < slots.length; si++) {
          for (let i = 0; i < 5; i++) {
            const day = addDays(mon, i);
            if (isDisabledDay(day, p)) continue;
            const dateIso = format(day, "yyyy-MM-dd");
            const key = cellKey(dateIso, si);
            const ids = amap[key] ?? [];
            ids.forEach((id) => {
              const s = subjects.find((x) => x.id === id);
              if (!s) return;
              const label = `${p.tipus} ${p.any} Q${p.quad}`;
              rows.push(
                [
                  label,
                  format(day, "dd/MM/yyyy"),
                  `${si + 1}`,
                  slots[si]?.start ?? "",
                  slots[si]?.end ?? "",
                  s.codigo, s.siglas, s.nivel,
                  s.curs ?? "",
                  s.quad ?? "",
                ].map(v => `"${String(v).replace(/"/g,'""')}"`).join(",")
              );
            });
          }
        }
      }
    }
    const blob = new Blob([rows.join("\n")], { type:"text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download="examenes.csv"; a.click();
    URL.revokeObjectURL(url);
  }
  function formatTxtLine(
    label: string, dateStr: string, slotIdx: number, start: string, end: string, s: Subject
  ) {
    const pad = (t: string, w: number) => (t || "").slice(0,w).padEnd(w," ");
    return (
      pad(label, 20) + pad(dateStr, 10) + pad(String(slotIdx), 2) +
      pad(start, 5) + pad(end, 5) + pad(s.codigo, 12) + pad(s.siglas, 12) +
      pad(s.nivel, 10) + pad(s.curs ?? "", 12) + pad(String(s.quad ?? ""), 1)
    );
  }
  function exportTXT() {
    const lines: string[] = [];
    lines.push("EXAMENS_EXPORT");
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      const label = `${p.tipus} ${p.any} Q${p.quad}`;
      for (const { mon } of eachWeek(mondayOfWeek(parseISO(p.startStr)), fridayOfWeek(parseISO(p.endStr)))) {
        for (let si = 0; si < slots.length; si++) {
          for (let i = 0; i < 5; i++) {
            const day = addDays(mon, i);
            if (isDisabledDay(day, p)) continue;
            const dateIso = format(day, "yyyy-MM-dd");
            const key = cellKey(dateIso, si);
            const ids = amap[key] ?? [];
            ids.forEach((id) => {
              const subj = subjects.find((x) => x.id === id);
              if (!subj) return;
              lines.push(
                formatTxtLine(label, format(day,"dd/MM/yyyy"), si+1, slots[si]?.start ?? "", slots[si]?.end ?? "", subj)
              );
            });
          }
        }
      }
    }
    const blob = new Blob([lines.join("\n")], { type:"text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href=url; a.download="examenes.txt"; a.click();
    URL.revokeObjectURL(url);
  }

  /* Import CSV (assignatures + períodes/franges) */
  const handleImportCSV: React.ChangeEventHandler<HTMLInputElement> = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;

    const parseDate = (raw: any): string | undefined => {
      if (!raw) return undefined;
      const s = String(raw).trim();
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;                 // yyyy-MM-dd
      const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);             // dd/MM/yyyy
      if (m) return `${m[3]}-${m[2]}-${m[1]}`;
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
          const pad = (h: string) =>
            h.split(":").map((x) => x.padStart(2, "0")).join(":");
          return { start: pad(a), end: pad(b) };
        })
        .filter(Boolean) as TimeSlot[];
    };
    const normQuad = (raw: any): 1 | 2 | undefined => {
      if (raw == null || raw === "") return undefined;
      const n = Number(String(raw).replace(/\D/g, ""));
      return n === 1 || n === 2 ? (n as 1 | 2) : undefined;
    };

    Papa.parse(f, {
      header: true,
      skipEmptyLines: true,
      complete: (res: Papa.ParseResult<any>) => {
        try {
          const rows = (res.data as any[]).filter(Boolean);

          const outSubjects: Subject[] = [];
          const periodMap = new Map<number, PeriodMeta>(); // per period_id
          const slotsMap: SlotsPerPeriod = {};

          for (const r of rows) {
            // Assignatura
            const codigo =
              r.codigo || r.CODIGO || r.Codi || r["CODI UPC"] || r.codi || r.CODI;
            const siglas =
              r.siglas || r.SIGLAS || r.sigles || r["sigles"] || r.SIGLES;
            const nivel =
              r.nivel || r.NIVEL || r.nivell || r.NIVELL || r.level || r.LEVEL;
            const curs =
              r.curs || r.curso || r.curso_academico || r.curs_academic || r.CURS || r.CURSO;
            const quad = normQuad(
              r.quadrimestre || r.quad || r.quarter || r.Q || r.QUADRIMESTRE || r.QUAD
            );

            if (codigo || siglas) {
              outSubjects.push({
                id: String(codigo || siglas),
                codigo: String(codigo || ""),
                siglas: String(siglas || ""),
                nivel: String(nivel || ""),
                curs: curs ? String(curs) : undefined,
                quad,
              });
            }

            // Període opcional
            const pidRaw = r.period_id ?? r.PERIOD_ID ?? r.PeriodId;
            const pid = pidRaw ? Number(pidRaw) : NaN;
            if (!Number.isFinite(pid)) continue;
            if (pid < 1 || pid > 5) continue;

            if (!periodMap.has(pid)) {
              const tipus = (r.period_tipus || r.PERIOD_TIPUS || r.tipo || r.TIPO || "").toString().toUpperCase();
              const tipusNorm: TipusPeriode =
                tipus === "FINAL" ? "FINAL" :
                tipus === "REAVALUACIO" || tipus === "REAVALUACIÓ" || tipus === "REAVALUACION" ? "REAVALUACIÓ" :
                "PARCIAL";

              const anyNum = Number(r.period_any || r.PERIOD_ANY || r.year || r.ANY);
              const any =
                Number.isFinite(anyNum) && anyNum >= 2025 && anyNum <= 2090
                  ? anyNum
                  : new Date().getFullYear();

              const pquad = normQuad(r.period_quad ?? r.PERIOD_QUAD ?? r.quadrimestre ?? r.quad) || 1;


              const startStr =
                parseDate(r.period_inici || r.PERIOD_INICI || r.start) ||
                format(mondayOfWeek(new Date()), "yyyy-MM-dd");
              const endStr =
                parseDate(r.period_fi || r.PERIOD_FI || r.end) ||
                format(fridayOfWeek(new Date()), "yyyy-MM-dd");

              const slots =
                parseSlots(r.period_slots || r.PERIOD_SLOTS || r.slots) ||
                [{ start: "08:00", end: "10:00" }];

              periodMap.set(pid, {
                id: pid,
                tipus: tipusNorm,
                any,
                quad: pquad,
                startStr,
                endStr,
              });
              slotsMap[pid] = slots;
            }
          }

          // IDs únics
          const seen = new Set<string>();
          const uniqueSubjects = outSubjects.map((s) => {
            let id = s.id;
            while (seen.has(id)) id = id + "-" + Math.random().toString(36).slice(2,5);
            seen.add(id);
            return { ...s, id };
          });
          setSubjects(uniqueSubjects);

          if (periodMap.size > 0) {
            const ordered = Array.from(periodMap.keys()).sort((a,b)=>a-b);
            const list = ordered.map((k)=> periodMap.get(k)!);
            setPeriods(list);
            setSlotsPerPeriod(slotsMap);
            setAssignedPerPeriod({});
            setActivePid(list[0].id);
            alert(`Importades ${uniqueSubjects.length} assignatures i ${list.length} períodes del CSV.`);
          } else {
            alert(`Importades ${uniqueSubjects.length} assignatures del CSV.`);
          }
        } catch (err) {
          console.error(err);
          alert("Error processant el CSV");
        }
      },
      error: () => alert("No s'ha pogut llegir el fitxer CSV"),
    });

    (e.currentTarget as HTMLInputElement).value = "";
  };

  /* Render */
  return (
    <div className="p-6 max-w-[1200px] mx-auto">
      <h1 className="text-2xl font-bold mb-2">
        Planificador d'exàmens (CSV períodes/franges)
      </h1>
      <p className="text-sm mb-6">
        Pots definir els períodes i les franges manualment o importar-los
        directament des del CSV (columnes period_*).
      </p>

      {/* Dades i intercanvi */}
      <div className="p-4 rounded-2xl border shadow-sm bg-white mb-6">
        <h2 className="font-semibold mb-3">Dades i intercanvi</h2>
        <div className="flex flex-wrap gap-3 items-center">
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar CSV (assignatures + períodes)
            <input
              type="file"
              accept=".csv,text/csv"
              className="hidden"
              onChange={handleImportCSV}
            />
          </label>

          <button onClick={exportCSV} className="px-3 py-2 border rounded-xl shadow-sm">Exportar CSV</button>
          <button onClick={exportTXT} className="px-3 py-2 border rounded-xl shadow-sm">Exportar TXT</button>
          <button onClick={exportJSON} className="px-3 py-2 border rounded-xl shadow-sm">Exportar JSON</button>
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar JSON
            <input type="file" accept="application/json" className="hidden" onChange={importJSON} />
          </label>

          <span className="text-xs text-gray-500 ml-auto">
            Disponibles: {availableSubjects.length}/{subjects.length}
          </span>
        </div>
      </div>

      {/* Pestanyes de períodes */}
      <div className="mb-4 flex items-center justify-between gap-2">
        <div className="flex flex-wrap gap-2">
          {periods.map((p) => (
            <button
              key={p.id}
              onClick={() => setActivePid(p.id)}
              className={`px-3 py-2 rounded-xl border shadow-sm ${
                p.id === activePid ? "bg-indigo-50 border-indigo-300" : "bg-white"
              }`}
              title="Canviar de període"
            >
              {p.tipus} {p.any} Q{p.quad}
            </button>
          ))}
        </div>
        <div className="flex gap-2">
          <button onClick={addPeriod} className="px-3 py-2 border rounded-xl shadow-sm">Afegir període</button>
          {periods.length > 1 && (
            <button onClick={()=>removePeriod(activePid)} className="px-3 py-2 border rounded-xl shadow-sm">
              Eliminar període actiu
            </button>
          )}
        </div>
      </div>

      {/* Configuració del període actiu + franges */}
      {activePeriod && (
        <div className="grid md:grid-cols-3 gap-4 mb-6">
          <div className="p-4 rounded-2xl border shadow-sm bg-white">
            <h2 className="font-semibold mb-3">Configuració del període</h2>

            <label className="block text-sm mb-1">Tipus</label>
            <select
              value={activePeriod.tipus}
              onChange={(e) => {
                const v = e.target.value as TipusPeriode;
                setPeriods(periods.map(p => p.id===activePid? {...p, tipus: v}: p));
              }}
              className="w-full border rounded-xl p-2"
            >
              <option>PARCIAL</option>
              <option>FINAL</option>
              <option>REAVALUACIÓ</option>
            </select>

            <label className="block text-sm mt-3 mb-1">Any</label>
            <select
              value={activePeriod.any}
              onChange={(e) => {
                const v = Number(e.target.value);
                setPeriods(periods.map(p => p.id===activePid? {...p, any: v}: p));
              }}
              className="w-full border rounded-xl p-2"
            >
              {Array.from({length: 2090-2025+1}, (_,i)=>2025+i).map(y=>(
                <option key={y} value={y}>{y}</option>
              ))}
            </select>

            <label className="block text-sm mt-3 mb-1">Quadrimestre</label>
            <select
              value={activePeriod.quad}
              onChange={(e) => {
                const v = Number(e.target.value) as 1|2;
                setPeriods(periods.map(p => p.id===activePid? {...p, quad: v}: p));
                setFilterQuad(v);
              }}
              className="w-full border rounded-xl p-2"
            >
              <option value={1}>1</option>
              <option value={2}>2</option>
            </select>

            <label className="block text-sm mt-3 mb-1">Inici</label>
            <input
              type="date"
              value={activePeriod.startStr}
              onChange={(e)=> setPeriods(periods.map(p => p.id===activePid? {...p, startStr: e.target.value}: p))}
              className="w-full border rounded-xl p-2"
            />
            <label className="block text-sm mt-3 mb-1">Fi</label>
            <input
              type="date"
              value={activePeriod.endStr}
              onChange={(e)=> setPeriods(periods.map(p => p.id===activePid? {...p, endStr: e.target.value}: p))}
              className="w-full border rounded-xl p-2"
            />
          </div>

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

      {/* DnD cobreix calaix + calendari */}
      <DndContext onDragEnd={onDragEnd} modifiers={[restrictToWindowEdges]}>
        {/* Calaix d'assignatures (amb filtres) */}
        <div className="p-4 rounded-2xl border shadow-sm bg-white mb-6">
          <h2 className="font-semibold mb-3">Assignatures (arrossega)</h2>

          <div className="flex flex-wrap items-center gap-3 mb-3">
            <div className="text-sm">
              <label className="mr-2">Curs:</label>
              <select
                value={filterCurs}
                onChange={(e)=> setFilterCurs(e.target.value)}
                className="border rounded-xl p-2"
              >
                <option value="">(Tots)</option>
                {allCursos.map(c => <option key={c} value={c}>{c}</option>)}
              </select>
            </div>

            <div className="text-sm">
              <label className="mr-2">Quadrimestre:</label>
              <select
                value={filterQuad}
                onChange={(e)=> setFilterQuad(Number(e.target.value) as 0|1|2)}
                className="border rounded-xl p-2"
              >
                <option value={0}>(Tots)</option>
                <option value={1}>1</option>
                <option value={2}>2</option>
              </select>
            </div>

            <button
              onClick={()=>{ setFilterCurs(""); setFilterQuad(0); }}
              className="text-sm px-3 py-2 border rounded-xl shadow-sm"
            >
              Neteja filtres
            </button>
          </div>

          <div className="flex flex-wrap gap-2">
            {availableSubjects.map((s) => (
              <Chip
                key={s.id}
                id={s.id}
                label={`${s.siglas} · ${s.codigo} · ${s.nivel}${s.curs ? " · " + s.curs : ""}${s.quad ? " · Q" + s.quad : ""}`}
              />
            ))}
            {availableSubjects.length === 0 && (
              <div className="text-xs text-gray-500 italic">
                No queden assignatures per programar amb els filtres actuals.
              </div>
            )}
          </div>
        </div>

        {/* Calendari */}
        {activePeriod && (
          <div className="mb-8">
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-lg font-semibold">
                {activePeriod.tipus} {activePeriod.any} Q{activePeriod.quad} — {format(parseISO(activePeriod.startStr), "dd/MM")} a {format(parseISO(activePeriod.endStr), "dd/MM")}
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
                            const dateIso = format(day, "yyyy-MM-dd");
                            const disabled = isDisabledDay(day, activePeriod);
                            const amap = assignedPerPeriod[activePid] ?? {};
                            const key = cellKey(dateIso, slotIndex);
                            const subjIds = amap[key] ?? [];
                            const assignedList = subjIds
                              .map((id) => subjects.find((x) => x.id === id))
                              .filter(Boolean) as Subject[];
                            return (
                              <DropCell
                                key={i}
                                id={`cell:${activePid}:${dateIso}:${slotIndex}`}
                                disabled={disabled}
                                assignedList={assignedList}
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
          <li>Fins a 5 períodes amb pestanyes; cada període té les seves franges i dates.</li>
          <li>Importa CSV amb columnes <code>period_*</code> per definir períodes i franges d’un sol cop.</li>
          <li>Les assignatures programades desapareixen del calaix per evitar duplicats.</li>
        </ul>
      </div>
    </div>
  );
}
