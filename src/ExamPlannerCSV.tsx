// src/ExamPlannerCSV.tsx
import * as XLSX from "xlsx";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
} from "docx";
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
import {
  compressToEncodedURIComponent,
  decompressFromEncodedURIComponent,
} from "lz-string";

/* ---------- Helpers ---------- */
function mondayOfWeek(d: Date) {
  const day = d.getDay(); // 0=dg … 6=ds
  const diff = (day + 6) % 7; // 0 si dilluns
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
function iso(d: Date) {
  return format(d, "yyyy-MM-dd");
}

/* ---------- Tipus i models ---------- */
type TipusPeriode = "PARCIAL" | "FINAL" | "REAVALUACIÓ";

interface Period {
  id: number;
  label: string;
  tipus: TipusPeriode;
  startStr: string; // "yyyy-MM-dd"
  endStr: string; // "yyyy-MM-dd"
  curs?: number; // any acadèmic (inici)
  quad?: 1 | 2; // quadrimestre del període
  blackouts?: string[];
}

interface Subject {
  id: string; // id estable (normalment codi o, si no, sigles)
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

interface TimeSlot {
  start: string;
  end: string;
} // "HH:mm"

type AssignedMap = Record<string, string[]>; // "YYYY-MM-DD|slotIndex" → [subjectId,...]
type AssignedPerPeriod = Record<number, AssignedMap>;
type SlotsPerPeriod = Record<number, TimeSlot[]>;

/** Informació d’aules i matriculats per cel·la i assignatura */
type RoomsEnroll = {
  rooms: string[];
  students?: number; // 👈 NOM coherent
};
type RoomsMapPerCell = Record<string, RoomsEnroll>; // subjectId → info
type RoomsDataPerPeriod = Record<number, Record<string, RoomsMapPerCell>>; // pid → (dateIso|slotIdx) → map

/* ---------- Subcomponents ---------- */
function MastersLines({ s }: { s: Subject }) {
  const hasAny = s.MET || s.MATT || s.MEE || s.MCYBERS;
  if (!hasAny) return null;

  return (
    <div className="mt-1 text-[10px] leading-tight space-y-0.5">
      {/* MET en rojo, solo el valor */}
      {s.MET && (
        <div className="text-red-700">
          <span>{s.MET}</span>
        </div>
      )}

      {/* MATT en azul, solo el valor */}
      {s.MATT && (
        <div className="text-blue-700">
          <span>{s.MATT}</span>
        </div>
      )}

      {/* MEE sin etiqueta, color neutro (ajústalo si quieres otro) */}
      {s.MEE && (
        <div className="text-gray-600">
          <span>{s.MEE}</span>
        </div>
      )}

      {/* MCYBERS en verde, solo el valor */}
      {s.MCYBERS && (
        <div className="text-green-700">
          <span>{s.MCYBERS}</span>
        </div>
      )}
    </div>
  );
}

function TrayChip({ id, s }: { id: string; s: Subject }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } =
    useDraggable({ id });

  return (
    <div
      ref={setNodeRef}
      {...listeners}
      {...attributes}
      className={`relative inline-flex flex-col px-3 py-2 rounded-2xl shadow-sm border text-sm select-none bg-white cursor-grab active:cursor-grabbing ${
        isDragging ? "opacity-70 ring-2 ring-indigo-300" : ""
      }`}
      style={{
        transform: transform
          ? `translate3d(${transform.x}px, ${transform.y}px, 0)`
          : undefined,
        maxWidth: 300,
      }}
      title={`${s.sigles} · ${s.codi}`}
    >
      <span className="font-medium truncate">
        {s.sigles} · {s.codi}
      </span>
      {s.nivell ? (
        <span className="text-xs opacity-80 leading-4">Nivell: {s.nivell}</span>
      ) : (
        <MastersLines s={s} />
      )}
    </div>
  );
}

function PlacedChip({
  pid,
  dateIso,
  slotIndex,
  s,
  extra,
}: {
  pid: number;
  dateIso: string;
  slotIndex: number;
  s: Subject;
  extra?: RoomsEnroll;
}) {
  // id especial per moure entre cel·les
  const dragId = `placed:${pid}:${dateIso}:${slotIndex}:${s.id}`;
  const { attributes, listeners, setNodeRef, transform, isDragging } =
    useDraggable({ id: dragId });

  const hasRooms = extra && extra.rooms && extra.rooms.length > 0;
  const hasStud =
    extra &&
    typeof extra.students === "number" &&
    !Number.isNaN(extra.students);

  return (
    <div
      ref={setNodeRef}
      {...listeners}
      {...attributes}
      className={`relative p-2 rounded-xl border shadow-sm bg-gray-50 ${
        isDragging ? "opacity-70 ring-2 ring-indigo-400" : ""
      }`}
      style={{
        transform: transform
          ? `translate3d(${transform.x}px, ${transform.y}px, 0)`
          : undefined,
      }}
      title="Arrossega per moure a una altra franja"
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
    </div>
  );
}

function DropCell({
  id,
  disabled,
  assignedList,
  extrasForSubjects,
  onRemoveOne,
  pid,
  dateIso,
  slotIndex,
}: {
  id: string;
  disabled?: boolean;
  assignedList?: Subject[];
  extrasForSubjects?: Record<string, RoomsEnroll>;
  onRemoveOne?: (subjectId: string) => void;
  pid: number;
  dateIso: string;
  slotIndex: number;
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
      {assignedList && assignedList.length ? (
        <div className="space-y-2">
          {assignedList.map((s) => {
            const extra = extrasForSubjects?.[s.id];

            return (
              <div key={s.id} className="relative">
                {/* Capseta arrossegable entre cel·les, AMB aules/estudiants a dins */}
                <PlacedChip
                  pid={pid}
                  dateIso={dateIso}
                  slotIndex={slotIndex}
                  s={s}
                  extra={extra}
                />

                {!disabled && onRemoveOne && (
                  <button
                    onClick={() => onRemoveOne(s.id)}
                    onPointerDown={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                    }}
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


function TrashBin() {
  const { setNodeRef, isOver } = useDroppable({ id: "trash:catalog" });
  return (
    <div
      ref={setNodeRef}
      className={`fixed bottom-6 right-6 z-50 flex items-center gap-2 px-4 py-3 rounded-2xl border shadow-md bg-white
        ${isOver ? "ring-2 ring-red-400 bg-red-50" : ""}`}
      title="Arrossega aquí per eliminar del catàleg"
    >
      <span className="text-lg">🗑️</span>
      <span className="text-sm font-medium">Elimina del catàleg</span>
    </div>
  );
}

/* ---------- Component principal ---------- */
export default function ExamPlannerCSV() {
  /* Assignatures (demo inicial – es sobreescriuran amb CSV/JSON) */
  const [subjects, setSubjects] = useState<Subject[]>([
    {
      id: "MAT101",
      codi: "MAT101",
      sigles: "CALC I",
      nivell: "GRAU",
      curs: "2025",
      quadrimestre: 1,
    },
    {
      id: "FIS201",
      codi: "FIS201",
      sigles: "FIS II",
      nivell: "GRAU",
      curs: "2025",
      quadrimestre: 1,
    },
    {
      id: "PRG150",
      codi: "PRG150",
      sigles: "PRG",
      nivell: "GRAU",
      curs: "2025",
      quadrimestre: 1,
    },
    {
      id: "TIC500",
      codi: "TIC500",
      sigles: "CIBER",
      curs: "2025",
      quadrimestre: 2,
      MCYBERS: "Sí",
      MET: "Optativa",
    },
  ]);

  /* Períodes */
  const [periods, setPeriods] = useState<Period[]>([
    {
      id: 1,
      label: "Període 1",
      tipus: "PARCIAL",
      startStr: format(mondayOfWeek(new Date()), "yyyy-MM-dd"),
      endStr: format(fridayOfWeek(new Date()), "yyyy-MM-dd"),
      curs: undefined,
      quad: undefined,
      blackouts: [],
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
  const [assignedPerPeriod, setAssignedPerPeriod] =
    useState<AssignedPerPeriod>({});

  /* Aules/Matriculats per període/cel·la/assignatura */
  const [roomsData, setRoomsData] = useState<RoomsDataPerPeriod>({});

  /* Períodes on cada assignatura està permesa (derivat del CSV d’import) */
  const [allowedPeriodsBySubject, setAllowedPeriodsBySubject] = useState<
    Record<string, number[]>
  >({});

  /* NOVETAT: llista d’ocultes (eliminades manualment de la safata) */
  const [hiddenSubjectIds, setHiddenSubjectIds] = useState<string[]>([]);

  /* --- Estat per Desfer l’eliminació definitiva --- */
  type DeletedSnapshot = {
    subject: Subject;
    allowedPeriods?: number[];
    placed: Record<number, string[]>; // pid -> llista de cellKeys on era present
    rooms: Record<number, Record<string, RoomsEnroll>>; // pid -> cellKey -> info
  };
  const [lastDeleted, setLastDeleted] = useState<DeletedSnapshot | null>(null);

  // Caducitat automàtica del banner "Desfer"
  useEffect(() => {
    if (!lastDeleted) return;
    const t = setTimeout(() => setLastDeleted(null), 20000); // 20 segons
    return () => clearTimeout(t);
  }, [lastDeleted]);

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

  /* Assignatures ja utilitzades — NOMÉS en el període actiu */
  const usedIds = useMemo(() => {
    const amap = assignedPerPeriod[activePid] ?? {};
    const s = new Set<string>();
    for (const list of Object.values(amap)) for (const id of list) s.add(id);
    return s;
  }, [assignedPerPeriod, activePid]);

  /* Filtrat de la safata: quadrimestre + pertinença al període + no usada + no oculta */
  const availableSubjects = useMemo(() => {
    const pcurs =
      activePeriod?.curs != null ? String(activePeriod.curs) : undefined;
    const pquad = activePeriod?.quad;
    const pid = activePid;

    return subjects
      .filter((s) => !usedIds.has(s.id))
      .filter((s) => !hiddenSubjectIds.includes(s.id))
      .filter((s) => (pcurs ? s.curs === pcurs : true))
      .filter((s) => (pquad ? s.quadrimestre === pquad : true))
      .filter((s) => {
        const allowed = allowedPeriodsBySubject[s.id];
        return Array.isArray(allowed) ? allowed.includes(pid) : true;
      });
  }, [
    subjects,
    usedIds,
    activePeriod?.curs,
    activePeriod?.quad,
    activePid,
    allowedPeriodsBySubject,
    hiddenSubjectIds,
  ]);

  /* Guardar/Carregar estat a URL (hash) */
  function saveStateToUrl() {
    const payload = {
      subjects,
      periods,
      slotsPerPeriod,
      assignedPerPeriod,
      activePid,
      roomsData,
      allowedPeriodsBySubject,
      hiddenSubjectIds,
    };
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
      if (data.allowedPeriodsBySubject)
        setAllowedPeriodsBySubject(data.allowedPeriodsBySubject);
      if (Array.isArray(data.hiddenSubjectIds))
        setHiddenSubjectIds(data.hiddenSubjectIds);
      if (typeof data.activePid === "number") setActivePid(data.activePid);
      return true;
    } catch {
      return false;
    }
  }
  function copyLinkToClipboard() {
    if (!window.location.hash.includes("state=")) {
      saveStateToUrl();
      return;
    }
    navigator.clipboard
      .writeText(window.location.href)
      .then(() => alert("Enllaç copiat!"))
      .catch(() => alert("No s’ha pogut copiar l’enllaç."));
  }
  useEffect(() => {
    loadStateFromUrl();
  }, []);

  /* DnD - arrossegar des de safata o moure entre cel·les */

  function onDragEnd(e: any) {
    const activeId = e.active?.id as string | undefined;
    const dropId = e.over?.id as string | undefined;
    if (!activeId || !dropId) return;

    // 🗑️ Paperera global: eliminar del catàleg
    if (dropId === "trash:catalog") {
      // subjectId pot venir de safata (id = subjectId) o de calendari (placed:pid:dateIso:slotIndex:subjectId)
      const subjectId = activeId.startsWith("placed:")
        ? activeId.split(":").slice(-1)[0]
        : activeId;

      deleteSubjectPermanently(subjectId);
      return;
    }

    // Resta: drop a una cel·la de calendari
    if (!dropId.startsWith("cell:")) return;

    // dropId = cell:periodId:YYYY-MM-DD:slotIndex
    const [, dropPidStr, dropDateIso, dropSlotIndexStr] = dropId.split(":");
    const dropPid = Number(dropPidStr);
    const dropKey = cellKey(dropDateIso, Number(dropSlotIndexStr));

    if (activeId.startsWith("placed:")) {
      // Moure entre cel·les: placed:pid:dateIso:slotIndex:subjectId
      const [, srcPidStr, srcDateIso, srcSlotIndexStr, subjectId] =
        activeId.split(":");
      const srcPid = Number(srcPidStr);
      const srcKey = cellKey(srcDateIso, Number(srcSlotIndexStr));
      if (srcPid !== dropPid) return; // només dins del període actiu
      setAssignedPerPeriod((prev) => {
        const amap = { ...(prev[srcPid] ?? {}) };
        amap[srcKey] = (amap[srcKey] ?? []).filter((id) => id !== subjectId);
        if (!amap[srcKey]?.length) delete amap[srcKey];
        const destList = new Set(amap[dropKey] ?? []);
        destList.add(subjectId);
        amap[dropKey] = Array.from(destList);
        return { ...prev, [srcPid]: amap };
      });
      return;
    }

    // Si no és 'placed', és un id d’assignatura de la safata
    const subjectId = activeId;
    setAssignedPerPeriod((prev) => {
      const prevMap = prev[dropPid] ?? {};
      const nextList = new Set(prevMap[dropKey] ?? []);
      nextList.add(subjectId);
      return {
        ...prev,
        [dropPid]: { ...prevMap, [dropKey]: Array.from(nextList) },
      };
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
      const next = (prevMap[key] ?? []).filter((id) => id !== subjectId);
      const copy: AssignedMap = { ...prevMap };
      if (next.length) copy[key] = next;
      else delete copy[key];
      return { ...prev, [pid]: copy };
    });
  }

  /* --- Eliminar definitivament un subjecte del catàleg (amb Desfer) --- */
  function deleteSubjectPermanently(subjectId: string) {
    const subj = subjects.find((s) => s.id === subjectId);
    if (!subj) return;
    if (
      !confirm(
        `Eliminar definitivament "${
          subj.sigles || subj.codi
        }" del catàleg?\nS’esborrarà de la safata, del calendari i de les dades d’aules/estudiants.`
      )
    ) {
      return;
    }

    // 1) Snapshot per Desfer
    const placed: Record<number, string[]> = {};
    for (const [pidStr, amap] of Object.entries(assignedPerPeriod)) {
      const pid = Number(pidStr);
      const cells: string[] = [];
      for (const [cell, ids] of Object.entries(amap)) {
        if (ids.includes(subjectId)) cells.push(cell);
      }
      if (cells.length) placed[pid] = cells;
    }
    const roomsSnap: Record<number, Record<string, RoomsEnroll>> = {};
    for (const [pidStr, per] of Object.entries(roomsData)) {
      const pid = Number(pidStr);
      const perOut: Record<string, RoomsEnroll> = {};
      for (const [cellKey, map] of Object.entries(per)) {
        const entry = map[subjectId];
        if (entry)
          perOut[cellKey] = {
            rooms: [...(entry.rooms || [])],
            students: entry.students,
          };
      }
      if (Object.keys(perOut).length) roomsSnap[pid] = perOut;
    }
    const allowed = allowedPeriodsBySubject[subjectId];
    setLastDeleted({
      subject: subj,
      allowedPeriods: allowed ? [...allowed] : undefined,
      placed,
      rooms: roomsSnap,
    });

    // 2) Elimina de l’estat
    setSubjects((prev) => prev.filter((s) => s.id !== subjectId));

    setAllowedPeriodsBySubject((prev) => {
      const { [subjectId]: _drop, ...rest } = prev;
      return rest;
    });

    setHiddenSubjectIds((prev) => prev.filter((id) => id !== subjectId));

    setAssignedPerPeriod((prev) => {
      const copy: AssignedPerPeriod = {};
      for (const [pidStr, amap] of Object.entries(prev)) {
        const newMap: AssignedMap = {};
        for (const [cell, ids] of Object.entries(amap)) {
          const next = ids.filter((id) => id !== subjectId);
          if (next.length) newMap[cell] = next;
        }
        copy[Number(pidStr)] = newMap;
      }
      return copy;
    });

    setRoomsData((prev) => {
      const out: RoomsDataPerPeriod = {};
      for (const [pidStr, per] of Object.entries(prev)) {
        const newPer: Record<string, RoomsMapPerCell> = {};
        for (const [cellKey, map] of Object.entries(per)) {
          const { [subjectId]: _drop, ...rest } = map;
          if (Object.keys(rest).length) newPer[cellKey] = rest;
        }
        out[Number(pidStr)] = newPer;
      }
      return out;
    });
  }

  function undoDelete() {
    if (!lastDeleted) return;
    const snap = lastDeleted;

    // 1) subjecte
    setSubjects((prev) => {
      if (prev.some((s) => s.id === snap.subject.id)) return prev;
      return [...prev, snap.subject];
    });

    // 2) allowedPeriods
    if (snap.allowedPeriods) {
      setAllowedPeriodsBySubject((prev) => ({
        ...prev,
        [snap.subject.id]: [...snap.allowedPeriods!],
      }));
    }

    // 3) col·locacions
    setAssignedPerPeriod((prev) => {
      const copy: AssignedPerPeriod = { ...prev };
      for (const [pidStr, cells] of Object.entries(snap.placed)) {
        const pid = Number(pidStr);
        const amap = { ...(copy[pid] ?? {}) };
        for (const cell of cells) {
          const setIds = new Set(amap[cell] ?? []);
          setIds.add(snap.subject.id);
          amap[cell] = Array.from(setIds);
        }
        copy[pid] = amap;
      }
      return copy;
    });

    // 4) rooms
    setRoomsData((prev) => {
      const out: RoomsDataPerPeriod = JSON.parse(
        JSON.stringify(prev || {})
      );
      for (const [pidStr, per] of Object.entries(snap.rooms)) {
        const pid = Number(pidStr);
        out[pid] = out[pid] || {};
        for (const [cellKey, info] of Object.entries(per)) {
          out[pid][cellKey] = out[pid][cellKey] || {};
          out[pid][cellKey][snap.subject.id] = {
            rooms: [...(info.rooms || [])],
            students: info.students,
          };
        }
      }
      return out;
    });

    setLastDeleted(null);
  }

  /* Gestió períodes */
  function addPeriod() {
    if (periods.length >= 5) {
      alert("Pots tenir com a màxim 5 períodes.");
      return;
    }
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
    setSlotsPerPeriod((sp) => ({
      ...sp,
      [newId]: [{ start: "08:00", end: "10:00" }],
    }));
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
    setRoomsData((rd) => {
      const c = { ...rd };
      delete c[id];
      return c;
    });
  }

  /* Exportacions */
  function exportJSON() {
    const data = {
      periods,
      slotsPerPeriod,
      assignedPerPeriod,
      subjects,
      roomsData,
      allowedPeriodsBySubject,
      hiddenSubjectIds,
    };
    const blob = new Blob([JSON.stringify(data, null, 2)], {
      type: "application/json",
    });
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
        if (data.assignedPerPeriod)
          setAssignedPerPeriod(data.assignedPerPeriod);
        if (Array.isArray(data.subjects)) setSubjects(data.subjects);
        if (data.roomsData) setRoomsData(data.roomsData);
        if (data.allowedPeriodsBySubject)
          setAllowedPeriodsBySubject(data.allowedPeriodsBySubject);
        if (Array.isArray(data.hiddenSubjectIds))
          setHiddenSubjectIds(data.hiddenSubjectIds);
        if (Array.isArray(data.periods) && data.periods.length)
          setActivePid(data.periods[0].id);
      } catch {
        alert("JSON no vàlid");
      }
    };
    reader.readAsText(f);
    ev.currentTarget.value = "";
  }

  // Inferència (per si falta curs/quad a alguna assignatura puntual)
  function inferCursFromDate(d: Date): string {
    const y = d.getFullYear(),
      m = d.getMonth() + 1;
    return (m >= 9 ? y : y - 1).toString(); // curs = any d'inici
  }
  function inferQuadFromDate(d: Date): 1 | 2 {
    const m = d.getMonth() + 1;
    return m >= 9 || m === 1 ? 1 : 2; // set–gen: Q1; feb–jul: Q2
  }

  function exportCSV() {
    const lines: string[] = [];
    for (const p of periods) {
      const slots = slotsPerPeriod[p.id] ?? [];
      const amap = assignedPerPeriod[p.id] ?? {};
      for (const { mon } of eachWeek(
        mondayOfWeek(parseISO(p.startStr)),
        fridayOfWeek(parseISO(p.endStr))
      )) {
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
              const CURS =
                s.curs?.toString() ??
                String(p.curs ?? inferCursFromDate(day));
              const QUADRIMESTRE = String(
                s.quadrimestre ?? p.quad ?? inferQuadFromDate(day)
              );
              const TIPUS_EXAMEN =
                p.tipus === "REAVALUACIÓ" ? "REAVALUACIO" : p.tipus;
              const DIA = format(day, "dd-MM-yyyy");
              const HORA_INICI = slots[si].start;
              const HORA_FI = slots[si].end;
              const UNITAT_DOCENT = s.codi;
              const GRUPS = "";
              lines.push(
                [
                  CENTRE,
                  CURS,
                  QUADRIMESTRE,
                  TIPUS_EXAMEN,
                  DIA,
                  HORA_INICI,
                  HORA_FI,
                  UNITAT_DOCENT,
                  GRUPS,
                ].join(",")
              );
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

function formatSubjectForCell(s: Subject, extra?: RoomsEnroll): string {
  const lines: string[] = [];

  // Línia principal
  lines.push(`${s.codi} · ${s.sigles}`);

  // Nivell + camps MastersTIC
  const extraLines: string[] = [];
  if (s.nivell) extraLines.push(s.nivell);
  if (s.MET) extraLines.push(s.MET);
  if (s.MATT) extraLines.push(s.MATT);
  if (s.MEE) extraLines.push(s.MEE);
  if (s.MCYBERS) extraLines.push(s.MCYBERS);
  if (extraLines.length) lines.push(extraLines.join(" · "));

  // Aules / Estudiants (procedents del CSV de roomsData)
  if (extra) {
    const hasRooms = extra.rooms && extra.rooms.length > 0;
    const hasStud =
      typeof extra.students === "number" && Number.isFinite(extra.students);

    if (hasRooms || hasStud) {
      if (hasRooms) {
        lines.push(`Aules/Rooms: ${extra.rooms.join(", ")}`);
      }
      if (hasStud) {
        lines.push(`Estudiants/Students: ${extra.students}`);
      }
    }
  }

  return lines.join("\n");
}

  // Construir los párrafos de una asignatura para Word (con colores)
  function buildSubjectParagraphsForWord(
    s: Subject,
    extra?: RoomsEnroll
  ): Paragraph[] {
    const paras: Paragraph[] = [];

    // 1) Línea principal: codi · sigles
    paras.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `${s.codi} · ${s.sigles}`,
            bold: true,
          }),
        ],
      })
    );

    // 2) Nivell (si existe)
    if (s.nivell) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({
              text: `Nivell: ${s.nivell}`,
            }),
          ],
        })
      );
    }

    // 3) Camps MastersTIC con colores:
    //   - MATT → azul
    //   - MET (si la quieres, de momento negro normal)
    //   - MCYBERS → verde
    //   - MEE → rojo
    if (s.MATT) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: s.MATT, color: "0000FF" }), // azul
          ],
        })
      );
    }

    if (s.MET) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: s.MET }), // sin color especial
          ],
        })
      );
    }

    if (s.MCYBERS) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: s.MCYBERS, color: "008000" }), // verde
          ],
        })
      );
    }

    if (s.MEE) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: s.MEE, color: "FF0000" }), // rojo
          ],
        })
      );
    }

    // 4) Aules / Estudiants (si hay datos importados de roomsData)
    if (extra?.rooms && extra.rooms.length > 0) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: "Aules/Rooms: ", bold: true }),
            new TextRun({ text: extra.rooms.join(", ") }),
          ],
        })
      );
    }

    if (
      extra &&
      typeof extra.students === "number" &&
      Number.isFinite(extra.students)
    ) {
      paras.push(
        new Paragraph({
          children: [
            new TextRun({ text: "Estudiants/Students: ", bold: true }),
            new TextRun({ text: String(extra.students) }),
          ],
        })
      );
    }

    return paras;
  }


  function exportExcel() {
    try {
      const wb = XLSX.utils.book_new();

      const slotColors = ["E3F2FD", "E8F5E9", "FFF3E0", "F3E5F5", "E0F7FA", "FBE9E7"];
      const dayLabelsCat = ["Dl", "Dt", "Dc", "Dj", "Dv"];
      const dayLabelsEn = ["Mon", "Tue", "Wed", "Thu", "Fri"];

      for (const p of periods) {
        const slots = slotsPerPeriod[p.id] ?? [];
        if (!slots.length) continue;

        const amap = assignedPerPeriod[p.id] ?? {};
        const roomsForPeriod = roomsData[p.id] ?? {};
        const rows: any[][] = [];
        const slotIndexPerRow: number[] = []; // -1 = capçalera/blank, >=0 = fila franja

        const start = mondayOfWeek(parseISO(p.startStr));
        const end = fridayOfWeek(parseISO(p.endStr));
        let weekStart = start;

        while (weekStart <= end) {
          if (rows.length > 0) {
            rows.push([]);
            slotIndexPerRow.push(-1);
          }

          const headerWeek: any[] = ["Franja horària / Time slot"];
          for (let di = 0; di < 5; di++) {
            const day = addDays(weekStart, di);
            headerWeek.push(
              `${dayLabelsCat[di]}/${dayLabelsEn[di]} ${format(
                day,
                "dd/MM"
              )}`
            );
          }
          rows.push(headerWeek);
          slotIndexPerRow.push(-1);

          slots.forEach((slot, si) => {
            const row: any[] = [`${slot.start}-${slot.end}`];

            for (let di = 0; di < 5; di++) {
              const day = addDays(weekStart, di);
              if (isDisabledDay(day, p)) {
                row.push("");
                continue;
              }
              const dateIso = format(day, "yyyy-MM-dd");
              const key = `${dateIso}|${si}`;
              const ids = amap[key] ?? [];
              const list = ids
                .map((id) => subjects.find((s) => s.id === id))
                .filter(Boolean) as Subject[];
             // Map amb la info d’aules/matriculats per assignatura en aquesta cel·la
             const roomsMap: RoomsMapPerCell = roomsForPeriod[key] ?? {};
              
             const cellText = list
                  .map((s) => formatSubjectForCell(s, roomsMap[s.id]))
                  .join("\n\n");
              row.push(cellText);
            }

            rows.push(row);
            slotIndexPerRow.push(si);
          });

          weekStart = addDays(weekStart, 7);
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        const range = XLSX.utils.decode_range(ws["!ref"] as string);

        const cols: any[] = [{ wch: 20 }];
        for (let i = 0; i < 5; i++) cols.push({ wch: 40 });
        (ws as any)["!cols"] = cols;
        (ws as any)["!rows"] = rows.map(() => ({ hpt: 36 }));

        for (let r = 0; r <= range.e.r; r++) {
          if (slotIndexPerRow[r] !== -1) continue;
          const first = rows[r]?.[0];
          if (first !== "Franja horària / Time slot") continue;
          for (let c = 0; c <= 5; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = (ws as any)[addr];
            if (!cell) continue;
            cell.s = {
              font: { bold: true },
              alignment: { horizontal: "center" },
              fill: { fgColor: { rgb: "E0E0E0" } },
            };
          }
        }

        for (let r = 0; r <= range.e.r; r++) {
          const si = slotIndexPerRow[r];
          if (si < 0) continue;
          const color = slotColors[si % slotColors.length];
          for (let c = 1; c <= 5; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = (ws as any)[addr];
            if (!cell) continue;
            const existing = cell.s ?? {};
            cell.s = {
              ...existing,
              alignment: {
                vertical: "top",
                wrapText: true,
                ...(existing.alignment || {}),
              },
              fill: { fgColor: { rgb: color } },
            };
          }
        }

        XLSX.utils.book_append_sheet(wb, ws, `${p.tipus}_id${p.id}`);
      }

      const wbout = XLSX.write(wb, {
        bookType: "xlsx",
        type: "array",
      });
      const blob = new Blob([wbout], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "calendari_examens.xlsx";
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      console.error("Error exportant l'Excel", err);
      alert(
        "No s'ha pogut exportar l'Excel. Revisa la consola del navegador per a més detalls."
      );
    }
  }

  async function exportWord() {
    try {
      const dayLabelsCat = ["Dl", "Dt", "Dc", "Dj", "Dv"];
      const dayLabelsEn = ["Mon", "Tue", "Wed", "Thu", "Fri"];

      const sectionChildren: (Paragraph | Table)[] = [];

      for (const p of periods) {
        const slots = slotsPerPeriod[p.id] ?? [];
        if (!slots.length) continue;

        const amap = assignedPerPeriod[p.id] ?? {};
        const roomsForPeriod = roomsData[p.id] ?? {};

        // Títol del període
        sectionChildren.push(
          new Paragraph({
            children: [
              new TextRun({
                text: `${p.label} — ${p.tipus}`,
                bold: true,
                size: 28, // ~14pt
              }),
            ],
            spacing: { before: 200, after: 200 },
          })
        );

        const start = mondayOfWeek(parseISO(p.startStr));
        const end = fridayOfWeek(parseISO(p.endStr));
        let weekStart = start;

        while (weekStart <= end) {
          // Títol de setmana
          sectionChildren.push(
            new Paragraph({
              children: [
                new TextRun({
                  text: `Setmana ${fmtDM(weekStart)} — ${fmtDM(
                    addDays(weekStart, 4)
                  )}`,
                  bold: true,
                }),
              ],
              spacing: { before: 200, after: 100 },
            })
          );

          const rows: TableRow[] = [];

          // Fila de capçalera
          const headerCells: TableCell[] = [];

          headerCells.push(
            new TableCell({
              children: [
                new Paragraph({
                  children: [
                    new TextRun({
                      text: "Franja horària / Time slot",
                      bold: true,
                    }),
                  ],
                }),
              ],
            })
          );

          for (let di = 0; di < 5; di++) {
            const day = addDays(weekStart, di);
            const label = `${dayLabelsCat[di]}/${dayLabelsEn[di]} ${fmtDM(
              day
            )}`;
            headerCells.push(
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({ text: label, bold: true }),
                    ],
                  }),
                ],
              })
            );
          }

          rows.push(
            new TableRow({
              children: headerCells,
            })
          );

          // Files per franja horària
          slots.forEach((slot, si) => {
            const rowCells: TableCell[] = [];

            // Primera columna: franja horària
            rowCells.push(
              new TableCell({
                children: [
                  new Paragraph({
                    children: [
                      new TextRun({
                        text: `${slot.start}-${slot.end}`,
                        bold: true,
                      }),
                    ],
                  }),
                ],
              })
            );

            // Columnes per dia
            for (let di = 0; di < 5; di++) {
              const day = addDays(weekStart, di);

              if (isDisabledDay(day, p)) {
                // Dia fora de període o blackout: cel·la buida
                rowCells.push(
                  new TableCell({
                    children: [new Paragraph({ text: "" })],
                  })
                );
                continue;
              }

              const dateIso = iso(day);
              const key = cellKey(dateIso, si);
              const ids = amap[key] ?? [];
              const list = ids
                .map((id) => subjects.find((s) => s.id === id))
                .filter(Boolean) as Subject[];

              const extrasCell = roomsForPeriod[key] ?? {};

              const cellParas: Paragraph[] = [];

              list.forEach((s, idx) => {
                const extra = extrasCell[s.id];
                const subjectParas = buildSubjectParagraphsForWord(
                  s,
                  extra
                );
                cellParas.push(...subjectParas);

                // Línia en blanc entre assignatures
                if (idx < list.length - 1) {
                  cellParas.push(new Paragraph({ text: "" }));
                }
              });

              if (!cellParas.length) {
                cellParas.push(new Paragraph({ text: "" }));
              }

              rowCells.push(
                new TableCell({
                  children: cellParas,
                })
              );
            }

            rows.push(
              new TableRow({
                children: rowCells,
              })
            );
          });

          const table = new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows,
            borders: {
              top: { style: BorderStyle.SINGLE, size: 2 },
              bottom: { style: BorderStyle.SINGLE, size: 2 },
              left: { style: BorderStyle.SINGLE, size: 2 },
              right: { style: BorderStyle.SINGLE, size: 2 },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 1 },
              insideVertical: { style: BorderStyle.SINGLE, size: 1 },
            },
          });

          sectionChildren.push(table);

          weekStart = addDays(weekStart, 7);
        }
      }

      const doc = new Document({
        sections: [
          {
            properties: {},
            children: sectionChildren,
          },
        ],
      });

      const blob = await Packer.toBlob(doc);
      const a = document.createElement("a");
      a.href = URL.createObjectURL(blob);
      a.download = "calendari_examens.docx";
      a.click();
      URL.revokeObjectURL(a.href);
    } catch (err) {
      console.error("Error exportant el Word", err);
      alert(
        "No s'ha pogut exportar el Word. Revisa la consola del navegador per a més detalls."
      );
    }
  }


  function exportTXT() {
    const LEN = {
      CODI: 10,
      CURS: 4,
      QUAD: 1,
      NOM: 120,
      DIA: 10,
      HORA: 5,
      DESC: 2000,
    } as const;
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
      for (const { mon } of eachWeek(
        mondayOfWeek(parseISO(p.startStr)),
        fridayOfWeek(parseISO(p.endStr))
      )) {
        for (let i = 0; i < 5; i++) {
          const day = addDays(mon, i);
          if (isDisabledDay(day, p)) continue;
          for (let si = 0; si < slots.length; si++) {
            const ids = amap[`${iso(day)}|${si}`] ?? [];
            if (!ids.length) continue;
            for (const id of ids) {
              const s = subjects.find((x) => x.id === id);
              if (!s) continue;
              const CODI = padText(s.codi, LEN.CODI);
              const CURS = padNum(
                s.curs ?? String(p.curs ?? inferCursFromDate(day)),
                LEN.CURS
              );
              const QUAD = padNum(
                s.quadrimestre ?? p.quad ?? inferQuadFromDate(day),
                LEN.QUAD
              );
              const NOM = padText(s.sigles, LEN.NOM);
              const DIA = padText(
                format(day, "dd-MM-yyyy"),
                LEN.DIA
              );
              const HORA = padText(
                (slots[si].start || "").replace(":", "-"),
                LEN.HORA
              );
              const DESC = padText(
                p.tipus === "REAVALUACIÓ" ? "REAVALUACIO" : p.tipus,
                LEN.DESC
              );
              lines.push(
                [CODI, CURS, QUAD, NOM, DIA, HORA, DESC].join(" ")
              );
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

  /* ---------- Import CSV (assignatures + períodes) — REPLACE ---------- */
  const handleImportCSV: React.ChangeEventHandler<HTMLInputElement> = (
    e
  ) => {
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
    const parseSlots = (raw: any): TimeSlot[] => {
      if (!raw) return [];
      return String(raw)
        .split(/[;,|]/)
        .map((p) => p.trim())
        .filter(Boolean)
        .map((pair) => {
          const mm = pair.match(
            /^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$/
          );
          if (!mm) return null;
          const [_, a, b] = mm;
          const pad = (h: string) =>
            h
              .split(":")
              .map((x) => x.padStart(2, "0"))
              .join(":");
          return { start: pad(a), end: pad(b) };
        })
        .filter(Boolean) as TimeSlot[];
    };
    const parseBlackouts = (raw: any): string[] => {
      if (!raw) return [];
      const toks = String(raw)
        .split(/[;,|]/)
        .map((s) => s.trim())
        .filter(Boolean);
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

          const subjByKey = new Map<string, Subject>();
          const periodsByKey = new Map<string, Set<number>>();
          const periodMap = new Map<number, Period>();
          const slotsMap: SlotsPerPeriod = {};
          const quadSeenPerPid = new Map<number, 1 | 2>();
          const cursSeenPerPid = new Map<number, number>();
          const keyOf = (codi: any, sigles: any) =>
            `${String(codi || "")
              .trim()
              .toLowerCase()}||${String(sigles || "")
              .trim()
              .toLowerCase()}`;

          for (const r of rows) {
            const codi =
              r.codi ?? r.codigo ?? r.CODI ?? r.CODIGO ?? r.code;
            const sigles =
              r.sigles ?? r.SIGLES ?? r.siglas ?? r.SIGLAS;
            const nivell = (
              r.nivell ?? r.NIVELL ?? r.nivel ?? r.NIVEL
            )?.toString();
            const curs = normCursAny(
              r.curs ?? r.CURS ?? r.curso ?? r.CURSO
            );
            const quadrimestre = normQuad(
              r.quadrimestre ?? r.QUADRIMESTRE ?? r.quad ?? r.QUAD
            );
            const MET = r.MET ?? r.met;
            const MATT = r.MATT ?? r.matt;
            const MEE = r.MEE ?? r.mee;
            const MCYBERS = r.MCYBERS ?? r.mcybers;

            if (codi || sigles) {
              const k = keyOf(codi, sigles);
              if (!subjByKey.has(k)) {
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
                const existed = subjByKey.get(k)!;
                if (!existed.nivell && nivell) existed.nivell = nivell;
                if (!existed.curs && curs) existed.curs = curs;
                if (!existed.quadrimestre && quadrimestre)
                  existed.quadrimestre = quadrimestre;
                if (!existed.MET && MET) existed.MET = String(MET);
                if (!existed.MATT && MATT) existed.MATT = String(MATT);
                if (!existed.MEE && MEE) existed.MEE = String(MEE);
                if (!existed.MCYBERS && MCYBERS)
                  existed.MCYBERS = String(MCYBERS);
              }
            }

            const pidRaw =
              r.period_id ??
              r.PERIOD_ID ??
              r.PeriodId ??
              r.periode ??
              r.PERIODO ??
              r.PERIOD;
            const pid = pidRaw ? Number(pidRaw) : NaN;
            if (Number.isFinite(pid) && pid >= 1) {
              const k = keyOf(codi, sigles);
              if (k.trim() !== "||") {
                if (!periodsByKey.has(k))
                  periodsByKey.set(k, new Set<number>());
                periodsByKey.get(k)!.add(pid);
              }
            }

            const tipusRaw = (
              r.period_tipus ??
              r.PERIOD_TIPUS ??
              r.tipo ??
              r.TIPO ??
              ""
            )
              .toString()
              .toUpperCase();
            const tipusNorm: TipusPeriode =
              tipusRaw === "FINAL"
                ? "FINAL"
                : tipusRaw === "REAVALUACIO" ||
                  tipusRaw === "REAVALUACIÓ" ||
                  tipusRaw === "REAVALUACION"
                ? "REAVALUACIÓ"
                : "PARCIAL";

            const startStr =
              parseDate(
                r.period_inici ?? r.PERIOD_INICI ?? r.start
              ) ||
              format(mondayOfWeek(new Date()), "yyyy-MM-dd");
            const endStr =
              parseDate(r.period_fi ?? r.PERIOD_FI ?? r.end) ||
              format(fridayOfWeek(new Date()), "yyyy-MM-dd");

            const parseSlotsLocal = (raw: any): TimeSlot[] => {
              if (!raw) return [];
              return String(raw)
                .split(/[;,|]/)
                .map((p) => p.trim())
                .filter(Boolean)
                .map((pair) => {
                  const mm = pair.match(
                    /^(\d{1,2}:\d{2})\s*-\s*(\d{1,2}:\d{2})$/
                  );
                  if (!mm) return null;
                  const [_, a, b] = mm;
                  const pad = (h: string) =>
                    h
                      .split(":")
                      .map((x) => x.padStart(2, "0"))
                      .join(":");
                  return { start: pad(a), end: pad(b) };
                })
                .filter(Boolean) as TimeSlot[];
            };
            const slots =
              parseSlotsLocal(
                r.period_slots ?? r.PERIOD_SLOTS ?? r.slots
              ) || [{ start: "08:00", end: "10:00" }];
            const blackouts = (() => {
              const raw =
                r.period_blackouts ??
                r.PERIOD_BLACKOUTS ??
                r.blackouts ??
                r.BLOCKED_DATES;
              if (!raw) return [];
              const toks = String(raw)
                .split(/[;,|]/)
                .map((s) => s.trim())
                .filter(Boolean);
              const out: string[] = [];
              for (const t of toks) {
                const d = parseDate(t);
                if (d) out.push(d);
              }
              return Array.from(new Set(out)).sort();
            })();

            const periodCurs = normCursAny(
              r.period_curs ?? r.PERIOD_CURS
            );
            const periodQuad = normQuad(
              r.period_quad ?? r.PERIOD_QUAD
            );

            const filaCurs = normCursAny(
              r.curs ?? r.CURS ?? r.curso ?? r.CURSO
            );
            const filaQuad = normQuad(
              r.quadrimestre ??
                r.QUADRIMESTRE ??
                r.quad ??
                r.QUAD
            );
            if (Number.isFinite(pid)) {
              if (filaQuad) quadSeenPerPid.set(pid, filaQuad);
              if (filaCurs) cursSeenPerPid.set(pid, Number(filaCurs));
            }

            if (Number.isFinite(pid) && pid >= 1) {
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
                if (!p.curs && periodCurs)
                  p.curs = Number(periodCurs);
                if (!p.quad && periodQuad) p.quad = periodQuad;
              }
            }
          }

          for (const [pid, p] of periodMap) {
            if (p.quad == null && quadSeenPerPid.has(pid))
              p.quad = quadSeenPerPid.get(pid)!;
            if (p.curs == null && cursSeenPerPid.has(pid))
              p.curs = cursSeenPerPid.get(pid)!;
          }

          const uniqueSubjects = Array.from(subjByKey.values());
          const nextAllowed: Record<string, number[]> = {};
          for (const s of uniqueSubjects) {
            const key = `${s.codi
              .trim()
              .toLowerCase()}||${s.sigles.trim().toLowerCase()}`;
            const set = periodsByKey.get(key);
            if (set && set.size)
              nextAllowed[s.id] = Array.from(set).sort(
                (a, b) => a - b
              );
          }

          setSubjects(uniqueSubjects);
          if (periodMap.size > 0) {
            const ordered = Array.from(periodMap.keys()).sort(
              (a, b) => a - b
            );
            const list = ordered.map((k) => periodMap.get(k)!);
            setPeriods(list);
            setSlotsPerPeriod(slotsMap);
            setAssignedPerPeriod({});
            setRoomsData({});
            setAllowedPeriodsBySubject(nextAllowed);
            setHiddenSubjectIds([]);
            setActivePid(list[0].id);
            alert(
              `Importades ${uniqueSubjects.length} assignatures i ${list.length} períodes del CSV.`
            );
          } else {
            setAllowedPeriodsBySubject(nextAllowed);
            setHiddenSubjectIds([]);
            alert(
              `Importades ${uniqueSubjects.length} assignatures del CSV.`
            );
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

/* ---------- Import CSV (assignatures + períodes) — MERGE ---------- */
const handleMergeSubjectsCSV: React.ChangeEventHandler<HTMLInputElement> = (e) => {
  const f = e.target.files?.[0];
  if (!f) return;

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

  Papa.parse(f, {
    header: true,
    skipEmptyLines: true,
    complete: (res: Papa.ParseResult<any>) => {
      try {
        const rows = (res.data as any[]).filter(Boolean);

        // Índexos auxiliars actuals
        const subjById = new Map(subjects.map(s => [s.id, s] as const));
        const subjKeyIndex = new Map<string, string>(); // key -> subjectId
        for (const s of subjects) {
          const key = `${s.codi.trim().toLowerCase()}||${s.sigles.trim().toLowerCase()}`;
          subjKeyIndex.set(key, s.id);
        }

        const nextSubjects = [...subjects];
        const nextAllowed = { ...allowedPeriodsBySubject };
        const nextPeriods = [...periods];
        const nextSlotsPerPeriod: SlotsPerPeriod = JSON.parse(JSON.stringify(slotsPerPeriod));

        const keyOf = (codi: any, sigles: any) =>
          `${String(codi || "").trim().toLowerCase()}||${String(sigles || "").trim().toLowerCase()}`;

        let addedSubjects = 0;
        let updatedSubjects = 0;
        let addedPeriods = 0;

        for (const r of rows) {
          const codi = r.codi ?? r.codigo ?? r.CODI ?? r.CODIGO ?? r.code;
          const sigles = r.sigles ?? r.SIGLES ?? r.siglas ?? r.SIGLAS;
          if (!codi && !sigles) continue;
          const k = keyOf(codi, sigles);

          const nivell = (r.nivell ?? r.NIVELL ?? r.nivel ?? r.NIVEL)?.toString();
          const curs = normCursAny(r.curs ?? r.CURS ?? r.curso ?? r.CURSO);
          const quadrimestre = normQuad(r.quadrimestre ?? r.QUADRIMESTRE ?? r.quad ?? r.QUAD);
          const MET = r.MET ?? r.met;
          const MATT = r.MATT ?? r.matt;
          const MEE = r.MEE ?? r.mee;
          const MCYBERS = r.MCYBERS ?? r.mcybers;

          let subjectId = subjKeyIndex.get(k);
          if (!subjectId) {
            subjectId = String(codi || sigles);
            subjKeyIndex.set(k, subjectId);
            nextSubjects.push({
              id: subjectId,
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
            addedSubjects++;
          } else {
            const s =
              subjById.get(subjectId) ||
              nextSubjects.find(x => x.id === subjectId)!;

            if (!s.nivell && nivell) { s.nivell = nivell; updatedSubjects++; }
            if (!s.curs && curs) { s.curs = curs; updatedSubjects++; }
            if (!s.quadrimestre && quadrimestre) { s.quadrimestre = quadrimestre; updatedSubjects++; }
            if (!s.MET && MET) { s.MET = String(MET); updatedSubjects++; }
            if (!s.MATT && MATT) { s.MATT = String(MATT); updatedSubjects++; }
            if (!s.MEE && MEE) { s.MEE = String(MEE); updatedSubjects++; }
            if (!s.MCYBERS && MCYBERS) { s.MCYBERS = String(MCYBERS); updatedSubjects++; }
          }

          const pidRaw =
            r.period_id ??
            r.PERIOD_ID ??
            r.PeriodId ??
            r.periode ??
            r.PERIODO ??
            r.PERIOD;
          const pid = pidRaw ? Number(pidRaw) : NaN;

          if (Number.isFinite(pid) && pid >= 1) {
            // Afegir període si no existeix
            if (!nextPeriods.find(p => p.id === pid)) {
              const tipusRaw = (
                r.period_tipus ??
                r.PERIOD_TIPUS ??
                r.tipo ??
                r.TIPO ??
                ""
              ).toString().toUpperCase();

              const tipus: TipusPeriode =
                tipusRaw === "FINAL"
                  ? "FINAL"
                  : (tipusRaw === "REAVALUACIO" ||
                     tipusRaw === "REAVALUACIÓ" ||
                     tipusRaw === "REAVALUACION")
                    ? "REAVALUACIÓ"
                    : "PARCIAL";

              const startStr =
                parseDate(r.period_inici ?? r.PERIOD_INICI ?? r.start) ||
                format(mondayOfWeek(new Date()), "yyyy-MM-dd");

              const endStr =
                parseDate(r.period_fi ?? r.PERIOD_FI ?? r.end) ||
                format(fridayOfWeek(new Date()), "yyyy-MM-dd");

              nextPeriods.push({
                id: pid,
                label: `Període ${pid}`,
                tipus,
                startStr,
                endStr,
                blackouts: [],
              });

              nextSlotsPerPeriod[pid] =
                nextSlotsPerPeriod[pid] ?? [{ start: "08:00", end: "10:00" }];

              addedPeriods++;
            }

            const arr = new Set(nextAllowed[subjectId] ?? []);
            arr.add(pid);
            nextAllowed[subjectId] = Array.from(arr).sort((a, b) => a - b);
          }
        }

        setSubjects(nextSubjects);
        setAllowedPeriodsBySubject(nextAllowed);
        setPeriods(nextPeriods.sort((a, b) => a.id - b.id));
        setSlotsPerPeriod(nextSlotsPerPeriod);

        alert(
          `Afegides ${addedSubjects} assignatures (actualitzades ${updatedSubjects}). Nous períodes: ${addedPeriods}.`
        );
      } catch (err) {
        console.error(err);
        alert("Error processant el CSV (merge).");
      }
    },
    error: () => alert("No s'ha pogut llegir el fitxer CSV (merge)"),
  });

  e.currentTarget.value = "";
};


  /* ---------- Helpers per dates d'aules ---------- */
  function normalizeDateToIso(raw: any): string | undefined {
    if (raw == null) return undefined;
    const s = String(raw).trim();
    if (!s) return undefined;

    // yyyy-MM-dd
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    // dd/MM/yyyy
    let m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;

    // dd-MM-yyyy
    m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;

    // Excel serial (dies des de 1899-12-30 aprox.)
    const n = Number(s);
    if (!Number.isNaN(n) && n > 30000 && n < 80000) {
      const excelEpoch = new Date(1899, 11, 30);
      const d = new Date(excelEpoch.getTime() + n * 86400000);
      const yy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yy}-${mm}-${dd}`;
    }

    return undefined;
  }

  function normalizeTime(raw: any): string | undefined {
    if (raw == null) return undefined;
    const s = String(raw).trim();
    if (!s) return undefined;
    const m = s.match(/^(\d{1,2}):(\d{2})/);
    if (!m) return undefined;
    const hh = m[1].padStart(2, "0");
    const mm = m[2];
    return `${hh}:${mm}`;
  }

  /* ---------- Import CSV (Aules + Matriculats) ---------- */
function handleImportRoomsCSV(ev: React.ChangeEvent<HTMLInputElement>) {
  const file = ev.target.files?.[0];
  if (!file) return;

  // Normalitza dates a "yyyy-MM-dd"
  const normDate = (raw: any): string | undefined => {
    if (raw == null) return undefined;
    const s = String(raw).trim();
    if (!s) return undefined;

    // Ja és ISO
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

    // dd/mm/yyyy o dd-mm-yyyy
    const m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})$/);
    if (m) return `${m[3]}-${m[2]}-${m[1]}`;

    // Número d'Excel (43831, etc.)
    const n = Number(s);
    if (!Number.isNaN(n) && n > 40000 && n < 70000) {
      const excelEpoch = new Date(1899, 11, 30); // 1899-12-30
      const d = new Date(excelEpoch.getTime() + n * 86400000);
      const yy = d.getFullYear();
      const mm = String(d.getMonth() + 1).padStart(2, "0");
      const dd = String(d.getDate()).padStart(2, "0");
      return `${yy}-${mm}-${dd}`;
    }

    return undefined;
  };

  // Normalitza hores a "HH:mm"
  const normTime = (raw: any): string | undefined => {
    if (raw == null) return undefined;
    let s = String(raw).trim();
    if (!s) return undefined;

    // Acceptem "14:45", "14:45:00", "14.45", "14-45"
    s = s.replace(".", ":").replace("-", ":");
    const m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (!m) return undefined;
    const hh = m[1].padStart(2, "0");
    const mm = m[2];
    return `${hh}:${mm}`;
  };

  const findSubjectByCodeOrSigles = (codi: string, sigles: string): Subject | undefined => {
    const c = codi.trim();
    const sgl = sigles.trim();
    return subjects.find(
      (s) => (c && s.codi === c) || (sgl && s.sigles === sgl)
    );
  };

  Papa.parse(file, {
    header: true,
    skipEmptyLines: "greedy",
    complete: (res: Papa.ParseResult<any>) => {
      try {
        const rows = (res.data as any[]).filter(Boolean);

        // Còpia profunda segura de roomsData (pid → cellKey → subjectId → {rooms,students})
        const nextRooms: RoomsDataPerPeriod =
          typeof structuredClone === "function"
            ? structuredClone(roomsData || {})
            : JSON.parse(JSON.stringify(roomsData || {}));

        let attached = 0;
        let skipped = 0;

        for (const r of rows) {
          // --- Codi / sigles ---
          const codi = String(
            r.codi ?? r.CODI ?? r.codigo ?? r.CODIGO ?? r.code ?? ""
          ).trim();
          const sigles = String(
            r.sigles ??
              r.SIGLES ??
              r.siglas ??
              r.SIGLAS ??
              r.nom ??
              r.NOM ??
              ""
          ).trim();
          if (!codi && !sigles) {
            skipped++;
            continue;
          }

          const subj = findSubjectByCodeOrSigles(codi, sigles);
          if (!subj) {
            skipped++;
            continue;
          }

          // --- Període ---
          const pidRaw =
            r.period_id ??
            r.PERIOD_ID ??
            r.PeriodId ??
            r.periode ??
            r.PERIODE ??
            r.PERIode ??
            r.PERIODO ??
            r.PERIOD ??
            r.Period;
          const pid = pidRaw ? Number(pidRaw) : NaN;
          if (!Number.isFinite(pid) || !periods.some((p) => p.id === pid)) {
            skipped++;
            continue;
          }

          // --- Data d'examen ---
          const dateIso = normDate(
            r["dia d'examen"] ??
              r["dia examen"] ??
              r.data_examen ??
              r.dia ??
              r.DIA ??
              r.fecha ??
              r.FECHA ??
              r.data ??
              r.DATA ??
              r.day
          );
          if (!dateIso) {
            skipped++;
            continue;
          }

          // --- Franja horària ---
          const start = normTime(
            r["hora d'inici de l'examen"] ??
              r["hora inici examen"] ??
              r.hora_inici_examen ??
              r.hora_inici ??
              r.inici ??
              r.start ??
              r.HORA_INICI ??
              r.HORA_INI
          );
          const end = normTime(
            r["hora de fi de l'examen"] ??
              r["hora fi examen"] ??
              r.hora_fi_examen ??
              r.hora_fi ??
              r.fi ??
              r.end ??
              r.HORA_FI
          );
          if (!start || !end) {
            skipped++;
            continue;
          }

          const slots = slotsPerPeriod[pid] ?? [];
          if (!slots.length) {
            skipped++;
            continue;
          }
          const slotIndex = slots.findIndex(
            (s) => s.start === start && s.end === end
          );
          if (slotIndex === -1) {
            skipped++;
            continue;
          }

          // --- Aula ---
          const aula = String(
            r.aula ?? r.AULA ?? r.sala ?? r.SALA ?? r.room ?? r.ROOM ?? ""
          ).trim();
          if (!aula) {
            skipped++;
            continue;
          }

          // --- Estudiants ---
          const nStudRaw =
            r["número d'estudiants matriculats"] ??
            r["num_estudiants"] ??
            r.estudiants ??
            r.ESTUDIANTS ??
            r.matriculats ??
            r.MATRICULATS ??
            r.matriculados ??
            r.MATRICULADOS ??
            r.students ??
            r.STUDENTS ??
            r.ENROLLED ??
            r.enrolled;
          const nStudents =
            nStudRaw != null && String(nStudRaw).trim() !== ""
              ? Number(String(nStudRaw).replace(/[^\d]/g, ""))
              : undefined;

          // --- Escriure a nextRooms[pid][dateIso|slotIndex][subjectId] ---
          if (!nextRooms[pid]) nextRooms[pid] = {};
          const key = cellKey(dateIso, slotIndex);
          if (!nextRooms[pid][key]) nextRooms[pid][key] = {};
          if (!nextRooms[pid][key][subj.id]) {
            nextRooms[pid][key][subj.id] = { rooms: [] };
          }

          const entry = nextRooms[pid][key][subj.id];
          if (aula && !entry.rooms.includes(aula)) {
            entry.rooms.push(aula);
          }
          if (
            typeof nStudents === "number" &&
            Number.isFinite(nStudents) &&
            entry.students == null
          ) {
            entry.students = nStudents;
          }

          attached++;
        }

        setRoomsData(nextRooms);
        alert(
          `Aules/Matrículats processats. Afegits: ${attached}. Omesos: ${skipped}.`
        );
      } catch (err) {
        console.error(err);
        alert("Error processant el CSV d’aules/matriculats");
      } finally {
        ev.target.value = "";
      }
    },
    error: () => {
      alert("No s'ha pogut llegir el fitxer CSV d’aules/matriculats");
      ev.target.value = "";
    },
  });
}

  /* ---------- Render ---------- */
  return (
    <div className="p-6 max-w-[1200px] mx-auto">
      <h1 className="text-2xl font-bold mb-2">
        Planificador d'exàmens — períodes amb curs/quadrimestre
      </h1>
      <p className="text-sm mb-6">
        CSV esperat (assignatures/períodes):{" "}
        <code>
          codi,sigles,nivell,curs,quadrimestre,period_id,period_tipus,period_inici,period_fi,period_slots,period_blackouts
        </code>
        . Opcional: <code>MET,MATT,MEE,MCYBERS</code>.
      </p>

      {/* Intercanvi de dades */}
      <div className="p-4 rounded-2xl border shadow-sm bg-white mb-6">
        <h2 className="font-semibold mb-3">Dades i intercanvi</h2>
        <div className="flex flex-wrap gap-3 items-center">
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar CSV (REEMPLAÇA)
            <input
              type="file"
              accept=".csv,text/csv"
              className="hidden"
              onChange={handleImportCSV}
            />
          </label>

          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Afegir assignatures (CSV) — MERGE
            <input
              type="file"
              accept=".csv,text/csv"
              className="hidden"
              onChange={handleMergeSubjectsCSV}
            />
          </label>

          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar Aules/Matriculats (CSV)
            <input
              type="file"
              accept=".csv,text/csv"
              className="hidden"
              onChange={handleImportRoomsCSV}
            />
          </label>

          <button
            onClick={exportCSV}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Exportar CSV
          </button>
          <button
            onClick={exportTXT}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Exportar TXT
          </button>
          <button
            onClick={exportExcel}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Exportar Excel
          </button>
          <button
            onClick={exportWord}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Exportar calendari en Word
          </button>
          <button
            onClick={exportJSON}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Exportar JSON
          </button>
          <label className="px-3 py-2 border rounded-xl shadow-sm cursor-pointer bg-white">
            Importar JSON
            <input
              type="file"
              accept="application/json"
              className="hidden"
              onChange={importJSON}
            />
          </label>

          <button
            onClick={saveStateToUrl}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Guardar a l’enllaç
          </button>
          <button
            onClick={() => {
              if (!loadStateFromUrl())
                alert("No s’ha trobat cap estat a l’enllaç.");
            }}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Carregar de l’enllaç
          </button>
          <button
            onClick={copyLinkToClipboard}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Copiar enllaç
          </button>

          <span className="text-xs text-gray-500 ml-auto">
            Disponibles: {availableSubjects.length}/{subjects.length}
          </span>
        </div>
      </div>

      {/* Banner Desfer eliminació */}
      {lastDeleted && (
        <div className="p-3 rounded-xl border shadow-sm bg-amber-50 mb-4 text-sm flex items-center gap-3">
          <span>
            S’ha eliminat{" "}
            <strong>
              {lastDeleted.subject.sigles || lastDeleted.subject.codi}
            </strong>
            .
          </span>
          <button
            onClick={undoDelete}
            className="px-2 py-1 border rounded-md bg-white"
            title="Desfer"
          >
            Desfer
          </button>
          <button
            onClick={() => setLastDeleted(null)}
            className="px-2 py-1 border rounded-md bg-white ml-1"
            title="Descartar"
          >
            Descarta
          </button>
        </div>
      )}

      {/* Pestanyes (sense mostrar id) */}
      <div className="mb-4 flex items-center justify-between gap-2">
        <div className="flex flex-wrap gap-2">
          {periods.map((p) => (
            <button
              key={p.id}
              onClick={() => setActivePid(p.id)}
              className={`px-3 py-2 rounded-xl border shadow-sm ${
                p.id === activePid
                  ? "bg-indigo-50 border-indigo-300"
                  : "bg-white"
              }`}
              title="Canviar de període"
            >
              {p.tipus}
            </button>
          ))}
        </div>
        <div className="flex gap-2">
          <button
            onClick={addPeriod}
            className="px-3 py-2 border rounded-xl shadow-sm"
          >
            Afegir període
          </button>
          {periods.length > 1 && (
            <button
              onClick={() => removePeriod(activePid)}
              className="px-3 py-2 border rounded-xl shadow-sm"
            >
              Eliminar període actiu
            </button>
          )}
        </div>
      </div>

      {/* Configuració del període actiu (resum) */}
      {activePeriod && (
        <div className="grid md:grid-cols-3 gap-4 mb-6">
          <div className="p-4 rounded-2xl border shadow-sm bg-white">
            <h2 className="font-semibold mb-3">
              Configuració del període
            </h2>

            <label className="block text-sm mb-1">Tipus</label>
            <select
              value={activePeriod.tipus}
              onChange={(e) => {
                const v = e.target.value as TipusPeriode;
                setPeriods((arr) =>
                  arr.map((p) =>
                    p.id === activePid ? { ...p, tipus: v } : p
                  )
                );
              }}
              className="w-full border rounded-xl p-2"
            >
              <option>PARCIAL</option>
              <option>FINAL</option>
              <option>REAVALUACIÓ</option>
            </select>

            <label className="block text-sm mt-3 mb-1">
              Curs (any d’inici)
            </label>
            <input
              type="number"
              placeholder="Ex. 2025"
              value={activePeriod.curs ?? ""}
              onChange={(e) => {
                const n = e.target.value
                  ? Number(e.target.value)
                  : undefined;
                setPeriods((arr) =>
                  arr.map((p) =>
                    p.id === activePid ? { ...p, curs: n } : p
                  )
                );
              }}
              className="w-full border rounded-xl p-2"
            />

            <label className="block text-sm mt-3 mb-1">
              Quadrimestre del període
            </label>
            <select
              value={activePeriod.quad ?? 0}
              onChange={(e) => {
                const v = Number(e.target.value) as 0 | 1 | 2;
                setPeriods((arr) =>
                  arr.map((p) =>
                    p.id === activePid
                      ? {
                          ...p,
                          quad:
                            v === 1 || v === 2
                              ? (v as 1 | 2)
                              : undefined,
                        }
                      : p
                  )
                );
              }}
              className="w-full border rounded-xl p-2"
            >
              <option value={0}>(Sense)</option>
              <option value={1}>1</option>
              <option value={2}>2</option>
            </select>
          </div>

          {/* Franges horàries */}
          <div className="p-4 rounded-2xl border shadow-sm bg-white md:col-span-2">
            <h2 className="font-semibold mb-3">
              Franges horàries (per a aquest període)
            </h2>
            <div className="space-y-2">
              {(slotsPerPeriod[activePid] ?? []).map((s, i) => (
                <div key={i} className="flex items-center gap-2">
                  <span className="text-sm w-6">{i + 1}.</span>
                  <input
                    value={s.start}
                    onChange={(e) => {
                      const v = e.target.value;
                      setSlotsPerPeriod((sp) => {
                        const arr = [...(sp[activePid] ?? [])];
                        arr[i] = { ...arr[i], start: v };
                        return { ...sp, [activePid]: arr };
                      });
                    }}
                    className="border rounded-xl p-2 w-28"
                    placeholder="HH:mm"
                  />
                  <span>–</span>
                  <input
                    value={s.end}
                    onChange={(e) => {
                      const v = e.target.value;
                      setSlotsPerPeriod((sp) => {
                        const arr = [...(sp[activePid] ?? [])];
                        arr[i] = { ...arr[i], end: v };
                        return { ...sp, [activePid]: arr };
                      });
                    }}
                    className="border rounded-xl p-2 w-28"
                    placeholder="HH:mm"
                  />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* Safata + Calendari */}
      <DndContext
        onDragEnd={onDragEnd}
        modifiers={[restrictToWindowEdges]}
      >
        {/* Safata d'assignatures */}
        <div className="p-4 rounded-2xl border shadow-sm bg-white mb-3">
          <h2 className="font-semibold mb-3">
            Assignatures (arrossega)
          </h2>

          <div className="flex flex-wrap gap-2">
            {availableSubjects.map((s) => (
              <TrayChip key={s.id} id={s.id} s={s} />
            ))}

            {!availableSubjects.length && (
              <div className="text-xs text-gray-500 italic">
                No hi ha assignatures per al curs/quadrimestre i període
                d’aquest calendari, o ja estan totes
                programades/ocultes.
              </div>
            )}
          </div>
        </div>

        {/* Llista d'eliminades (amb restauració) */}
        {hiddenSubjectIds.length > 0 && (
          <div className="p-3 rounded-xl border shadow-sm bg-yellow-50 mb-6 text-sm">
            <div className="font-semibold mb-2">
              Assignatures eliminades de la safata
            </div>
            <div className="flex flex-wrap gap-2">
              {hiddenSubjectIds.map((id) => {
                const s = subjects.find((x) => x.id === id);
                if (!s) return null;
                return (
                  <span
                    key={id}
                    className="inline-flex items-center gap-2 px-2 py-1 rounded-lg border bg-white"
                  >
                    {s.sigles || s.codi}
                    <button
                      className="text-xs px-2 py-0.5 border rounded-md"
                      onClick={() =>
                        setHiddenSubjectIds((prev) =>
                          prev.filter((x) => x !== id)
                        )
                      }
                      title="Restaurar a la safata"
                    >
                      Restaurar
                    </button>
                  </span>
                );
              })}
              <button
                className="ml-2 text-xs px-2 py-0.5 border rounded-md bg-white"
                onClick={() => setHiddenSubjectIds([])}
                title="Restaurar totes"
              >
                Restaurar totes
              </button>
            </div>
          </div>
        )}

        {/* Calendari del període actiu */}
        {activePeriod && (
          <div className="mb-8">
            <div className="flex items-center gap-3 mb-2">
              <h3 className="text-lg font-semibold">
                {activePeriod.tipus} —{" "}
                {format(
                  parseISO(activePeriod.startStr),
                  "dd/MM"
                )}{" "}
                a{" "}
                {format(
                  parseISO(activePeriod.endStr),
                  "dd/MM"
                )}
              </h3>
              <span className="text-sm text-gray-500">(dl–dv)</span>
            </div>

            {[...eachWeek(
              mondayOfWeek(parseISO(activePeriod.startStr)),
              fridayOfWeek(parseISO(activePeriod.endStr))
            )].map(({ mon, fri }, wIdx) => (
              <div key={wIdx} className="mt-6">
                <div className="flex items-center gap-3 mb-2">
                  <h4 className="font-semibold">
                    Setmana {format(mon, "dd/MM")} —{" "}
                    {format(fri, "dd/MM")}
                  </h4>
                  <span className="text-xs text-gray-500">(dl–dv)</span>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse text-sm">
                    <thead>
                      <tr>
                        <th className="border p-2 w-[160px] text-left">
                          franja horària/Time slot
                        </th>
                        {Array.from({ length: 5 }).map((_, i) => (
                          <th
                            key={i}
                            className="border p-2 min-w-[170px] text-left"
                          >
                            <div className="font-semibold">
                              {
                                [
                                  "Dl/Mon",
                                  "Dt/Tu",
                                  "Dc/Wed",
                                  "Dj/Thu",
                                  "Dv/Fri",
                                ][i]
                              }
                            </div>
                            <div className="text-xs text-gray-500">
                              {fmtDM(addDays(mon, i))}
                            </div>
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(slotsPerPeriod[activePid] ?? []).map(
                        (s, slotIndex) => (
                          <tr key={slotIndex}>
                            <td className="border p-2 align-top font-medium whitespace-nowrap">
                              {s.start}-{s.end}
                            </td>
                            {Array.from({ length: 5 }).map((_, i) => {
                              const day = addDays(mon, i);
                              const dateIso = iso(day);
                              const disabled = isDisabledDay(
                                day,
                                activePeriod
                              );
                              const amap =
                                assignedPerPeriod[activePid] ?? {};
                              const subjIds =
                                amap[
                                  cellKey(dateIso, slotIndex)
                                ] ?? [];

                              const assignedList = subjIds
                                .map((id) =>
                                  subjects.find(
                                    (x) => x.id === id
                                  )
                                )
                                .filter(Boolean) as Subject[];

                              const extrasForSubjects: Record<
                                string,
                                RoomsEnroll
                              > = {};
                              const extrasCell =
                                roomsData?.[activePid]?.[
                                  cellKey(
                                    dateIso,
                                    slotIndex
                                  )
                                ] ?? {};
                              for (const sid of subjIds) {
                                if (extrasCell[sid])
                                  extrasForSubjects[sid] =
                                    extrasCell[sid];
                              }

                              return (
                                <DropCell
                                  key={i}
                                  id={`cell:${activePid}:${dateIso}:${slotIndex}`}
                                  disabled={disabled}
                                  assignedList={assignedList}
                                  extrasForSubjects={
                                    extrasForSubjects
                                  }
                                  onRemoveOne={(subjectId) =>
                                    removeOneFromCell(
                                      activePid,
                                      dateIso,
                                      slotIndex,
                                      subjectId
                                    )
                                  }
                                  pid={activePid}
                                  dateIso={dateIso}
                                  slotIndex={slotIndex}
                                />
                              );
                            })}
                          </tr>
                        )
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            ))}
          </div>
        )}

        <TrashBin />
      </DndContext>
    </div>
  );
}
