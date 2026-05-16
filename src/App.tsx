import {
  closestCenter,
  DndContext,
  DragOverlay,
  PointerSensor,
  useDraggable,
  useDroppable,
  useSensor,
  useSensors,
  type DragEndEvent,
  type DragStartEvent,
} from "@dnd-kit/core";
import { readSheet, type Row } from "read-excel-file/browser";
import {
  ChevronLeft,
  Circle,
  Download,
  Minus,
  Move,
  PanelTop,
  Pencil,
  Plus,
  Printer,
  RectangleHorizontal,
  Search,
  Square,
  Trash2,
  Upload,
  UserPlus,
  UserRound,
  UsersRound,
  X,
} from "lucide-react";
import { useEffect, useMemo, useReducer, useRef, useState, type CSSProperties } from "react";

type TableShape = "round" | "rectangle" | "square" | "head";
type GuestFilter = "all" | "unseated" | "seated";
type PlannerView = "sketch" | "table";

type Guest = {
  id: string;
  name: string;
  group: string;
};

type Seat = {
  id: string;
  guestId: string | null;
};

type PlannerTable = {
  id: string;
  name: string;
  shape: TableShape;
  scale: number;
  x: number;
  y: number;
  seats: Seat[];
};

type PlannerState = {
  guests: Guest[];
  tables: PlannerTable[];
  selectedTableId: string | null;
};

type PlannerAction =
  | { type: "add_guest"; guest: Guest }
  | { type: "add_guests"; guests: Guest[] }
  | { type: "import_state"; state: PlannerState }
  | { type: "remove_guest"; guestId: string }
  | { type: "update_guest"; guestId: string; name: string }
  | { type: "add_table"; table: PlannerTable }
  | { type: "delete_table"; tableId: string }
  | { type: "select_table"; tableId: string | null }
  | { type: "move_table"; tableId: string; x: number; y: number }
  | { type: "update_table"; table: PlannerTable }
  | { type: "assign_guest"; guestId: string; tableId: string; seatId: string }
  | { type: "unassign_guest"; guestId: string }
  | { type: "reset_demo" };

type SeatPosition = {
  x: number;
  y: number;
};

type TableLayout = {
  width: number;
  height: number;
  body: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  seats: SeatPosition[];
};

type TableDrag = {
  tableId: string;
  startX: number;
  startY: number;
  originX: number;
  originY: number;
  width: number;
  height: number;
};

type RoomConfigFile = {
  app: "wedding-planner";
  version: 1;
  exportedAt: string;
  canvas: {
    height: number;
    width: number;
    zoom: number;
  };
  state: PlannerState;
};

type GuestContextMenuState = {
  guestId: string;
  x: number;
  y: number;
};

const STORAGE_KEY = "wedding-planner-state-v1";
const CANVAS_WIDTH = 2100;
const CANVAS_HEIGHT = 1400;
const GRID_SIZE = 20;
const SEAT_SIZE = 48;
const TABLE_SCALE_MIN = 0.65;
const TABLE_SCALE_MAX = 2.2;
const TABLE_SCALE_STEP = 0.1;
const CANVAS_ZOOM_MIN = 0.35;
const CANVAS_ZOOM_MAX = 2.5;
const CANVAS_ZOOM_STEP = 0.1;

const tableShapeLabels: Record<TableShape, string> = {
  round: "Masa rotunda",
  rectangle: "Masa lunga",
  square: "Masa patrata",
  head: "Prezidiu",
};

const starterGuests: Guest[] = [
  { id: "guest-nasa", name: "Nasa", group: "Familie" },
  { id: "guest-mireasa", name: "Mireasa", group: "Prezidiu" },
  { id: "guest-mire", name: "Mire", group: "Prezidiu" },
  { id: "guest-nasu", name: "Nasu", group: "Familie" },
];

function makeId(prefix: string) {
  const value =
    typeof globalThis.crypto?.randomUUID === "function"
      ? globalThis.crypto.randomUUID()
      : Math.random().toString(36).slice(2);
  return `${prefix}-${value}`;
}

function makeSeats(count: number, existing: Seat[] = []) {
  return Array.from({ length: count }, (_, index) => {
    return existing[index] ?? { id: makeId("seat"), guestId: null };
  });
}

function normalizeTableScale(value: unknown) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 1;
  }

  return Math.max(TABLE_SCALE_MIN, Math.min(TABLE_SCALE_MAX, Math.round(numeric * 10) / 10));
}

function normalizeCanvasZoom(value: unknown) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 1;
  }

  return Math.max(CANVAS_ZOOM_MIN, Math.min(CANVAS_ZOOM_MAX, Math.round(numeric * 100) / 100));
}

function createInitialState(): PlannerState {
  return {
    guests: starterGuests,
    selectedTableId: "table-main",
    tables: [
      {
        id: "table-head",
        name: "prezidiu",
        shape: "head",
        scale: 1,
        x: 720,
        y: 105,
        seats: [
          { id: "seat-head-1", guestId: "guest-nasa" },
          { id: "seat-head-2", guestId: "guest-mireasa" },
          { id: "seat-head-3", guestId: "guest-mire" },
          { id: "seat-head-4", guestId: "guest-nasu" },
        ],
      },
    ],
  };
}

function loadState() {
  if (typeof window === "undefined") {
    return createInitialState();
  }

  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return createInitialState();
    }

    const parsed = JSON.parse(raw) as PlannerState;
    if (!Array.isArray(parsed.guests) || !Array.isArray(parsed.tables)) {
      return createInitialState();
    }

    return {
      guests: parsed.guests,
      tables: parsed.tables.map((table) => ({
        ...table,
        scale: normalizeTableScale(table.scale),
      })),
      selectedTableId: parsed.selectedTableId ?? null,
    };
  } catch {
    return createInitialState();
  }
}

function plannerReducer(state: PlannerState, action: PlannerAction): PlannerState {
  switch (action.type) {
    case "add_guest":
      return {
        ...state,
        guests: [...state.guests, action.guest],
      };
    case "add_guests":
      return {
        ...state,
        guests: [...state.guests, ...action.guests],
      };
    case "import_state":
      return action.state;
    case "remove_guest":
      return {
        ...state,
        guests: state.guests.filter((guest) => guest.id !== action.guestId),
        tables: state.tables.map((table) => ({
          ...table,
          seats: table.seats.map((seat) =>
            seat.guestId === action.guestId ? { ...seat, guestId: null } : seat,
          ),
        })),
      };
    case "update_guest":
      return {
        ...state,
        guests: state.guests.map((guest) =>
          guest.id === action.guestId ? { ...guest, name: action.name } : guest,
        ),
      };
    case "add_table":
      return {
        ...state,
        selectedTableId: action.table.id,
        tables: [...state.tables, action.table],
      };
    case "delete_table":
      return {
        ...state,
        selectedTableId: state.selectedTableId === action.tableId ? null : state.selectedTableId,
        tables: state.tables.filter((table) => table.id !== action.tableId),
      };
    case "select_table":
      return {
        ...state,
        selectedTableId: action.tableId,
      };
    case "move_table":
      return {
        ...state,
        tables: state.tables.map((table) =>
          table.id === action.tableId ? { ...table, x: action.x, y: action.y } : table,
        ),
      };
    case "update_table":
      return {
        ...state,
        tables: state.tables.map((table) =>
          table.id === action.table.id ? action.table : table,
        ),
      };
    case "assign_guest":
      return assignGuestToSeat(state, action.guestId, action.tableId, action.seatId);
    case "unassign_guest":
      return {
        ...state,
        tables: state.tables.map((table) => ({
          ...table,
          seats: table.seats.map((seat) =>
            seat.guestId === action.guestId ? { ...seat, guestId: null } : seat,
          ),
        })),
      };
    case "reset_demo":
      return createInitialState();
    default:
      return state;
  }
}

function assignGuestToSeat(
  state: PlannerState,
  guestId: string,
  targetTableId: string,
  targetSeatId: string,
) {
  let source:
    | {
        tableId: string;
        seatId: string;
      }
    | null = null;
  let replacedGuestId: string | null = null;

  state.tables.forEach((table) => {
    table.seats.forEach((seat) => {
      if (seat.guestId === guestId) {
        source = { tableId: table.id, seatId: seat.id };
      }
      if (table.id === targetTableId && seat.id === targetSeatId) {
        replacedGuestId = seat.guestId;
      }
    });
  });

  const nextTables = state.tables.map((table) => {
    return {
      ...table,
      seats: table.seats.map((seat) => {
        if (seat.guestId === guestId) {
          return { ...seat, guestId: null };
        }

        if (table.id === targetTableId && seat.id === targetSeatId) {
          return { ...seat, guestId };
        }

        return seat;
      }),
    };
  });

  if (replacedGuestId && replacedGuestId !== guestId && source) {
    return {
      ...state,
      tables: nextTables.map((table) => ({
        ...table,
        seats: table.seats.map((seat) =>
          table.id === source?.tableId && seat.id === source.seatId
            ? { ...seat, guestId: replacedGuestId }
            : seat,
        ),
      })),
    };
  }

  return {
    ...state,
    tables: nextTables,
  };
}

function createTable(shape: TableShape, index: number): PlannerTable {
  const isHead = shape === "head";
  const step = index % 4;
  return {
    id: makeId("table"),
    name: isHead ? "prezidiu" : `masa ${index + 1}`,
    shape,
    scale: 1,
    x: 360 + step * 90,
    y: 320 + step * 70,
    seats: makeSeats(isHead ? 4 : shape === "rectangle" ? 8 : 10),
  };
}

function getAssignedGuestIds(tables: PlannerTable[]) {
  const ids = new Set<string>();
  tables.forEach((table) => {
    table.seats.forEach((seat) => {
      if (seat.guestId) {
        ids.add(seat.guestId);
      }
    });
  });
  return ids;
}

function getTableLayout(table: PlannerTable): TableLayout {
  const scale = normalizeTableScale(table.scale);

  if (table.shape === "head") {
    const baseBodyWidth = Math.max(320, table.seats.length * 78 + 20);
    const body = {
      x: 30,
      y: 66,
      width: Math.max(table.seats.length * 76, baseBodyWidth * scale),
      height: Math.max(42, 58 * scale),
    };
    return {
      width: body.width + body.x * 2,
      height: body.y + body.height + 26,
      body,
      seats: topOnlyPositions(table.seats.length, body),
    };
  }

  if (table.shape === "rectangle") {
    const body = { x: 72, y: 88, width: 232 * scale, height: 86 * scale };
    return {
      width: body.width + body.x * 2,
      height: body.height + body.y * 2,
      body,
      seats: rectanglePositions(table.seats.length, body),
    };
  }

  if (table.shape === "square") {
    const body = { x: 82, y: 82, width: 136 * scale, height: 136 * scale };
    return {
      width: body.width + body.x * 2,
      height: body.height + body.y * 2,
      body,
      seats: rectanglePositions(table.seats.length, body),
    };
  }

  const bodyDiameter = 128 * scale;
  const bodyMargin = 86;
  const width = bodyDiameter + bodyMargin * 2;
  const center = { x: width / 2, y: width / 2 };
  const orbitRadius = bodyDiameter / 2 + SEAT_SIZE / 2 + 16;
  return {
    width,
    height: width,
    body: { x: bodyMargin, y: bodyMargin, width: bodyDiameter, height: bodyDiameter },
    seats: table.seats.map((_, index) => {
      const angle = -Math.PI / 2 + (Math.PI * 2 * index) / table.seats.length;
      return {
        x: center.x + Math.cos(angle) * orbitRadius - SEAT_SIZE / 2,
        y: center.y + Math.sin(angle) * orbitRadius - SEAT_SIZE / 2,
      };
    }),
  };
}

function topOnlyPositions(count: number, body: TableLayout["body"]) {
  return Array.from({ length: count }, (_, index) => ({
    x: body.x + ((index + 1) / (count + 1)) * body.width - SEAT_SIZE / 2,
    y: 12,
  }));
}

function rectanglePositions(count: number, body: TableLayout["body"]) {
  const sideCounts = [0, 0, 0, 0];
  Array.from({ length: count }).forEach((_, index) => {
    sideCounts[index % 4] += 1;
  });

  const positions: SeatPosition[] = [];
  const pushSide = (side: number, amount: number) => {
    Array.from({ length: amount }).forEach((_, index) => {
      const offset = (index + 1) / (amount + 1);
      if (side === 0) {
        positions.push({
          x: body.x + body.width * offset - SEAT_SIZE / 2,
          y: body.y - SEAT_SIZE - 12,
        });
      }
      if (side === 1) {
        positions.push({
          x: body.x + body.width + 12,
          y: body.y + body.height * offset - SEAT_SIZE / 2,
        });
      }
      if (side === 2) {
        positions.push({
          x: body.x + body.width * (1 - offset) - SEAT_SIZE / 2,
          y: body.y + body.height + 12,
        });
      }
      if (side === 3) {
        positions.push({
          x: body.x - SEAT_SIZE - 12,
          y: body.y + body.height * (1 - offset) - SEAT_SIZE / 2,
        });
      }
    });
  };

  sideCounts.forEach((amount, side) => pushSide(side, amount));
  return positions;
}

function seatDropId(tableId: string, seatId: string) {
  return `seat:${tableId}:${seatId}`;
}

function parseSeatDropId(id: string) {
  const [, tableId, seatId] = id.split(":");
  if (!tableId || !seatId) {
    return null;
  }
  return { tableId, seatId };
}

function guestDragId(guestId: string) {
  return `guest:${guestId}`;
}

function parseGuestDragId(id: string) {
  return id.startsWith("guest:") ? id.slice("guest:".length) : null;
}

function initials(name: string) {
  return name
    .split(" ")
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part[0])
    .join("")
    .toUpperCase();
}

function cellToText(value: unknown) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  return String(value).trim();
}

function normalizeHeader(value: unknown) {
  return cellToText(value).toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findColumn(headers: Row, candidates: string[], fallback: number) {
  const normalizedCandidates = new Set(candidates.map((candidate) => normalizeHeader(candidate)));
  const index = headers.findIndex((header) => normalizedCandidates.has(normalizeHeader(header)));
  return index >= 0 ? index : fallback;
}

function parseGuestRows(rows: Row[]) {
  const headerIndex = rows.findIndex((row) => row.some((cell) => normalizeHeader(cell) === "firstname"));
  const headers = rows[headerIndex] ?? [];
  const startIndex = headerIndex >= 0 ? headerIndex + 1 : 0;
  const firstNameColumn = findColumn(headers, ["First Name", "Firstname", "Name"], 0);
  const lastNameColumn = findColumn(headers, ["Last Name", "Lastname"], 1);
  const addressColumn = findColumn(headers, ["Address"], 4);

  return rows
    .slice(startIndex)
    .map((row) => {
      const firstName = cellToText(row[firstNameColumn]);
      const lastName = cellToText(row[lastNameColumn]);
      const fallbackName = row.map(cellToText).find(Boolean) ?? "";
      const name = [firstName, lastName].filter(Boolean).join(" ").trim() || fallbackName;
      const group = cellToText(row[addressColumn]) || "Import Excel";
      return name ? { id: makeId("guest"), name, group } : null;
    })
    .filter((guest): guest is Guest => Boolean(guest));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isTableShape(value: unknown): value is TableShape {
  return value === "round" || value === "rectangle" || value === "square" || value === "head";
}

function finiteNumber(value: unknown, fallback: number) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function readPlannerStateFromConfig(value: unknown): PlannerState | null {
  const root = isRecord(value) && isRecord(value.state) ? value.state : value;
  if (!isRecord(root) || !Array.isArray(root.guests) || !Array.isArray(root.tables)) {
    return null;
  }

  const guests = root.guests
    .filter(isRecord)
    .map((guest, index) => ({
      id: typeof guest.id === "string" && guest.id ? guest.id : makeId(`guest-import-${index}`),
      name: typeof guest.name === "string" && guest.name.trim() ? guest.name.trim() : `Invitat ${index + 1}`,
      group: typeof guest.group === "string" && guest.group.trim() ? guest.group.trim() : "Import config",
    }));
  const guestIds = new Set(guests.map((guest) => guest.id));

  const tables = root.tables
    .filter(isRecord)
    .map((table, tableIndex) => {
      const rawSeats = Array.isArray(table.seats) ? table.seats : [];
      const seats = rawSeats.filter(isRecord).map((seat, seatIndex) => {
        const guestId = typeof seat.guestId === "string" && guestIds.has(seat.guestId) ? seat.guestId : null;
        return {
          id: typeof seat.id === "string" && seat.id ? seat.id : makeId(`seat-import-${tableIndex}-${seatIndex}`),
          guestId,
        };
      });

      return {
        id: typeof table.id === "string" && table.id ? table.id : makeId(`table-import-${tableIndex}`),
        name: typeof table.name === "string" && table.name.trim() ? table.name.trim() : `masa ${tableIndex + 1}`,
        shape: isTableShape(table.shape) ? table.shape : "round",
        scale: normalizeTableScale(table.scale),
        x: finiteNumber(table.x, 320 + tableIndex * 40),
        y: finiteNumber(table.y, 260 + tableIndex * 40),
        seats,
      };
    });

  const tableIds = new Set(tables.map((table) => table.id));
  const selectedTableId =
    typeof root.selectedTableId === "string" && tableIds.has(root.selectedTableId)
      ? root.selectedTableId
      : null;

  return { guests, tables, selectedTableId };
}

function readCanvasZoomFromConfig(value: unknown) {
  if (isRecord(value)) {
    if (isRecord(value.canvas) && value.canvas.zoom !== undefined) {
      return normalizeCanvasZoom(value.canvas.zoom);
    }

    if (value.canvasZoom !== undefined) {
      return normalizeCanvasZoom(value.canvasZoom);
    }
  }

  return 1;
}

function App() {
  const [state, dispatch] = useReducer(plannerReducer, undefined, loadState);
  const [view, setView] = useState<PlannerView>("sketch");
  const [query, setQuery] = useState("");
  const [filter, setFilter] = useState<GuestFilter>("all");
  const [guestName, setGuestName] = useState("");
  const [guestGroup, setGuestGroup] = useState("Invitati");
  const [importMessage, setImportMessage] = useState("");
  const [configMessage, setConfigMessage] = useState("");
  const [canvasZoom, setCanvasZoom] = useState(1);
  const [activeGuestId, setActiveGuestId] = useState<string | null>(null);
  const [guestMenu, setGuestMenu] = useState<GuestContextMenuState | null>(null);
  const [guestMenuName, setGuestMenuName] = useState("");
  const [tableDrag, setTableDrag] = useState<TableDrag | null>(null);

  const assignedGuestIds = useMemo(() => getAssignedGuestIds(state.tables), [state.tables]);
  const guestsById = useMemo(() => {
    return new Map(state.guests.map((guest) => [guest.id, guest]));
  }, [state.guests]);
  const selectedTable = state.tables.find((table) => table.id === state.selectedTableId) ?? null;
  const activeGuest = activeGuestId ? guestsById.get(activeGuestId) ?? null : null;
  const menuGuest = guestMenu ? guestsById.get(guestMenu.guestId) ?? null : null;
  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: { distance: 6 },
    }),
  );

  const visibleGuests = useMemo(() => {
    const normalizedQuery = query.trim().toLowerCase();
    return state.guests.filter((guest) => {
      const matchesQuery =
        !normalizedQuery ||
        guest.name.toLowerCase().includes(normalizedQuery) ||
        guest.group.toLowerCase().includes(normalizedQuery);
      const seated = assignedGuestIds.has(guest.id);
      const matchesFilter =
        filter === "all" || (filter === "seated" && seated) || (filter === "unseated" && !seated);
      return matchesQuery && matchesFilter;
    });
  }, [assignedGuestIds, filter, query, state.guests]);

  useEffect(() => {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }, [state]);

  useEffect(() => {
    if (!tableDrag) {
      return undefined;
    }

    const handlePointerMove = (event: PointerEvent) => {
      const nextX = tableDrag.originX + (event.clientX - tableDrag.startX) / canvasZoom;
      const nextY = tableDrag.originY + (event.clientY - tableDrag.startY) / canvasZoom;
      dispatch({
        type: "move_table",
        tableId: tableDrag.tableId,
        x: Math.max(
          0,
          Math.min(CANVAS_WIDTH - tableDrag.width, Math.round(nextX / GRID_SIZE) * GRID_SIZE),
        ),
        y: Math.max(
          0,
          Math.min(CANVAS_HEIGHT - tableDrag.height, Math.round(nextY / GRID_SIZE) * GRID_SIZE),
        ),
      });
    };

    const handlePointerUp = () => setTableDrag(null);

    window.addEventListener("pointermove", handlePointerMove);
    window.addEventListener("pointerup", handlePointerUp, { once: true });
    return () => {
      window.removeEventListener("pointermove", handlePointerMove);
      window.removeEventListener("pointerup", handlePointerUp);
    };
  }, [canvasZoom, tableDrag]);

  function setBoardZoom(value: number) {
    setCanvasZoom(normalizeCanvasZoom(value));
  }

  function handleAddGuest(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const trimmedName = guestName.trim();
    if (!trimmedName) {
      return;
    }
    dispatch({
      type: "add_guest",
      guest: {
        id: makeId("guest"),
        name: trimmedName,
        group: guestGroup.trim() || "Invitati",
      },
    });
    setGuestName("");
  }

  function handleAddTable(shape: TableShape) {
    dispatch({
      type: "add_table",
      table: createTable(shape, state.tables.length),
    });
  }

  async function handleImportGuests(file: File) {
    setImportMessage("Se importa...");

    try {
      const rows = await readSheet(file);
      const guests = parseGuestRows(rows);

      if (!guests.length) {
        setImportMessage("Nu am gasit invitati in fisier.");
        return;
      }

      dispatch({ type: "add_guests", guests });
      setImportMessage(`${guests.length} invitati importati`);
    } catch {
      setImportMessage("Fisierul nu a putut fi importat.");
    }
  }

  function handleExportRoomConfig() {
    const config: RoomConfigFile = {
      app: "wedding-planner",
      version: 1,
      exportedAt: new Date().toISOString(),
      canvas: {
        width: CANVAS_WIDTH,
        height: CANVAS_HEIGHT,
        zoom: canvasZoom,
      },
      state,
    };
    const blob = new Blob([JSON.stringify(config, null, 2)], {
      type: "application/json",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const datePart = new Date().toISOString().slice(0, 10);
    link.href = url;
    link.download = `wedding-room-config-${datePart}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
    setConfigMessage("Config exportat");
  }

  async function handleImportRoomConfig(file: File) {
    setConfigMessage("Se importa config...");

    try {
      const parsed = JSON.parse(await file.text()) as unknown;
      const importedState = readPlannerStateFromConfig(parsed);
      if (!importedState) {
        setConfigMessage("Config invalid.");
        return;
      }

      dispatch({ type: "import_state", state: importedState });
      setBoardZoom(readCanvasZoomFromConfig(parsed));
      setView("sketch");
      setTableDrag(null);
      setActiveGuestId(null);
      setConfigMessage("Config importat");
    } catch {
      setConfigMessage("Config invalid.");
    }
  }

  function handleDragStart(event: DragStartEvent) {
    setGuestMenu(null);
    setActiveGuestId(parseGuestDragId(String(event.active.id)));
  }

  function handleDragEnd(event: DragEndEvent) {
    const guestId = parseGuestDragId(String(event.active.id));
    const overId = event.over ? String(event.over.id) : "";
    setActiveGuestId(null);

    if (!guestId || !overId) {
      return;
    }

    if (overId === "guest-list") {
      dispatch({ type: "unassign_guest", guestId });
      return;
    }

    if (overId.startsWith("seat:")) {
      const drop = parseSeatDropId(overId);
      if (drop) {
        dispatch({
          type: "assign_guest",
          guestId,
          tableId: drop.tableId,
          seatId: drop.seatId,
        });
      }
    }
  }

  function updateTable(table: PlannerTable) {
    dispatch({ type: "update_table", table });
  }

  function openGuestMenu(guestId: string, x: number, y: number) {
    const guest = guestsById.get(guestId);
    if (!guest) {
      return;
    }

    setGuestMenu({ guestId, x, y });
    setGuestMenuName(guest.name);
  }

  function saveGuestMenuName() {
    if (!guestMenu) {
      return;
    }

    const trimmedName = guestMenuName.trim();
    if (trimmedName) {
      dispatch({ type: "update_guest", guestId: guestMenu.guestId, name: trimmedName });
    }
    setGuestMenu(null);
  }

  return (
    <DndContext
      sensors={sensors}
      collisionDetection={closestCenter}
      onDragStart={handleDragStart}
      onDragEnd={handleDragEnd}
      onDragCancel={() => setActiveGuestId(null)}
    >
      <div className="app-shell">
        <Sidebar
          assignedGuestIds={assignedGuestIds}
          filter={filter}
          guestGroup={guestGroup}
          guestName={guestName}
          guests={visibleGuests}
          onAddGuest={handleAddGuest}
          onAddTable={handleAddTable}
          onFilterChange={setFilter}
          onGuestGroupChange={setGuestGroup}
          onGuestNameChange={setGuestName}
          onImportGuests={handleImportGuests}
          onExportRoomConfig={handleExportRoomConfig}
          onImportRoomConfig={handleImportRoomConfig}
          onRemoveGuest={(guestId) => dispatch({ type: "remove_guest", guestId })}
          onReset={() => dispatch({ type: "reset_demo" })}
          configMessage={configMessage}
          importMessage={importMessage}
          query={query}
          setQuery={setQuery}
          totals={{ all: state.guests.length, seated: assignedGuestIds.size }}
        />

        <main className="planner-main">
          <Topbar
            canvasZoom={canvasZoom}
            selectedTable={selectedTable}
            view={view}
            onBack={() => {
              dispatch({ type: "select_table", tableId: null });
              setView("sketch");
            }}
            onZoomChange={setBoardZoom}
            onPrint={() => window.print()}
            onViewChange={setView}
          />

          {view === "sketch" ? (
            <CanvasView
              guestsById={guestsById}
              onDeleteTable={(tableId) => dispatch({ type: "delete_table", tableId })}
              onSelectTable={(tableId) => dispatch({ type: "select_table", tableId })}
              onStartTableDrag={(drag) => setTableDrag(drag)}
              onGuestContextMenu={openGuestMenu}
              onZoomChange={setBoardZoom}
              zoom={canvasZoom}
              selectedTableId={state.selectedTableId}
              tables={state.tables}
            />
          ) : (
            <SeatingTableView assignedGuestIds={assignedGuestIds} guests={state.guests} tables={state.tables} />
          )}
        </main>

        {view === "sketch" && selectedTable ? (
          <TableInspector
            table={selectedTable}
            onClose={() => dispatch({ type: "select_table", tableId: null })}
            onDelete={() => dispatch({ type: "delete_table", tableId: selectedTable.id })}
            onUpdate={updateTable}
          />
        ) : null}

        {guestMenu && menuGuest ? (
          <GuestContextMenu
            guest={menuGuest}
            name={guestMenuName}
            x={guestMenu.x}
            y={guestMenu.y}
            onClose={() => setGuestMenu(null)}
            onNameChange={setGuestMenuName}
            onRemove={() => {
              dispatch({ type: "unassign_guest", guestId: guestMenu.guestId });
              setGuestMenu(null);
            }}
            onSave={saveGuestMenuName}
          />
        ) : null}
      </div>

      <DragOverlay>{activeGuest ? <GuestChip guest={activeGuest} floating /> : null}</DragOverlay>
    </DndContext>
  );
}

type SidebarProps = {
  assignedGuestIds: Set<string>;
  configMessage: string;
  filter: GuestFilter;
  guestGroup: string;
  guestName: string;
  guests: Guest[];
  importMessage: string;
  onAddGuest: (event: React.FormEvent<HTMLFormElement>) => void;
  onAddTable: (shape: TableShape) => void;
  onFilterChange: (filter: GuestFilter) => void;
  onExportRoomConfig: () => void;
  onGuestGroupChange: (group: string) => void;
  onGuestNameChange: (name: string) => void;
  onImportGuests: (file: File) => void;
  onImportRoomConfig: (file: File) => void;
  onRemoveGuest: (guestId: string) => void;
  onReset: () => void;
  query: string;
  setQuery: (query: string) => void;
  totals: {
    all: number;
    seated: number;
  };
};

function Sidebar({
  assignedGuestIds,
  configMessage,
  filter,
  guestGroup,
  guestName,
  guests,
  importMessage,
  onAddGuest,
  onAddTable,
  onFilterChange,
  onExportRoomConfig,
  onGuestGroupChange,
  onGuestNameChange,
  onImportGuests,
  onImportRoomConfig,
  onRemoveGuest,
  onReset,
  query,
  setQuery,
  totals,
}: SidebarProps) {
  const { isOver, setNodeRef } = useDroppable({ id: "guest-list" });

  return (
    <aside className="sidebar">
      <section className="table-tools">
        <div className="section-title">
          <h2>Adaugare Masa</h2>
        </div>
        <div className="shape-grid">
          <button className="shape-button" title={tableShapeLabels.rectangle} onClick={() => onAddTable("rectangle")}>
            <RectangleHorizontal size={30} />
            <Plus size={15} className="shape-plus" />
          </button>
          <button className="shape-button" title={tableShapeLabels.square} onClick={() => onAddTable("square")}>
            <Square size={30} />
            <Plus size={15} className="shape-plus" />
          </button>
          <button className="shape-button" title={tableShapeLabels.round} onClick={() => onAddTable("round")}>
            <Circle size={34} />
            <Plus size={15} className="shape-plus" />
          </button>
          <button className="shape-button" title={tableShapeLabels.head} onClick={() => onAddTable("head")}>
            <PanelTop size={31} />
            <Plus size={15} className="shape-plus" />
          </button>
        </div>
        <div className="config-tools">
          <button type="button" onClick={onExportRoomConfig}>
            <Download size={18} />
            Export room-configs
          </button>
          <label>
            <Upload size={18} />
            Import room-configs
            <input
              accept="application/json,.json"
              aria-label="Importa room-configs"
              type="file"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) {
                  onImportRoomConfig(file);
                }
                event.currentTarget.value = "";
              }}
            />
          </label>
          {configMessage ? <span>{configMessage}</span> : null}
        </div>
      </section>

      <section className="guest-tools">
        <div className="section-title guest-title">
          <h2>Lista Invitati</h2>
          <div className="filter-tabs" aria-label="Filtru invitati">
            <button className={filter === "all" ? "active" : ""} onClick={() => onFilterChange("all")}>
              Toti
            </button>
            <button className={filter === "unseated" ? "active" : ""} onClick={() => onFilterChange("unseated")}>
              Neasezati
            </button>
            <button className={filter === "seated" ? "active" : ""} onClick={() => onFilterChange("seated")}>
              Asezati
            </button>
          </div>
        </div>

        <form className="guest-form" onSubmit={onAddGuest}>
          <div className="form-row">
            <input
              aria-label="Nume invitat"
              placeholder="Nume invitat"
              value={guestName}
              onChange={(event) => onGuestNameChange(event.target.value)}
            />
            <input
              aria-label="Grup"
              placeholder="Grup"
              value={guestGroup}
              onChange={(event) => onGuestGroupChange(event.target.value)}
            />
          </div>
          <button type="submit">
            <UserPlus size={20} />
            Adauga Invitat
          </button>
        </form>

        <div className="import-area">
          <label className="import-button">
            <Upload size={19} />
            Import Excel
            <input
              accept=".xlsx,.xls"
              aria-label="Importa invitati din Excel"
              type="file"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) {
                  onImportGuests(file);
                }
                event.currentTarget.value = "";
              }}
            />
          </label>
          {importMessage ? <span>{importMessage}</span> : null}
        </div>

        <label className="search-box">
          <Search size={21} />
          <input
            aria-label="Cauta invitat"
            placeholder="Cauta invitat"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
          />
        </label>

        <div className="guest-stats">
          <span>{totals.all} invitati</span>
          <span>{totals.seated} asezati</span>
          <button onClick={onReset}>Demo</button>
        </div>

        <div ref={setNodeRef} className={`guest-list ${isOver ? "is-over" : ""}`}>
          {guests.map((guest) => (
            <GuestRow
              assigned={assignedGuestIds.has(guest.id)}
              guest={guest}
              key={guest.id}
              onRemove={() => onRemoveGuest(guest.id)}
            />
          ))}
          {guests.length === 0 ? <p className="empty-state">Niciun invitat</p> : null}
        </div>
      </section>
    </aside>
  );
}

type TopbarProps = {
  canvasZoom: number;
  selectedTable: PlannerTable | null;
  view: PlannerView;
  onBack: () => void;
  onPrint: () => void;
  onViewChange: (view: PlannerView) => void;
  onZoomChange: (zoom: number) => void;
};

function Topbar({ canvasZoom, selectedTable, view, onBack, onPrint, onViewChange, onZoomChange }: TopbarProps) {
  return (
    <header className="topbar">
      <button className="outline-button" onClick={onBack}>
        <ChevronLeft size={22} />
        Inapoi
      </button>
      <div className="view-tabs" aria-label="Vizualizare">
        <button className={view === "sketch" ? "active" : ""} onClick={() => onViewChange("sketch")}>
          Schita
        </button>
        <button className={view === "table" ? "active" : ""} onClick={() => onViewChange("table")}>
          Tabel
        </button>
      </div>
      {view === "sketch" ? (
        <div className="board-zoom" aria-label="Zoom schita">
          <button
            type="button"
            title="Zoom out"
            onClick={() => onZoomChange(canvasZoom - CANVAS_ZOOM_STEP)}
          >
            <Minus size={18} />
          </button>
          <button type="button" className="zoom-value" onClick={() => onZoomChange(1)}>
            {Math.round(canvasZoom * 100)}%
          </button>
          <button
            type="button"
            title="Zoom in"
            onClick={() => onZoomChange(canvasZoom + CANVAS_ZOOM_STEP)}
          >
            <Plus size={18} />
          </button>
        </div>
      ) : null}
      <button className="primary-button" onClick={onPrint}>
        <Printer size={21} />
        Print PDF
      </button>
      {selectedTable ? <span className="selected-pill">{selectedTable.name}</span> : null}
    </header>
  );
}

type CanvasViewProps = {
  guestsById: Map<string, Guest>;
  onDeleteTable: (tableId: string) => void;
  onGuestContextMenu: (guestId: string, x: number, y: number) => void;
  onSelectTable: (tableId: string) => void;
  onStartTableDrag: (drag: TableDrag) => void;
  onZoomChange: (zoom: number) => void;
  selectedTableId: string | null;
  tables: PlannerTable[];
  zoom: number;
};

function CanvasView({
  guestsById,
  onDeleteTable,
  onGuestContextMenu,
  onSelectTable,
  onStartTableDrag,
  onZoomChange,
  selectedTableId,
  tables,
  zoom,
}: CanvasViewProps) {
  const shellRef = useRef<HTMLElement | null>(null);
  const panRef = useRef<{
    pointerId: number;
    scrollLeft: number;
    scrollTop: number;
    startX: number;
    startY: number;
  } | null>(null);
  const [isPanning, setIsPanning] = useState(false);

  function handleWheel(event: React.WheelEvent<HTMLElement>) {
    if (!event.ctrlKey && !event.metaKey) {
      return;
    }

    event.preventDefault();

    const shell = shellRef.current;
    if (!shell) {
      return;
    }

    const rect = shell.getBoundingClientRect();
    const viewportX = event.clientX - rect.left;
    const viewportY = event.clientY - rect.top;
    const boardX = (shell.scrollLeft + viewportX) / zoom;
    const boardY = (shell.scrollTop + viewportY) / zoom;
    const nextZoom = normalizeCanvasZoom(zoom * (event.deltaY > 0 ? 0.9 : 1.1));

    onZoomChange(nextZoom);

    requestAnimationFrame(() => {
      shell.scrollLeft = boardX * nextZoom - viewportX;
      shell.scrollTop = boardY * nextZoom - viewportY;
    });
  }

  function handlePanPointerDown(event: React.PointerEvent<HTMLDivElement>) {
    if (event.button !== 0 || event.target !== event.currentTarget) {
      return;
    }

    const shell = shellRef.current;
    if (!shell) {
      return;
    }

    panRef.current = {
      pointerId: event.pointerId,
      scrollLeft: shell.scrollLeft,
      scrollTop: shell.scrollTop,
      startX: event.clientX,
      startY: event.clientY,
    };
    event.currentTarget.setPointerCapture(event.pointerId);
    setIsPanning(true);
  }

  function handlePanPointerMove(event: React.PointerEvent<HTMLDivElement>) {
    const pan = panRef.current;
    const shell = shellRef.current;
    if (!pan || !shell) {
      return;
    }

    shell.scrollLeft = pan.scrollLeft - (event.clientX - pan.startX);
    shell.scrollTop = pan.scrollTop - (event.clientY - pan.startY);
  }

  function stopPanning(event: React.PointerEvent<HTMLDivElement>) {
    if (panRef.current?.pointerId === event.pointerId) {
      event.currentTarget.releasePointerCapture(event.pointerId);
      panRef.current = null;
      setIsPanning(false);
    }
  }

  return (
    <section className="canvas-shell" ref={shellRef} onWheel={handleWheel}>
      <div
        className="canvas-zoom-space"
        style={{ width: CANVAS_WIDTH * zoom, height: CANVAS_HEIGHT * zoom }}
      >
        <div
          className={`canvas ${isPanning ? "is-panning" : ""}`}
          onPointerDown={handlePanPointerDown}
          onPointerMove={handlePanPointerMove}
          onPointerUp={stopPanning}
          onPointerCancel={stopPanning}
          style={{
            width: CANVAS_WIDTH,
            height: CANVAS_HEIGHT,
            transform: `scale(${zoom})`,
          }}
        >
          {tables.map((table) => {
            const layout = getTableLayout(table);
            return (
              <TableNode
                guestsById={guestsById}
                key={table.id}
                onDelete={() => onDeleteTable(table.id)}
                onSelect={() => onSelectTable(table.id)}
                onStartDrag={(event) =>
                  onStartTableDrag({
                    tableId: table.id,
                    startX: event.clientX,
                    startY: event.clientY,
                    originX: table.x,
                    originY: table.y,
                    width: layout.width,
                    height: layout.height,
                  })
                }
                selected={table.id === selectedTableId}
                onGuestContextMenu={onGuestContextMenu}
                table={table}
              />
            );
          })}
        </div>
      </div>
    </section>
  );
}

type TableNodeProps = {
  guestsById: Map<string, Guest>;
  onDelete: () => void;
  onGuestContextMenu: (guestId: string, x: number, y: number) => void;
  onSelect: () => void;
  onStartDrag: (event: React.PointerEvent<HTMLDivElement>) => void;
  selected: boolean;
  table: PlannerTable;
};

function TableNode({
  guestsById,
  onDelete,
  onGuestContextMenu,
  onSelect,
  onStartDrag,
  selected,
  table,
}: TableNodeProps) {
  const layout = getTableLayout(table);
  const occupied = table.seats.filter((seat) => seat.guestId).length;

  function handleBodyPointerDown(event: React.PointerEvent<HTMLDivElement>) {
    if (event.button !== 0) {
      return;
    }
    onSelect();
    onStartDrag(event);
  }

  return (
    <article
      className={`table-node table-node-${table.shape} ${selected ? "selected" : ""}`}
      style={{ left: table.x, top: table.y, width: layout.width, height: layout.height }}
      onClick={onSelect}
    >
      {selected ? (
        <div className="node-actions">
          <button title="Editeaza masa" onClick={onSelect}>
            <Pencil size={18} />
          </button>
          <button
            title="Sterge masa"
            onClick={(event) => {
              event.stopPropagation();
              onDelete();
            }}
          >
            <Trash2 size={18} />
          </button>
        </div>
      ) : null}

      <div
        className={`table-body table-body-${table.shape}`}
        onPointerDown={handleBodyPointerDown}
        style={{
          left: layout.body.x,
          top: layout.body.y,
          width: layout.body.width,
          height: layout.body.height,
        }}
      >
        <Move size={16} className="move-mark" />
        <strong>{table.name}</strong>
        <span>
          {occupied}/{table.seats.length}
        </span>
      </div>

      {table.seats.map((seat, index) => (
        <SeatDrop
          guest={seat.guestId ? guestsById.get(seat.guestId) ?? null : null}
          index={index}
          key={seat.id}
          position={layout.seats[index]}
          seat={seat}
          table={table}
          onGuestContextMenu={onGuestContextMenu}
        />
      ))}
    </article>
  );
}

type SeatDropProps = {
  guest: Guest | null;
  index: number;
  onGuestContextMenu: (guestId: string, x: number, y: number) => void;
  position: SeatPosition | undefined;
  seat: Seat;
  table: PlannerTable;
};

function SeatDrop({ guest, index, onGuestContextMenu, position, seat, table }: SeatDropProps) {
  const { isOver, setNodeRef } = useDroppable({ id: seatDropId(table.id, seat.id) });
  const style: CSSProperties = {
    left: position?.x ?? 0,
    top: position?.y ?? 0,
  };

  return (
    <div
      ref={setNodeRef}
      className={`seat ${guest ? "filled" : ""} ${isOver ? "is-over" : ""}`}
      style={style}
      title={guest?.name ?? `Loc ${index + 1}`}
      onContextMenu={(event) => {
        if (!guest) {
          return;
        }

        event.preventDefault();
        event.stopPropagation();
        onGuestContextMenu(guest.id, event.clientX, event.clientY);
      }}
    >
      {guest ? <GuestSeatToken guest={guest} /> : <span>{index + 1}</span>}
    </div>
  );
}

type GuestRowProps = {
  assigned: boolean;
  guest: Guest;
  onRemove: () => void;
};

function GuestRow({ assigned, guest, onRemove }: GuestRowProps) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: guestDragId(guest.id),
  });
  const style: CSSProperties | undefined = transform
    ? { transform: `translate3d(${transform.x}px, ${transform.y}px, 0)` }
    : undefined;

  return (
    <div
      ref={setNodeRef}
      className={`guest-row ${assigned ? "assigned" : ""} ${isDragging ? "dragging" : ""}`}
      style={style}
    >
      <button className="drag-avatar" type="button" {...listeners} {...attributes}>
        <UserRound size={25} />
      </button>
      <div>
        <strong>{guest.name}</strong>
        <span>{guest.group}</span>
      </div>
      <em>{assigned ? "Asezat" : "Liber"}</em>
      <button className="remove-guest" type="button" title="Sterge invitat" onClick={onRemove}>
        <X size={17} />
      </button>
    </div>
  );
}

function GuestSeatToken({ guest }: { guest: Guest }) {
  const { attributes, listeners, setNodeRef, transform, isDragging } = useDraggable({
    id: guestDragId(guest.id),
  });
  const style: CSSProperties | undefined = transform
    ? { transform: `translate3d(${transform.x}px, ${transform.y}px, 0)` }
    : undefined;

  return (
    <button
      ref={setNodeRef}
      className={`seat-token ${isDragging ? "dragging" : ""}`}
      style={style}
      type="button"
      title={guest.name}
      {...listeners}
      {...attributes}
    >
      <span className="seat-avatar">{initials(guest.name)}</span>
      <span className="seat-name">{guest.name}</span>
    </button>
  );
}

function GuestChip({ guest, floating = false }: { guest: Guest; floating?: boolean }) {
  return (
    <div className={`guest-chip ${floating ? "floating" : ""}`}>
      <UserRound size={20} />
      <span>{guest.name}</span>
    </div>
  );
}

type GuestContextMenuProps = {
  guest: Guest;
  name: string;
  x: number;
  y: number;
  onClose: () => void;
  onNameChange: (name: string) => void;
  onRemove: () => void;
  onSave: () => void;
};

function GuestContextMenu({
  guest,
  name,
  x,
  y,
  onClose,
  onNameChange,
  onRemove,
  onSave,
}: GuestContextMenuProps) {
  const menuStyle: CSSProperties = {
    left: Math.max(12, Math.min(x, window.innerWidth - 276)),
    top: Math.max(12, Math.min(y, window.innerHeight - 190)),
  };

  return (
    <div className="context-menu-layer" onMouseDown={onClose} onContextMenu={(event) => event.preventDefault()}>
      <form
        className="guest-context-menu"
        style={menuStyle}
        onSubmit={(event) => {
          event.preventDefault();
          onSave();
        }}
        onMouseDown={(event) => event.stopPropagation()}
      >
        <div className="guest-context-title">
          <UserRound size={18} />
          <span>{guest.name}</span>
        </div>
        <label>
          Edit name
          <input autoFocus value={name} onChange={(event) => onNameChange(event.target.value)} />
        </label>
        <button className="context-action" type="submit">
          <Pencil size={17} />
          Save name
        </button>
        <button className="context-action danger" type="button" onClick={onRemove}>
          <Trash2 size={17} />
          Remove from table
        </button>
      </form>
    </div>
  );
}

type TableInspectorProps = {
  table: PlannerTable;
  onClose: () => void;
  onDelete: () => void;
  onUpdate: (table: PlannerTable) => void;
};

function TableInspector({ table, onClose, onDelete, onUpdate }: TableInspectorProps) {
  const seated = table.seats.filter((seat) => seat.guestId).length;
  const tableScale = normalizeTableScale(table.scale);

  function setSeatCount(value: number) {
    const nextCount = Math.max(2, Math.min(24, Number.isFinite(value) ? Math.round(value) : table.seats.length));
    onUpdate({ ...table, seats: makeSeats(nextCount, table.seats) });
  }

  function setTableScale(value: number) {
    onUpdate({ ...table, scale: normalizeTableScale(value) });
  }

  return (
    <aside className="inspector">
      <div className="inspector-header">
        <div>
          <span>Masa selectata</span>
          <strong>{table.name}</strong>
        </div>
        <button type="button" title="Inchide" onClick={onClose}>
          <X size={18} />
        </button>
      </div>

      <label>
        Nume
        <input
          value={table.name}
          onChange={(event) => onUpdate({ ...table, name: event.target.value || "masa" })}
        />
      </label>

      <label>
        Locuri
        <input
          max={24}
          min={2}
          type="number"
          value={table.seats.length}
          onChange={(event) => setSeatCount(Number(event.target.value))}
        />
      </label>

      <label>
        Forma
        <select
          value={table.shape}
          onChange={(event) => onUpdate({ ...table, shape: event.target.value as TableShape })}
        >
          <option value="round">{tableShapeLabels.round}</option>
          <option value="rectangle">{tableShapeLabels.rectangle}</option>
          <option value="square">{tableShapeLabels.square}</option>
          <option value="head">{tableShapeLabels.head}</option>
        </select>
      </label>

      <div className="size-control">
        <div>
          <span>Marime masa</span>
          <strong>{Math.round(tableScale * 100)}%</strong>
        </div>
        <div className="zoom-controls">
          <button
            type="button"
            title="Micsoreaza masa"
            onClick={() => setTableScale(tableScale - TABLE_SCALE_STEP)}
          >
            <Minus size={18} />
          </button>
          <input
            aria-label="Marime masa"
            max={TABLE_SCALE_MAX}
            min={TABLE_SCALE_MIN}
            step={TABLE_SCALE_STEP}
            type="range"
            value={tableScale}
            onChange={(event) => setTableScale(Number(event.target.value))}
          />
          <button
            type="button"
            title="Mareste masa"
            onClick={() => setTableScale(tableScale + TABLE_SCALE_STEP)}
          >
            <Plus size={18} />
          </button>
        </div>
      </div>

      <div className="inspector-meter">
        <span>{seated} asezati</span>
        <span>{table.seats.length - seated} libere</span>
      </div>

      <button className="danger-button" type="button" onClick={onDelete}>
        <Trash2 size={18} />
        Sterge masa
      </button>
    </aside>
  );
}

type SeatingTableViewProps = {
  assignedGuestIds: Set<string>;
  guests: Guest[];
  tables: PlannerTable[];
};

function SeatingTableView({ assignedGuestIds, guests, tables }: SeatingTableViewProps) {
  const guestsById = new Map(guests.map((guest) => [guest.id, guest]));
  const unseated = guests.filter((guest) => !assignedGuestIds.has(guest.id));

  return (
    <section className="table-view">
      <div className="table-summary">
        <div>
          <UsersRound size={24} />
          <strong>{guests.length}</strong>
          <span>invitati</span>
        </div>
        <div>
          <Circle size={22} />
          <strong>{tables.length}</strong>
          <span>mese</span>
        </div>
        <div>
          <UserRound size={23} />
          <strong>{unseated.length}</strong>
          <span>neasezati</span>
        </div>
      </div>

      <div className="table-list-view">
        {tables.map((table) => {
          const seated = table.seats
            .map((seat, index) => ({
              index,
              guest: seat.guestId ? guestsById.get(seat.guestId) ?? null : null,
            }))
            .filter((item) => item.guest);

          return (
            <article className="table-report" key={table.id}>
              <header>
                <strong>{table.name}</strong>
                <span>
                  {seated.length}/{table.seats.length} locuri
                </span>
              </header>
              <ol>
                {table.seats.map((seat, index) => {
                  const guest = seat.guestId ? guestsById.get(seat.guestId) : null;
                  return (
                    <li key={seat.id}>
                      <span>Loc {index + 1}</span>
                      <strong>{guest?.name ?? "Liber"}</strong>
                    </li>
                  );
                })}
              </ol>
            </article>
          );
        })}
      </div>

      {unseated.length ? (
        <section className="unseated-report">
          <h2>Neasezati</h2>
          <div>
            {unseated.map((guest) => (
              <span key={guest.id}>{guest.name}</span>
            ))}
          </div>
        </section>
      ) : null}
    </section>
  );
}

export default App;
