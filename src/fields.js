export const PREVIO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "partidas", label: "Partidas" },
  { k: "numPrevio", label: "N° previo", select: ["", "1er previo", "2do previo", "3er previo"] },
  { k: "tipoOperacion", label: "Tipo de operación", select: ["", "IMPORTACION", "EXPORTACION"] },
  { k: "tipoPrevio", label: "Dificultad (opcional)", select: ["", "A", "B", "C", "D"] },
  {
    k: "sector",
    label: "Sector",
    select: ["", "QUIMICO", "PERECEDERO", "METALURGICO", "TEXTIL", "AGRICULTURA", "MANUFACTURA", "SALUD", "DIGITAL"],
  },
  { k: "ejecutivo", label: "Ejecutivo", cat: "ejecutivos" },
  { k: "observaciones", label: "Observaciones" },
];

export const DESPACHO_FIELDS = [
  { k: "pedimento", label: "Pedimento" },
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "tipoOperacion", label: "Tipo de operación", select: ["", "IMPORTACION", "EXPORTACION"] },
  { k: "dificultad", label: "Dificultad (opcional)", select: ["", "A", "B", "C", "D"] },
  { k: "resultado", label: "Resultado", cat: "resultados" },
  { k: "ejecutivo", label: "Ejecutivo", cat: "ejecutivos" },
  { k: "observaciones", label: "Observaciones" },
];

export const REVALIDADA_FIELDS = [
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
];

export const PENDIENTE_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "pedimento", label: "Pedimento" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "tipo", label: "Tipo (previo/despacho/revalidada)", select: ["", "previo", "despacho", "revalidada"] },
  { k: "ejecutivo", label: "Ejecutivo", cat: "ejecutivos" },
  { k: "observaciones", label: "Observaciones" },
];

export function emptyCaptura() {
  return {
    id: "c_" + Date.now() + "_" + Math.random().toString(36).slice(2, 8),
    uploadedBy: null,
    horaCaptura: null,
    sourceType: null,
    sourceFileName: "",
    fecha: null,
    tramitador: "",
    aduana: "",
    previos: [],
    despachos: [],
    revalidadas: [],
    pendientes: [],
    otrasActividades: ["", "", "", ""],
    historial: [],
  };
}
