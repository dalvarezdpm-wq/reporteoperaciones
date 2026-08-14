export const PREVIO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "partidas", label: "Partidas" },
  { k: "numPrevio", label: "N° previo", select: ["", "1er previo", "2do previo", "3er previo"] },
  { k: "tipoPrevio", label: "Dificultad", select: ["", "A", "B", "C", "D"] },
  {
    k: "sector",
    label: "Sector",
    select: ["", "QUIMICO", "PERECEDERO", "METALURGICO", "TEXTIL", "AGRICULTURA", "MANUFACTURA", "SALUD", "DIGITAL"],
  },
];

export const DESPACHO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "pedimento", label: "Pedimento" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "numDespacho", label: "N° despacho", select: ["", "1er despacho", "2do despacho"] },
  { k: "tipoDespacho", label: "Tipo despacho" },
  { k: "resultado", label: "Resultado", cat: "resultados" },
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
