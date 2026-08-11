export const PREVIO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "partidas", label: "Partidas" },
  { k: "numPrevio", label: "N° previo" },
  { k: "tipoPrevio", label: "Dificultad", select: ["", "A", "B", "C", "D"] },
];

export const DESPACHO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "pedimento", label: "Pedimento" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "numDespacho", label: "N° despacho" },
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
