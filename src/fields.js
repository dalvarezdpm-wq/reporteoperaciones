export const PREVIO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "partidas", label: "Partidas" },
  { k: "tipoPrevio", label: "Tipo de previo" },
];

export const DESPACHO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "pedimento", label: "Pedimento" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "tipoDespacho", label: "Tipo despacho" },
  { k: "resultado", label: "Resultado", cat: "resultados" },
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
    otrasActividades: ["", "", "", ""],
  };
}
