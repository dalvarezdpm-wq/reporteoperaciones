export const PREVIO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "partidas", label: "Partidas" },
  { k: "tipoPrevio", label: "Tipo de previo" },
  { k: "asignacion", label: "Asignación" },
  { k: "inicio", label: "Inicio", type: "time" },
  { k: "fin", label: "Fin", type: "time" },
];

export const DESPACHO_FIELDS = [
  { k: "ref", label: "Ref." },
  { k: "guia", label: "Guía" },
  { k: "pedimento", label: "Pedimento" },
  { k: "cliente", label: "Cliente", cat: "clientes" },
  { k: "almacen", label: "Almacén", cat: "almacenes" },
  { k: "tipoDespacho", label: "Tipo despacho" },
  { k: "inicio", label: "Inicio", type: "time" },
  { k: "liberacion", label: "Liberación", type: "time" },
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
    coordinadorOp: "",
    coordinador: "",
    tramitador: "",
    horaInicio: "",
    horaCierre: "",
    aduana: "",
    previos: [],
    despachos: [],
    otrasActividades: ["", "", "", ""],
    elaboro: "",
    reviso: "",
    autorizo: "",
  };
}
