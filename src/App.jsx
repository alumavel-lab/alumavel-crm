import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Users, Briefcase, Search, Plus, X, Pencil, Trash2, ChevronLeft,
  Building2, Phone, Mail, MapPin, Euro, Clock, FileText, CheckCircle2,
  AlertCircle, Circle, Loader2, Hash, ClipboardList, Receipt, Timer,
  ChevronRight, Save, Truck, Boxes, AlertTriangle, ArrowDownCircle, ArrowUpCircle, Package, AlertOctagon,
  CalendarDays, Layers, Ruler, LogIn, LogOut, Coffee, Download, FileSpreadsheet, Wallet, Lock, UserCog, ShieldCheck,
  Globe, MessageCircle, BarChart3, Factory, Wrench, Copy, Image as ImageIcon, Menu, Send, Printer
} from "lucide-react";
import * as XLSX from "xlsx";
import { PieChart, Pie, Cell, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, LineChart, Line } from "recharts";
import { initializeApp } from "firebase/app";
import { getDatabase, ref, get as fbGet, set as fbSet } from "firebase/database";

// Configuración de tu proyecto de Firebase (crmalumavel). La apiKey de Firebase
// no es un secreto — el acceso real se controla con las reglas de seguridad de
// la base de datos, no ocultando esta clave.
const firebaseConfig = {
  apiKey: "AIzaSyBf51Gy2drHdUyJ4kCstNdvvV2h3uYe4RM",
  authDomain: "crmalumavel.firebaseapp.com",
  databaseURL: "https://crmalumavel-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "crmalumavel",
  storageBucket: "crmalumavel.firebasestorage.app",
  messagingSenderId: "14276528333",
  appId: "1:14276528333:web:e85e21696c3b9b82a18184",
  measurementId: "G-4E5XMQ3149",
};
const fbApp = initializeApp(firebaseConfig);
const fbDb = getDatabase(fbApp);

/* ---------------------------------------------------------------
   ALUMAVEL · CRM — Núcleo Fase I: Clientes + Proyectos/Obras
   Datos persistidos en Firebase Realtime Database (compartidos con el equipo)
--------------------------------------------------------------- */

const K_CLIENTES = "alumavel:clientes";
const K_PROYECTOS = "alumavel:proyectos";
const K_PROVEEDORES = "alumavel:proveedores";
const K_MATERIALES = "alumavel:materiales";
const K_PEDIDOS = "alumavel:pedidos";
const K_INCIDENCIAS = "alumavel:incidencias";
const K_ARTICULOS = "alumavel:articulos";
const K_FACTURAS = "alumavel:facturas";
const K_FICHAJES = "alumavel:fichajes";
const K_USUARIOS = "alumavel:usuarios";
const K_SESION = "alumavel:sesion_usuario_id";
const ROLES = ["Administrador", "Usuario"];
const MODULOS_DISPONIBLES = [
  { id: "proyectos", label: "Proyectos / Obras" },
  { id: "clientes", label: "Clientes" },
  { id: "proveedores", label: "Proveedores" },
  { id: "stock", label: "Gestión de Stock" },
  { id: "pedidos", label: "Pedidos" },
  { id: "incidencias", label: "Incidencias" },
  { id: "calendario", label: "Calendario" },
  { id: "articulos", label: "Artículos" },
  { id: "facturas", label: "Facturas" },
  { id: "presupuestos", label: "Presupuestos" },
  { id: "mediciones", label: "Mediciones" },
  { id: "chat", label: "Chat" },
  { id: "tareas", label: "Tareas" },
  { id: "ingresos", label: "Entrada de dinero" },
  { id: "informes", label: "Informes" },
  { id: "fabrica", label: "Fábrica" },
  { id: "instalaciones", label: "Instalaciones" },
  { id: "fichajes", label: "Fichajes" },
];

// Configuración física del almacén de cristales dentro de Fábrica. La ubicación se
// asigna a nivel de CABALLETE (rack), no por cristal individual — cada hueco guarda
// un caballete entero. Arriba son los caballetes de Uxcar, abajo los de ALUMAVEL.
const ZONAS_CRISTALES = {
  arriba: { label: "Arriba (Uxcar)", filas: 3, huecos: 15 },
  abajo: { label: "Abajo (ALUMAVEL)", filas: 2, huecos: 15 },
};
// Fila que se deja siempre libre como "colchón": solo se usa cuando el resto de
// filas de esa zona ya están completamente llenas.
const FILA_RESERVA = { arriba: 3 };
const ubicacionTexto = (u) => (u ? `${u.zona === "arriba" ? "Arriba" : "Abajo"} · Fila ${u.fila} · Hueco ${u.hueco}` : "Sin ubicar");

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
// Firebase Realtime Database a veces devuelve un objeto en vez de un array (por huecos
// en los índices, o listas vacías) — esto lo normaliza siempre a un array de verdad.
const toArray = (val) => {
  if (Array.isArray(val)) return val;
  if (val && typeof val === "object") return Object.values(val);
  return [];
};
const escapeHtml = (s) => (s || "").toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

// Lector genérico de Excel/CSV para importaciones masivas (contactos, tarifas...).
// A diferencia del importador de "control de montaje" (que espera una estructura
// fija de bloques/viviendas), este simplemente lee la primera fila como cabeceras
// y el resto como filas de datos, tolerando que las columnas tengan nombres
// distintos (Nombre/Cliente/Empresa, Teléfono/Móvil/Tel, etc.).
function leerFilasExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const hoja = workbook.Sheets[workbook.SheetNames[0]];
  // Algunas hojas tienen un título o notas en las primeras filas antes de las
  // cabeceras reales (como "Control de Reparaciones" en la fila 1). Buscamos
  // la primera fila con al menos 2 celdas rellenas y la usamos como cabecera,
  // en vez de asumir siempre que las cabeceras están en la fila 1.
  const filasCrudas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: "" });
  let indiceCabecera = 0;
  for (let i = 0; i < Math.min(filasCrudas.length, 15); i++) {
    const noVacias = (filasCrudas[i] || []).filter((c) => String(c).trim() !== "").length;
    if (noVacias >= 2) { indiceCabecera = i; break; }
  }
  const cabeceras = (filasCrudas[indiceCabecera] || []).map((c) => String(c || "").trim());
  const filas = [];
  for (let i = indiceCabecera + 1; i < filasCrudas.length; i++) {
    const fila = filasCrudas[i];
    if (!fila || fila.every((c) => String(c).trim() === "")) continue;
    const obj = {};
    cabeceras.forEach((cab, idx) => { if (cab) obj[cab] = fila[idx] !== undefined ? fila[idx] : ""; });
    filas.push(obj);
  }
  return filas;
}
function normalizarCabecera(s) {
  return (s || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
}
function valorPorCabeceras(fila, candidatos) {
  const claves = Object.keys(fila);
  // Primero probamos coincidencia exacta (más fiable), y si no hay ninguna,
  // probamos que la cabecera "contenga" al candidato (ej: "Cliente / Referencia"
  // contiene "cliente"), para no depender de que el texto sea idéntico.
  for (const candidato of candidatos) {
    const claveExacta = claves.find((k) => normalizarCabecera(k) === candidato);
    if (claveExacta !== undefined && String(fila[claveExacta]).trim() !== "") return fila[claveExacta];
  }
  for (const candidato of candidatos) {
    const claveParcial = claves.find((k) => normalizarCabecera(k).includes(candidato));
    if (claveParcial !== undefined && String(fila[claveParcial]).trim() !== "") return fila[claveParcial];
  }
  return "";
}
const money = (n) => (isNaN(n) ? "0,00 €" : Number(n).toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + " €");
const fmtDate = (d) => {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return "—";
  return dt.toLocaleDateString("es-ES", { day: "2-digit", month: "2-digit", year: "numeric" });
};
const daysDiff = (a, b) => Math.round((new Date(b) - new Date(a)) / 86400000);

// Descarga cualquier listado como un documento Word (en realidad HTML con extensión
// .doc, que Word abre sin problema) — usado por el botón "Descargar esta vista (Word)"
// en varios módulos. `columnas` es un array de textos de cabecera, `filas` un array de
// arrays con el mismo número de columnas.
const descargarListaComoWord = (titulo, columnas, filas) => {
  const filaTabla = (cols) => `<tr>${cols.map((c) => `<td style="border:1px solid #ccc;padding:6px 10px;">${c}</td>`).join("")}</tr>`;
  const cabeceraTabla = (cols) => `<tr>${cols.map((c) => `<th style="border:1px solid #ccc;padding:6px 10px;background:#f1f1f1;text-align:left;">${c}</th>`).join("")}</tr>`;
  const htmlDoc = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset="utf-8"><title>${titulo}</title></head>
    <body style="font-family: Calibri, Arial, sans-serif; color:#1a1a1a;">
      <h1 style="color:#2E8B57;">${titulo}</h1>
      <p style="color:#666;">Generado el ${fmtDate(new Date().toISOString().slice(0, 10))}</p>
      <table style="border-collapse:collapse;width:100%;">
        ${cabeceraTabla(columnas)}
        ${filas.map(filaTabla).join("")}
      </table>
    </body>
    </html>
  `;
  const blob = new Blob(["\ufeff", htmlDoc], { type: "application/msword" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `${titulo.toLowerCase().replace(/[^a-z0-9]+/g, "_")}_${new Date().toISOString().slice(0, 10)}.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
};

const ESTADO_PRESUPUESTO = ["Esperando presupuesto", "Presupuesto enviado", "Presupuesto aceptado", "Presupuesto rechazado"];
const ESTADO_TRABAJO = ["Pendiente de aceptación", "En proceso", "Albarán de carga firmado", "Listo para reparto/recogida", "Entregado", "Cancelado"];
const CATEGORIAS_CHECKLIST_MATERIALES = ["Cristal", "Persiana", "Perfil", "Postigo", "Panel de puerta", "Tirador / Manilla", "Mosquitera", "Guía de persiana"];
const checklistMaterialesPorDefecto = () => CATEGORIAS_CHECKLIST_MATERIALES.map((nombre) => ({ id: uid(), nombre, estado: null }));
const normalizarChecklist = (lista) => {
  if (!lista || !lista.length) return checklistMaterialesPorDefecto();
  return lista.map((item) => ({
    ...item,
    estado: item.estado !== undefined ? item.estado : (item.marcado ? "si" : null),
  }));
};
const ESTADO_LOGISTICA = ["Sin definir", "Reparto (camión)", "Recogida en fábrica"];
const PROVINCIAS_REPARTO = ["Almería", "Granada", "Murcia", "Valencia", "Alicante", "Otra ciudad..."];
const TIPO_CLIENTE = ["Cliente", "Distribuidor"];
const FORMA_PAGO = ["Contado", "Transferencia", "Pago a 30 días", "Pago a 60 días"];
const TIPO_VENTA = ["Preventa", "Postventa"];
const UNIDAD_COMPRA = ["Unidad", "Metro lineal", "m²", "Kg", "Barra", "Caja", "Palet"];
const ESTADO_MOVIMIENTO = ["Previsto", "Recibido", "Cancelado"];
const ESTADO_MOVIMIENTO_STYLE = {
  "Previsto": "bg-sky-50 text-sky-700 ring-sky-200",
  "Recibido": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "Cancelado": "bg-slate-100 text-slate-500 ring-slate-200",
};
const ESTADO_PEDIDO = ["Pendiente", "Realizado", "Recibido", "Reclamado", "Cancelado"];
const ESTADO_PEDIDO_STYLE = {
  "Pendiente": "bg-slate-100 text-slate-600 ring-slate-200",
  "Realizado": "bg-sky-50 text-sky-700 ring-sky-200",
  "Recibido": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "Reclamado": "bg-amber-50 text-amber-700 ring-amber-200",
  "Cancelado": "bg-rose-50 text-rose-700 ring-rose-200",
};
const ESTADO_LINEA_PEDIDO_STYLE = {
  "Solicitado": "bg-sky-50 text-sky-700 ring-sky-200",
  "Recibido": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "No recibido": "bg-rose-50 text-rose-700 ring-rose-200",
};
const ESTADO_TRABAJO_INCIDENCIA = ["Pendiente revisión", "En proceso", "Entregado"];
const ESTADO_INCIDENCIA = ["Pendiente revisión", "En proceso", "Entregado", "Solucionado"];
const ESTADO_INCIDENCIA_STYLE = {
  "Pendiente revisión": "bg-rose-50 text-rose-700 ring-rose-200",
  "En proceso": "bg-amber-50 text-amber-700 ring-amber-200",
  "Entregado": "bg-sky-50 text-sky-700 ring-sky-200",
  "Solucionado": "bg-emerald-50 text-emerald-700 ring-emerald-200",
};
const ESTADO_TRABAJO_INCIDENCIA_STYLE = {
  "Pendiente revisión": "bg-rose-50 text-rose-700 ring-rose-200",
  "En proceso": "bg-amber-50 text-amber-700 ring-amber-200",
  "Entregado": "bg-emerald-50 text-emerald-700 ring-emerald-200",
};
const TIPO_FACTURA = ["Definitiva", "Proforma", "Abono"];
const ESTADO_FACTURA_STYLE = {
  "Pendiente": "bg-rose-50 text-rose-700 ring-rose-200",
  "Parcial": "bg-amber-50 text-amber-700 ring-amber-200",
  "Pagada": "bg-emerald-50 text-emerald-700 ring-emerald-200",
};
const estadoFacturaCalc = (factura) => {
  const total = parseFloat(factura.total) || 0;
  const pagado = (factura.pagos || []).reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
  if (pagado <= 0) return "Pendiente";
  if (pagado < total) return "Parcial";
  return "Pagada";
};

const ESTADO_TRABAJO_STYLE = {
  "Pendiente de aceptación": "bg-rose-50 text-rose-700 ring-rose-200",
  "En proceso": "bg-amber-50 text-amber-700 ring-amber-200",
  "Entregado": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "Cancelado": "bg-slate-100 text-slate-500 ring-slate-200",
};
const ESTADO_PRESUPUESTO_STYLE = {
  "Esperando presupuesto": "bg-slate-100 text-slate-600 ring-slate-200",
  "Presupuesto enviado": "bg-sky-50 text-sky-700 ring-sky-200",
  "Presupuesto aceptado": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "Presupuesto rechazado": "bg-rose-50 text-rose-700 ring-rose-200",
};

function Badge({ children, className = "" }) {
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-medium ring-1 whitespace-nowrap ${className}`}>
      {children}
    </span>
  );
}

function CornerFrame({ children, className = "" }) {
  return (
    <div className={`relative ${className}`}>
      <span className="absolute -top-px -left-px w-3 h-3 border-t-2 border-l-2 border-[#2E8B57]" />
      <span className="absolute -top-px -right-px w-3 h-3 border-t-2 border-r-2 border-[#2E8B57]" />
      <span className="absolute -bottom-px -left-px w-3 h-3 border-b-2 border-l-2 border-[#2E8B57]" />
      <span className="absolute -bottom-px -right-px w-3 h-3 border-b-2 border-r-2 border-[#2E8B57]" />
      {children}
    </div>
  );
}

function Field({ label, children, required }) {
  return (
    <label className="block">
      <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-1">
        {label}{required && <span className="text-[#C2542A]"> *</span>}
      </span>
      {children}
    </label>
  );
}

const inputCls =
  "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 placeholder:text-slate-400 focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]";

function TextInput(props) { return <input {...props} className={inputCls + " " + (props.className || "")} />; }
function Select({ children, ...props }) { return <select {...props} className={inputCls + " " + (props.className || "")}>{children}</select>; }
function TextArea(props) { return <textarea {...props} className={inputCls + " " + (props.className || "")} />; }

/* ================= APP ================= */

export default function App() {
  const [loading, setLoading] = useState(true);
  const [clientes, setClientes] = useState([]);
  const [proyectos, setProyectos] = useState([]);
  const [proveedores, setProveedores] = useState([]);
  const [materiales, setMateriales] = useState([]);
  const [pedidos, setPedidos] = useState([]);
  const [incidencias, setIncidencias] = useState([]);
  const [articulos, setArticulos] = useState([]);
  const [facturas, setFacturas] = useState([]);
  const [presupuestos, setPresupuestos] = useState([]);
  const [ingresos, setIngresos] = useState([]);
  const [solicitudesPedido, setSolicitudesPedido] = useState([]);
  const [instalaciones, setInstalaciones] = useState([]);
  const [mediciones, setMediciones] = useState([]);
  const [tareas, setTareas] = useState([]);
  const [archivosEmpresa, setArchivosEmpresa] = useState([]);
  const [vehiculos, setVehiculos] = useState([]);
  const [cristales, setCristales] = useState([]);
  const [fichajes, setFichajes] = useState([]);
  const [usuarios, setUsuarios] = useState([]);
  const [sesionesUsuario, setSesionesUsuario] = useState({});
  const [sesionUsuarioId, setSesionUsuarioId] = useState(null);
  const [sesionClienteId, setSesionClienteId] = useState(null);
  const [modulo, setModulo] = useState("proyectos"); // clientes | proyectos | proveedores | stock | pedidos | incidencias | calendario | articulos | facturas | presupuestos | fichajes | administracion
  const [menuMovilAbierto, setMenuMovilAbierto] = useState(false);
  const [toast, setToast] = useState(null);

  // clientes view state
  const [clienteView, setClienteView] = useState("list"); // list | form | detail
  const [clienteEditId, setClienteEditId] = useState(null);
  const [clienteDetailId, setClienteDetailId] = useState(null);

  // proyectos view state
  const [proyectoView, setProyectoView] = useState("list");
  const [proyectoEditId, setProyectoEditId] = useState(null);
  const [proyectoDetailId, setProyectoDetailId] = useState(null);

  // proveedores view state
  const [proveedorView, setProveedorView] = useState("list");
  const [proveedorEditId, setProveedorEditId] = useState(null);
  const [proveedorDetailId, setProveedorDetailId] = useState(null);

  // stock view state
  const [materialView, setMaterialView] = useState("list");
  const [materialEditId, setMaterialEditId] = useState(null);
  const [materialDetailId, setMaterialDetailId] = useState(null);

  // pedidos view state
  const [pedidoView, setPedidoView] = useState("list");
  const [pedidoEditId, setPedidoEditId] = useState(null);
  const [pedidoDetailId, setPedidoDetailId] = useState(null);
  const [pedidoPrefill, setPedidoPrefill] = useState(null);
  const [ingresoPrefill, setIngresoPrefill] = useState(null);
  const [pedidoTabPrincipal, setPedidoTabPrincipal] = useState("pedidos");
  const [solicitudPrefill, setSolicitudPrefill] = useState(null);

  // incidencias view state
  const [incidenciaView, setIncidenciaView] = useState("list");
  const [incidenciaEditId, setIncidenciaEditId] = useState(null);
  const [incidenciaDetailId, setIncidenciaDetailId] = useState(null);

  // artículos view state
  const [articuloView, setArticuloView] = useState("list");
  const [articuloEditId, setArticuloEditId] = useState(null);
  const [articuloDetailId, setArticuloDetailId] = useState(null);

  // facturas view state
  const [facturaView, setFacturaView] = useState("list");
  const [facturaEditId, setFacturaEditId] = useState(null);
  const [facturaDetailId, setFacturaDetailId] = useState(null);

  // presupuestos view state
  const [presupuestoView, setPresupuestoView] = useState("list");
  const [presupuestoEditId, setPresupuestoEditId] = useState(null);
  const [presupuestoDetailId, setPresupuestoDetailId] = useState(null);

  // ingresos (entradas de dinero) view state
  const [ingresoView, setIngresoView] = useState("list");
  const [ingresoEditId, setIngresoEditId] = useState(null);

  // solicitudes de pedido (empleados) view state
  const [solicitudPedidoView, setSolicitudPedidoView] = useState("list");
  const [solicitudPedidoEditId, setSolicitudPedidoEditId] = useState(null);
  const [solicitudPedidoDetailId, setSolicitudPedidoDetailId] = useState(null);

  // instalaciones view state
  const [instalacionView, setInstalacionView] = useState("list");
  const [instalacionDetailId, setInstalacionDetailId] = useState(null);

  // mediciones view state
  const [medicionView, setMedicionView] = useState("list");
  const [medicionEditId, setMedicionEditId] = useState(null);
  const [medicionDetailId, setMedicionDetailId] = useState(null);
  const [presupuestoPrefill, setPresupuestoPrefill] = useState(null);

  // fichajes
  const [empleadoActual, setEmpleadoActual] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const claves = ["clientes", "proyectos", "proveedores", "materiales", "pedidos", "incidencias",
          "articulos", "facturas", "presupuestos", "ingresos", "solicitudes_pedido", "instalaciones",
          "vehiculos", "fichajes", "usuarios", "cristales", "mediciones", "sesionesUsuario", "tareas", "archivosEmpresa"];
        const resultados = {};
        await Promise.all(claves.map(async (k) => {
          const snap = await fbGet(ref(fbDb, k)).catch(() => null);
          resultados[k] = snap && snap.exists() ? snap.val() : null;
        }));
        if (resultados.clientes) setClientes(toArray(resultados.clientes));
        if (resultados.proyectos) setProyectos(toArray(resultados.proyectos));
        if (resultados.proveedores) setProveedores(toArray(resultados.proveedores));
        if (resultados.materiales) setMateriales(toArray(resultados.materiales));
        if (resultados.pedidos) setPedidos(toArray(resultados.pedidos));
        if (resultados.incidencias) setIncidencias(toArray(resultados.incidencias));
        if (resultados.articulos) setArticulos(toArray(resultados.articulos));
        if (resultados.facturas) setFacturas(toArray(resultados.facturas));
        if (resultados.presupuestos) setPresupuestos(toArray(resultados.presupuestos));
        if (resultados.ingresos) setIngresos(toArray(resultados.ingresos));
        if (resultados.solicitudes_pedido) setSolicitudesPedido(toArray(resultados.solicitudes_pedido));
        if (resultados.instalaciones) setInstalaciones(toArray(resultados.instalaciones));
        if (resultados.vehiculos) {
          setVehiculos(toArray(resultados.vehiculos));
        } else {
          const flotaPorDefecto = [
            { id: uid(), nombre: "Furgoneta pequeña 1", tipo: "Furgoneta pequeña" },
            { id: uid(), nombre: "Furgoneta pequeña 2", tipo: "Furgoneta pequeña" },
            { id: uid(), nombre: "Furgoneta pequeña 3", tipo: "Furgoneta pequeña" },
            { id: uid(), nombre: "Furgoneta grande", tipo: "Furgoneta grande" },
            { id: uid(), nombre: "Camión 3.500", tipo: "Camión" },
            { id: uid(), nombre: "Camión 8.500", tipo: "Camión" },
          ];
          setVehiculos(flotaPorDefecto);
          fbSet(ref(fbDb, "vehiculos"), flotaPorDefecto).catch(() => {});
        }
        if (resultados.fichajes) setFichajes(toArray(resultados.fichajes));
        if (resultados.usuarios) setUsuarios(toArray(resultados.usuarios));
        if (resultados.cristales) setCristales(toArray(resultados.cristales));
        if (resultados.mediciones) setMediciones(toArray(resultados.mediciones));
        if (resultados.sesionesUsuario) setSesionesUsuario(resultados.sesionesUsuario);
        if (resultados.tareas) setTareas(toArray(resultados.tareas));
        if (resultados.archivosEmpresa) setArchivosEmpresa(toArray(resultados.archivosEmpresa));

        // Estas son locales de este navegador/dispositivo, no compartidas — cada persona
        // mantiene su propia sesión iniciada en su propio ordenador o móvil.
        const empLocal = localStorage.getItem("alumavel_empleado_actual");
        if (empLocal) setEmpleadoActual(empLocal);
        const sesLocal = localStorage.getItem("alumavel_sesion_usuario_id");
        if (sesLocal) setSesionUsuarioId(sesLocal);
        const sesCliLocal = localStorage.getItem("alumavel_sesion_cliente_id");
        if (sesCliLocal) setSesionClienteId(sesCliLocal);
      } catch (e) {
        console.error(e);
      } finally {
        setLoading(false);
      }
    })();
  }, []);

  const persistTimers = useRef({});
  const persistLatest = useRef({});
  const persistInFlight = useRef({});

  const writeToStorage = useCallback(async (key) => {
    if (persistInFlight.current[key]) {
      persistTimers.current[key] = setTimeout(() => writeToStorage(key), 200);
      return;
    }
    persistInFlight.current[key] = true;
    const payload = persistLatest.current[key];
    for (let intento = 1; intento <= 3; intento++) {
      try {
        await fbSet(ref(fbDb, key), payload);
        persistInFlight.current[key] = false;
        return;
      } catch (e) {
        console.error(`storage error (intento ${intento})`, e);
        if (intento < 3) {
          await new Promise((r) => setTimeout(r, 500 * intento));
        }
      }
    }
    persistInFlight.current[key] = false;
    showToast("No se pudo guardar. Comprueba tu conexión y vuelve a intentarlo.", "error");
  }, []);

  const persist = useCallback((key, value) => {
    persistLatest.current[key] = value;
    if (persistTimers.current[key]) clearTimeout(persistTimers.current[key]);
    persistTimers.current[key] = setTimeout(() => writeToStorage(key), 350);
  }, [writeToStorage]);

  const showToast = (msg, kind = "ok") => {
    setToast({ msg, kind });
    setTimeout(() => setToast(null), 2600);
  };

  // Al crear un proyecto nuevo, se genera automáticamente una tarjeta en el
  // tablero de Trello "BETA - Flujo Pedidos ALUMAVEL", en la lista
  // "1. Presupuesto". Las listas de destino están fijadas (no se
  // buscan dinámicamente) para que sea rápido y fiable.
  const TRELLO_LIST_PRESUPUESTO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c8559d524e872245b15";
  const TRELLO_LIST_PRESUPUESTO_ACEPTADO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c88ccd1454328ee7bdb";
  const TRELLO_LIST_PRESUPUESTO_NO_ACEPTADO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c8b012703c3f4d05960";
  const TRELLO_LIST_PEDIDO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c8d3a955c40e82419ea";
  const TRELLO_LIST_FABRICA = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1fde6c1c9fa356948902";
  const TRELLO_LIST_ALBARAN_FIRMADO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b25bfd19de306990759e6";
  const TRELLO_LIST_INCIDENCIA = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c93ca0a4bdfc7864ecf";
  const TRELLO_LIST_TERMINADO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1c9627043753aababaf2";
  const TRELLO_LIST_REPARTO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1d1886784010eff5f400";
  const TRELLO_LIST_RECOGIDA = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1d1a4e0416054c67d219";
  const TRELLO_LIST_FACTURADO = "ari:cloud:trello::list/workspace/660f99b0d1c36157c36cae6d/6a8b1f44235d22fcea51c83f";

  const crearTarjetaTrello = async (proyecto, clienteNombre, onCreated) => {
    try {
      const desc = [
        `Cliente: ${clienteNombre || "—"}`,
        `Presupuesto: ${proyecto.importePresupuesto ? money(proyecto.importePresupuesto) : "—"}`,
        `Entrega prevista: ${proyecto.fechaEntregaPrevista || "—"}`,
        `Referencia CRM: Proyecto #${proyecto.numero} — ${proyecto.nombre || ""}`,
      ].join("\n");

      const prompt = `Crea una tarjeta nueva en Trello usando la herramienta trelloWriteCard con action="create", listId="${TRELLO_LIST_PRESUPUESTO}", name="#${proyecto.numero} — ${proyecto.nombre || "Sin nombre"} (${clienteNombre || "sin cliente"})", y desc="${desc.replace(/"/g, "'")}". No hagas nada más, no busques el tablero ni la lista, usa exactamente el listId indicado. Al terminar, responde ÚNICAMENTE con el id de la tarjeta creada (el campo "id" que devuelve la herramienta), sin ningún otro texto.`;

      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-5",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }],
          mcp_servers: [{ type: "url", url: "https://mcp.trello.com/v1", name: "trello-mcp" }],
        }),
      });
      if (!response.ok) throw new Error("Respuesta no válida de la API");
      const data = await response.json();
      showToast("Tarjeta creada en Trello (Presupuesto)");

      if (onCreated) {
        const toolResults = (data.content || []).filter((c) => c.type === "mcp_tool_result");
        let cardId = null;
        for (const block of toolResults) {
          const text = block?.content?.[0]?.text || "";
          const match = text.match(/ari:cloud:trello::card\/[^\s"'\\]+/);
          if (match) { cardId = match[0]; break; }
        }
        if (!cardId) {
          const finalText = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join(" ");
          const match2 = finalText.match(/ari:cloud:trello::card\/[^\s"'\\]+/);
          if (match2) cardId = match2[0];
        }
        if (cardId) onCreated(cardId);
      }
    } catch (err) {
      showToast("No se pudo crear la tarjeta en Trello (revisa la conexión)", "error");
    }
  };

  const moverTarjetaTrello = async (cardId, listId, etiqueta) => {
    if (!cardId) return;
    try {
      const prompt = `Mueve la tarjeta de Trello con id="${cardId}" a la lista con listId="${listId}" usando la herramienta trelloWriteCard con action="move". No hagas nada más ni busques nada, usa exactamente esos identificadores.`;
      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-5",
          max_tokens: 500,
          messages: [{ role: "user", content: prompt }],
          mcp_servers: [{ type: "url", url: "https://mcp.trello.com/v1", name: "trello-mcp" }],
        }),
      });
      if (!response.ok) throw new Error("Respuesta no válida de la API");
      showToast(`Tarjeta de Trello movida${etiqueta ? ` a "${etiqueta}"` : ""}`);
    } catch (err) {
      showToast("No se pudo mover la tarjeta en Trello (revisa la conexión)", "error");
    }
  };

  const crearTarjetaTrelloPresupuesto = async (presupuesto, onCreated) => {
    try {
      const desc = [
        `Cliente: ${presupuesto.clienteNombre || "—"}`,
        `Importe: ${presupuesto.importe ? money(presupuesto.importe) : "—"}`,
        `Descripción: ${presupuesto.descripcion || "—"}`,
        `Teléfono: ${presupuesto.telefono || "—"}`,
        `Referencia CRM: Presupuesto ${presupuesto.numero}`,
      ].join("\n");
      const nombreTarjeta = `${presupuesto.numero} — ${presupuesto.clienteNombre || "Sin cliente"}`;
      const prompt = `Crea una tarjeta nueva en Trello usando la herramienta trelloWriteCard con action="create", listId="${TRELLO_LIST_PRESUPUESTO}", name="${nombreTarjeta.replace(/"/g, "'")}", y desc="${desc.replace(/"/g, "'")}". No hagas nada más, no busques el tablero ni la lista, usa exactamente el listId indicado. Al terminar, responde ÚNICAMENTE con el id de la tarjeta creada (el campo "id" que devuelve la herramienta), sin ningún otro texto.`;

      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-5",
          max_tokens: 1000,
          messages: [{ role: "user", content: prompt }],
          mcp_servers: [{ type: "url", url: "https://mcp.trello.com/v1", name: "trello-mcp" }],
        }),
      });
      if (!response.ok) throw new Error("Respuesta no válida de la API");
      const data = await response.json();
      showToast("Tarjeta creada en Trello (Presupuesto)");

      if (onCreated) {
        const toolResults = (data.content || []).filter((c) => c.type === "mcp_tool_result");
        let cardId = null;
        for (const block of toolResults) {
          const text = block?.content?.[0]?.text || "";
          const match = text.match(/ari:cloud:trello::card\/[^\s"'\\]+/);
          if (match) { cardId = match[0]; break; }
        }
        if (!cardId) {
          const finalText = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join(" ");
          const match2 = finalText.match(/ari:cloud:trello::card\/[^\s"'\\]+/);
          if (match2) cardId = match2[0];
        }
        if (cardId) onCreated(cardId);
      }
    } catch (err) {
      showToast("No se pudo crear la tarjeta en Trello (revisa la conexión)", "error");
    }
  };

  const saveClientes = (next) => { setClientes(next); persist("clientes", next); };

  const importarContactosMasivo = (nuevosClientes) => {
    saveClientes([...nuevosClientes, ...clientes]);
    showToast(`${nuevosClientes.length} contacto(s) importado(s)`);
  };
  const saveProyectos = (next) => { setProyectos(next); persist("proyectos", next); };

  const upsertCliente = (data) => {
    let next;
    if (data.id) {
      next = clientes.map((c) => (c.id === data.id ? data : c));
      showToast("Cliente actualizado");
    } else {
      const nc = { ...data, id: uid() };
      next = [nc, ...clientes];
      showToast("Cliente dado de alta");
    }
    saveClientes(next);
    setClienteView("list");
  };

  // Alta rápida de cliente desde dentro del formulario de Presupuesto, sin navegar
  // fuera de él — solo pide lo mínimo obligatorio (nombre y límite de crédito).
  const crearClienteRapido = (nombre, limiteCredito) => {
    const nc = { id: uid(), nombre, limiteCredito: parseFloat(limiteCredito) || 0, tipo: "Cliente", formaPago: "Contado" };
    saveClientes([nc, ...clientes]);
    showToast(`Cliente "${nombre}" dado de alta`);
  };

  const deleteCliente = (id) => {
    if (proyectos.some((p) => p.clienteId === id)) {
      showToast("No se puede eliminar: tiene proyectos asociados", "error");
      return;
    }
    saveClientes(clientes.filter((c) => c.id !== id));
    setClienteView("list");
    showToast("Cliente eliminado");
  };

  const nextNumeroProyecto = () => {
    const nums = proyectos.map((p) => parseInt(String(p.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    const max = nums.length ? Math.max(...nums) : 4189;
    return String(max + 1);
  };

  const saveInstalaciones = (next) => { setInstalaciones(next); persist("instalaciones", next); };

  const saveMediciones = (next) => { setMediciones(next); persist("mediciones", next); };

  const upsertMedicion = (data) => {
    if (data.id && mediciones.some((m) => m.id === data.id)) {
      saveMediciones(mediciones.map((m) => (m.id === data.id ? { ...m, ...data } : m)));
      showToast("Medición actualizada");
    } else {
      const nueva = { ...data, id: data.id || uid() };
      saveMediciones([nueva, ...mediciones]);
      showToast("Medición creada");
    }
  };

  const deleteMedicion = (id) => {
    saveMediciones(mediciones.filter((m) => m.id !== id));
    showToast("Medición borrada");
  };

  // Al pulsar "Pasar a presupuesto" dentro de una medición: se abre el
  // formulario de Presupuestos ya con el cliente y una descripción con el
  // resumen de todo lo medido, para que solo falte poner el número y el importe.
  const pasarMedicionAPresupuesto = (medicion, resumenGlobal) => {
    saveMediciones(mediciones.map((m) => (m.id === medicion.id ? { ...m, presupuestoCreado: true } : m)));
    setPresupuestoPrefill({
      clienteNombre: medicion.clienteNombre,
      direccionEnvio: medicion.direccion || "",
      descripcion: `Medición realizada el ${medicion.fecha || ""} en ${medicion.direccion || "la obra"}:\n\n${resumenGlobal || "(Todavía no se han registrado elementos medidos en esta medición.)"}`,
      techos: medicion.techos || [],
    });
    setModulo("presupuestos");
    setPresupuestoEditId(null);
    setPresupuestoView("form");
  };

  // ---- Tareas (asignadas desde el chat interno o directamente) ----
  const saveTareas = (next) => { setTareas(next); persist("tareas", next); };

  const crearTarea = (data) => {
    const nueva = {
      id: uid(),
      titulo: data.titulo || "",
      descripcion: data.descripcion || "",
      asignadoA: data.asignadoA,
      asignadoANombre: data.asignadoANombre || "",
      asignadoPor: currentUser?.id || null,
      asignadoPorNombre: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "",
      requiereConfirmacion: !!data.requiereConfirmacion,
      fechaLimite: data.fechaLimite || "",
      salaId: data.salaId || null,
      salaNombre: data.salaNombre || "",
      estado: "Pendiente",
      fechaCreacion: Date.now(),
      fechaCompletada: null,
    };
    saveTareas([nueva, ...tareas]);
    showToast("Tarea creada");
  };

  const marcarTareaHecha = (id) => {
    saveTareas(tareas.map((t) => {
      if (t.id !== id) return t;
      return t.requiereConfirmacion
        ? { ...t, estado: "Pendiente de confirmar", fechaCompletada: Date.now() }
        : { ...t, estado: "Hecha", fechaCompletada: Date.now() };
    }));
    showToast("Tarea marcada como hecha");
  };

  const confirmarTarea = (id, ok) => {
    saveTareas(tareas.map((t) => (t.id === id ? { ...t, estado: ok ? "Hecha" : "Pendiente" } : t)));
    showToast(ok ? "Tarea confirmada" : "Tarea devuelta a pendiente");
  };

  const deleteTarea = (id) => {
    saveTareas(tareas.filter((t) => t.id !== id));
    showToast("Tarea borrada");
  };

  // ---- Archivos de empresa (tarifas, contactos, plantillas... no ligados a una obra) ----
  const saveArchivosEmpresa = (next) => { setArchivosEmpresa(next); persist("archivosEmpresa", next); };

  const subirArchivoEmpresa = (file, categoria) => {
    if (!file) return;
    if (file.size > 8 * 1024 * 1024) {
      showToast("El archivo pesa más de 8 MB, no se puede subir", "error");
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      const nuevo = {
        id: uid(), nombre: file.name, categoria: categoria || "Otros", url: e.target.result,
        subidoEn: Date.now(), subidoPor: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "",
      };
      saveArchivosEmpresa([nuevo, ...archivosEmpresa]);
      showToast("Archivo guardado");
    };
    reader.readAsDataURL(file);
  };

  const deleteArchivoEmpresa = (id) => {
    saveArchivosEmpresa(archivosEmpresa.filter((a) => a.id !== id));
    showToast("Archivo borrado");
  };

  const crearInstalacionParaProyecto = (proyecto) => {
    const nueva = {
      id: uid(), proyectoId: proyecto.id, estado: "Pendiente de instalación",
      presupuestoInstalacion: "", registroHoras: [], gastos: [], materialFurgoneta: [], notas: "",
      fechaMontaje: proyecto.fechaMontaje || "", vehiculoId: null,
    };
    setInstalaciones((prev) => {
      const updated = [nueva, ...prev];
      persist("instalaciones", updated);
      return updated;
    });
  };

  const updateInstalacion = (id, patch) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  };

  const crearInstalacionManual = (data) => {
    const nueva = {
      id: uid(), proyectoId: data.proyectoId || null, nombre: data.nombre || "", clienteNombre: data.clienteNombre || "",
      estado: "Pendiente de instalación", presupuestoInstalacion: data.presupuestoInstalacion || "",
      registroHoras: [], gastos: [], materialFurgoneta: [], notas: data.notas || "",
      fechaMontaje: "", vehiculoId: null,
    };
    saveInstalaciones([nueva, ...instalaciones]);
    showToast("Instalación creada");
    setInstalacionDetailId(nueva.id);
    setInstalacionView("detail");
  };

  const addMaterialFurgoneta = (instId, nombre) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === instId ? { ...i, materialFurgoneta: [{ id: uid(), nombre, estado: "Pendiente de cargar" }, ...(i.materialFurgoneta || [])] } : i)));
  };

  const cicloEstadoMaterialFurgoneta = (instId, itemId) => {
    const orden = ["Pendiente de cargar", "En la furgoneta", "En la obra"];
    saveInstalaciones(instalaciones.map((i) => {
      if (i.id !== instId) return i;
      return { ...i, materialFurgoneta: (i.materialFurgoneta || []).map((m) => m.id === itemId ? { ...m, estado: orden[(orden.indexOf(m.estado) + 1) % orden.length] } : m) };
    }));
  };

  const deleteMaterialFurgoneta = (instId, itemId) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === instId ? { ...i, materialFurgoneta: (i.materialFurgoneta || []).filter((m) => m.id !== itemId) } : i)));
  };

  const saveVehiculos = (next) => { setVehiculos(next); persist("vehiculos", next); };

  const upsertVehiculo = (data) => {
    let next;
    if (data.id) {
      next = vehiculos.map((v) => (v.id === data.id ? { ...v, ...data } : v));
    } else {
      next = [...vehiculos, { ...data, id: uid() }];
    }
    saveVehiculos(next);
  };

  const deleteVehiculo = (id) => {
    saveVehiculos(vehiculos.filter((v) => v.id !== id));
    saveInstalaciones(instalaciones.map((i) => (i.vehiculoId === id ? { ...i, vehiculoId: null } : i)));
  };

  const addHoraInstalacion = (id, registro) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === id ? { ...i, registroHoras: [{ id: uid(), ...registro }, ...(i.registroHoras || [])] } : i)));
    showToast("Horas registradas");
  };

  const deleteHoraInstalacion = (id, registroId) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === id ? { ...i, registroHoras: (i.registroHoras || []).filter((r) => r.id !== registroId) } : i)));
  };

  const addGastoInstalacion = (id, gasto) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === id ? { ...i, gastos: [{ id: uid(), ...gasto }, ...(i.gastos || [])] } : i)));
    showToast("Gasto registrado");
  };

  const deleteGastoInstalacion = (id, gastoId) => {
    saveInstalaciones(instalaciones.map((i) => (i.id === id ? { ...i, gastos: (i.gastos || []).filter((g) => g.id !== gastoId) } : i)));
  };

  const upsertProyecto = (data) => {
    let next;
    if (data.id) {
      const anterior = proyectos.find((p) => p.id === data.id);

      // Al aceptar el presupuesto por primera vez, sugerimos la fecha de fabricación
      // según los días de plazo de materiales indicados, si todavía no se ha puesto a mano.
      if (
        data.estadoPresupuesto === "Presupuesto aceptado" &&
        anterior?.estadoPresupuesto !== "Presupuesto aceptado" &&
        !data.fechaFabricacion && !anterior?.fechaFabricacion
      ) {
        const dias = parseInt(data.diasPlazoMateriales ?? anterior?.diasPlazoMateriales, 10) || 0;
        const fecha = new Date();
        fecha.setDate(fecha.getDate() + dias);
        data = { ...data, fechaFabricacion: fecha.toISOString().slice(0, 10) };
      }

      next = proyectos.map((p) => (p.id === data.id ? { ...p, ...data } : p));
      showToast("Proyecto actualizado");

      if (data.llevaInstalacion && !anterior?.llevaInstalacion && !instalaciones.some((i) => i.proyectoId === data.id)) {
        crearInstalacionParaProyecto({ ...anterior, ...data });
        showToast("Ficha de instalación creada — pendiente de instalación");
      }

      if (anterior?.trelloCardId && data.estadoPresupuesto && data.estadoPresupuesto !== anterior.estadoPresupuesto) {
        if (data.estadoPresupuesto === "Presupuesto aceptado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_PRESUPUESTO_ACEPTADO, "Presupuesto aceptado");
        } else if (data.estadoPresupuesto === "Presupuesto rechazado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_PRESUPUESTO_NO_ACEPTADO, "Presupuesto no aceptado");
        }
      }
      if (anterior?.trelloCardId && data.estadoTrabajo && data.estadoTrabajo !== anterior.estadoTrabajo) {
        if (data.estadoTrabajo === "Albarán de carga firmado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_ALBARAN_FIRMADO, "Albarán de carga firmado");
        } else if (data.estadoTrabajo === "Listo para reparto/recogida") {
          if (data.estadoLogistica === "Recogida en fábrica") {
            moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_RECOGIDA, "Recogida");
          } else if (data.estadoLogistica === "Reparto (camión)") {
            moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_REPARTO, "Reparto");
          } else {
            showToast('Elige "Reparto" o "Recogida" en el proyecto para que la tarjeta se mueva bien', "error");
          }
        } else if (data.estadoTrabajo === "Entregado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_TERMINADO, "Terminado");
        }
      }
    } else {
      const np = {
        ...data,
        id: uid(),
        numero: nextNumeroProyecto(),
        gastos: [],
        registroHorario: [],
        checklistMateriales: checklistMaterialesPorDefecto(),
      };
      next = [np, ...proyectos];
      showToast("Proyecto dado de alta");
      if (np.llevaInstalacion) {
        crearInstalacionParaProyecto(np);
      }
      const clienteNombre = clientes.find((c) => c.id === np.clienteId)?.nombre;
      crearTarjetaTrello(np, clienteNombre, (cardId) => {
        setProyectos((prev) => {
          const updated = prev.map((p) => (p.id === np.id ? { ...p, trelloCardId: cardId } : p));
          persist("proyectos", updated);
          return updated;
        });
      });
    }
    saveProyectos(next);
    setProyectoView("list");
  };

  const deleteProyecto = (id) => {
    saveProyectos(proyectos.filter((p) => p.id !== id));
    setProyectoView("list");
    showToast("Proyecto eliminado");
  };

  const updateProyectoInline = (id, patch) => {
    const next = proyectos.map((p) => (p.id === id ? { ...p, ...patch } : p));
    saveProyectos(next);
  };

  // Usa un artículo (de la lista de Artículos) en un proyecto: descuenta automáticamente
  // del stock cada material que compone ese artículo, multiplicado por la cantidad usada.
  const usarArticuloEnProyecto = (proyectoId, articuloId, cantidadUsada) => {
    const proyecto = proyectos.find((p) => p.id === proyectoId);
    const articulo = articulos.find((a) => a.id === articuloId);
    if (!proyecto || !articulo || cantidadUsada <= 0) return;
    const usoId = uid();
    const hoy = new Date().toISOString().slice(0, 10);

    const materialesNext = materiales.map((m) => {
      const lineas = (articulo.materiales || []).filter((l) => l.materialId === m.id);
      if (lineas.length === 0) return m;
      let stockReal = m.stockReal || 0;
      const nuevosMovimientos = lineas.map((l) => {
        const cantidad = (parseFloat(l.cantidad) || 0) * cantidadUsada;
        stockReal -= cantidad;
        return {
          id: uid(), tipo: "salida", cantidad, contacto: `Proyecto #${proyecto.numero} — ${articulo.nombre}`,
          estado: "Recibido", fecha: hoy, usoId,
        };
      });
      return { ...m, stockReal, movimientos: [...(m.movimientos || []), ...nuevosMovimientos] };
    });
    saveMateriales(materialesNext);

    const articulosUsados = [...(proyecto.articulosUsados || []), { id: usoId, articuloId, articuloNombre: articulo.nombre, cantidad: cantidadUsada, fecha: hoy }];
    updateProyectoInline(proyectoId, { articulosUsados });
    showToast(`${articulo.nombre} usado en el proyecto: stock descontado automáticamente`);
  };

  // Deshace el uso de un artículo: repone en stock los materiales que se habían descontado.
  const quitarArticuloUsado = (proyectoId, usoId) => {
    const proyecto = proyectos.find((p) => p.id === proyectoId);
    if (!proyecto) return;

    const materialesNext = materiales.map((m) => {
      const movs = (m.movimientos || []).filter((mv) => mv.usoId === usoId);
      if (movs.length === 0) return m;
      const totalRevertir = movs.reduce((s, mv) => s + (parseFloat(mv.cantidad) || 0), 0);
      return {
        ...m,
        stockReal: (m.stockReal || 0) + totalRevertir,
        movimientos: (m.movimientos || []).filter((mv) => mv.usoId !== usoId),
      };
    });
    saveMateriales(materialesNext);

    updateProyectoInline(proyectoId, { articulosUsados: (proyecto.articulosUsados || []).filter((u) => u.id !== usoId) });
    showToast("Uso del artículo eliminado: stock repuesto");
  };

  const saveProveedores = (next) => { setProveedores(next); persist("proveedores", next); };
  const saveMateriales = (next) => { setMateriales(next); persist("materiales", next); };

  // Importación masiva de tarifas de materiales: si el código o la descripción
  // ya existen, actualiza el precio; si no, da de alta el material nuevo.
  // Se hace todo en una sola escritura para no perder filas del lote.
  const importarTarifasMateriales = (filas) => {
    let working = [...materiales];
    let actualizados = 0, creados = 0;
    filas.forEach((fila) => {
      const idx = working.findIndex((m) =>
        (fila.codigo && m.codigo && m.codigo.toLowerCase() === String(fila.codigo).toLowerCase()) ||
        (fila.descripcion && m.descripcion && m.descripcion.toLowerCase() === String(fila.descripcion).toLowerCase())
      );
      if (idx >= 0) {
        working[idx] = {
          ...working[idx],
          ...(fila.precioVenta !== "" && fila.precioVenta != null ? { precioVenta: fila.precioVenta } : {}),
          ...(fila.precioCompra !== "" && fila.precioCompra != null ? { precioCompra: fila.precioCompra } : {}),
        };
        actualizados++;
      } else if (fila.codigo || fila.descripcion) {
        working = [{
          id: uid(), codigo: fila.codigo || "", descripcion: fila.descripcion || fila.codigo, proveedorId: "",
          stockReal: 0, stockMinimo: 0, stockOptimo: 0, color: "", acabadoDescripcion: "",
          longitud: "", ancho: "", alto: "", grueso: "",
          precioCompra: fila.precioCompra || "", precioVenta: fila.precioVenta || "", unidadCompra: "Unidad", categoria: "", familia: "",
        }, ...working];
        creados++;
      }
    });
    saveMateriales(working);
    showToast(`${actualizados} tarifa(s) actualizada(s), ${creados} material(es) nuevo(s)`);
  };

  const upsertProveedor = (data) => {
    let next;
    if (data.id) {
      next = proveedores.map((p) => (p.id === data.id ? data : p));
      showToast("Proveedor actualizado");
    } else {
      next = [{ ...data, id: uid(), historico: [] }, ...proveedores];
      showToast("Proveedor dado de alta");
    }
    saveProveedores(next);
    setProveedorView("list");
  };

  const deleteProveedor = (id) => {
    if (materiales.some((m) => m.proveedorId === id)) {
      showToast("No se puede eliminar: tiene materiales asociados", "error");
      return;
    }
    saveProveedores(proveedores.filter((p) => p.id !== id));
    setProveedorView("list");
    showToast("Proveedor eliminado");
  };

  const updateProveedorInline = (id, patch) => {
    saveProveedores(proveedores.map((p) => (p.id === id ? { ...p, ...patch } : p)));
  };

  const upsertMaterial = (data) => {
    let next;
    if (data.id) {
      next = materiales.map((m) => (m.id === data.id ? { ...m, ...data } : m));
      showToast("Material actualizado");
    } else {
      next = [{ ...data, id: uid(), movimientos: [] }, ...materiales];
      showToast("Material dado de alta");
    }
    saveMateriales(next);
    setMaterialView("list");
  };

  const deleteMaterial = (id) => {
    saveMateriales(materiales.filter((m) => m.id !== id));
    setMaterialView("list");
    showToast("Material eliminado");
  };

  const addMovimiento = (materialId, mov) => {
    const next = materiales.map((m) => {
      if (m.id !== materialId) return m;
      const movimientos = [...(m.movimientos || []), mov];
      let stockReal = m.stockReal || 0;
      if (mov.estado === "Recibido") {
        stockReal += mov.tipo === "entrada" ? mov.cantidad : -mov.cantidad;
      }
      return { ...m, movimientos, stockReal };
    });
    saveMateriales(next);
  };

  const removeMovimiento = (materialId, movId) => {
    const next = materiales.map((m) => {
      if (m.id !== materialId) return m;
      const mov = (m.movimientos || []).find((x) => x.id === movId);
      const movimientos = (m.movimientos || []).filter((x) => x.id !== movId);
      let stockReal = m.stockReal || 0;
      if (mov && mov.estado === "Recibido") {
        stockReal -= mov.tipo === "entrada" ? mov.cantidad : -mov.cantidad;
      }
      return { ...m, movimientos, stockReal };
    });
    saveMateriales(next);
  };

  const savePedidos = (next) => { setPedidos(next); persist("pedidos", next); };

  const nextNumeroPedido = () => {
    const nums = pedidos.map((p) => parseInt(String(p.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    const max = nums.length ? Math.max(...nums) : 599;
    return String(max + 1);
  };

  const upsertPedido = (data) => {
    let next;
    if (data.id) {
      next = pedidos.map((p) => (p.id === data.id ? { ...p, ...data } : p));
      showToast("Pedido actualizado");
    } else {
      if (data.proyectoId) {
        const proyectoDestino = proyectos.find((p) => p.id === data.proyectoId);
        const checklist = normalizarChecklist(proyectoDestino?.checklistMateriales);
        const incompleto = checklist.some((c) => !c.estado);
        if (incompleto) {
          showToast('No puedes crear el pedido: completa antes el checklist "Qué lleva la obra" en la ficha del proyecto (Sí/No en todas las categorías).', "error");
          return;
        }
      }
      next = [{ ...data, id: uid(), numero: nextNumeroPedido() }, ...pedidos];
      showToast("Pedido dado de alta");
      if (data.proyectoId) {
        const proyecto = proyectos.find((p) => p.id === data.proyectoId);
        if (proyecto?.trelloCardId) {
          moverTarjetaTrello(proyecto.trelloCardId, TRELLO_LIST_PEDIDO, "Pedido");
        }
      }
    }
    savePedidos(next);
    setPedidoView("list");
  };

  const deletePedido = (id) => {
    savePedidos(pedidos.filter((p) => p.id !== id));
    setPedidoView("list");
    showToast("Pedido eliminado");
  };

  // Marca un pedido como recibido: cada línea genera automáticamente un
  // movimiento de entrada en el material correspondiente y suma al stock real.
  const recibirPedido = (pedidoId) => {
    const pedido = pedidos.find((p) => p.id === pedidoId);
    if (!pedido) return;
    const proveedorNombre = proveedores.find((pr) => pr.id === pedido.proveedorId)?.nombre || "Proveedor";
    const hoy = new Date().toISOString().slice(0, 10);

    const materialesNext = materiales.map((m) => {
      const lineasMaterial = pedido.lineas.filter((l) => l.materialId === m.id && l.estado !== "Recibido");
      if (lineasMaterial.length === 0) return m;
      let stockReal = m.stockReal || 0;
      const nuevosMovimientos = lineasMaterial.map((l) => {
        stockReal += l.cantidad;
        return {
          id: uid(), tipo: "entrada", cantidad: l.cantidad, contacto: proveedorNombre,
          estado: "Recibido", fecha: hoy, pedidoNumero: pedido.numero,
        };
      });
      return { ...m, stockReal, movimientos: [...(m.movimientos || []), ...nuevosMovimientos] };
    });
    saveMateriales(materialesNext);

    const pedidosNext = pedidos.map((p) =>
      p.id === pedidoId
        ? { ...p, estado: "Recibido", fechaRecibido: hoy, lineas: p.lineas.map((l) => ({ ...l, estado: "Recibido" })) }
        : p
    );
    savePedidos(pedidosNext);
    showToast("Pedido recibido: stock actualizado automáticamente");

    if (pedido.proyectoId) {
      const proyecto = proyectos.find((p) => p.id === pedido.proyectoId);
      const pedidosDelProyecto = pedidosNext.filter((p) => p.proyectoId === pedido.proyectoId);
      const todosCerrados = pedidosDelProyecto.every((p) => p.estado === "Recibido" || p.estado === "Cancelado");
      const algunoRecibido = pedidosDelProyecto.some((p) => p.estado === "Recibido");
      if (proyecto?.trelloCardId && todosCerrados && algunoRecibido) {
        moverTarjetaTrello(proyecto.trelloCardId, TRELLO_LIST_FABRICA, "Fábrica");
      }
    }
  };

  // Guarda el resultado de la comprobación de un albarán en el pedido, SOLO cuando
  // el usuario confirma a mano — no toca el stock ni el estado del pedido.
  const confirmarAlbaranPedido = (pedidoId, resultado) => {
    savePedidos(pedidos.map((p) => (p.id === pedidoId ? { ...p, albaranComprobado: resultado } : p)));
    showToast("Comprobación de albarán guardada");
  };

  // Marca que el pedido ya se ha comunicado de verdad al proveedor — por email (al pulsar
  // "Enviar por email") o a mano (si se ha llamado/hablado de otra forma). Hasta que esto
  // no esté marcado, no se puede recibir el pedido — para no marcar como recibido algo
  // que en realidad nunca se llegó a pedir.
  const marcarPedidoEnviado = (pedidoId, metodo) => {
    savePedidos(pedidos.map((p) => (p.id === pedidoId ? { ...p, envioConfirmado: true, envioMetodo: metodo, fechaEnvioConfirmado: new Date().toISOString() } : p)));
    showToast(metodo === "email" ? "Marcado como enviado por email" : "Marcado como pedido realizado");
  };

  // Confirmación INDEPENDIENTE desde el almacén/fábrica de que un material ha llegado.
  // No toca el estado de oficina (pedido.estado / línea.estado) ni el stock — es un
  // segundo control aparte, pensado para comparar luego contra lo que dice oficina.
  const confirmarLineaFabrica = (pedidoId, lineaId, confirmado) => {
    const next = pedidos.map((p) => {
      if (p.id !== pedidoId) return p;
      return {
        ...p,
        lineas: p.lineas.map((l) => (l.id === lineaId ? {
          ...l,
          confirmadoFabrica: confirmado,
          fechaConfirmadoFabrica: confirmado ? new Date().toISOString().slice(0, 10) : "",
          confirmadoPorFabrica: confirmado ? (currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "Fábrica") : "",
        } : l)),
      };
    });
    savePedidos(next);
  };

  // ---------- Cristales (almacén de vidrio dentro de Fábrica) ----------
  const saveCristales = (next) => { setCristales(next); persist("cristales", next); };

  const addCristal = (data) => {
    const nuevo = { id: uid(), estado: "Pendiente", ubicacion: null, fechaColocado: "", fechaLlegada: new Date().toISOString().slice(0, 10), ...data };
    saveCristales([nuevo, ...cristales]);
    return nuevo.id;
  };

  const updateCristal = (id, patch) => {
    saveCristales(cristales.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  };

  const deleteCristal = (id) => {
    saveCristales(cristales.filter((c) => c.id !== id));
  };

  // Asigna una ubicación concreta (zona/fila/hueco) a un caballete, comprobando que
  // no esté ya ocupada por otro.
  const ubicarCristal = (id, ubicacion) => {
    const otros = cristales.filter((c) => c.id !== id && c.ubicacion &&
      c.ubicacion.zona === ubicacion.zona && c.ubicacion.fila === ubicacion.fila && c.ubicacion.hueco === ubicacion.hueco);
    updateCristal(id, { ubicacion, estado: "Colocado", fechaColocado: new Date().toISOString().slice(0, 10) });
    showToast(otros.length > 0 ? `Caballete añadido a un hueco que ya tenía ${otros.length} expediente(s)` : "Caballete ubicado en el almacén");
    return true;
  };

  const liberarCristal = (id) => {
    updateCristal(id, { ubicacion: null, estado: "Pendiente", fechaColocado: "" });
  };

  const saveIncidencias = (next) => { setIncidencias(next); persist("incidencias", next); };

  // Importa incidencias masivamente desde un Excel tipo "Control de reparaciones".
  // Cada fila necesita encajar con un cliente que ya exista y que tenga al
  // menos un proyecto (las incidencias van siempre ligadas a un proyecto).
  // Las filas que no encuentran coincidencia se reportan al final para
  // revisarlas a mano.
  const importarIncidenciasMasivo = (filas) => {
    const nums = incidencias.map((i) => parseInt(String(i.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    let siguienteNum = (nums.length ? Math.max(...nums) : 0) + 1;
    const normalizar = (s) => (s || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\(.*?\)/g, "").trim();

    const nuevas = [];
    const sinCoincidencia = [];
    filas.forEach((fila) => {
      const nombreBuscado = normalizar(fila.clienteTexto);
      if (!nombreBuscado) return;
      const clienteMatch = clientes.find((c) => {
        const n = normalizar(c.nombre);
        return n && (n.includes(nombreBuscado) || nombreBuscado.includes(n));
      });
      const proyectoMatch = clienteMatch ? proyectos.find((p) => p.clienteId === clienteMatch.id) : null;
      if (!proyectoMatch) {
        sinCoincidencia.push(fila.clienteTexto);
        return;
      }
      const especificaciones = [
        fila.tareas,
        fila.materiales ? `Materiales a pedir: ${fila.materiales}` : "",
        fila.dudas ? `Dudas / a confirmar: ${fila.dudas}` : "",
      ].filter(Boolean).join("\n\n");
      nuevas.push({
        id: uid(), numero: String(siguienteNum++), proyectoId: proyectoMatch.id,
        fecha: new Date().toISOString().slice(0, 10),
        especificaciones: especificaciones || "(Importado desde Excel sin descripción)",
        observaciones: [fila.telefono ? `Teléfono: ${fila.telefono}` : "", fila.notas || ""].filter(Boolean).join(" — "),
        responsableInicial: "", comercialAsociado: "", responsableActual: "",
        estadoTrabajo: "Pendiente revisión", estadoIncidencia: "Pendiente revisión",
        fechaEntregaPrevista: "", fechaEntregado: "",
      });
    });

    if (nuevas.length > 0) saveIncidencias([...nuevas, ...incidencias]);

    if (nuevas.length === 0 && sinCoincidencia.length === 0) {
      showToast("No se encontró ninguna fila válida en el Excel", "error");
    } else {
      let resumen = `${nuevas.length} incidencia(s) creada(s).`;
      if (sinCoincidencia.length > 0) {
        resumen += `\n\n${sinCoincidencia.length} fila(s) NO se pudieron importar porque no encontré un cliente con proyecto que coincida:\n${sinCoincidencia.slice(0, 15).join("\n")}${sinCoincidencia.length > 15 ? "\n..." : ""}\n\nDa de alta a esos clientes (con al menos un proyecto) y vuelve a intentarlo con esas filas.`;
      }
      alert(resumen);
    }
  };

  const nextNumeroIncidencia = () => {
    const nums = incidencias.map((i) => parseInt(String(i.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    const max = nums.length ? Math.max(...nums) : 0;
    return String(max + 1);
  };

  const upsertIncidencia = (data) => {
    let next;
    if (data.id) {
      next = incidencias.map((i) => (i.id === data.id ? { ...i, ...data } : i));
      showToast("Incidencia actualizada");
    } else {
      next = [{ ...data, id: uid(), numero: nextNumeroIncidencia(), gastos: [] }, ...incidencias];
      showToast("Incidencia dada de alta");
      if (data.proyectoId) {
        const proyecto = proyectos.find((p) => p.id === data.proyectoId);
        if (proyecto?.trelloCardId) {
          moverTarjetaTrello(proyecto.trelloCardId, TRELLO_LIST_INCIDENCIA, "Incidencia");
        }
      }
    }
    saveIncidencias(next);
    setIncidenciaView("list");
  };

  const deleteIncidencia = (id) => {
    saveIncidencias(incidencias.filter((i) => i.id !== id));
    setIncidenciaView("list");
    showToast("Incidencia eliminada");
  };

  const updateIncidenciaInline = (id, patch) => {
    saveIncidencias(incidencias.map((i) => (i.id === id ? { ...i, ...patch } : i)));
  };

  const openProyectoFromCalendar = (id) => { setModulo("proyectos"); setProyectoDetailId(id); setProyectoView("detail"); };
  const openPedidoFromCalendar = (id) => { setModulo("pedidos"); setPedidoDetailId(id); setPedidoView("detail"); };
  // Cuando el usuario toca el aviso de "han llegado" pedidos, los marcamos como
  // vistos para que no se lo sigamos recordando cada vez que entra al CRM.
  const marcarAvisosPedidosVistos = (lista) => {
    if (!lista || lista.length === 0) return;
    const ids = new Set(lista.map((p) => p.id));
    savePedidos(pedidos.map((p) => (ids.has(p.id) ? { ...p, avisoLlegadaVisto: true } : p)));
  };

  const enviarAPedido = (proveedorId, lineas, proyectoId, comentarios) => {
    setPedidoPrefill({ proveedorId, lineas, proyectoId, comentarios });
    setPedidoEditId(null);
    setPedidoView("form");
    setModulo("pedidos");
  };

  // Cuando el albarán no cuadra con lo pedido, esto crea (y abre listo para
  // revisar y enviar) un pedido nuevo al mismo proveedor solo con lo que falta.
  const crearPedidoFaltanteDesdeAlbaran = (pedidoOriginal, lineasFaltantes) => {
    if (!lineasFaltantes || lineasFaltantes.length === 0) {
      showToast("No hay líneas pendientes que pedir", "error");
      return;
    }
    enviarAPedido(
      pedidoOriginal.proveedorId,
      lineasFaltantes,
      pedidoOriginal.proyectoId,
      `Pedido de reposición: esto es lo que faltó (o no coincidió) en el albarán del pedido #${pedidoOriginal.numero}.`
    );
    showToast("Pedido con lo que falta creado — revísalo y envíalo");
  };
  const irARegistrarIngreso = (proyecto, clienteNombre, importeSugerido) => {
    setIngresoPrefill({
      id: null, fecha: new Date().toISOString().slice(0, 10), clienteNombre: clienteNombre || "",
      importe: importeSugerido ? importeSugerido.toFixed(2) : "", concepto: "Resto del presupuesto",
      formaPago: "Transferencia", proyectoId: proyecto.id, notas: "",
    });
    setIngresoEditId(null);
    setIngresoView("form");
    setModulo("ingresos");
  };
  const irASolicitarArticulo = (articuloNombre) => {
    setSolicitudPrefill({
      id: null, proyectoId: "", tipo: "Artículo nuevo para catálogo",
      descripcion: articuloNombre ? `Solicito dar de alta un artículo nuevo: ${articuloNombre}` : "",
      urgencia: "Normal",
    });
    setSolicitudPedidoEditId(null);
    setSolicitudPedidoView("form");
    setPedidoTabPrincipal("solicitudes");
    setModulo("pedidos");
  };
  const openIncidenciaFromCalendar = (id) => { setModulo("incidencias"); setIncidenciaDetailId(id); setIncidenciaView("detail"); };
  const cambiarFechaProyecto = (id, campo, nuevaFecha) => {
    saveProyectos(proyectos.map((p) => (p.id === id ? { ...p, [campo]: nuevaFecha } : p)));
    showToast("Fecha actualizada");
  };
  const cambiarFechaPedido = (id, campo, nuevaFecha) => {
    savePedidos(pedidos.map((p) => (p.id === id ? { ...p, [campo]: nuevaFecha } : p)));
    showToast("Fecha actualizada");
  };
  const cambiarFechaIncidencia = (id, campo, nuevaFecha) => {
    saveIncidencias(incidencias.map((i) => (i.id === id ? { ...i, [campo]: nuevaFecha } : i)));
    showToast("Fecha actualizada");
  };
  const irAInstalacion = (instalacionId) => { setModulo("instalaciones"); setInstalacionDetailId(instalacionId); setInstalacionView("detail"); };

  const saveArticulos = (next) => { setArticulos(next); persist("articulos", next); };

  const importarTarifasArticulos = (filas) => {
    let working = [...articulos];
    let actualizados = 0, creados = 0;
    filas.forEach((fila) => {
      const idx = working.findIndex((a) => fila.nombre && a.nombre && a.nombre.toLowerCase() === String(fila.nombre).toLowerCase());
      if (idx >= 0) {
        working[idx] = { ...working[idx], ...(fila.precioVenta !== "" && fila.precioVenta != null ? { precioVenta: fila.precioVenta } : {}) };
        actualizados++;
      } else if (fila.nombre) {
        working = [{
          id: uid(), nombre: fila.nombre, descripcion: "", proveedorId: "", precioVenta: fila.precioVenta || "",
          materiales: [], fases: [], tamano: "", medidas: "", volumen: "", importeEnvio: "", importeMontaje: "",
        }, ...working];
        creados++;
      }
    });
    saveArticulos(working);
    showToast(`${actualizados} tarifa(s) actualizada(s), ${creados} artículo(s) nuevo(s)`);
  };

  const nextNumeroArticulo = () => {
    const nums = articulos.map((a) => parseInt(String(a.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    const max = nums.length ? Math.max(...nums) : 0;
    return String(max + 1);
  };

  const upsertArticulo = (data) => {
    let next;
    if (data.id) {
      next = articulos.map((a) => (a.id === data.id ? { ...a, ...data } : a));
      showToast("Artículo actualizado");
    } else {
      next = [{ ...data, id: uid(), numero: nextNumeroArticulo(), gastos: [] }, ...articulos];
      showToast("Artículo dado de alta");
    }
    saveArticulos(next);
    setArticuloView("list");
  };

  const deleteArticulo = (id) => {
    saveArticulos(articulos.filter((a) => a.id !== id));
    setArticuloView("list");
    showToast("Artículo eliminado");
  };

  const updateArticuloInline = (id, patch) => {
    saveArticulos(articulos.map((a) => (a.id === id ? { ...a, ...patch } : a)));
  };

  const saveFacturas = (next) => { setFacturas(next); persist("facturas", next); };

  const nextNumeroFactura = () => {
    const nums = facturas.map((f) => parseInt(String(f.numero).replace(/\D/g, ""), 10)).filter((n) => !isNaN(n));
    const max = nums.length ? Math.max(...nums) : 100;
    return String(max + 1);
  };

  const upsertFactura = (data) => {
    let next;
    if (data.id) {
      next = facturas.map((f) => (f.id === data.id ? { ...f, ...data } : f));
      showToast("Factura actualizada");
    } else {
      next = [{ ...data, id: uid(), numero: nextNumeroFactura(), pagos: [] }, ...facturas];
      showToast("Factura emitida");
    }
    saveFacturas(next);
    setFacturaView("list");
  };

  const deleteFactura = (id) => {
    saveFacturas(facturas.filter((f) => f.id !== id));
    setFacturaView("list");
    showToast("Factura eliminada");
  };

  const addPago = (facturaId, pago) => {
    const next = facturas.map((f) => (f.id === facturaId ? { ...f, pagos: [...(f.pagos || []), pago] } : f));
    saveFacturas(next);
  };
  const removePago = (facturaId, pagoId) => {
    const next = facturas.map((f) => (f.id === facturaId ? { ...f, pagos: (f.pagos || []).filter((p) => p.id !== pagoId) } : f));
    saveFacturas(next);
  };

  // Registra un pago de un proyecto. Cada ingreso genera su propia factura por el
  // importe exacto recibido (no por el total del presupuesto), quedando ya pagada.
  const registrarPagoProyecto = (proyectoId, { importe, fecha, formaPago, tipo }) => {
    const proyecto = proyectos.find((p) => p.id === proyectoId);
    if (!proyecto) return;
    const importeNum = parseFloat(importe) || 0;
    const pago = { id: uid(), importe: importeNum, fecha, formaPago };
    const nueva = {
      id: uid(),
      numero: nextNumeroFactura(),
      clienteId: proyecto.clienteId,
      tipo: tipo || "Definitiva",
      fecha,
      proyectosIds: [proyectoId],
      total: importeNum,
      pagos: [pago],
    };
    saveFacturas([nueva, ...facturas]);
    showToast(`Factura ${nueva.numero} generada por ${money(importeNum)}`);
    if (proyecto.trelloCardId) {
      moverTarjetaTrello(proyecto.trelloCardId, TRELLO_LIST_FACTURADO, "Facturado");
    }
  };

  const savePresupuestos = (next) => { setPresupuestos(next); persist("presupuestos", next); };

  const upsertPresupuesto = (data) => {
    let next;
    if (data.id) {
      const anterior = presupuestos.find((p) => p.id === data.id);
      const dataNormalizada = { ...data, estado: data.estado || anterior?.estado || "Pendiente" };
      next = presupuestos.map((p) => (p.id === data.id ? { ...p, ...dataNormalizada } : p));
      showToast("Presupuesto actualizado");

      if (anterior?.trelloCardId && dataNormalizada.estado && dataNormalizada.estado !== anterior.estado) {
        if (dataNormalizada.estado === "Aceptado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_PRESUPUESTO_ACEPTADO, "Presupuesto aceptado");
        } else if (dataNormalizada.estado === "Rechazado") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_PRESUPUESTO_NO_ACEPTADO, "Presupuesto no aceptado");
        } else if (dataNormalizada.estado === "Pendiente" || dataNormalizada.estado === "En espera") {
          moverTarjetaTrello(anterior.trelloCardId, TRELLO_LIST_PRESUPUESTO, "Presupuesto");
        }
      }
    } else {
      const np = { estado: "Pendiente", ...data, id: uid() };
      next = [np, ...presupuestos];
      showToast("Presupuesto dado de alta");
      crearTarjetaTrelloPresupuesto(np, (cardId) => {
        setPresupuestos((prev) => {
          const updated = prev.map((p) => (p.id === np.id ? { ...p, trelloCardId: cardId } : p));
          persist("presupuestos", updated);
          return updated;
        });
      });
    }
    savePresupuestos(next);
    setPresupuestoView("list");
  };

  const deletePresupuesto = (id) => {
    savePresupuestos(presupuestos.filter((p) => p.id !== id));
    setPresupuestoView("list");
    showToast("Presupuesto eliminado");
  };

  // Crea un Proyecto a partir de un Presupuesto ya aceptado, REUTILIZANDO la misma
  // tarjeta de Trello (no se crea una tarjeta nueva) para que el hilo no se corte.
  const crearProyectoDesdePresupuesto = (presupuesto) => {
    let clienteId = clientes.find((c) => c.nombre.trim().toLowerCase() === (presupuesto.clienteNombre || "").trim().toLowerCase())?.id || "";
    const np = {
      id: uid(),
      numero: nextNumeroProyecto(),
      nombre: presupuesto.descripcion || presupuesto.numero,
      clienteId,
      importePresupuesto: presupuesto.importe || 0,
      estadoPresupuesto: "Presupuesto aceptado",
      estadoTrabajo: "Pendiente de aceptación",
      trelloCardId: presupuesto.trelloCardId || null,
      gastos: [],
      registroHorario: [],
      checklistMateriales: checklistMaterialesPorDefecto(),
      techos: presupuesto.techos || [],
    };
    saveProyectos([np, ...proyectos]);
    savePresupuestos(presupuestos.map((p) => (p.id === presupuesto.id ? { ...p, proyectoCreadoId: np.id } : p)));
    showToast(`Proyecto #${np.numero} creado desde el presupuesto (misma tarjeta de Trello)`);
    return np.id;
  };

  // Crea una nueva versión del mismo presupuesto (ej. 4192 -> 4192.1 -> 4192.2),
  // copiando los datos del cliente pero dejándolo listo para ajustar precio/descripción.
  // Guarda un "punto de guardado" del presupuesto actual en su propio historial de
  // versiones (mismo registro, sin duplicar nada en el listado), y lo deja abierto
  // para editar el importe/descripción de la nueva revisión.
  // Crea una réplica del presupuesto (mismo número base + guion + siguiente número:
  // 4192 -> 4192-1 -> 4192-2), como una fila más en el listado, con su propia
  // tarjeta de Trello, lista para abrir y modificar.
  const duplicarPresupuesto = (presupuesto) => {
    const numeroBase = String(presupuesto.numero).split("-")[0].trim();
    const sufijos = presupuestos
      .filter((p) => String(p.numero).split("-")[0].trim() === numeroBase)
      .map((p) => {
        const partes = String(p.numero).split("-");
        return partes.length > 1 ? parseInt(partes[1], 10) || 0 : 0;
      });
    const siguienteSufijo = Math.max(0, ...sufijos) + 1;
    const nuevoNumero = `${numeroBase}-${siguienteSufijo}`;

    const nueva = {
      id: uid(),
      numero: nuevoNumero,
      fechaEnvio: new Date().toISOString().slice(0, 10),
      clienteNombre: presupuesto.clienteNombre,
      telefono: presupuesto.telefono || "",
      descripcion: presupuesto.descripcion || "",
      importe: presupuesto.importe || "",
      estado: "Pendiente",
      motivoRechazo: "",
      fechaRespuesta: "",
      comentarios: `Réplica del presupuesto ${presupuesto.numero}.`,
      fechaPrevistaConfirmacion: "",
      envio: presupuesto.envio || false,
      direccionEnvio: presupuesto.direccionEnvio || "",
      montaje: presupuesto.montaje || false,
      zona: presupuesto.zona || "",
      llamadas: [],
    };
    savePresupuestos([nueva, ...presupuestos]);
    crearTarjetaTrelloPresupuesto(nueva, (cardId) => {
      setPresupuestos((prev) => {
        const updated = prev.map((p) => (p.id === nueva.id ? { ...p, trelloCardId: cardId } : p));
        persist("presupuestos", updated);
        return updated;
      });
    });
    showToast(`Réplica creada: ${nuevoNumero} — ya la puedes modificar`);
    return nueva.id;
  };

  const addLlamadaPresupuesto = (presupuestoId, llamada) => {
    const next = presupuestos.map((p) => (p.id === presupuestoId ? { ...p, llamadas: [llamada, ...(p.llamadas || [])] } : p));
    savePresupuestos(next);
    showToast("Llamada registrada");
  };

  const deleteLlamadaPresupuesto = (presupuestoId, llamadaId) => {
    const next = presupuestos.map((p) => (p.id === presupuestoId ? { ...p, llamadas: (p.llamadas || []).filter((l) => l.id !== llamadaId) } : p));
    savePresupuestos(next);
  };

  const saveIngresos = (next) => { setIngresos(next); persist("ingresos", next); };

  const upsertIngreso = (data) => {
    let next;
    if (data.id) {
      const anterior = ingresos.find((i) => i.id === data.id);
      let actualizado = { ...anterior, ...data };
      if (!anterior?.proyectoId && data.proyectoId) {
        registrarPagoProyecto(data.proyectoId, { importe: actualizado.importe, fecha: actualizado.fecha, formaPago: actualizado.formaPago, tipo: "Anticipo" });
        actualizado.vinculado = true;
        showToast("Entrada vinculada al proyecto y factura generada");
      } else {
        showToast("Entrada de dinero actualizada");
      }
      next = ingresos.map((i) => (i.id === data.id ? actualizado : i));
    } else {
      const nuevo = { proyectoId: null, vinculado: false, ...data, id: uid() };
      if (nuevo.proyectoId) {
        registrarPagoProyecto(nuevo.proyectoId, { importe: nuevo.importe, fecha: nuevo.fecha, formaPago: nuevo.formaPago, tipo: "Anticipo" });
        nuevo.vinculado = true;
      }
      next = [nuevo, ...ingresos];
      showToast(nuevo.proyectoId ? "Entrada registrada y factura generada" : "Entrada de dinero registrada");
    }
    saveIngresos(next);
    setIngresoView("list");
  };

  const deleteIngreso = (id) => {
    saveIngresos(ingresos.filter((i) => i.id !== id));
    showToast("Entrada de dinero eliminada");
  };

  // Vincula una entrada de dinero (recibida antes de crear el proyecto) a un proyecto
  // ya existente. Genera además la factura correspondiente, igual que un pago normal.
  const vincularIngresoAProyecto = (ingresoId, proyectoId) => {
    const ingreso = ingresos.find((i) => i.id === ingresoId);
    if (!ingreso || !proyectoId) return;
    registrarPagoProyecto(proyectoId, { importe: ingreso.importe, fecha: ingreso.fecha, formaPago: ingreso.formaPago, tipo: "Anticipo" });
    saveIngresos(ingresos.map((i) => (i.id === ingresoId ? { ...i, proyectoId, vinculado: true } : i)));
    showToast("Entrada vinculada al proyecto y factura generada");
  };

  const saveSolicitudesPedido = (next) => { setSolicitudesPedido(next); persist("solicitudes_pedido", next); };

  const upsertSolicitudPedido = (data) => {
    let next;
    if (data.id) {
      next = solicitudesPedido.map((s) => (s.id === data.id ? { ...s, ...data } : s));
      showToast("Solicitud actualizada");
    } else {
      const nueva = {
        id: uid(), estado: "Pendiente", comentarios: [], motivoRechazo: "",
        solicitanteId: currentUser?.id || null,
        solicitanteNombre: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "Desconocido",
        fecha: new Date().toISOString().slice(0, 10),
        ...data,
      };
      next = [nueva, ...solicitudesPedido];
      showToast("Solicitud de pedido enviada");
    }
    saveSolicitudesPedido(next);
    setSolicitudPedidoView("list");
  };

  const deleteSolicitudPedido = (id) => {
    saveSolicitudesPedido(solicitudesPedido.filter((s) => s.id !== id));
    setSolicitudPedidoView("list");
    showToast("Solicitud eliminada");
  };

  const addComentarioSolicitud = (id, texto) => {
    if (!texto.trim()) return;
    const comentario = {
      id: uid(), texto, fecha: new Date().toISOString(),
      autor: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "Desconocido",
    };
    saveSolicitudesPedido(solicitudesPedido.map((s) => (s.id === id ? { ...s, comentarios: [...(s.comentarios || []), comentario] } : s)));
  };

  const rechazarSolicitudPedido = (id, motivo) => {
    saveSolicitudesPedido(solicitudesPedido.map((s) => (s.id === id ? { ...s, estado: "Rechazada", motivoRechazo: motivo } : s)));
    showToast("Solicitud rechazada");
  };

  // Aprueba la solicitud y lleva al gestor directamente a crear el pedido real,
  // con el proyecto y la descripción ya pre-rellenados (reutiliza pedidoPrefill).
  const aprobarSolicitudPedido = (solicitud) => {
    saveSolicitudesPedido(solicitudesPedido.map((s) => (s.id === solicitud.id ? { ...s, estado: "Aprobada" } : s)));
    setPedidoPrefill({
      proyectoId: solicitud.proyectoId || "",
      comentarios: `Solicitado por ${solicitud.solicitanteNombre} el ${fmtDate(solicitud.fecha)}: ${solicitud.descripcion}`,
      lineas: [],
    });
    setPedidoEditId(null);
    setPedidoView("form");
    setModulo("pedidos");
    showToast("Solicitud aprobada — completa ahora el pedido real");
  };

  const saveFichajes = (next) => { setFichajes(next); persist("fichajes", next); };

  const setEmpleadoActualPersist = (nombre) => {
    setEmpleadoActual(nombre);
    localStorage.setItem("alumavel_empleado_actual", nombre);
  };

  const registrarFichaje = (empleado, accion) => {
    const now = new Date();
    const hoy = now.toISOString().slice(0, 10);
    const hora = now.toISOString();
    let registro = fichajes.find((f) => f.empleado === empleado && f.fecha === hoy);

    if (accion === "entrada") {
      if (registro) { showToast("Ya existe una entrada hoy para este empleado", "error"); return; }
      registro = { id: uid(), empleado, fecha: hoy, entrada: hora, descansos: [], salida: null, estado: "trabajando" };
      saveFichajes([registro, ...fichajes]);
      showToast("Entrada registrada");
      return;
    }
    if (!registro) { showToast("No hay una entrada registrada hoy", "error"); return; }

    let next = { ...registro };
    if (accion === "inicio_descanso") {
      if (next.estado !== "trabajando") { showToast("No se puede iniciar descanso ahora", "error"); return; }
      next.descansos = [...next.descansos, { id: uid(), inicio: hora, fin: null }];
      next.estado = "descanso";
      showToast("Descanso iniciado");
    } else if (accion === "fin_descanso") {
      if (next.estado !== "descanso") { showToast("No hay un descanso en curso", "error"); return; }
      next.descansos = next.descansos.map((d, i) => (i === next.descansos.length - 1 ? { ...d, fin: hora } : d));
      next.estado = "trabajando";
      showToast("Descanso finalizado");
    } else if (accion === "salida") {
      if (next.estado !== "trabajando") { showToast("No se puede fichar salida ahora", "error"); return; }
      next.salida = hora;
      next.estado = "finalizado";
      showToast("Salida registrada");
    }
    saveFichajes(fichajes.map((f) => (f.id === registro.id ? next : f)));
  };

  const deleteFichaje = (id) => {
    saveFichajes(fichajes.filter((f) => f.id !== id));
    showToast("Fichaje eliminado");
  };

  const saveUsuarios = (next) => { setUsuarios(next); persist("usuarios", next); };

  const upsertUsuario = (data) => {
    let next;
    if (data.id) {
      next = usuarios.map((u) => (u.id === data.id ? { ...u, ...data } : u));
      showToast("Usuario actualizado");
    } else {
      if (usuarios.some((u) => u.email.toLowerCase() === data.email.toLowerCase())) {
        showToast("Ya existe un usuario con ese email", "error");
        return;
      }
      next = [{ ...data, id: uid() }, ...usuarios];
      showToast("Usuario dado de alta");
    }
    saveUsuarios(next);
  };

  const deleteUsuario = (id) => {
    if (id === sesionUsuarioId) { showToast("No puedes eliminar tu propio usuario", "error"); return; }
    saveUsuarios(usuarios.filter((u) => u.id !== id));
    showToast("Usuario eliminado");
  };

  const iniciarSesion = (id) => {
    setSesionUsuarioId(id);
    localStorage.setItem("alumavel_sesion_usuario_id", id);
  };
  const cerrarSesion = () => {
    setSesionUsuarioId(null);
    localStorage.removeItem("alumavel_sesion_usuario_id");
  };

  const currentUser = usuarios.find((u) => u.id === sesionUsuarioId) || null;

  // Seguimiento de tiempo conectado por usuario: mientras haya un usuario con
  // sesión iniciada (currentUser), guardamos una "sesión" en Firebase con hora
  // de inicio y hora de fin. La hora de fin se va actualizando cada 2 minutos
  // (y al cerrar la pestaña/sesión) para que, aunque cierren el navegador sin
  // avisar, quede un fin aproximado y no una sesión "infinita".
  const sesionActivaIdRef = useRef(null);
  useEffect(() => {
    if (!currentUser) return;
    const sessionId = uid();
    sesionActivaIdRef.current = sessionId;
    const path = `sesionesUsuario/${currentUser.id}/${sessionId}`;
    const inicio = Date.now();
    fbSet(ref(fbDb, path), { inicio, fin: inicio }).catch(() => {});

    const actualizarFin = () => {
      fbSet(ref(fbDb, `${path}/fin`), Date.now()).catch(() => {});
    };
    const intervalo = setInterval(actualizarFin, 120000);
    const onVisibilidad = () => { if (document.visibilityState === "hidden") actualizarFin(); };
    window.addEventListener("beforeunload", actualizarFin);
    document.addEventListener("visibilitychange", onVisibilidad);

    return () => {
      clearInterval(intervalo);
      actualizarFin();
      window.removeEventListener("beforeunload", actualizarFin);
      document.removeEventListener("visibilitychange", onVisibilidad);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [currentUser?.id]);

  const isAdmin = currentUser?.rol === "Administrador";
  const modulosPermitidos = isAdmin
    ? MODULOS_DISPONIBLES.map((m) => m.id)
    : (currentUser?.modulos || MODULOS_DISPONIBLES.map((m) => m.id));
  const tieneAcceso = (id) => modulosPermitidos.includes(id);
  const misTareasPendientes = tareas.filter((t) =>
    (t.asignadoA === currentUser?.id && t.estado === "Pendiente") ||
    (t.asignadoPor === currentUser?.id && t.estado === "Pendiente de confirmar")
  ).length;

  useEffect(() => {
    if (currentUser && modulo !== "administracion" && !tieneAcceso(modulo)) {
      setModulo(modulosPermitidos[0] || "proyectos");
    }
  }, [currentUser, modulo]);

  const iniciarSesionCliente = (id) => {
    setSesionClienteId(id);
    localStorage.setItem("alumavel_sesion_cliente_id", id);
  };
  const cerrarSesionCliente = () => {
    setSesionClienteId(null);
    localStorage.removeItem("alumavel_sesion_cliente_id");
  };
  const currentCliente = clientes.find((c) => c.id === sesionClienteId) || null;

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-[#F4F5F3]">
        <div className="flex items-center gap-2 text-slate-500 text-sm">
          <Loader2 className="animate-spin" size={18} /> Cargando CRM…
        </div>
      </div>
    );
  }

  if (!currentUser && !currentCliente) {
    return (
      <LoginGate
        usuarios={usuarios}
        clientes={clientes}
        onCreateFirstAdmin={(data) => { const u = { ...data, id: uid(), rol: "Administrador" }; saveUsuarios([u, ...usuarios]); iniciarSesion(u.id); }}
        onLogin={(id) => iniciarSesion(id)}
        onLoginCliente={(id) => iniciarSesionCliente(id)}
      />
    );
  }

  if (!currentUser && currentCliente) {
    return (
      <ClientePortal
        cliente={currentCliente}
        proyectos={proyectos.filter((p) => p.clienteId === currentCliente.id)}
        facturas={facturas.filter((f) => f.clienteId === currentCliente.id)}
        incidencias={incidencias.filter((i) => proyectos.find((p) => p.id === i.proyectoId && p.clienteId === currentCliente.id))}
        onLogout={cerrarSesionCliente}
      />
    );
  }

  return (
    <div className="h-screen bg-[#F4F5F3] flex text-slate-800 overflow-hidden" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .font-mono-num{ font-family:'IBM Plex Mono', monospace; }
        .font-display{ font-family:'Inter', system-ui, sans-serif; letter-spacing:-0.02em; }
      `}</style>

      {/* Fondo oscuro al abrir el menú en móvil */}
      {menuMovilAbierto && (
        <div className="fixed inset-0 bg-black/40 z-30 md:hidden" onClick={() => setMenuMovilAbierto(false)} />
      )}

      {/* Sidebar */}
      <aside className={`w-60 shrink-0 bg-[#2A1F3D] text-slate-200 flex flex-col h-full overflow-y-auto fixed md:static inset-y-0 left-0 z-40 transform transition-transform duration-200 ${menuMovilAbierto ? "translate-x-0" : "-translate-x-full"} md:translate-x-0`}>
        <div className="px-5 py-5 border-b border-white/10">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded bg-[#2E8B57] flex items-center justify-center font-mono-num font-bold text-sm">A</div>
            <div>
              <div className="font-display font-bold text-[15px] leading-none">ALUMAVEL</div>
              <div className="text-[10px] tracking-[0.15em] uppercase text-slate-400 mt-1">Panel de gestión</div>
            </div>
          </div>
        </div>
        <nav className="flex-1 px-3 py-4 space-y-1" onClick={() => setMenuMovilAbierto(false)}>
          {tieneAcceso("clientes") && (
          <button
            onClick={() => setModulo("clientes")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "clientes" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Users size={16} /> Clientes
          </button>
          )}
          {tieneAcceso("proyectos") && (
          <button
            onClick={() => setModulo("proyectos")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "proyectos" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Briefcase size={16} /> Proyectos / Obras
          </button>
          )}
          {tieneAcceso("mediciones") && (
          <button
            onClick={() => setModulo("mediciones")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "mediciones" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Ruler size={16} /> Mediciones
          </button>
          )}
          {tieneAcceso("presupuestos") && (
          <button
            onClick={() => setModulo("presupuestos")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "presupuestos" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <FileSpreadsheet size={16} /> Presupuestos
          </button>
          )}
          {tieneAcceso("instalaciones") && (
          <button
            onClick={() => setModulo("instalaciones")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "instalaciones" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Wrench size={16} /> Instalaciones
          </button>
          )}
          {tieneAcceso("proveedores") && (
          <button
            onClick={() => setModulo("proveedores")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "proveedores" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Truck size={16} /> Proveedores
          </button>
          )}
          {tieneAcceso("pedidos") && (
          <button
            onClick={() => setModulo("pedidos")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "pedidos" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <ClipboardList size={16} /> Pedidos
          </button>
          )}
          {tieneAcceso("stock") && (
          <button
            onClick={() => setModulo("stock")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "stock" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Boxes size={16} /> Gestión de Stock
          </button>
          )}
          {tieneAcceso("articulos") && (
          <button
            onClick={() => setModulo("articulos")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "articulos" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Layers size={16} /> Artículos
            <Badge className="ml-auto bg-white/10 text-slate-300 ring-white/10 !py-0">Fase II</Badge>
          </button>
          )}
          {tieneAcceso("incidencias") && (
          <button
            onClick={() => setModulo("incidencias")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "incidencias" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <AlertOctagon size={16} /> Incidencias
          </button>
          )}
          {tieneAcceso("calendario") && (
          <button
            onClick={() => setModulo("calendario")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "calendario" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <CalendarDays size={16} /> Calendario
          </button>
          )}
          {tieneAcceso("ingresos") && (
          <button
            onClick={() => setModulo("ingresos")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "ingresos" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Wallet size={16} /> Entrada de dinero
          </button>
          )}
          {tieneAcceso("facturas") && (
          <button
            onClick={() => setModulo("facturas")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "facturas" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Receipt size={16} /> Facturas
          </button>
          )}
          {tieneAcceso("fabrica") && (
          <button
            onClick={() => setModulo("fabrica")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "fabrica" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Factory size={16} /> Fábrica
          </button>
          )}
          {tieneAcceso("informes") && (
          <button
            onClick={() => setModulo("informes")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "informes" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <BarChart3 size={16} /> Informes
          </button>
          )}
          {tieneAcceso("fichajes") && (
          <button
            onClick={() => setModulo("fichajes")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "fichajes" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <Timer size={16} /> Fichajes
          </button>
          )}
          <div className="pt-2 mt-2 border-t border-white/10" />
          {isAdmin && (
            <>
              <button
                onClick={() => setModulo("administracion")}
                className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
                  modulo === "administracion" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
                }`}
              >
                <UserCog size={16} /> Administración
              </button>
            </>
          )}
          <div className="pt-2 mt-2 border-t border-white/10" />
          {tieneAcceso("chat") && (
          <button
            onClick={() => setModulo("chat")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "chat" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <MessageCircle size={16} /> Chat
          </button>
          )}
          {tieneAcceso("tareas") && (
          <button
            onClick={() => setModulo("tareas")}
            className={`w-full flex items-center gap-2.5 px-3 py-2.5 rounded-md text-sm font-medium transition ${
              modulo === "tareas" ? "bg-[#2E8B57] text-white" : "text-slate-300 hover:bg-white/5"
            }`}
          >
            <ClipboardList size={16} /> Tareas
            {misTareasPendientes > 0 && (
              <span className="ml-auto bg-rose-500 text-white text-[10px] font-bold rounded-full w-5 h-5 flex items-center justify-center">{misTareasPendientes}</span>
            )}
          </button>
          )}
        </nav>
        <div className="px-5 py-3 border-t border-white/10">
          <div className="flex items-center gap-2 mb-2">
            <div className="w-7 h-7 rounded-full bg-[#2E8B57]/30 flex items-center justify-center text-xs font-bold text-slate-100 shrink-0">
              {currentUser.nombre?.[0]?.toUpperCase() || "?"}
            </div>
            <div className="min-w-0">
              <div className="text-xs font-semibold text-slate-200 truncate">{currentUser.nombre} {currentUser.apellidos}</div>
              <div className="text-[10px] text-slate-500 flex items-center gap-1">
                {isAdmin && <ShieldCheck size={10} />} {currentUser.rol}
              </div>
            </div>
          </div>
          <button onClick={cerrarSesion} className="w-full flex items-center justify-center gap-1.5 text-xs font-semibold text-slate-400 hover:text-white border border-white/10 hover:border-white/20 rounded-md py-1.5 transition">
            <LogOut size={13} /> Cerrar sesión
          </button>
        </div>
      </aside>

      {/* Main */}
      <main className="flex-1 min-w-0 h-full overflow-y-auto">
        {/* Barra superior con botón de menú, solo visible en móvil */}
        <div className="md:hidden sticky top-0 z-20 bg-white border-b border-slate-200 px-4 py-3 flex items-center gap-3">
          <button onClick={() => setMenuMovilAbierto(true)} className="p-1.5 -ml-1.5 rounded-md hover:bg-slate-100">
            <Menu size={22} className="text-slate-700" />
          </button>
          <span className="font-display font-bold text-slate-800">ALUMAVEL</span>
        </div>
        {(() => {
          const incidenciasAbiertas = incidencias.filter((i) => i.estadoIncidencia !== "Solucionado");
          const incidenciasAntiguas = incidenciasAbiertas.filter((i) => i.fecha && daysDiff(i.fecha, new Date().toISOString().slice(0, 10)) > 7);
          if (incidenciasAbiertas.length === 0 || modulo === "incidencias") return null;
          return (
            <button
              onClick={() => setModulo("incidencias")}
              className={`w-full flex items-center gap-2 px-6 py-2.5 text-sm font-semibold text-left transition ${incidenciasAntiguas.length > 0 ? "bg-rose-50 text-rose-700 hover:bg-rose-100" : "bg-amber-50 text-amber-700 hover:bg-amber-100"}`}
            >
              <AlertOctagon size={15} />
              Tienes {incidenciasAbiertas.length} incidencia{incidenciasAbiertas.length === 1 ? "" : "s"} sin resolver
              {incidenciasAntiguas.length > 0 && ` — ${incidenciasAntiguas.length} lleva${incidenciasAntiguas.length === 1 ? "" : "n"} más de 7 días abierta${incidenciasAntiguas.length === 1 ? "" : "s"}`}
              . Toca para verlas →
            </button>
          );
        })()}
        {(() => {
          if (!currentUser) return null;
          const hoy = new Date().toISOString().slice(0, 10);
          const pedidosRetrasados = pedidos.filter((p) =>
            p.avisos?.retraso && p.avisos?.usuarioId === currentUser.id &&
            p.estado !== "Recibido" && p.estado !== "Cancelado" &&
            p.fechaEntregaPrevista && p.fechaEntregaPrevista < hoy
          );
          const pedidosLlegados = pedidos.filter((p) =>
            p.avisos?.llegada && p.avisos?.usuarioId === currentUser.id &&
            p.estado === "Recibido" && !p.avisoLlegadaVisto
          );
          const total = pedidosRetrasados.length + pedidosLlegados.length;
          if (total === 0 || modulo === "pedidos") return null;
          return (
            <button
              onClick={() => { marcarAvisosPedidosVistos(pedidosLlegados); setModulo("pedidos"); }}
              className="w-full flex items-center gap-2 px-6 py-2.5 text-sm font-semibold text-left transition bg-sky-50 text-sky-700 hover:bg-sky-100"
            >
              <Truck size={15} />
              {pedidosLlegados.length > 0 && `${pedidosLlegados.length} pedido${pedidosLlegados.length === 1 ? "" : "s"} han llegado`}
              {pedidosLlegados.length > 0 && pedidosRetrasados.length > 0 && " · "}
              {pedidosRetrasados.length > 0 && `${pedidosRetrasados.length} pedido${pedidosRetrasados.length === 1 ? "" : "s"} van con retraso`}
              . Toca para verlos →
            </button>
          );
        })()}
        {modulo === "clientes" && (
          <ClientesModulo
            clientes={clientes}
            proyectos={proyectos}
            ingresos={ingresos}
            view={clienteView}
            setView={setClienteView}
            editId={clienteEditId}
            setEditId={setClienteEditId}
            detailId={clienteDetailId}
            setDetailId={setClienteDetailId}
            onUpsert={upsertCliente}
            onDelete={deleteCliente}
            onImportarMasivo={importarContactosMasivo}
            isAdmin={isAdmin}
            openProyecto={(pid) => {
              setModulo("proyectos");
              setProyectoDetailId(pid);
              setProyectoView("detail");
            }}
          />
        )}
        {modulo === "proyectos" && (
          <ProyectosModulo
            proyectos={proyectos}
            clientes={clientes}
            facturas={facturas}
            ingresos={ingresos}
            materiales={materiales}
            articulos={articulos}
            pedidos={pedidos}
            proveedores={proveedores}
            openPedido={openPedidoFromCalendar}
            view={proyectoView}
            setView={setProyectoView}
            editId={proyectoEditId}
            setEditId={setProyectoEditId}
            detailId={proyectoDetailId}
            setDetailId={setProyectoDetailId}
            onUpsert={upsertProyecto}
            onDelete={deleteProyecto}
            isAdmin={isAdmin}
            onInlineUpdate={updateProyectoInline}
            nextNumero={nextNumeroProyecto}
            onRegistrarPago={registrarPagoProyecto}
            onRemovePago={removePago}
            onUsarArticulo={usarArticuloEnProyecto}
            onQuitarArticuloUsado={quitarArticuloUsado}
            onRegistrarIngreso={irARegistrarIngreso}
            instalaciones={instalaciones}
            onVerInstalacion={irAInstalacion}
            onGenerarPedidoFaltante={enviarAPedido}
          />
        )}
        {modulo === "proveedores" && (
          <ProveedoresModulo
            proveedores={proveedores}
            materiales={materiales}
            view={proveedorView}
            setView={setProveedorView}
            editId={proveedorEditId}
            setEditId={setProveedorEditId}
            detailId={proveedorDetailId}
            setDetailId={setProveedorDetailId}
            onUpsert={upsertProveedor}
            onDelete={deleteProveedor}
            isAdmin={isAdmin}
            onInlineUpdate={updateProveedorInline}
          />
        )}
        {modulo === "stock" && (
          <StockModulo
            materiales={materiales}
            proveedores={proveedores}
            view={materialView}
            setView={setMaterialView}
            editId={materialEditId}
            setEditId={setMaterialEditId}
            detailId={materialDetailId}
            setDetailId={setMaterialDetailId}
            onUpsert={upsertMaterial}
            onDelete={deleteMaterial}
            isAdmin={isAdmin}
            onAddMovimiento={addMovimiento}
            onRemoveMovimiento={removeMovimiento}
            onEnviarAPedido={enviarAPedido}
            onImportarTarifas={importarTarifasMateriales}
          />
        )}
        {modulo === "pedidos" && (
          <PedidosModulo
            pedidos={pedidos}
            proveedores={proveedores}
            materiales={materiales}
            proyectos={proyectos}
            view={pedidoView}
            setView={setPedidoView}
            editId={pedidoEditId}
            setEditId={setPedidoEditId}
            detailId={pedidoDetailId}
            setDetailId={setPedidoDetailId}
            onUpsert={upsertPedido}
            onDelete={deletePedido}
            isAdmin={isAdmin}
            onRecibir={recibirPedido}
            nextNumero={nextNumeroPedido}
            prefill={pedidoPrefill}
            onClearPrefill={() => setPedidoPrefill(null)}
            onConfirmarAlbaran={confirmarAlbaranPedido}
            onMarcarEnviado={marcarPedidoEnviado}
            onCrearDesdeFoto={(lineas, comentario) => enviarAPedido(undefined, lineas, "", comentario)}
            onCrearPedidoFaltante={crearPedidoFaltanteDesdeAlbaran}
            tabPrincipal={pedidoTabPrincipal}
            setTabPrincipal={setPedidoTabPrincipal}
            solicitudes={solicitudesPedido}
            currentUser={currentUser}
            solicitudView={solicitudPedidoView}
            setSolicitudView={setSolicitudPedidoView}
            solicitudEditId={solicitudPedidoEditId}
            setSolicitudEditId={setSolicitudPedidoEditId}
            solicitudDetailId={solicitudPedidoDetailId}
            setSolicitudDetailId={setSolicitudPedidoDetailId}
            onUpsertSolicitud={upsertSolicitudPedido}
            onDeleteSolicitud={deleteSolicitudPedido}
            onAprobarSolicitud={aprobarSolicitudPedido}
            onRechazarSolicitud={rechazarSolicitudPedido}
            onComentarSolicitud={addComentarioSolicitud}
            solicitudPrefill={solicitudPrefill}
            onClearSolicitudPrefill={() => setSolicitudPrefill(null)}
          />
        )}
        {modulo === "incidencias" && (
          <IncidenciasModulo
            incidencias={incidencias}
            proyectos={proyectos}
            clientes={clientes}
            pedidos={pedidos}
            proveedores={proveedores}
            openPedido={openPedidoFromCalendar}
            onPedirMateriales={(incidencia) => enviarAPedido(
              undefined,
              [],
              incidencia.proyectoId,
              `Pedido generado desde la incidencia #${incidencia.numero}.`
            )}
            view={incidenciaView}
            setView={setIncidenciaView}
            editId={incidenciaEditId}
            setEditId={setIncidenciaEditId}
            detailId={incidenciaDetailId}
            setDetailId={setIncidenciaDetailId}
            onUpsert={upsertIncidencia}
            onDelete={deleteIncidencia}
            isAdmin={isAdmin}
            onInlineUpdate={updateIncidenciaInline}
            nextNumero={nextNumeroIncidencia}
            onImportarMasivo={importarIncidenciasMasivo}
          />
        )}
        {modulo === "calendario" && (
          <CalendarioModulo
            proyectos={proyectos}
            clientes={clientes}
            pedidos={pedidos}
            incidencias={incidencias}
            openProyecto={openProyectoFromCalendar}
            openPedido={openPedidoFromCalendar}
            openIncidencia={openIncidenciaFromCalendar}
            onCambiarFechaProyecto={cambiarFechaProyecto}
            onCambiarFechaPedido={cambiarFechaPedido}
            onCambiarFechaIncidencia={cambiarFechaIncidencia}
          />
        )}
        {modulo === "articulos" && (
          <ArticulosModulo
            articulos={articulos}
            proveedores={proveedores}
            materiales={materiales}
            view={articuloView}
            setView={setArticuloView}
            editId={articuloEditId}
            setEditId={setArticuloEditId}
            detailId={articuloDetailId}
            setDetailId={setArticuloDetailId}
            onUpsert={upsertArticulo}
            onDelete={deleteArticulo}
            isAdmin={isAdmin}
            onInlineUpdate={updateArticuloInline}
            nextNumero={nextNumeroArticulo}
            onSolicitarArticulo={() => irASolicitarArticulo()}
            onImportarTarifas={importarTarifasArticulos}
          />
        )}
        {modulo === "facturas" && (
          <FacturasModulo
            facturas={facturas}
            clientes={clientes}
            proyectos={proyectos}
            view={facturaView}
            setView={setFacturaView}
            editId={facturaEditId}
            setEditId={setFacturaEditId}
            detailId={facturaDetailId}
            setDetailId={setFacturaDetailId}
            onUpsert={upsertFactura}
            onDelete={deleteFactura}
            isAdmin={isAdmin}
            onAddPago={addPago}
            onRemovePago={removePago}
            nextNumero={nextNumeroFactura}
          />
        )}
        {modulo === "presupuestos" && (
          <PresupuestosModulo
            presupuestos={presupuestos}
            clientes={clientes}
            onCrearClienteRapido={crearClienteRapido}
            view={presupuestoView}
            setView={setPresupuestoView}
            editId={presupuestoEditId}
            setEditId={setPresupuestoEditId}
            detailId={presupuestoDetailId}
            setDetailId={setPresupuestoDetailId}
            onUpsert={upsertPresupuesto}
            onDelete={deletePresupuesto}
            onAddLlamada={addLlamadaPresupuesto}
            onDeleteLlamada={deleteLlamadaPresupuesto}
            onCrearProyecto={(presupuesto) => {
              const proyectoId = crearProyectoDesdePresupuesto(presupuesto);
              setModulo("proyectos");
              setProyectoDetailId(proyectoId);
              setProyectoView("detail");
            }}
            onDuplicar={(presupuesto) => {
              const nuevaId = duplicarPresupuesto(presupuesto);
              setPresupuestoDetailId(nuevaId);
              setPresupuestoEditId(nuevaId);
              setPresupuestoView("form");
            }}
            isAdmin={isAdmin}
            prefill={presupuestoPrefill}
            onClearPrefill={() => setPresupuestoPrefill(null)}
          />
        )}
        {modulo === "mediciones" && (
          <MedicionesModulo
            mediciones={mediciones}
            view={medicionView}
            setView={setMedicionView}
            editId={medicionEditId}
            setEditId={setMedicionEditId}
            detailId={medicionDetailId}
            setDetailId={setMedicionDetailId}
            onUpsert={upsertMedicion}
            onDelete={deleteMedicion}
            incidencias={incidencias}
            onUpsertIncidencia={upsertIncidencia}
            onPasarAPresupuesto={pasarMedicionAPresupuesto}
          />
        )}
        {modulo === "chat" && (
          <ChatModulo
            usuarios={usuarios}
            currentUser={currentUser}
            proyectos={proyectos}
            onCrearTarea={crearTarea}
          />
        )}
        {modulo === "tareas" && (
          <TareasModulo
            tareas={tareas}
            usuarios={usuarios}
            currentUser={currentUser}
            onMarcarHecha={marcarTareaHecha}
            onConfirmar={confirmarTarea}
            onDelete={deleteTarea}
          />
        )}
        {modulo === "ingresos" && (
          <IngresosModulo
            ingresos={ingresos}
            clientes={clientes}
            proyectos={proyectos}
            view={ingresoView}
            setView={setIngresoView}
            editId={ingresoEditId}
            setEditId={setIngresoEditId}
            onUpsert={upsertIngreso}
            onDelete={deleteIngreso}
            onVincular={vincularIngresoAProyecto}
            onRegistrarIngreso={irARegistrarIngreso}
            prefill={ingresoPrefill}
            onClearPrefill={() => setIngresoPrefill(null)}
            isAdmin={isAdmin}
          />
        )}
        {modulo === "informes" && (
          <InformesModulo
            proyectos={proyectos}
            presupuestos={presupuestos}
            ingresos={ingresos}
            facturas={facturas}
            incidencias={incidencias}
            pedidos={pedidos}
            clientes={clientes}
            materiales={materiales}
            instalaciones={instalaciones}
            usuarios={usuarios}
            sesionesUsuario={sesionesUsuario}
            isAdmin={isAdmin}
          />
        )}
        {modulo === "fabrica" && (
          <FabricaModulo
            proyectos={proyectos}
            pedidos={pedidos}
            proveedores={proveedores}
            materiales={materiales}
            clientes={clientes}
            onConfirmarLinea={confirmarLineaFabrica}
            onIniciarFabricacion={(proyectoId) => {
              updateProyectoInline(proyectoId, { estadoTrabajo: "En proceso" });
              showToast("Proyecto marcado como en fabricación");
            }}
            cristales={cristales}
            onAddCristal={addCristal}
            onUpdateCristal={updateCristal}
            onDeleteCristal={deleteCristal}
            onUbicarCristal={ubicarCristal}
            onLiberarCristal={liberarCristal}
          />
        )}
        {modulo === "instalaciones" && (
          <InstalacionesModulo
            instalaciones={instalaciones}
            proyectos={proyectos}
            clientes={clientes}
            vehiculos={vehiculos}
            onUpsertVehiculo={upsertVehiculo}
            onDeleteVehiculo={deleteVehiculo}
            view={instalacionView}
            setView={setInstalacionView}
            detailId={instalacionDetailId}
            setDetailId={setInstalacionDetailId}
            onUpdate={updateInstalacion}
            onCrearManual={crearInstalacionManual}
            onAddHora={addHoraInstalacion}
            onDeleteHora={deleteHoraInstalacion}
            onAddGasto={addGastoInstalacion}
            onDeleteGasto={deleteGastoInstalacion}
            onAddMaterialFurgoneta={addMaterialFurgoneta}
            onCicloMaterialFurgoneta={cicloEstadoMaterialFurgoneta}
            onDeleteMaterialFurgoneta={deleteMaterialFurgoneta}
            isAdmin={isAdmin}
            incidencias={incidencias}
            onUpsertIncidencia={upsertIncidencia}
            materiales={materiales}
            usuarios={usuarios}
            onCrearTarea={crearTarea}
            currentUser={currentUser}
          />
        )}
        {modulo === "fichajes" && (
          <FichajesModulo
            fichajes={fichajes}
            empleadoActual={empleadoActual}
            setEmpleadoActual={setEmpleadoActualPersist}
            onFichar={registrarFichaje}
            onDeleteFichaje={deleteFichaje}
            isAdmin={isAdmin}
          />
        )}
        {modulo === "administracion" && isAdmin && (
          <AdministracionModulo
            usuarios={usuarios}
            currentUser={currentUser}
            onUpsert={upsertUsuario}
            onDelete={deleteUsuario}
          />
        )}
      </main>

      {toast && (
        <div
          className={`fixed bottom-5 right-5 px-4 py-2.5 rounded-md shadow-lg text-sm font-medium text-white flex items-center gap-2 ${
            toast.kind === "error" ? "bg-rose-600" : "bg-[#256E46]"
          }`}
        >
          {toast.kind === "error" ? <AlertCircle size={16} /> : <CheckCircle2 size={16} />}
          {toast.msg}
        </div>
      )}
    </div>
  );
}

/* ================= CLIENTES ================= */

function ClientesModulo({ clientes, proyectos, ingresos, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onImportarMasivo, openProyecto, isAdmin }) {
  const [q, setQ] = useState("");
  const [tipo, setTipo] = useState("");
  const inputContactosRef = useRef(null);

  const manejarImportarContactos = async (file) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const filas = leerFilasExcel(buffer);
    const nuevos = [];
    filas.forEach((fila) => {
      const nombre = String(valorPorCabeceras(fila, ["nombre", "cliente", "empresa", "razonsocial", "nombrecliente"])).trim();
      if (!nombre) return;
      nuevos.push({
        id: uid(), nombre, codigo: "", tipo: "Cliente",
        cif: String(valorPorCabeceras(fila, ["cif", "nif", "dni"])).trim(),
        direccion: String(valorPorCabeceras(fila, ["direccion", "domicilio"])).trim(),
        direccionFiscal: "", provincia: "", provinciaManual: "",
        pueblo: String(valorPorCabeceras(fila, ["pueblo", "poblacion", "ciudad", "localidad"])).trim(),
        cp: String(valorPorCabeceras(fila, ["cp", "codigopostal"])).trim(),
        email: String(valorPorCabeceras(fila, ["email", "correo", "mail"])).trim(),
        movil: String(valorPorCabeceras(fila, ["telefono", "movil", "tel", "phone"])).trim(),
        observaciones: String(valorPorCabeceras(fila, ["notas", "observaciones", "comentarios"])).trim(),
        primerPago: false, formaPago: "Contado", limiteCredito: "", tipoTarifa: "",
        portalActivo: false, portalPassword: "",
      });
    });
    if (nuevos.length === 0) {
      alert("No se ha encontrado ninguna fila con nombre. Revisa que el Excel tenga una columna \"Nombre\" (o \"Cliente\"/\"Empresa\").");
      return;
    }
    onImportarMasivo(nuevos);
  };

  const filtered = useMemo(() => {
    return clientes.filter((c) => {
      if (tipo && c.tipo !== tipo) return false;
      if (!q) return true;
      const hay = `${c.nombre} ${c.cif} ${c.email} ${c.movil} ${c.provincia} ${c.provinciaManual || ""} ${c.pueblo || ""} ${c.codigo}`.toLowerCase();
      return hay.includes(q.toLowerCase());
    });
  }, [clientes, q, tipo]);

  if (view === "form") {
    const editing = clientes.find((c) => c.id === editId) || null;
    return (
      <ClienteForm
        initial={editing}
        clientes={clientes}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={(data) => onUpsert(data)}
      />
    );
  }

  if (view === "detail") {
    const cliente = clientes.find((c) => c.id === detailId);
    if (!cliente) { setView("list"); return null; }
    const trabajos = proyectos.filter((p) => p.clienteId === cliente.id);
    return (
      <ClienteDetail
        cliente={cliente}
        trabajos={trabajos}
        ingresos={ingresos}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(cliente.id); setView("form"); }}
        onDelete={() => onDelete(cliente.id)}
        isAdmin={isAdmin}
        openProyecto={openProyecto}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<Users size={20} className="text-[#2E8B57]" />}
        title="Clientes"
        manualKey="clientes"
        subtitle={`${clientes.length} cliente${clientes.length === 1 ? "" : "s"} registrado${clientes.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-3 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO CLIENTE
      </button>

      <input
        ref={inputContactosRef}
        type="file"
        accept=".xlsx,.xls,.ods,.csv"
        className="hidden"
        onChange={(e) => { manejarImportarContactos(e.target.files[0]); e.target.value = ""; }}
      />
      <button
        onClick={() => inputContactosRef.current?.click()}
        className="w-full flex items-center justify-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 py-2.5 rounded-md mb-6 hover:bg-slate-50"
      >
        Importar contactos desde Excel (masivo)
      </button>

      <div className="flex gap-3 mb-4">
        <div className="relative flex-1 max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Buscar por nombre, CIF, email, teléfono…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]"
          />
        </div>
        <Select value={tipo} onChange={(e) => setTipo(e.target.value)} className="max-w-[180px]">
          <option value="">Todos los tipos</option>
          {TIPO_CLIENTE.map((t) => <option key={t} value={t}>{t}</option>)}
        </Select>
        <button
          onClick={() => descargarListaComoWord(
            "Clientes",
            ["Nombre / Empresa", "CIF", "Email", "Móvil", "Proyectos", "Riesgo"],
            filtered.map((c) => {
              const limiteCredito = parseFloat(c.limiteCredito) || 0;
              const riesgoActual = proyectos.filter((p) => p.clienteId === c.id).reduce((s, p) => {
                const recibido = (ingresos || []).filter((i) => i.proyectoId === p.id).reduce((s2, i) => s2 + (parseFloat(i.importe) || 0), 0);
                return s + Math.max((parseFloat(p.importePresupuesto) || 0) - recibido, 0);
              }, 0);
              const pct = limiteCredito ? `${((riesgoActual / limiteCredito) * 100).toFixed(0)}%` : "—";
              return [c.nombre, c.cif || "—", c.email || "—", c.movil || "—", String(proyectos.filter((p) => p.clienteId === c.id).length), pct];
            })
          )}
          className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
        >
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nombre / Empresa</th>
              <th className="px-4 py-3 font-semibold">CIF</th>
              <th className="px-4 py-3 font-semibold">Email</th>
              <th className="px-4 py-3 font-semibold">Móvil</th>
              <th className="px-4 py-3 font-semibold text-right">Proyectos</th>
              <th className="px-4 py-3 font-semibold">Riesgo</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay clientes que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((c) => (
              <tr
                key={c.id}
                onClick={() => { setDetailId(c.id); setView("detail"); }}
                className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition"
              >
                <td className="px-4 py-3 font-medium text-slate-800">{c.nombre}
                  {c.tipo === "Distribuidor" && <Badge className="ml-2 bg-violet-50 text-violet-700 ring-violet-200">Distribuidor</Badge>}
                </td>
                <td className="px-4 py-3 font-mono-num text-slate-500">{c.cif || "—"}</td>
                <td className="px-4 py-3 text-slate-600">{c.email || "—"}</td>
                <td className="px-4 py-3 text-slate-600">{c.movil || "—"}</td>
                <td className="px-4 py-3 text-right font-mono-num text-slate-500">{proyectos.filter((p) => p.clienteId === c.id).length}</td>
                <td className="px-4 py-3">
                  <RiesgoClienteBadge cliente={c} proyectos={proyectos} ingresos={ingresos} />
                </td>
                <td className="px-4 py-3 text-right">
                  {isAdmin && (
                    <button
                      onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar a ${c.nombre}?`)) onDelete(c.id); }}
                      className="text-slate-300 hover:text-rose-500 transition"
                    >
                      <Trash2 size={15} />
                    </button>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// Texto de ayuda para cada pestaña. Se muestra al pulsar el botón "?" que
// aparece junto al título de cada sección (ver componente Header más abajo).
const MANUALES = {
  clientes: {
    puntos: [
      "Aquí tienes el listado de todos tus clientes. Usa el buscador para filtrar por nombre, CIF, email o teléfono.",
      "Pulsa \"NUEVO CLIENTE\" para dar de alta uno.",
      "Haz clic en una fila para abrir el detalle: sus proyectos, sus pagos y sus datos de contacto.",
      "Desde el detalle puedes editar sus datos o borrarlo.",
      "La columna \"Riesgo\" muestra qué porcentaje de su límite de crédito tiene pendiente de pagar. Si sale en rojo, ese cliente ya ha superado el límite que le pusiste.",
    ],
  },
  proyectos: {
    puntos: [
      "Aquí ves todos los proyectos/obras, con su cliente, importe y estado.",
      "Un proyecto normalmente nace de un presupuesto aceptado (desde la pestaña Presupuestos), pero también puedes crear uno directamente con el botón de nuevo proyecto.",
      "Dentro de cada proyecto puedes registrar los pagos que te van llegando, los materiales/artículos usados, y ver los pedidos e instalaciones relacionados con esa obra.",
      "Desde aquí también puedes generar el pedido de materiales de un proyecto directamente hacia Pedidos.",
    ],
  },
  proveedores: {
    puntos: [
      "Listado de tus proveedores de materiales.",
      "Pulsa para añadir uno nuevo con sus datos de contacto.",
      "Al entrar en el detalle de un proveedor puedes ver qué materiales le compras y editar sus datos.",
    ],
  },
  stock: {
    puntos: [
      "Catálogo de todos los materiales (perfiles, herrajes, cristal, etc.) que manejas, con su stock actual.",
      "Puedes registrar entradas (cuando llega material) y salidas (cuando se usa) de cada material, y el historial de movimientos queda guardado.",
      "Si un material se queda corto, puedes enviarlo directamente a un pedido a proveedor desde aquí.",
    ],
  },
  pedidos: {
    puntos: [
      "Aquí gestionas los pedidos que haces a tus proveedores.",
      "Puedes crear un pedido a mano, o generarlo automáticamente desde una foto de un albarán/presupuesto de proveedor.",
      "Cuando el proveedor confirma o envía el pedido, marca esos estados aquí para llevar el seguimiento.",
      "Cuando llega el material, usa \"Recibir\" para meterlo en el stock automáticamente.",
      "La pestaña \"Solicitudes\" (si la ves) es donde tus empleados piden materiales y tú, como administrador, las apruebas o rechazas.",
    ],
  },
  solicitudesPedido: {
    puntos: [
      "Aquí los empleados solicitan materiales que necesitan para una obra.",
      "Como administrador, puedes aprobar la solicitud (y se convierte en pedido a proveedor) o rechazarla explicando el motivo.",
      "El empleado puede ver el estado de sus propias solicitudes.",
    ],
  },
  fabrica: {
    puntos: [
      "Este módulo controla lo que pasa en el almacén/taller, independientemente de la parte de oficina.",
      "\"Listo para fabricar\": líneas de pedido ya recibidas y listas para empezar a fabricar.",
      "\"Materiales pendientes\": lo que aún falta para poder fabricar.",
      "\"En fabricación\": lo que ya se está fabricando.",
      "\"Cristales\": aquí se gestiona la ubicación física de los caballetes de cristal en el almacén (zona Arriba/Uxcar y Abajo/ALUMAVEL). Puedes importar un packing list en foto o PDF y el sistema coloca automáticamente cada caballete en un hueco libre.",
    ],
  },
  instalaciones: {
    puntos: [
      "Aquí ves las instalaciones (obras) que están en marcha o completadas, con las horas y gastos de cada una.",
      "Puedes añadir horas trabajadas, gastos, y materiales cargados en la furgoneta para cada instalación.",
      "El \"Control de montaje por vivienda\" (dentro del detalle de una instalación con varios bloques/viviendas) te permite marcar qué elementos están instalados en cada vivienda, subir fotos, y generar incidencias directamente si algo falla.",
      "Desde aquí también se accede a \"Vehículos\" (la flota) y \"Furgoneta\" (qué material va en cada furgoneta).",
    ],
  },
  vehiculos: {
    puntos: [
      "Listado de la flota de vehículos (furgonetas, camiones) disponibles para asignar a instalaciones.",
      "Aquí das de alta o borras vehículos.",
    ],
  },
  furgoneta: {
    puntos: [
      "Vista de qué material falta cargar, qué va ya en la furgoneta, y qué está en la obra, para todas las instalaciones activas a la vez.",
      "Te ayuda a preparar la furgoneta antes de salir a una obra sin olvidar nada.",
    ],
  },
  incidencias: {
    puntos: [
      "Aquí registras cualquier problema o incidencia (una rotura, un fallo de medida, una reclamación, etc.).",
      "Puedes vincular la incidencia a un proyecto y, si hace falta material para resolverla, pedirlo directamente a proveedor desde la propia incidencia.",
      "El estado de la incidencia (abierta, en curso, resuelta) se actualiza desde el detalle.",
    ],
  },
  calendario: {
    puntos: [
      "Vista mensual que junta automáticamente fechas de proyectos, pedidos e incidencias, para que veas todo en un solo calendario.",
      "Puedes arrastrar un elemento a otro día para cambiarle la fecha directamente desde aquí.",
      "Haz clic en un elemento para ir directamente a su detalle (proyecto, pedido o incidencia).",
    ],
  },
  articulos: {
    puntos: [
      "Catálogo de artículos que fabricas (por ejemplo, un tipo de ventana concreto), compuestos a partir de los materiales de Stock.",
      "Cuando usas un artículo en un proyecto, se descuentan automáticamente los materiales que lo componen.",
    ],
  },
  facturas: {
    puntos: [
      "Aquí gestionas las facturas que emites a tus clientes.",
      "Puedes registrar los pagos que vas recibiendo de cada factura y ver cuánto queda pendiente.",
    ],
  },
  presupuestos: {
    puntos: [
      "Aquí creas y gestionas los presupuestos que envías a clientes (nuevos o ya existentes).",
      "Puedes registrar las llamadas de seguimiento que haces a un cliente sobre su presupuesto.",
      "Cuando el cliente lo acepta, pulsa \"Crear proyecto\" para convertir ese presupuesto en un proyecto/obra real.",
      "También puedes duplicar un presupuesto para no escribirlo todo de nuevo si es parecido a otro.",
      "Si vienes de la pestaña Mediciones y pulsaste \"Pasar a presupuesto\", el formulario se abre aquí ya con el cliente y una descripción con el resumen de las medidas — solo te falta poner el número y el importe.",
    ],
  },
  mediciones: {
    puntos: [
      "Aquí registras las medidas que tomas en una visita, antes incluso de tener un presupuesto o, a veces, antes de tener el cliente dado de alta en el CRM.",
      "Puedes importar un Excel/ODS con la plantilla, o pulsar \"+ Añadir a mano (sin Excel)\" para crear una vivienda/habitación escribiendo su nombre directamente.",
      "Dentro de cada vivienda, además de lo que venga del Excel, puedes añadir una habitación o elemento a mano poniéndole un nombre (ej: \"Cocina\") y su medida.",
      "En cada elemento marcas qué lleva: marco, hojas, persiana, mosquitera, cajón de obra, montaje, tapajuntas, postigo, silicona... y puedes añadir cualquier otro con el \"+\".",
      "Puedes añadir foto por elemento y documentos (planos, Excel, PDF...) clasificados por categoría.",
      "Cuando termines de medir, pulsa \"Pasar a presupuesto\" dentro de la medición: se abre un presupuesto nuevo con el cliente puesto y una descripción con el resumen de todo lo medido.",
    ],
  },
  chat: {
    puntos: [
      "Chat interno de la empresa: hay una sala \"General\" para todos, y una sala automática para cada proyecto/obra.",
      "Los mensajes se guardan — puedes cerrar el CRM y al volver seguirán ahí.",
      "Pasa el ratón por encima de un mensaje y pulsa \"→ Crear tarea\" para convertirlo directamente en una tarea asignada a un compañero, sin tener que escribirla de nuevo.",
      "Pulsa \"Exportar a PDF\", marca los mensajes que quieras (no hace falta que sean todos) y se abrirá una ventana lista para imprimir o guardar como PDF.",
    ],
  },
  tareas: {
    puntos: [
      "Aquí ves las tareas que te han asignado (pestaña \"Asignadas a mí\") y las que tú le has asignado a otros (\"Creadas por mí\").",
      "Las tareas normalmente se crean desde un mensaje del Chat, con el botón \"→ Crear tarea\".",
      "Al crear una tarea puedes elegir si necesitas confirmarla tú mismo cuando la marquen como hecha, o si con que la marquen como hecha ya es suficiente.",
      "Si pediste confirmación, cuando tu compañero la marque como hecha te aparecerá en \"Pendiente de confirmar\" dentro de \"Creadas por mí\", y podrás confirmarla o devolverla si no está bien.",
    ],
  },
  ingresos: {
    puntos: [
      "Aquí ves el dinero que ha entrado (pagos de clientes), aunque no venga directamente de una factura.",
      "Puedes vincular una entrada de dinero a un proyecto concreto para que se reste de lo pendiente de ese proyecto.",
    ],
  },
  informes: {
    puntos: [
      "Vista general con gráficos: proyectos, presupuestos, dinero cobrado y pendiente, incidencias, y productividad de instalación.",
      "Al final (solo para administradores) verás el tiempo que cada usuario ha tenido el CRM abierto: hoy, últimos 7 días y total acumulado.",
      "Te sirve para hacerte una idea rápida de cómo va el negocio sin entrar pestaña por pestaña.",
    ],
  },
  fichajes: {
    puntos: [
      "Control horario de los empleados: fichar entrada, salida, y pausas.",
      "Como administrador puedes ver y corregir los fichajes de todos los empleados.",
    ],
  },
  administracion: {
    puntos: [
      "Aquí gestionas quién tiene acceso al CRM y con qué permisos (administrador o empleado normal).",
      "Puedes dar de alta nuevos usuarios o quitarle el acceso a alguien.",
      "Botón \"Descargar copia de seguridad ahora\": te descargas al momento todos los datos del CRM en un archivo. Además, todos los días se envía automáticamente una copia por email a alumavel@alumavel.es, sin que tengas que hacer nada.",
    ],
  },
};

function Header({ icon, title, subtitle, action, manualKey }) {
  const [ayudaAbierta, setAyudaAbierta] = useState(false);
  const manual = manualKey ? MANUALES[manualKey] : null;
  return (
    <div className="flex items-start justify-between flex-wrap gap-3 mb-6">
      <div>
        <div className="flex items-center gap-2">
          {icon}
          <h1 className="font-display text-2xl font-extrabold text-slate-900">{title}</h1>
          {manual && (
            <button
              type="button"
              onClick={() => setAyudaAbierta(true)}
              title="Cómo funciona esta pestaña"
              className="w-6 h-6 flex items-center justify-center rounded-full border border-slate-300 text-slate-500 hover:bg-slate-100 hover:text-slate-700 text-xs font-bold shrink-0"
            >
              ?
            </button>
          )}
        </div>
        {subtitle && <p className="text-sm text-slate-500 mt-1">{subtitle}</p>}
      </div>
      {action}
      {manual && ayudaAbierta && (
        <div
          className="fixed inset-0 bg-black/40 flex items-center justify-center z-50 p-4"
          onClick={() => setAyudaAbierta(false)}
        >
          <div
            className="bg-white rounded-lg max-w-lg w-full max-h-[80vh] overflow-y-auto p-6"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2">
                {icon}
                <h2 className="font-display text-lg font-bold text-slate-900">Cómo funciona: {title}</h2>
              </div>
              <button onClick={() => setAyudaAbierta(false)} className="text-slate-400 hover:text-slate-600">
                <X size={18} />
              </button>
            </div>
            <ul className="space-y-2.5 text-sm text-slate-700 list-disc pl-5">
              {manual.puntos.map((p, i) => <li key={i}>{p}</li>)}
            </ul>
          </div>
        </div>
      )}
    </div>
  );
}

function RiesgoClienteBadge({ cliente, proyectos, ingresos }) {
  const limiteCredito = parseFloat(cliente.limiteCredito) || 0;
  if (!limiteCredito) return <span className="text-xs text-slate-300">—</span>;
  const riesgoActual = proyectos.filter((p) => p.clienteId === cliente.id).reduce((s, p) => {
    const importePresupuesto = parseFloat(p.importePresupuesto) || 0;
    const recibido = (ingresos || []).filter((i) => i.proyectoId === p.id).reduce((s2, i) => s2 + (parseFloat(i.importe) || 0), 0);
    return s + Math.max(importePresupuesto - recibido, 0);
  }, 0);
  const pct = (riesgoActual / limiteCredito) * 100;
  const cls = pct >= 100 ? "bg-rose-50 text-rose-700 ring-rose-200" : pct >= 80 ? "bg-amber-50 text-amber-700 ring-amber-200" : "bg-emerald-50 text-emerald-700 ring-emerald-200";
  return <Badge className={cls}>{pct.toFixed(0)}%{pct >= 100 ? " ⚠" : ""}</Badge>;
}

function ClienteForm({ initial, clientes, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, nombre: "", codigo: "", tipo: "Cliente", cif: "", direccion: "", direccionFiscal: "",
      provincia: "", provinciaManual: "", pueblo: "", cp: "", email: "", movil: "", observaciones: "",
      primerPago: false, formaPago: "Contado", limiteCredito: "", tipoTarifa: "",
      portalActivo: false, portalPassword: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value });
  const pueblosDisponibles = useMemo(() => {
    const uniq = new Set((clientes || []).map((c) => (c.pueblo || "").trim()).filter(Boolean));
    return Array.from(uniq).sort();
  }, [clientes]);
  const tarifasDisponibles = useMemo(() => {
    const uniq = new Set(["Tarifa tiendas", "PVP1", "PVP2", "PVP3", ...(clientes || []).map((c) => (c.tipoTarifa || "").trim())].filter(Boolean));
    return Array.from(uniq).sort();
  }, [clientes]);

  const [errorMsg, setErrorMsg] = useState("");
  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.nombre.trim()) {
      setErrorMsg("Falta el campo Nombre / Empresa (es obligatorio).");
      alert("Falta el campo Nombre / Empresa. Ese campo es obligatorio para guardar el cliente.");
      return;
    }
    if (!f.limiteCredito || parseFloat(f.limiteCredito) <= 0) {
      setErrorMsg("Falta el Límite de crédito asegurado (es obligatorio y tiene que ser mayor que 0).");
      return;
    }
    if (f.portalActivo && (!f.email.trim() || !f.portalPassword.trim())) {
      alert("Para activar el acceso al portal, el cliente necesita email y contraseña.");
      return;
    }
    setErrorMsg("");
    onSave({ ...f, limiteCredito: parseFloat(f.limiteCredito) || 0 });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5">
        <ChevronLeft size={16} /> Volver
      </button>

      {/* Título del formulario */}
      <h1 className="font-display text-xl font-extrabold text-slate-900 mb-4">
        {initial ? "Editar cliente" : "Alta de cliente"}
      </h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Nombre / Empresa" required>
            <TextInput value={f.nombre} onChange={set("nombre")} placeholder="Nombre y apellidos o razón social" required />
          </Field>
          <Field label="Código de cliente / tienda">
            <TextInput value={f.codigo} onChange={set("codigo")} placeholder="Opcional" />
          </Field>
        </div>
        <div className="grid grid-cols-3 gap-4">
          <Field label="Tipo">
            <Select value={f.tipo} onChange={set("tipo")}>
              {TIPO_CLIENTE.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="DNI / CIF">
            <TextInput value={f.cif} onChange={set("cif")} />
          </Field>
          <Field label="Móvil">
            <TextInput value={f.movil} onChange={set("movil")} />
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-4">
          <Field label="Dirección de entrega">
            <TextInput value={f.direccion} onChange={set("direccion")} placeholder="Dónde se entrega el material" />
          </Field>
          <Field label="Dirección fiscal">
            <TextInput value={f.direccionFiscal} onChange={set("direccionFiscal")} placeholder="Para facturación (si es distinta)" />
          </Field>
        </div>
        <div className="grid grid-cols-3 gap-4">
          <Field label="Provincia">
            <Select value={f.provincia || ""} onChange={set("provincia")}>
              <option value="">Selecciona...</option>
              {PROVINCIAS_REPARTO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          {f.provincia === "Otra ciudad..." && (
            <Field label="Escribe la provincia/ciudad">
              <TextInput value={f.provinciaManual} onChange={set("provinciaManual")} placeholder="Ej: Córdoba" />
            </Field>
          )}
          <Field label="Pueblo / localidad">
            <TextInput value={f.pueblo} onChange={set("pueblo")} list="pueblos-datalist" placeholder="Escribe el pueblo" />
            <datalist id="pueblos-datalist">
              {pueblosDisponibles.map((p) => <option key={p} value={p} />)}
            </datalist>
          </Field>
          <Field label="Código postal">
            <TextInput value={f.cp} onChange={set("cp")} />
          </Field>
          <Field label="Email">
            <TextInput type="email" value={f.email} onChange={set("email")} />
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-4 items-end">
          <Field label="Forma de pago">
            <Select value={f.formaPago} onChange={set("formaPago")}>
              {FORMA_PAGO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Tipo de tarifa">
            <TextInput value={f.tipoTarifa} onChange={set("tipoTarifa")} list="tarifas-datalist" placeholder="Ej: Tarifa tiendas" />
            <datalist id="tarifas-datalist">
              {tarifasDisponibles.map((t) => <option key={t} value={t} />)}
            </datalist>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-4 items-end">
          <Field label="Límite de crédito asegurado (€)" required>
            <TextInput type="number" step="0.01" value={f.limiteCredito} onChange={set("limiteCredito")} placeholder="Ej: 40000" />
          </Field>
        </div>
        <div className="flex items-end">
          <label className="flex items-center gap-2 pb-2 text-sm text-slate-700">
            <input type="checkbox" checked={f.primerPago} onChange={set("primerPago")} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
            Primer pago realizado
          </label>
        </div>
        <p className="text-xs text-slate-400 -mt-3">Importe máximo que tienes asegurado en compras de este cliente. El CRM avisará cuando lo que le queda pendiente de cobro en sus proyectos se acerque o supere este límite.</p>
        <Field label="Observaciones">
          <TextArea rows={3} value={f.observaciones} onChange={set("observaciones")} />
        </Field>

        <div className="border-t border-slate-100 pt-4">
          <label className="flex items-center gap-2 text-sm font-semibold text-slate-700 mb-3">
            <input type="checkbox" checked={f.portalActivo} onChange={set("portalActivo")} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
            <Globe size={15} className="text-[#2E8B57]" /> Dar acceso al portal de cliente
          </label>
          {f.portalActivo && (
            <div className="grid grid-cols-2 gap-4 bg-slate-50 border border-slate-200 rounded-md p-3">
              <Field label="Email de acceso" required>
                <TextInput type="email" value={f.email} onChange={set("email")} required placeholder="Usa el email del cliente" />
              </Field>
              <Field label="Contraseña del portal" required>
                <TextInput type="password" value={f.portalPassword} onChange={set("portalPassword")} required />
              </Field>
            </div>
          )}
        </div>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100 border border-slate-300">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="whitespace-nowrap flex items-center gap-1.5 text-sm font-bold px-5 py-2.5 rounded-md">
            <Save size={15} /> <span>Guardar cliente</span>
          </button>
        </div>
      </form>
    </div>
  );
}

function ClienteDetail({ cliente, trabajos, ingresos, onBack, onEdit, onDelete, openProyecto, isAdmin }) {
  const limiteCredito = parseFloat(cliente.limiteCredito) || 0;
  const riesgoActual = trabajos.reduce((s, p) => {
    const importePresupuesto = parseFloat(p.importePresupuesto) || 0;
    const recibido = (ingresos || []).filter((i) => i.proyectoId === p.id).reduce((s2, i) => s2 + (parseFloat(i.importe) || 0), 0);
    return s + Math.max(importePresupuesto - recibido, 0);
  }, 0);
  const pctRiesgo = limiteCredito ? (riesgoActual / limiteCredito) * 100 : 0;

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5">
        <ChevronLeft size={16} /> Volver al listado
      </button>

      <div className="flex items-start justify-between mb-6">
        <div>
          <div className="flex items-center gap-2">
            <h1 className="font-display text-2xl font-extrabold text-slate-900">{cliente.nombre}</h1>
            {cliente.tipo === "Distribuidor" && <Badge className="bg-violet-50 text-violet-700 ring-violet-200">Distribuidor</Badge>}
            {cliente.portalActivo ? (
              <Badge className="bg-sky-50 text-sky-700 ring-sky-200"><Globe size={11} /> Portal activo</Badge>
            ) : (
              <Badge className="bg-slate-100 text-slate-500 ring-slate-200"><Globe size={11} /> Sin acceso al portal</Badge>
            )}
          </div>
          <p className="text-sm text-slate-500 mt-1 font-mono-num">{cliente.codigo && `Código ${cliente.codigo} · `}{cliente.cif || "Sin CIF"}</p>
        </div>
        <div className="flex gap-2">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
            <Pencil size={14} /> Editar
          </button>
          {isAdmin && (
          <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50">
            <Trash2 size={14} /> Eliminar
          </button>
          )}
        </div>
      </div>

      {limiteCredito > 0 && (
        <div className={`flex flex-wrap items-center gap-2 px-4 py-3 rounded-md mb-6 text-sm font-semibold ${
          pctRiesgo >= 100 ? "bg-rose-50 border border-rose-300 text-rose-700" : pctRiesgo >= 80 ? "bg-amber-50 border border-amber-300 text-amber-800" : "bg-emerald-50 border border-emerald-200 text-emerald-700"
        }`}>
          <ShieldCheck size={16} />
          Riesgo actual: {money(riesgoActual)} de {money(limiteCredito)} asegurados ({pctRiesgo.toFixed(0)}%).
          {pctRiesgo >= 100 && <span className="text-base"> ⚠ Límite superado — no se le puede vender más sin cobrar antes.</span>}
          {pctRiesgo >= 80 && pctRiesgo < 100 && <span> Cerca del límite, ve con cuidado con nuevos pedidos.</span>}
        </div>
      )}

      <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6 mb-6">
        <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
          <InfoRow icon={<MapPin size={14} />} label="Dirección de entrega" value={[cliente.direccion, cliente.pueblo, cliente.provincia === "Otra ciudad..." ? cliente.provinciaManual : cliente.provincia, cliente.cp].filter(Boolean).join(", ") || "—"} />
          {cliente.direccionFiscal && <InfoRow icon={<MapPin size={14} />} label="Dirección fiscal" value={cliente.direccionFiscal} />}
          <InfoRow icon={<Mail size={14} />} label="Email" value={cliente.email || "—"} />
          <InfoRow icon={<Phone size={14} />} label="Móvil" value={cliente.movil || "—"} />
          <InfoRow icon={<Euro size={14} />} label="Forma de pago" value={cliente.formaPago + (cliente.primerPago ? " · Primer pago realizado" : "")} />
          <InfoRow icon={<FileText size={14} />} label="Tipo de tarifa" value={cliente.tipoTarifa || "—"} />
          <InfoRow icon={<ShieldCheck size={14} />} label="Límite de crédito asegurado" value={limiteCredito ? money(limiteCredito) : "—"} />
        </div>
        {cliente.observaciones && (
          <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Observaciones</span>
            {cliente.observaciones}
          </div>
        )}
      </CornerFrame>

      <div className="flex items-center justify-between mb-3">
        <h2 className="font-display font-bold text-slate-800 flex items-center gap-2">
          <ClipboardList size={16} className="text-[#2E8B57]" /> Histórico de trabajos
        </h2>
        <span className="text-xs text-slate-400">{trabajos.length} proyecto{trabajos.length === 1 ? "" : "s"}</span>
      </div>
      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        {trabajos.length === 0 ? (
          <p className="px-4 py-8 text-center text-sm text-slate-400">Este cliente aún no tiene proyectos asociados.</p>
        ) : (
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                <th className="px-4 py-2.5 font-semibold">Nº</th>
                <th className="px-4 py-2.5 font-semibold">Proyecto</th>
                <th className="px-4 py-2.5 font-semibold">Estado presupuesto</th>
                <th className="px-4 py-2.5 font-semibold">Estado trabajo</th>
                <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                <th className="px-4 py-2.5 font-semibold">Entrega prevista</th>
              </tr>
            </thead>
            <tbody>
              {trabajos.map((t) => (
                <tr key={t.id} onClick={() => openProyecto(t.id)} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer">
                  <td className="px-4 py-2.5 font-mono-num text-slate-500">#{t.numero}</td>
                  <td className="px-4 py-2.5 font-medium text-slate-800">{t.nombre}</td>
                  <td className="px-4 py-2.5"><Badge className={ESTADO_PRESUPUESTO_STYLE[t.estadoPresupuesto]}>{t.estadoPresupuesto}</Badge></td>
                  <td className="px-4 py-2.5"><Badge className={ESTADO_TRABAJO_STYLE[t.estadoTrabajo]}>{t.estadoTrabajo}</Badge></td>
                  <td className="px-4 py-2.5 text-right font-mono-num">{money(t.importePresupuesto)}</td>
                  <td className="px-4 py-2.5 text-slate-500">{fmtDate(t.fechaEntregaPrevista)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function InfoRow({ icon, label, value }) {
  return (
    <div>
      <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 flex items-center gap-1.5 mb-1">{icon}{label}</span>
      <span className="text-slate-700">{value}</span>
    </div>
  );
}

/* ================= PROYECTOS ================= */

function ProyectosModulo({ proyectos, clientes, facturas, ingresos, materiales, articulos, pedidos, proveedores, openPedido, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onInlineUpdate, nextNumero, isAdmin, onRegistrarPago, onRemovePago, onUsarArticulo, onQuitarArticuloUsado, onRegistrarIngreso, instalaciones, onVerInstalacion, onGenerarPedidoFaltante }) {
  const [q, setQ] = useState("");
  const [estadoTrabajo, setEstadoTrabajo] = useState("");
  const [clienteFiltro, setClienteFiltro] = useState("");
  const [ciudadFiltro, setCiudadFiltro] = useState("");

  const clienteNombre = (id) => clientes.find((c) => c.id === id)?.nombre || "—";
  const ciudadRepartoDe = (p) => p.estadoLogistica === "Reparto (camión)"
    ? (p.provinciaReparto === "Otra ciudad..." ? (p.ciudadRepartoManual || "").trim() : p.provinciaReparto)
    : "";

  const ciudadesDisponibles = useMemo(() => {
    const set = new Set(proyectos.map(ciudadRepartoDe).filter(Boolean));
    return Array.from(set).sort();
  }, [proyectos]);

  const filtered = useMemo(() => {
    return proyectos.filter((p) => {
      if (estadoTrabajo && p.estadoTrabajo !== estadoTrabajo) return false;
      if (clienteFiltro && p.clienteId !== clienteFiltro) return false;
      if (ciudadFiltro && ciudadRepartoDe(p) !== ciudadFiltro) return false;
      if (!q) return true;
      const hay = `${p.numero} ${p.nombre} ${clienteNombre(p.clienteId)} ${p.ubicacion}`.toLowerCase();
      return hay.includes(q.toLowerCase());
    });
  }, [proyectos, q, estadoTrabajo, clienteFiltro, ciudadFiltro, clientes]);

  if (view === "form") {
    const editing = proyectos.find((p) => p.id === editId) || null;
    return (
      <ProyectoForm
        initial={editing}
        clientes={clientes}
        nextNumero={nextNumero}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={(data) => onUpsert(data)}
      />
    );
  }

  if (view === "detail") {
    const proyecto = proyectos.find((p) => p.id === detailId);
    if (!proyecto) { setView("list"); return null; }
    return (
      <ProyectoDetail
        proyecto={proyecto}
        cliente={clientes.find((c) => c.id === proyecto.clienteId)}
        facturas={facturas.filter((f) => (f.proyectosIds || []).includes(proyecto.id))}
        ingresos={ingresos.filter((i) => i.proyectoId === proyecto.id)}
        materiales={materiales}
        articulos={articulos}
        pedidos={pedidos.filter((pd) => pd.proyectoId === proyecto.id)}
        proveedores={proveedores}
        openPedido={openPedido}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(proyecto.id); setView("form"); }}
        onDelete={() => onDelete(proyecto.id)}
        isAdmin={isAdmin}
        onInlineUpdate={onInlineUpdate}
        onRegistrarPago={(data) => onRegistrarPago(proyecto.id, data)}
        onRemovePago={onRemovePago}
        onUsarArticulo={(articuloId, cantidad) => onUsarArticulo(proyecto.id, articuloId, cantidad)}
        onQuitarArticuloUsado={(usoId) => onQuitarArticuloUsado(proyecto.id, usoId)}
        onRegistrarIngreso={onRegistrarIngreso}
        instalacion={instalaciones.find((i) => i.proyectoId === proyecto.id)}
        onVerInstalacion={onVerInstalacion}
        onGenerarPedidoFaltante={onGenerarPedidoFaltante}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<Briefcase size={20} className="text-[#2E8B57]" />}
        title="Proyectos / Obras"
        manualKey="proyectos"
        subtitle={`${proyectos.length} proyecto${proyectos.length === 1 ? "" : "s"} registrado${proyectos.length === 1 ? "" : "s"}`}
      />

      {clientes.length === 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
          ⚠ No puedes crear un proyecto todavía: primero da de alta al menos un cliente en la pestaña Clientes.
        </div>
      )}
      <button
        onClick={() => { setEditId(null); setView("form"); }}
        disabled={clientes.length === 0}
        title={clientes.length === 0 ? "Da de alta un cliente primero" : ""}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO PROYECTO
      </button>

      {clientes.length === 0 && (
        <div className="mb-4 px-4 py-3 rounded-md bg-amber-50 text-amber-800 text-sm border border-amber-200">
          Necesitas dar de alta al menos un cliente antes de crear un proyecto.
        </div>
      )}

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 min-w-[220px] max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Buscar por nº, nombre, cliente, ubicación…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]"
          />
        </div>
        <Select value={clienteFiltro} onChange={(e) => setClienteFiltro(e.target.value)} className="max-w-[200px]">
          <option value="">Todos los clientes</option>
          {clientes.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
        </Select>
        <Select value={estadoTrabajo} onChange={(e) => setEstadoTrabajo(e.target.value)} className="max-w-[200px]">
          <option value="">Todos los estados</option>
          {ESTADO_TRABAJO.map((t) => <option key={t}>{t}</option>)}
        </Select>
        {ciudadesDisponibles.length > 0 && (
          <Select value={ciudadFiltro} onChange={(e) => setCiudadFiltro(e.target.value)} className="max-w-[200px]">
            <option value="">Todas las ciudades (reparto)</option>
            {ciudadesDisponibles.map((c) => <option key={c}>{c}</option>)}
          </Select>
        )}
        <button
          onClick={() => descargarListaComoWord(
            "Proyectos",
            ["Nº", "Proyecto", "Cliente", "Estado presupuesto", "Estado trabajo", "Importe"],
            filtered.map((p) => [String(p.numero), p.nombre, clienteNombre(p.clienteId), p.estadoPresupuesto, p.estadoTrabajo, money(p.importePresupuesto)])
          )}
          className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
        >
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nº</th>
              <th className="px-4 py-3 font-semibold">Proyecto</th>
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Presupuesto</th>
              <th className="px-4 py-3 font-semibold">Estado trabajo</th>
              <th className="px-4 py-3 font-semibold">Ciudad reparto</th>
              <th className="px-4 py-3 font-semibold text-right">Importe</th>
              <th className="px-4 py-3 font-semibold">Entrega prevista</th>
              <th className="px-4 py-3 font-semibold">Diferencia</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={10} className="px-4 py-10 text-center text-slate-400 text-sm">No hay proyectos que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((p) => {
              const diff = p.fechaEntregado && p.fechaEntregaPrevista ? daysDiff(p.fechaEntregaPrevista, p.fechaEntregado) : null;
              const previstaVencida = !p.fechaEntregado && p.fechaEntregaPrevista && new Date(p.fechaEntregaPrevista) < new Date();
              return (
                <tr
                  key={p.id}
                  onClick={() => { setDetailId(p.id); setView("detail"); }}
                  className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition"
                >
                  <td className="px-4 py-3 font-mono-num text-slate-500">#{p.numero}</td>
                  <td className="px-4 py-3 font-medium text-slate-800">{p.nombre}</td>
                  <td className="px-4 py-3 text-slate-600">{clienteNombre(p.clienteId)}</td>
                  <td className="px-4 py-3"><Badge className={ESTADO_PRESUPUESTO_STYLE[p.estadoPresupuesto]}>{p.estadoPresupuesto}</Badge></td>
                  <td className="px-4 py-3"><Badge className={ESTADO_TRABAJO_STYLE[p.estadoTrabajo]}>{p.estadoTrabajo}</Badge></td>
                  <td className="px-4 py-3 text-slate-600">{ciudadRepartoDe(p) || (p.estadoLogistica === "Recogida en fábrica" ? "Recogida" : "—")}</td>
                  <td className="px-4 py-3 text-right font-mono-num">{money(p.importePresupuesto)}</td>
                  <td className={`px-4 py-3 ${previstaVencida ? "text-rose-600 font-semibold" : "text-slate-500"}`}>{fmtDate(p.fechaEntregaPrevista)}</td>
                  <td className="px-4 py-3">
                    {diff === null ? <span className="text-slate-300">—</span> : diff === 0 ? (
                      <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">A tiempo</Badge>
                    ) : diff < 0 ? (
                      <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">{Math.abs(diff)}d antes</Badge>
                    ) : (
                      <Badge className="bg-rose-50 text-rose-700 ring-rose-200">{diff}d después</Badge>
                    )}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {isAdmin && (
                      <button
                        onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar el proyecto #${p.numero}?`)) onDelete(p.id); }}
                        className="text-slate-300 hover:text-rose-500 transition"
                      >
                        <Trash2 size={15} />
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ProyectoForm({ initial, clientes, nextNumero, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, nombre: "", clienteId: clientes[0]?.id || "", tipoVenta: "Preventa",
      especificaciones: "", estadoPresupuesto: "Esperando presupuesto", presupuestoFirmado: false,
      condicionesCumplidas: false, importePresupuesto: "", estadoTrabajo: "Pendiente de aceptación",
      ubicacion: "", fechaSolicitud: new Date().toISOString().slice(0, 10), fechaEntregaPrevista: "", fechaEntregado: "",
      provinciaReparto: "", ciudadRepartoManual: "", llevaInstalacion: false,
      diasPlazoMateriales: "", fechaFabricacion: "", fechaMontaje: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value });
  const [errorMsg, setErrorMsg] = useState("");

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.nombre.trim() || !f.clienteId) {
      setErrorMsg("Faltan campos obligatorios: Cliente y Nombre / descripción.");
      alert("Faltan campos obligatorios: Cliente y Nombre / descripción del proyecto.");
      return;
    }
    setErrorMsg("");
    onSave({ ...f, importePresupuesto: parseFloat(f.importePresupuesto) || 0, diasPlazoMateriales: parseInt(f.diasPlazoMateriales, 10) || 0 });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5">
        <ChevronLeft size={16} /> Volver
      </button>

      {/* Barra superior grande y visible con el botón de acción */}
      <div className="bg-white border-2 border-slate-300 rounded-lg px-5 py-5 mb-4">
        <div className="flex items-center gap-3 mb-4">
          <h1 className="font-display text-xl font-extrabold text-slate-900">
            {initial ? `Editar proyecto #${initial.numero}` : "Alta de proyecto/obra"}
          </h1>
          {!initial && <Badge className="bg-slate-100 text-slate-500 ring-slate-200 font-mono-num">Nº {nextNumero()} (automático)</Badge>}
        </div>
        <div className="flex flex-wrap gap-3">
          <button type="button" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "3px solid #256E46" }} className="whitespace-nowrap flex items-center justify-center gap-2 text-base font-extrabold px-6 py-4 rounded-md flex-1 min-w-[220px]">
            <Save size={20} /> <span>Guardar proyecto</span>
          </button>
          <button type="button" onClick={onCancel} className="px-6 py-4 rounded-md text-base font-bold text-slate-700 hover:bg-slate-100 border-2 border-slate-300">Cancelar</button>
        </div>
      </div>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Cliente" required>
            <Select value={f.clienteId} onChange={set("clienteId")} required>
              {clientes.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
            </Select>
          </Field>
          <Field label="Nombre / descripción breve" required>
            <TextInput value={f.nombre} onChange={set("nombre")} required />
          </Field>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Tipo">
            <Select value={f.tipoVenta} onChange={set("tipoVenta")}>
              {TIPO_VENTA.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Ubicación / obra">
            <TextInput value={f.ubicacion} onChange={set("ubicacion")} />
          </Field>
          <Field label="Importe presupuesto (€)">
            <TextInput type="number" step="0.01" min="0" value={f.importePresupuesto} onChange={set("importePresupuesto")} />
          </Field>
        </div>

        <Field label="Especificaciones del trabajo">
          <TextArea rows={3} value={f.especificaciones} onChange={set("especificaciones")} />
        </Field>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Estado del presupuesto">
            <Select value={f.estadoPresupuesto} onChange={set("estadoPresupuesto")}>
              {ESTADO_PRESUPUESTO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Estado del trabajo">
            <Select value={f.estadoTrabajo} onChange={set("estadoTrabajo")}>
              {ESTADO_TRABAJO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Reparto / Recogida">
            <Select value={f.estadoLogistica || "Sin definir"} onChange={set("estadoLogistica")}>
              {ESTADO_LOGISTICA.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
        </div>

        {f.estadoLogistica === "Reparto (camión)" && (
          <div className="grid grid-cols-2 gap-4">
            <Field label="Provincia / zona de reparto">
              <Select value={f.provinciaReparto || ""} onChange={set("provinciaReparto")}>
                <option value="">Selecciona...</option>
                {PROVINCIAS_REPARTO.map((t) => <option key={t}>{t}</option>)}
              </Select>
            </Field>
            {f.provinciaReparto === "Otra ciudad..." && (
              <Field label="Escribe la ciudad">
                <TextInput value={f.ciudadRepartoManual} onChange={set("ciudadRepartoManual")} placeholder="Ej: Córdoba" />
              </Field>
            )}
          </div>
        )}

        <div className="flex gap-6">
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={f.presupuestoFirmado} onChange={set("presupuestoFirmado")} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
            Presupuesto firmado
          </label>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={f.condicionesCumplidas} onChange={set("condicionesCumplidas")} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
            Condiciones cumplidas
          </label>
          <label className="flex items-center gap-2 text-sm text-slate-700">
            <input type="checkbox" checked={f.llevaInstalacion} onChange={set("llevaInstalacion")} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
            Este proyecto lleva instalación
          </label>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Fecha de solicitud">
            <TextInput type="date" value={f.fechaSolicitud} onChange={set("fechaSolicitud")} />
          </Field>
          <Field label="Fecha de entrega prevista">
            <TextInput type="date" value={f.fechaEntregaPrevista} onChange={set("fechaEntregaPrevista")} />
          </Field>
          <Field label="Fecha entregado">
            <TextInput type="date" value={f.fechaEntregado} onChange={set("fechaEntregado")} />
          </Field>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Días que tardan en llegar los materiales">
            <TextInput type="number" value={f.diasPlazoMateriales} onChange={set("diasPlazoMateriales")} placeholder="Ej: 15" />
          </Field>
          <Field label="Fecha de fabricación">
            <TextInput type="date" value={f.fechaFabricacion} onChange={set("fechaFabricacion")} />
            <p className="text-xs text-slate-400 mt-1">Se sugiere sola al aceptar el presupuesto, según los días de arriba. Puedes cambiarla cuando quieras.</p>
          </Field>
          <Field label="Fecha de montaje">
            <TextInput type="date" value={f.fechaMontaje} onChange={set("fechaMontaje")} />
          </Field>
        </div>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100 border border-slate-300">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="whitespace-nowrap flex items-center gap-1.5 text-sm font-bold px-5 py-2.5 rounded-md">
            <Save size={15} /> <span>Guardar proyecto</span>
          </button>
        </div>
      </form>
    </div>
  );
}

function ProyectoDetail({ proyecto, cliente, facturas, ingresos, materiales, articulos, pedidos, proveedores, openPedido, onBack, onEdit, onDelete, onInlineUpdate, isAdmin, onRegistrarPago, onRemovePago, onUsarArticulo, onQuitarArticuloUsado, onRegistrarIngreso, instalacion, onVerInstalacion, onGenerarPedidoFaltante }) {
  const [tab, setTab] = useState("datos");
  const gastos = proyecto.gastos || [];
  const horas = proyecto.registroHorario || [];
  const totalGastos = gastos.reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
  const totalHoras = horas.reduce((s, h) => s + (parseFloat(h.tiempo) || 0), 0);
  const diff = proyecto.fechaEntregado && proyecto.fechaEntregaPrevista ? daysDiff(proyecto.fechaEntregaPrevista, proyecto.fechaEntregado) : null;
  const ciudadReparto = proyecto.estadoLogistica === "Reparto (camión)"
    ? (proyecto.provinciaReparto === "Otra ciudad..." ? proyecto.ciudadRepartoManual : proyecto.provinciaReparto)
    : null;
  const checklist = normalizarChecklist(proyecto.checklistMateriales);
  const despieceAgrupado = calcularDespieceConjuntoProyecto(proyecto.techos);
  const despieceConStock = compararDespieceConStock(despieceAgrupado, materiales);
  const generarPedidoDeFaltantes = () => {
    const faltantes = despieceConStock.filter((d) => d.falta > 0.01);
    if (faltantes.length === 0) { alert("No falta nada: el stock cubre todo el despiece calculado."); return; }
    const lineas = faltantes.map((d) => ({
      id: uid(),
      modo: d.material ? "catalogo" : "libre",
      materialId: d.material ? d.material.id : "",
      referencia: d.material ? "" : d.perfil,
      ancho: "", alto: "",
      cantidad: Math.ceil(d.falta),
      precio: d.material ? (d.material.precioCompra || "") : "",
      estado: "Solicitado",
    }));
    onGenerarPedidoFaltante(null, lineas, proyecto.id, `Pedido de faltantes generado desde el despiece conjunto del proyecto #${proyecto.numero}. Revisa proveedor y precios antes de enviarlo.`);
  };
  const totalRecibido = (ingresos || []).reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);
  const importePresupuesto = parseFloat(proyecto.importePresupuesto) || 0;
  const saldoPendiente = importePresupuesto - totalRecibido;
  const pedidosMarcados = checklist.filter((c) => c.estado === "si").length;
  const setChecklistEstado = (itemId, valor) => {
    const next = checklist.map((c) => (c.id === itemId ? { ...c, estado: c.estado === valor ? null : valor } : c));
    onInlineUpdate(proyecto.id, { checklistMateriales: next });
  };

  const [gForm, setGForm] = useState({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  const [hForm, setHForm] = useState({ tarea: "", tiempo: "", empleado: "", costeHora: "" });
  const [pForm, setPForm] = useState({ importe: "", fecha: new Date().toISOString().slice(0, 10), formaPago: "Transferencia", tipo: "Definitiva" });

  const totalFacturado = facturas.reduce((s, f) => s + (parseFloat(f.total) || 0), 0);
  const saldoPresupuesto = (proyecto.importePresupuesto || 0) - totalFacturado;

  const submitPago = (e) => {
    e.preventDefault();
    const importe = parseFloat(pForm.importe) || 0;
    if (importe <= 0) return;
    onRegistrarPago({ importe, fecha: pForm.fecha, formaPago: pForm.formaPago, tipo: pForm.tipo });
    setPForm({ importe: "", fecha: new Date().toISOString().slice(0, 10), formaPago: "Transferencia", tipo: "Definitiva" });
  };


  const addGasto = (e) => {
    e.preventDefault();
    if (!gForm.proveedor.trim()) return;
    const next = [...gastos, { ...gForm, id: uid(), importe: parseFloat(gForm.importe) || 0 }];
    onInlineUpdate(proyecto.id, { gastos: next });
    setGForm({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  };
  const removeGasto = (id) => onInlineUpdate(proyecto.id, { gastos: gastos.filter((g) => g.id !== id) });

  const addHora = (e) => {
    e.preventDefault();
    if (!hForm.tarea.trim()) return;
    const next = [...horas, { ...hForm, id: uid(), tiempo: parseFloat(hForm.tiempo) || 0, costeHora: parseFloat(hForm.costeHora) || 0 }];
    onInlineUpdate(proyecto.id, { registroHorario: next });
    setHForm({ tarea: "", tiempo: "", empleado: "", costeHora: "" });
  };
  const removeHora = (id) => onInlineUpdate(proyecto.id, { registroHorario: horas.filter((h) => h.id !== id) });
  const costeManoObra = horas.reduce((s, h) => s + (parseFloat(h.tiempo) || 0) * (parseFloat(h.costeHora) || 0), 0);

  const articulosUsados = proyecto.articulosUsados || [];
  const [uForm, setUForm] = useState({ articuloId: articulos?.[0]?.id || "", cantidad: "1" });
  const articuloInfo = (id) => articulos.find((a) => a.id === id);
  const costeArticulo = (art) => (art?.materiales || []).reduce((s, l) => s + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0);
  const totalCosteUsado = articulosUsados.reduce((s, u) => {
    const art = articuloInfo(u.articuloId);
    return s + costeArticulo(art) * u.cantidad;
  }, 0);

  const submitUsoArticulo = (e) => {
    e.preventDefault();
    const cantidad = parseFloat(uForm.cantidad) || 0;
    if (!uForm.articuloId || cantidad <= 0) return;
    onUsarArticulo(uForm.articuloId, cantidad);
    setUForm({ articuloId: uForm.articuloId, cantidad: "1" });
  };

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5">
        <ChevronLeft size={16} /> Volver al listado
      </button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2 flex-wrap">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">#{proyecto.numero}</span>
            <h1 className="font-display text-2xl font-extrabold text-slate-900">{proyecto.nombre}</h1>
          </div>
          <p className="text-sm text-slate-500 mt-1">
            {cliente?.nombre || "Cliente no encontrado"} {proyecto.ubicacion && `· ${proyecto.ubicacion}`}
          </p>
        </div>
        <div className="flex gap-2 shrink-0">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
            <Pencil size={14} /> Editar
          </button>
          {isAdmin && (
          <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50">
            <Trash2 size={14} /> Eliminar
          </button>
          )}
        </div>
      </div>

      <div className="flex flex-wrap gap-2 mb-6">
        <Badge className={ESTADO_PRESUPUESTO_STYLE[proyecto.estadoPresupuesto]}>{proyecto.estadoPresupuesto}</Badge>
        <Badge className={ESTADO_TRABAJO_STYLE[proyecto.estadoTrabajo]}>{proyecto.estadoTrabajo}</Badge>
        {proyecto.presupuestoFirmado && <Badge className="bg-sky-50 text-sky-700 ring-sky-200">Presupuesto firmado</Badge>}
        {proyecto.condicionesCumplidas && <Badge className="bg-sky-50 text-sky-700 ring-sky-200">Condiciones cumplidas</Badge>}
        {ciudadReparto && <Badge className="bg-amber-50 text-amber-700 ring-amber-200">🚚 Reparto: {ciudadReparto}</Badge>}
        {proyecto.estadoLogistica === "Recogida en fábrica" && <Badge className="bg-slate-50 text-slate-700 ring-slate-200">Recogida en fábrica</Badge>}
      </div>

      {/* KPI strip */}
      <div className="grid grid-cols-4 gap-3 mb-6">
        <Kpi label="Importe presupuesto" value={money(proyecto.importePresupuesto)} />
        <Kpi label="Gastos asociados" value={money(totalGastos)} sub={totalGastos > proyecto.importePresupuesto ? "Por encima del presupuesto" : null} />
        <Kpi label="Horas registradas" value={`${totalHoras.toFixed(1)} h`} />
        <Kpi
          label="Diferencia entrega"
          value={diff === null ? "—" : diff === 0 ? "A tiempo" : diff < 0 ? `${Math.abs(diff)}d antes` : `${diff}d después`}
          tone={diff !== null && diff > 0 ? "bad" : diff !== null ? "good" : "neutral"}
        />
      </div>

      {importePresupuesto > 0 && (
        <div className={`flex flex-wrap items-center gap-2 px-4 py-3 rounded-md mb-6 text-sm font-semibold ${saldoPendiente > 0 ? "bg-amber-50 border border-amber-300 text-amber-800" : "bg-emerald-50 border border-emerald-200 text-emerald-700"}`}>
          <Wallet size={16} />
          {saldoPendiente > 0 ? (
            <>Recibido {money(totalRecibido)} de {money(importePresupuesto)} — quedan <span className="text-base">{money(saldoPendiente)}</span> por cobrar.</>
          ) : (
            <>Presupuesto cobrado por completo ({money(totalRecibido)} de {money(importePresupuesto)}).</>
          )}
          {saldoPendiente > 0 && (
            <button
              onClick={() => onRegistrarIngreso(proyecto, cliente?.nombre, saldoPendiente)}
              style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
              className="ml-auto flex items-center gap-1.5 text-xs font-semibold hover:opacity-90 px-3 py-1.5 rounded-md shrink-0"
            >
              <Plus size={13} /> Meter dinero
            </button>
          )}
        </div>
      )}
      {proyecto.llevaInstalacion && (
        <div className="flex flex-wrap items-center gap-2 px-4 py-3 rounded-md mb-6 text-sm font-semibold bg-violet-50 border border-violet-200 text-violet-700">
          <Wrench size={16} />
          Este proyecto lleva instalación.
          {instalacion ? (
            <button onClick={() => onVerInstalacion(instalacion.id)} className="ml-auto flex items-center gap-1.5 text-xs font-semibold text-white bg-violet-600 hover:bg-violet-700 px-3 py-1.5 rounded-md shrink-0">
              Ver ficha de instalación ({instalacion.estado})
            </button>
          ) : (
            <span className="ml-auto text-xs text-violet-500">Creando ficha…</span>
          )}
        </div>
      )}

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        {[
          { id: "datos", label: "Datos", icon: FileText },
          { id: "articulos", label: `Artículos usados (${articulosUsados.length})`, icon: Layers },
          { id: "comparativa", label: "Presupuesto vs Real", icon: Wallet },
          { id: "gastos", label: `Gastos asociados (${gastos.length})`, icon: Receipt },
          { id: "horas", label: `Registro horario (${horas.length})`, icon: Timer },
          { id: "pagos", label: `Pagos / Facturas (${facturas.length})`, icon: Wallet },
          { id: "pedidos", label: `Pedidos de materiales (${pedidos.length})`, icon: ClipboardList },
          { id: "checklist", label: `Qué lleva la obra (${checklist.filter((c) => c.estado).length}/${checklist.length})`, icon: CheckCircle2 },
          { id: "despiece", label: `Despiece de techos (${(proyecto.techos || []).length})`, icon: Ruler },
        ].map((t) => (
          <button
            key={t.id}
            onClick={() => setTab(t.id)}
            className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${
              tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"
            }`}
          >
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "datos" && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
            <InfoRow icon={<Hash size={14} />} label="Tipo" value={proyecto.tipoVenta} />
            <InfoRow icon={<MapPin size={14} />} label="Ubicación" value={proyecto.ubicacion || "—"} />
            <InfoRow icon={<FileText size={14} />} label="Fecha solicitud" value={fmtDate(proyecto.fechaSolicitud)} />
            <InfoRow icon={<FileText size={14} />} label="Entrega prevista" value={fmtDate(proyecto.fechaEntregaPrevista)} />
            <InfoRow icon={<CheckCircle2 size={14} />} label="Fecha entregado" value={fmtDate(proyecto.fechaEntregado)} />
            <InfoRow icon={<MapPin size={14} />} label="Ciudad de reparto" value={ciudadReparto || "—"} />
          </div>
          {proyecto.especificaciones && (
            <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Especificaciones</span>
              {proyecto.especificaciones}
            </div>
          )}
        </CornerFrame>
      )}

      {tab === "articulos" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {articulosUsados.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Todavía no se ha usado ningún artículo en este proyecto.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Artículo</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Cantidad</th>
                    <th className="px-4 py-2.5 font-semibold">Fecha</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Coste materiales</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {articulosUsados.map((u) => {
                    const art = articuloInfo(u.articuloId);
                    return (
                      <tr key={u.id} className="border-b border-slate-100 last:border-0">
                        <td className="px-4 py-2.5 font-medium text-slate-800">{u.articuloNombre}{!art && <span className="text-slate-400"> (eliminado del catálogo)</span>}</td>
                        <td className="px-4 py-2.5 text-right font-mono-num">{u.cantidad}</td>
                        <td className="px-4 py-2.5 text-slate-500">{fmtDate(u.fecha)}</td>
                        <td className="px-4 py-2.5 text-right font-mono-num">{money(costeArticulo(art) * u.cantidad)}</td>
                        <td className="px-4 py-2.5 text-right">
                          <button onClick={() => { if (confirm("¿Deshacer este uso? Se repondrá el stock de los materiales descontados.")) onQuitarArticuloUsado(u.id); }} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5" colSpan={3}>Coste total de materiales usados</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(totalCosteUsado)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>

          {(!articulos || articulos.length === 0) ? (
            <p className="text-sm text-slate-400 px-1">No hay artículos dados de alta todavía en el módulo Artículos.</p>
          ) : (
            <form onSubmit={submitUsoArticulo} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-6 gap-2 items-end">
              <div className="col-span-4">
                <Field label="Artículo del catálogo">
                  <Select value={uForm.articuloId} onChange={(e) => setUForm({ ...uForm, articuloId: e.target.value })}>
                    {articulos.map((a) => <option key={a.id} value={a.id}>{a.numero} — {a.nombre}</option>)}
                  </Select>
                </Field>
              </div>
              <div className="col-span-1">
                <Field label="Cantidad"><TextInput type="number" min="0.01" step="0.01" value={uForm.cantidad} onChange={(e) => setUForm({ ...uForm, cantidad: e.target.value })} /></Field>
              </div>
              <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="whitespace-nowrap flex items-center justify-center gap-1 text-sm font-bold px-3 py-2 rounded-md h-[38px] col-span-1">
                <Plus size={15} /> Usar
              </button>
            </form>
          )}
          <p className="text-xs text-slate-400 px-1">Al usar un artículo aquí, se descuenta automáticamente del Stock cada material que lo compone, multiplicado por la cantidad.</p>
        </div>
      )}

      {tab === "comparativa" && (() => {
        const costeMateriales = totalGastos + totalCosteUsado;
        const costeTotalReal = costeMateriales + costeManoObra;
        const presupuesto = parseFloat(proyecto.importePresupuesto) || 0;
        const margen = presupuesto - costeTotalReal;
        const margenPct = presupuesto ? (margen / presupuesto) * 100 : null;
        return (
          <div className="space-y-4">
            <div className="grid grid-cols-2 gap-3">
              <Kpi label="Presupuesto" value={money(presupuesto)} />
              <Kpi label="Coste real total" value={money(costeTotalReal)} tone={costeTotalReal > presupuesto ? "bad" : "good"} />
              <Kpi label="Coste materiales (gastos + artículos)" value={money(costeMateriales)} />
              <Kpi label="Coste mano de obra" value={money(costeManoObra)} />
            </div>

            <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
              <div className="flex items-center justify-between mb-4">
                <span className="text-sm font-semibold text-slate-600">Margen</span>
                <span className={`font-mono-num text-2xl font-extrabold ${margen >= 0 ? "text-emerald-600" : "text-rose-600"}`}>
                  {money(margen)} {margenPct !== null && <span className="text-base">({margenPct.toFixed(1)}%)</span>}
                </span>
              </div>
              <div className="space-y-2 text-sm">
                <div className="flex justify-between"><span className="text-slate-500">Gastos asociados (proveedores/facturas)</span><span className="font-mono-num">{money(totalGastos)}</span></div>
                <div className="flex justify-between"><span className="text-slate-500">Coste de artículos usados (materiales)</span><span className="font-mono-num">{money(totalCosteUsado)}</span></div>
                <div className="flex justify-between pt-2 border-t border-slate-100"><span className="text-slate-500">Coste mano de obra ({totalHoras.toFixed(1)} h)</span><span className="font-mono-num">{money(costeManoObra)}</span></div>
                <div className="flex justify-between pt-2 border-t border-slate-200 font-semibold"><span className="text-slate-700">Coste real total</span><span className="font-mono-num">{money(costeTotalReal)}</span></div>
              </div>
            </CornerFrame>
            <p className="text-xs text-slate-400 px-1">El coste de mano de obra se calcula a partir de las horas y el coste €/hora que introduces en la pestaña "Registro horario" de cada tarea.</p>
          </div>
        );
      })()}

      {tab === "gastos" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {gastos.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin gastos registrados todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                    <th className="px-4 py-2.5 font-semibold">Producto / Tarea</th>
                    <th className="px-4 py-2.5 font-semibold">Factura</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {gastos.map((g) => (
                    <tr key={g.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-medium text-slate-800">{g.proveedor}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.producto || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.facturaAsociada || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.estadoFactura}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money(g.importe)}</td>
                      <td className="px-4 py-2.5 text-right">
                        <button onClick={() => removeGasto(g.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5" colSpan={4}>Total gastos asociados</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(totalGastos)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>

          <form onSubmit={addGasto} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-6 gap-2 items-end">
            <Field label="Proveedor"><TextInput value={gForm.proveedor} onChange={(e) => setGForm({ ...gForm, proveedor: e.target.value })} /></Field>
            <Field label="Producto / Tarea"><TextInput value={gForm.producto} onChange={(e) => setGForm({ ...gForm, producto: e.target.value })} /></Field>
            <Field label="Factura asociada"><TextInput value={gForm.facturaAsociada} onChange={(e) => setGForm({ ...gForm, facturaAsociada: e.target.value })} /></Field>
            <Field label="Estado factura">
              <Select value={gForm.estadoFactura} onChange={(e) => setGForm({ ...gForm, estadoFactura: e.target.value })}>
                <option>Pendiente</option><option>Pagada</option>
              </Select>
            </Field>
            <Field label="Importe (€)"><TextInput type="number" step="0.01" value={gForm.importe} onChange={(e) => setGForm({ ...gForm, importe: e.target.value })} /></Field>
            <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]">
              <Plus size={15} /> Añadir
            </button>
          </form>
        </div>
      )}

      {tab === "horas" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {horas.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin registro horario todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Tarea</th>
                    <th className="px-4 py-2.5 font-semibold">Empleado</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Tiempo (h)</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Coste €/h</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Subtotal</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {horas.map((h) => (
                    <tr key={h.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-medium text-slate-800">{h.tarea}</td>
                      <td className="px-4 py-2.5 text-slate-600">{h.empleado || "—"}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{Number(h.tiempo).toFixed(1)}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money(h.costeHora)}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money((parseFloat(h.tiempo) || 0) * (parseFloat(h.costeHora) || 0))}</td>
                      <td className="px-4 py-2.5 text-right">
                        <button onClick={() => removeHora(h.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5">Total</td>
                    <td></td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{totalHoras.toFixed(1)}</td>
                    <td></td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(costeManoObra)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>

          <form onSubmit={addHora} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-5 gap-2 items-end">
            <Field label="Tarea"><TextInput value={hForm.tarea} onChange={(e) => setHForm({ ...hForm, tarea: e.target.value })} /></Field>
            <Field label="Empleado"><TextInput value={hForm.empleado} onChange={(e) => setHForm({ ...hForm, empleado: e.target.value })} /></Field>
            <Field label="Tiempo (horas)"><TextInput type="number" step="0.25" value={hForm.tiempo} onChange={(e) => setHForm({ ...hForm, tiempo: e.target.value })} /></Field>
            <Field label="Coste €/hora"><TextInput type="number" step="0.5" value={hForm.costeHora} onChange={(e) => setHForm({ ...hForm, costeHora: e.target.value })} placeholder="Ej. 15" /></Field>
            <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]">
              <Plus size={15} /> Añadir

            </button>
          </form>
        </div>
      )}

      {tab === "pagos" && (
        <div className="space-y-4">
          <div className="grid grid-cols-3 gap-3">
            <Kpi label="Importe presupuesto" value={money(proyecto.importePresupuesto)} />
            <Kpi label="Total facturado" value={money(totalFacturado)} tone="good" />
            <Kpi label="Pendiente de facturar" value={money(Math.max(saldoPresupuesto, 0))} tone={saldoPresupuesto > 0 ? "bad" : "neutral"} />
          </div>

          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
            Cada ingreso genera su propia factura por el importe exacto recibido (no por el total del presupuesto).
          </div>

          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {facturas.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin facturas ni pagos registrados todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Nº Factura</th>
                    <th className="px-4 py-2.5 font-semibold">Tipo</th>
                    <th className="px-4 py-2.5 font-semibold">Fecha</th>
                    <th className="px-4 py-2.5 font-semibold">Forma de pago</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {facturas.map((f) => (
                    <tr key={f.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-mono-num text-slate-500">{f.numero}</td>
                      <td className="px-4 py-2.5 text-slate-600">{f.tipo}</td>
                      <td className="px-4 py-2.5 text-slate-500">{fmtDate(f.fecha)}</td>
                      <td className="px-4 py-2.5 text-slate-600">{f.pagos?.[0]?.formaPago || "—"}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num font-semibold">{money(f.total)}</td>
                      <td className="px-4 py-2.5 text-right">
                        {f.pagos?.[0] && (
                          <button onClick={() => onRemovePago(f.id, f.pagos[0].id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5" colSpan={4}>Total facturado</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(totalFacturado)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>

          <form onSubmit={submitPago} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-5 gap-2 items-end">
            <Field label="Importe ingresado (€)"><TextInput type="number" step="0.01" value={pForm.importe} onChange={(e) => setPForm({ ...pForm, importe: e.target.value })} required /></Field>
            <Field label="Fecha"><TextInput type="date" value={pForm.fecha} onChange={(e) => setPForm({ ...pForm, fecha: e.target.value })} /></Field>
            <Field label="Forma de pago">
              <Select value={pForm.formaPago} onChange={(e) => setPForm({ ...pForm, formaPago: e.target.value })}>
                {FORMA_PAGO.map((t) => <option key={t}>{t}</option>)}
              </Select>
            </Field>
            <Field label="Tipo de factura">
              <Select value={pForm.tipo} onChange={(e) => setPForm({ ...pForm, tipo: e.target.value })}>
                {TIPO_FACTURA.map((t) => <option key={t}>{t}</option>)}
              </Select>
            </Field>
            <button type="submit" className="flex items-center justify-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold px-3 py-2 rounded-md h-[38px]">
              <Plus size={15} /> Generar factura
            </button>
          </form>
        </div>
      )}

      {tab === "pedidos" && (
        <div className="space-y-4">
          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
            Estos son los pedidos de materiales dados de alta en el módulo Pedidos y vinculados a este proyecto. Para crear uno nuevo, ve a Pedidos → Nuevo pedido y selecciona este proyecto.
          </div>
          {checklist.some((c) => !c.estado) && (
            <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold">
              ⚠ El checklist "Qué lleva la obra" no está completo todavía. No podrás crear un pedido nuevo para este proyecto hasta rellenarlo (pestaña "Qué lleva la obra").
            </div>
          )}
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {pedidos.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin pedidos de materiales vinculados a este proyecto todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Nº Pedido</th>
                    <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                    <th className="px-4 py-2.5 font-semibold">Líneas</th>
                    <th className="px-4 py-2.5 font-semibold">Entrega prevista</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {pedidos.map((p) => (
                    <tr key={p.id} onClick={() => openPedido(p.id)} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                      <td className="px-4 py-2.5 font-mono-num text-slate-500">#{p.numero}</td>
                      <td className="px-4 py-2.5 font-medium text-slate-800">{proveedores.find((pr) => pr.id === p.proveedorId)?.nombre || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{(p.lineas || []).length} línea{(p.lineas || []).length === 1 ? "" : "s"}</td>
                      <td className="px-4 py-2.5 text-slate-500">{fmtDate(p.fechaEntregaPrevista)}</td>
                      <td className="px-4 py-2.5"><Badge className={ESTADO_PEDIDO_STYLE[p.estado]}>{p.estado}</Badge></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {tab === "checklist" && (
        <div className="space-y-4">
          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
            Marca "Sí" si esta obra lleva esa categoría, o "No" si no la lleva — así no hay confusión entre "no decidido todavía" y "no lleva".
          </div>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden divide-y divide-slate-100">
            {checklist.map((item) => (
              <div key={item.id} className="flex items-center gap-3 px-4 py-3.5">
                <span className="text-sm font-medium text-slate-800 flex-1">{item.nombre}</span>
                <button
                  type="button"
                  onClick={() => setChecklistEstado(item.id, "si")}
                  className={`px-4 py-1.5 rounded-md text-sm font-semibold border transition cursor-pointer select-none ${item.estado === "si" ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-300 text-slate-500 hover:bg-slate-50"}`}
                >
                  Sí
                </button>
                <button
                  type="button"
                  onClick={() => setChecklistEstado(item.id, "no")}
                  className={`px-4 py-1.5 rounded-md text-sm font-semibold border transition cursor-pointer select-none ${item.estado === "no" ? "bg-rose-600 border-rose-600 text-white" : "bg-white border-slate-300 text-slate-500 hover:bg-slate-50"}`}
                >
                  No
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {tab === "despiece" && (
        <div className="space-y-4">
          {(proyecto.techos || []).length === 0 ? (
            <p className="text-sm text-slate-400">
              Esta obra todavía no tiene techos calculados. Se añaden desde Mediciones (sección "Techos") y, al pasar la medición a presupuesto y aceptarlo, se traen aquí solos.
            </p>
          ) : (
            <>
              <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
                Suma del despiece de los {proyecto.techos.length} techo(s) calculados en esta obra, comparado con el Stock actual. La comparación es solo por nombre y cantidad total — todavía no sabe en qué almacén está cada cosa (eso solo existe hoy para Cristales).
              </div>
              <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
                <table className="w-full text-sm">
                  <thead className="bg-slate-50 text-slate-500 text-xs">
                    <tr>
                      <th className="text-left px-4 py-2">Perfil</th>
                      <th className="text-left px-4 py-2">Piezas</th>
                      <th className="text-left px-4 py-2">Metros necesarios</th>
                      <th className="text-left px-4 py-2">En stock</th>
                      <th className="text-left px-4 py-2">Estado</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {despieceConStock.map((d) => (
                      <tr key={d.perfil}>
                        <td className="px-4 py-2.5 font-medium text-slate-700">{d.perfil}</td>
                        <td className="px-4 py-2.5 text-slate-500">{d.piezas}</td>
                        <td className="px-4 py-2.5 text-slate-500">{d.metros.toFixed(2)} m</td>
                        <td className="px-4 py-2.5 text-slate-500">{d.disponible === null ? "no está en Stock" : `${d.disponible} m`}</td>
                        <td className="px-4 py-2.5">
                          {d.falta > 0.01 ? (
                            <Badge className="bg-rose-50 text-rose-700 ring-rose-200">Faltan {d.falta.toFixed(2)} m</Badge>
                          ) : (
                            <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Cubierto</Badge>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              <button
                onClick={generarPedidoDeFaltantes}
                style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
                className="text-sm font-semibold px-4 py-2.5 rounded-md"
              >
                Generar pedido con lo que falta
              </button>
              <p className="text-xs text-slate-400">Esto abre el formulario de pedido ya relleno — revísalo, pon el proveedor y guárdalo tú. No se envía nada solo.</p>
            </>
          )}
        </div>
      )}
    </div>
  );
}

function Kpi({ label, value, sub, tone = "neutral" }) {
  const toneCls = tone === "bad" ? "text-rose-600" : tone === "good" ? "text-emerald-600" : "text-slate-800";
  return (
    <div className="bg-white border border-slate-200 rounded-lg px-4 py-3">
      <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 mb-1">{label}</div>
      <div className={`font-mono-num text-lg font-bold ${toneCls}`}>{value}</div>
      {sub && <div className="text-[11px] text-rose-500 mt-0.5">{sub}</div>}
    </div>
  );
}

/* ================= PROVEEDORES ================= */

function ProveedoresModulo({ proveedores, materiales, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onInlineUpdate, isAdmin }) {
  const [q, setQ] = useState("");

  const filtered = useMemo(() => {
    if (!q) return proveedores;
    const s = q.toLowerCase();
    return proveedores.filter((p) => `${p.nombre} ${p.apellidos} ${p.email} ${p.movil} ${p.cif}`.toLowerCase().includes(s));
  }, [proveedores, q]);

  if (view === "form") {
    const editing = proveedores.find((p) => p.id === editId) || null;
    return <ProveedorForm initial={editing} onCancel={() => setView(editId ? "detail" : "list")} onSave={onUpsert} />;
  }

  if (view === "detail") {
    const proveedor = proveedores.find((p) => p.id === detailId);
    if (!proveedor) { setView("list"); return null; }
    const mats = materiales.filter((m) => m.proveedorId === proveedor.id);
    return (
      <ProveedorDetail
        proveedor={proveedor}
        materiales={mats}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(proveedor.id); setView("form"); }}
        onDelete={() => onDelete(proveedor.id)}
        isAdmin={isAdmin}
        onInlineUpdate={onInlineUpdate}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<Truck size={20} className="text-[#2E8B57]" />}
        title="Proveedores"
        manualKey="proveedores"
        subtitle={`${proveedores.length} proveedor${proveedores.length === 1 ? "" : "es"} registrado${proveedores.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO PROVEEDOR
      </button>

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nombre, CIF, email, teléfono…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
        </div>
        <button
          onClick={() => descargarListaComoWord(
            "Proveedores",
            ["Empresa / Nombre", "Email", "Teléfono", "Materiales"],
            filtered.map((p) => [p.nombre, p.email || "—", p.movil || "—", String(materiales.filter((m) => m.proveedorId === p.id).length)])
          )}
          className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
        >
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Empresa / Nombre</th>
              <th className="px-4 py-3 font-semibold">Email</th>
              <th className="px-4 py-3 font-semibold">Teléfono</th>
              <th className="px-4 py-3 font-semibold text-right">Materiales</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={5} className="px-4 py-10 text-center text-slate-400 text-sm">No hay proveedores que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((p) => (
              <tr key={p.id} onClick={() => { setDetailId(p.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                <td className="px-4 py-3 font-medium text-slate-800">{p.nombre} {p.apellidos}</td>
                <td className="px-4 py-3 text-slate-600">{p.email || "—"}</td>
                <td className="px-4 py-3 text-slate-600">{p.movil || "—"}</td>
                <td className="px-4 py-3 text-right font-mono-num text-slate-500">{materiales.filter((m) => m.proveedorId === p.id).length}</td>
                <td className="px-4 py-3 text-right">
                  {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar a ${p.nombre}?`)) onDelete(p.id); }} className="text-slate-300 hover:text-rose-500 transition">
                    <Trash2 size={15} />
                  </button>)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ProveedorForm({ initial, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || { id: null, nombre: "", apellidos: "", cif: "", direccion: "", provincia: "", cp: "", email: "", movil: "", comentarios: "" }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const submit = (e) => { e.preventDefault(); if (!f.nombre.trim()) { setErrorMsg("Falta el campo Nombre / Empresa (es obligatorio)."); return; } setErrorMsg(""); onSave(f); };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-4">{initial ? "Editar proveedor" : "Alta de proveedor"}</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Nombre / Empresa" required><TextInput value={f.nombre} onChange={set("nombre")} required /></Field>
          <Field label="Apellidos"><TextInput value={f.apellidos} onChange={set("apellidos")} /></Field>
        </div>
        <div className="grid grid-cols-3 gap-4">
          <Field label="DNI / CIF"><TextInput value={f.cif} onChange={set("cif")} /></Field>
          <Field label="Email"><TextInput type="email" value={f.email} onChange={set("email")} /></Field>
          <Field label="Móvil"><TextInput value={f.movil} onChange={set("movil")} /></Field>
        </div>
        <Field label="Dirección"><TextInput value={f.direccion} onChange={set("direccion")} /></Field>
        <div className="grid grid-cols-2 gap-4">
          <Field label="Provincia"><TextInput value={f.provincia} onChange={set("provincia")} /></Field>
          <Field label="Código postal"><TextInput value={f.cp} onChange={set("cp")} /></Field>
        </div>
        <Field label="Comentarios"><TextArea rows={3} value={f.comentarios} onChange={set("comentarios")} /></Field>
        {errorMsg && (
          <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
            ⚠ {errorMsg}
          </div>
        )}
        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md cursor-pointer select-none"><Save size={15} /> <span>Guardar proveedor</span></button>
        </div>
      </form>
    </div>
  );
}

function ProveedorDetail({ proveedor, materiales, onBack, onEdit, onDelete, onInlineUpdate, isAdmin }) {
  const [tab, setTab] = useState("datos");
  const historico = proveedor.historico || [];
  const [hForm, setHForm] = useState({ articulo: "", cantidad: "", fecha: new Date().toISOString().slice(0, 10), precio: "", comentarios: "" });

  const addHistorico = (e) => {
    e.preventDefault();
    if (!hForm.articulo.trim()) return;
    const next = [{ ...hForm, id: uid(), cantidad: parseFloat(hForm.cantidad) || 0, precio: parseFloat(hForm.precio) || 0 }, ...historico];
    onInlineUpdate(proveedor.id, { historico: next });
    setHForm({ articulo: "", cantidad: "", fecha: new Date().toISOString().slice(0, 10), precio: "", comentarios: "" });
  };
  const removeHistorico = (id) => onInlineUpdate(proveedor.id, { historico: historico.filter((h) => h.id !== id) });

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>
      <div className="flex items-start justify-between mb-6">
        <div>
          <h1 className="font-display text-2xl font-extrabold text-slate-900">{proveedor.nombre} {proveedor.apellidos}</h1>
          <p className="text-sm text-slate-500 mt-1 font-mono-num">{proveedor.cif || "Sin CIF"}</p>
        </div>
        <div className="flex gap-2">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        {[
          { id: "datos", label: "Datos proveedor", icon: FileText },
          { id: "materiales", label: `Materiales suministrados (${materiales.length})`, icon: Package },
          { id: "historico", label: `Histórico artículos/servicios (${historico.length})`, icon: ClipboardList },
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "datos" && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
            <InfoRow icon={<MapPin size={14} />} label="Dirección" value={[proveedor.direccion, proveedor.provincia, proveedor.cp].filter(Boolean).join(", ") || "—"} />
            <InfoRow icon={<Mail size={14} />} label="Email" value={proveedor.email || "—"} />
            <InfoRow icon={<Phone size={14} />} label="Móvil" value={proveedor.movil || "—"} />
          </div>
          {proveedor.comentarios && (
            <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Comentarios</span>
              {proveedor.comentarios}
            </div>
          )}
        </CornerFrame>
      )}

      {tab === "materiales" && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          {materiales.length === 0 ? (
            <p className="px-4 py-8 text-center text-sm text-slate-400">Este proveedor aún no suministra materiales en el catálogo.</p>
          ) : (
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Código</th>
                  <th className="px-4 py-2.5 font-semibold">Descripción</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Stock</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Precio compra</th>
                </tr>
              </thead>
              <tbody>
                {materiales.map((m) => (
                  <tr key={m.id} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-mono-num text-slate-500">{m.codigo}</td>
                    <td className="px-4 py-2.5 font-medium text-slate-800">{m.descripcion}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{m.stockReal ?? 0}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(m.precioCompra)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}

      {tab === "historico" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {historico.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin compras registradas todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Artículo / Servicio</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Cantidad</th>
                    <th className="px-4 py-2.5 font-semibold">Fecha</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Precio</th>
                    <th className="px-4 py-2.5 font-semibold">Comentarios</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {historico.map((h) => (
                    <tr key={h.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-medium text-slate-800">{h.articulo}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{h.cantidad}</td>
                      <td className="px-4 py-2.5 text-slate-500">{fmtDate(h.fecha)}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money(h.precio)}</td>
                      <td className="px-4 py-2.5 text-slate-500">{h.comentarios || "—"}</td>
                      <td className="px-4 py-2.5 text-right"><button onClick={() => removeHistorico(h.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
          <form onSubmit={addHistorico} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-6 gap-2 items-end">
            <Field label="Artículo / Servicio"><TextInput value={hForm.articulo} onChange={(e) => setHForm({ ...hForm, articulo: e.target.value })} /></Field>
            <Field label="Cantidad"><TextInput type="number" value={hForm.cantidad} onChange={(e) => setHForm({ ...hForm, cantidad: e.target.value })} /></Field>
            <Field label="Fecha"><TextInput type="date" value={hForm.fecha} onChange={(e) => setHForm({ ...hForm, fecha: e.target.value })} /></Field>
            <Field label="Precio (€)"><TextInput type="number" step="0.01" value={hForm.precio} onChange={(e) => setHForm({ ...hForm, precio: e.target.value })} /></Field>
            <Field label="Comentarios"><TextInput value={hForm.comentarios} onChange={(e) => setHForm({ ...hForm, comentarios: e.target.value })} /></Field>
            <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Añadir</button>
          </form>
        </div>
      )}
    </div>
  );
}

/* ================= STOCK ================= */

function StockModulo({ materiales, proveedores, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onAddMovimiento, onRemoveMovimiento, isAdmin, onEnviarAPedido, onImportarTarifas }) {
  const [q, setQ] = useState("");
  const [subview, setSubview] = useState("catalogo"); // catalogo | reponer
  const inputTarifasRef = useRef(null);

  const manejarImportarTarifas = async (file) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const filasExcel = leerFilasExcel(buffer);
    const filas = filasExcel.map((fila) => ({
      codigo: String(valorPorCabeceras(fila, ["codigo", "cod", "referencia", "ref"])).trim(),
      descripcion: String(valorPorCabeceras(fila, ["descripcion", "nombre", "material"])).trim(),
      precioVenta: valorPorCabeceras(fila, ["precioventa", "pventa", "pvp", "precio", "tarifa"]),
      precioCompra: valorPorCabeceras(fila, ["preciocompra", "pcompra", "coste", "costo"]),
    })).filter((f) => f.codigo || f.descripcion);
    if (filas.length === 0) {
      alert("No se ha encontrado ninguna fila con código o descripción. Revisa las cabeceras del Excel.");
      return;
    }
    onImportarTarifas(filas);
  };

  const proveedorNombre = (id) => proveedores.find((p) => p.id === id)?.nombre || "—";

  const filtered = useMemo(() => {
    return materiales.filter((m) => {
      if (!q) return true;
      const s = `${m.codigo} ${m.descripcion} ${m.categoria} ${m.familia} ${proveedorNombre(m.proveedorId)}`.toLowerCase();
      return s.includes(q.toLowerCase());
    });
  }, [materiales, q, proveedores]);

  const necesitaReponer = materiales.filter((m) => m.stockReal <= m.stockMinimo);
  const bajoMinimoCount = necesitaReponer.length;

  const gruposPorProveedor = useMemo(() => {
    const map = {};
    necesitaReponer.forEach((m) => {
      const pid = m.proveedorId || "sin-proveedor";
      if (!map[pid]) map[pid] = [];
      map[pid].push(m);
    });
    return map;
  }, [necesitaReponer]);

  const cantidadSugerida = (m) => {
    const objetivo = m.stockOptimo && m.stockOptimo > m.stockMinimo ? m.stockOptimo : m.stockMinimo;
    const falta = objetivo - (m.stockReal || 0);
    return falta > 0 ? falta : 1;
  };

  const crearPedidoGrupo = (proveedorId, mats) => {
    const lineas = mats.map((m) => ({
      id: uid(), modo: "catalogo", materialId: m.id, referencia: "", ancho: "", alto: "",
      cantidad: cantidadSugerida(m), estado: "Solicitado",
    }));
    onEnviarAPedido(proveedorId, lineas);
  };

  const exportarReponerExcel = () => {
    const rows = necesitaReponer.map((m) => ({
      Código: m.codigo,
      Descripción: m.descripcion,
      Proveedor: proveedorNombre(m.proveedorId),
      "Stock real": m.stockReal ?? 0,
      "Stock mínimo": m.stockMinimo ?? 0,
      "Cantidad sugerida": cantidadSugerida(m),
      "Precio compra": m.precioCompra ?? 0,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "A reponer");
    XLSX.writeFile(wb, `materiales_a_reponer_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  if (view === "form") {
    const editing = materiales.find((m) => m.id === editId) || null;
    return <MaterialForm initial={editing} proveedores={proveedores} onCancel={() => setView(editId ? "detail" : "list")} onSave={onUpsert} />;
  }

  if (view === "detail") {
    const material = materiales.find((m) => m.id === detailId);
    if (!material) { setView("list"); return null; }
    return (
      <MaterialDetail
        material={material}
        proveedor={proveedores.find((p) => p.id === material.proveedorId)}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(material.id); setView("form"); }}
        onDelete={() => onDelete(material.id)}
        isAdmin={isAdmin}
        onRemoveMovimiento={(movId) => onRemoveMovimiento(material.id, movId)}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<Boxes size={20} className="text-[#2E8B57]" />}
        title="Gestión de Stock"
        manualKey="stock"
        subtitle={`${materiales.length} material${materiales.length === 1 ? "" : "es"} en catálogo`}
      />

      {proveedores.length === 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
          ⚠ No puedes crear un material todavía: primero da de alta al menos un proveedor en la pestaña Proveedores.
        </div>
      )}
      <button
        onClick={() => { setEditId(null); setView("form"); }}
        disabled={proveedores.length === 0}
        title={proveedores.length === 0 ? "Da de alta un proveedor primero" : ""}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO MATERIAL
      </button>

      <input
        ref={inputTarifasRef}
        type="file"
        accept=".xlsx,.xls,.ods,.csv"
        className="hidden"
        onChange={(e) => { manejarImportarTarifas(e.target.files[0]); e.target.value = ""; }}
      />
      <button
        onClick={() => inputTarifasRef.current?.click()}
        className="w-full flex items-center justify-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 py-2.5 rounded-md mb-6 hover:bg-slate-50"
      >
        Importar tarifas desde Excel (masivo) — actualiza precios existentes o da de alta materiales nuevos
      </button>

      {proveedores.length === 0 && (
        <div className="mb-4 px-4 py-3 rounded-md bg-amber-50 text-amber-800 text-sm border border-amber-200">
          Necesitas dar de alta al menos un proveedor antes de crear un material.
        </div>
      )}

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        <button onClick={() => setSubview("catalogo")} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${subview === "catalogo" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          <Boxes size={14} /> Catálogo
        </button>
        <button onClick={() => setSubview("reponer")} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${subview === "reponer" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          <AlertTriangle size={14} /> A reponer {bajoMinimoCount > 0 && `(${bajoMinimoCount})`}
        </button>
      </div>

      {subview === "catalogo" && (
        <>
          <div className="relative flex-1 max-w-sm mb-4">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por código, descripción, categoría, proveedor…"
              className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
          </div>

          <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-3 font-semibold">Código</th>
                  <th className="px-4 py-3 font-semibold">Descripción</th>
                  <th className="px-4 py-3 font-semibold">Proveedor</th>
                  <th className="px-4 py-3 font-semibold text-right">Stock real</th>
                  <th className="px-4 py-3 font-semibold text-right">Mínimo</th>
                  <th className="px-4 py-3 font-semibold text-right">Precio compra</th>
                  <th className="px-4 py-3"></th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 && (
                  <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay materiales que coincidan con la búsqueda.</td></tr>
                )}
                {filtered.map((m) => {
                  const bajo = m.stockReal <= m.stockMinimo;
                  return (
                    <tr key={m.id} onClick={() => { setDetailId(m.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                      <td className="px-4 py-3 font-mono-num text-slate-500">{m.codigo}</td>
                      <td className="px-4 py-3 font-medium text-slate-800">{m.descripcion}</td>
                      <td className="px-4 py-3 text-slate-600">{proveedorNombre(m.proveedorId)}</td>
                      <td className={`px-4 py-3 text-right font-mono-num font-semibold ${bajo ? "text-rose-600" : "text-slate-700"}`}>{m.stockReal ?? 0}</td>
                      <td className="px-4 py-3 text-right font-mono-num text-slate-500">{m.stockMinimo ?? 0}</td>
                      <td className="px-4 py-3 text-right font-mono-num">{money(m.precioCompra)}</td>
                      <td className="px-4 py-3 text-right">
                        {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar el material ${m.codigo}?`)) onDelete(m.id); }} className="text-slate-300 hover:text-rose-500 transition">
                          <Trash2 size={15} />
                        </button>)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}

      {subview === "reponer" && (
        <div>
          <div className="flex items-center justify-between mb-4">
            <p className="text-sm text-slate-500">Materiales cuyo stock real ha caído a su mínimo o por debajo. Genera un pedido directamente o exporta la lista a Excel.</p>
            <button onClick={exportarReponerExcel} disabled={bajoMinimoCount === 0} className="flex items-center gap-1.5 text-sm font-semibold text-white bg-slate-700 hover:bg-slate-800 disabled:bg-slate-300 disabled:cursor-not-allowed px-3.5 py-2 rounded-md shrink-0">
              <FileSpreadsheet size={15} /> Exportar a Excel
            </button>
          </div>

          {bajoMinimoCount === 0 ? (
            <div className="bg-white border border-slate-200 rounded-lg px-4 py-10 text-center text-sm text-slate-400">
              Ningún material está por debajo de su stock mínimo ahora mismo.
            </div>
          ) : (
            <div className="space-y-6">
              {Object.entries(gruposPorProveedor).map(([proveedorId, mats]) => (
                <div key={proveedorId} className="bg-white border border-slate-200 rounded-lg overflow-hidden">
                  <div className="flex items-center justify-between px-4 py-3 bg-slate-50 border-b border-slate-200">
                    <span className="font-semibold text-slate-800 text-sm">{proveedorNombre(proveedorId === "sin-proveedor" ? null : proveedorId)}</span>
                    {proveedorId !== "sin-proveedor" && (
                      <button onClick={() => crearPedidoGrupo(proveedorId, mats)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1.5 text-xs font-semibold hover:opacity-90 px-3 py-1.5 rounded-md">
                        <ClipboardList size={13} /> Crear pedido con estos {mats.length} material{mats.length === 1 ? "" : "es"}
                      </button>
                    )}
                  </div>
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                        <th className="px-4 py-2.5 font-semibold">Código</th>
                        <th className="px-4 py-2.5 font-semibold">Descripción</th>
                        <th className="px-4 py-2.5 font-semibold text-right">Stock real</th>
                        <th className="px-4 py-2.5 font-semibold text-right">Mínimo</th>
                        <th className="px-4 py-2.5 font-semibold text-right">Cantidad sugerida</th>
                        <th className="px-4 py-2.5 font-semibold text-right">Precio compra</th>
                      </tr>
                    </thead>
                    <tbody>
                      {mats.map((m) => (
                        <tr key={m.id} className="border-b border-slate-100 last:border-0">
                          <td className="px-4 py-2.5 font-mono-num text-slate-500">{m.codigo}</td>
                          <td className="px-4 py-2.5 font-medium text-slate-800">{m.descripcion}</td>
                          <td className="px-4 py-2.5 text-right font-mono-num font-semibold text-rose-600">{m.stockReal ?? 0}</td>
                          <td className="px-4 py-2.5 text-right font-mono-num text-slate-500">{m.stockMinimo ?? 0}</td>
                          <td className="px-4 py-2.5 text-right font-mono-num font-semibold text-[#2E8B57]">{cantidadSugerida(m)}</td>
                          <td className="px-4 py-2.5 text-right font-mono-num">{money(m.precioCompra)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function MaterialForm({ initial, proveedores, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, codigo: "", descripcion: "", proveedorId: proveedores[0]?.id || "",
      stockReal: 0, stockMinimo: 0, stockOptimo: 0, color: "", acabadoDescripcion: "",
      longitud: "", ancho: "", alto: "", grueso: "",
      precioCompra: "", precioVenta: "", unidadCompra: "Unidad", categoria: "", familia: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const submit = (e) => {
    e.preventDefault();
    if (!f.codigo.trim() || !f.descripcion.trim()) {
      setErrorMsg("Faltan campos obligatorios: Código de material y/o Descripción.");
      return;
    }
    setErrorMsg("");
    onSave({
      ...f,
      stockReal: parseFloat(f.stockReal) || 0,
      stockMinimo: parseFloat(f.stockMinimo) || 0,
      stockOptimo: parseFloat(f.stockOptimo) || 0,
      precioCompra: parseFloat(f.precioCompra) || 0,
      precioVenta: parseFloat(f.precioVenta) || 0,
    });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">{initial ? "Editar material" : "Alta de material"}</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-3 gap-4">
          <Field label="Código de material" required><TextInput value={f.codigo} onChange={set("codigo")} required /></Field>
          <Field label="Descripción" required><TextInput value={f.descripcion} onChange={set("descripcion")} required /></Field>
          <Field label="Proveedor" required>
            <Select value={f.proveedorId} onChange={set("proveedorId")} required>
              {proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </Select>
          </Field>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Categoría"><TextInput value={f.categoria} onChange={set("categoria")} /></Field>
          <Field label="Familia"><TextInput value={f.familia} onChange={set("familia")} /></Field>
        </div>

        <div className="grid grid-cols-3 gap-4">
          {!initial && <Field label="Stock inicial (unidades)"><TextInput type="number" value={f.stockReal} onChange={set("stockReal")} /></Field>}
          <Field label="Stock mínimo"><TextInput type="number" value={f.stockMinimo} onChange={set("stockMinimo")} /></Field>
          <Field label="Stock óptimo"><TextInput type="number" value={f.stockOptimo} onChange={set("stockOptimo")} /></Field>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Acabado — color"><TextInput value={f.color} onChange={set("color")} /></Field>
          <Field label="Acabado — descripción"><TextInput value={f.acabadoDescripcion} onChange={set("acabadoDescripcion")} /></Field>
        </div>

        <div>
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-1">Dimensiones</span>
          <div className="grid grid-cols-4 gap-3">
            <TextInput type="number" placeholder="Longitud" value={f.longitud} onChange={set("longitud")} />
            <TextInput type="number" placeholder="Ancho" value={f.ancho} onChange={set("ancho")} />
            <TextInput type="number" placeholder="Alto" value={f.alto} onChange={set("alto")} />
            <TextInput type="number" placeholder="Grueso" value={f.grueso} onChange={set("grueso")} />
          </div>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Precio compra (€)"><TextInput type="number" step="0.01" value={f.precioCompra} onChange={set("precioCompra")} /></Field>
          <Field label="Precio venta (€)"><TextInput type="number" step="0.01" value={f.precioVenta} onChange={set("precioVenta")} /></Field>
          <Field label="Unidad de compra">
            <Select value={f.unidadCompra} onChange={set("unidadCompra")}>{UNIDAD_COMPRA.map((u) => <option key={u}>{u}</option>)}</Select>
          </Field>
        </div>

        {errorMsg && (
          <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
            ⚠ {errorMsg}
          </div>
        )}
        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md cursor-pointer select-none"><Save size={15} /> <span>Guardar material</span></button>
        </div>
      </form>
    </div>
  );
}

function MaterialDetail({ material, proveedor, onBack, onEdit, onDelete, onRemoveMovimiento, isAdmin }) {
  const [tab, setTab] = useState("datos");
  const movimientos = material.movimientos || [];
  const bajo = material.stockReal <= material.stockMinimo;

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">{material.codigo}</span>
            <h1 className="font-display text-2xl font-extrabold text-slate-900">{material.descripcion}</h1>
          </div>
          <p className="text-sm text-slate-500 mt-1">{proveedor?.nombre || "Proveedor no encontrado"} {material.categoria && `· ${material.categoria}`}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      {bajo && (
        <div className="mb-4 flex items-center gap-2 px-4 py-2.5 rounded-md bg-rose-50 border border-rose-200 text-rose-700 text-sm">
          <AlertTriangle size={16} /> Stock por debajo del mínimo establecido ({material.stockMinimo})
        </div>
      )}

      <div className="grid grid-cols-4 gap-3 mb-6">
        <Kpi label="Stock real" value={String(material.stockReal ?? 0)} tone={bajo ? "bad" : "good"} />
        <Kpi label="Stock mínimo" value={String(material.stockMinimo ?? 0)} />
        <Kpi label="Stock óptimo" value={String(material.stockOptimo ?? 0)} />
        <Kpi label="Precio compra" value={money(material.precioCompra)} />
      </div>

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        {[
          { id: "datos", label: "Datos", icon: FileText },
          { id: "movimientos", label: `Histórico de movimientos (${movimientos.length})`, icon: Package },
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "datos" && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
            <InfoRow icon={<Hash size={14} />} label="Familia" value={material.familia || "—"} />
            <InfoRow icon={<Euro size={14} />} label="Precio venta" value={money(material.precioVenta)} />
            <InfoRow icon={<Package size={14} />} label="Unidad de compra" value={material.unidadCompra || "—"} />
            <InfoRow icon={<Package size={14} />} label="Acabado" value={[material.color, material.acabadoDescripcion].filter(Boolean).join(" — ") || "—"} />
            <InfoRow icon={<Hash size={14} />} label="Dimensiones (L×A×H×Grueso)" value={[material.longitud, material.ancho, material.alto, material.grueso].filter((v) => v !== "" && v != null).length ? `${material.longitud || 0} × ${material.ancho || 0} × ${material.alto || 0} × ${material.grueso || 0}` : "—"} />
          </div>
        </CornerFrame>
      )}

      {tab === "movimientos" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {movimientos.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin movimientos registrados todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Tipo</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Cantidad</th>
                    <th className="px-4 py-2.5 font-semibold">Proveedor / Cliente</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                    <th className="px-4 py-2.5 font-semibold">Fecha</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {[...movimientos].reverse().map((m) => (
                    <tr key={m.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5">
                        {m.tipo === "entrada" ? (
                          <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200"><ArrowDownCircle size={12} /> Entrada</Badge>
                        ) : (
                          <Badge className="bg-amber-50 text-amber-700 ring-amber-200"><ArrowUpCircle size={12} /> Salida</Badge>
                        )}
                      </td>
                      <td className="px-4 py-2.5 text-right font-mono-num font-semibold">{m.tipo === "entrada" ? "+" : "−"}{m.cantidad}</td>
                      <td className="px-4 py-2.5 text-slate-600">{m.contacto || "—"}{m.pedidoNumero && <span className="text-slate-400"> · Pedido #{m.pedidoNumero}</span>}</td>
                      <td className="px-4 py-2.5"><Badge className={ESTADO_MOVIMIENTO_STYLE[m.estado]}>{m.estado}</Badge></td>
                      <td className="px-4 py-2.5 text-slate-500">{fmtDate(m.fecha)}</td>
                      <td className="px-4 py-2.5 text-right">{isAdmin && <button onClick={() => onRemoveMovimiento(m.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
            Este histórico es solo informativo. Los movimientos de stock se generan automáticamente al marcar un pedido como "Recibido" desde el módulo Pedidos — no se registran a mano aquí.
          </div>
        </div>
      )}
    </div>
  );
}

/* ================= PEDIDOS ================= */

function PedidosModulo({ pedidos, proveedores, materiales, proyectos, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onRecibir, nextNumero, isAdmin, prefill, onClearPrefill, onConfirmarAlbaran, onMarcarEnviado, onCrearDesdeFoto, onCrearPedidoFaltante, tabPrincipal, setTabPrincipal, solicitudes, currentUser, solicitudView, setSolicitudView, solicitudEditId, setSolicitudEditId, solicitudDetailId, setSolicitudDetailId, onUpsertSolicitud, onDeleteSolicitud, onAprobarSolicitud, onRechazarSolicitud, onComentarSolicitud, solicitudPrefill, onClearSolicitudPrefill }) {
  const [q, setQ] = useState("");
  const [leyendoFoto, setLeyendoFoto] = useState(false);
  const [errorFoto, setErrorFoto] = useState("");
  const inputFotoRef = useRef(null);

  const leerFotoPedido = async (file) => {
    setLeyendoFoto(true);
    setErrorFoto("");
    try {
      const base64Data = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("No se pudo leer el archivo"));
        r.readAsDataURL(file);
      });
      const esPdf = file.type === "application/pdf";
      const contentBlock = esPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } }
        : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64Data } };

      const prompt = 'Esto es una lista de materiales o un pedido escrito/impreso (puede ser una foto de notas a mano, una lista de un proveedor, etc). Puede tener varias líneas. Es MUY IMPORTANTE que revises el documento entero, de arriba a abajo, y devuelvas TODAS las líneas, sin saltarte ninguna ni resumir. Devuelve ÚNICAMENTE un JSON válido (sin texto adicional, sin backticks) con este formato exacto: [{"referencia":"nombre o descripción tal cual aparece","ancho":"","alto":"","cantidad":numero}]. Si hay medidas (ancho x alto) inclúyelas, si no, deja esos campos vacíos. No omitas ninguna línea.';

      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 8000,
          messages: [{ role: "user", content: [contentBlock, { type: "text", text: prompt }] }],
        }),
      });
      if (!response.ok) {
        const errBody = await response.text();
        console.error("anthropic-proxy respuesta no válida:", response.status, errBody);
        throw new Error("Respuesta no válida de la API: " + response.status);
      }
      const data = await response.json();
      if (data.error) {
        console.error("Error devuelto por la API:", data.error);
        throw new Error(data.error.message || "Error de la API");
      }
      const textoRespuesta = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
      const limpio = textoRespuesta.replace(/```json|```/g, "").trim();
      const inicio = limpio.indexOf("[");
      const fin = limpio.lastIndexOf("]");
      const jsonCandidato = inicio !== -1 && fin !== -1 ? limpio.slice(inicio, fin + 1) : limpio;
      let items;
      try {
        items = JSON.parse(jsonCandidato);
      } catch (parseErr) {
        console.error("No se pudo parsear el JSON de la IA. Texto recibido:", textoRespuesta);
        throw new Error("La respuesta de la IA no tenía formato válido");
      }
      const lineas = items.map((it) => ({
        id: uid(), modo: "libre", materialId: "", referencia: it.referencia || "",
        ancho: it.ancho || "", alto: it.alto || "", cantidad: it.cantidad || "", precio: "", estado: "Solicitado",
      })).filter((l) => l.referencia);

      if (lineas.length === 0) {
        setErrorFoto("No he podido leer ninguna línea clara en la imagen. Prueba con una foto más nítida.");
        setLeyendoFoto(false);
        return;
      }
      onCrearDesdeFoto(lineas, `Pedido creado a partir de una foto/PDF subida (${file.name}). Revisa las líneas antes de guardar.`);
    } catch (err) {
      console.error("Error leyendo foto de pedido:", err);
      setErrorFoto("No se pudo leer el archivo. Prueba con una foto más clara, con más luz, o inténtalo de nuevo. (" + err.message + ")");
    } finally {
      setLeyendoFoto(false);
    }
  };

  const [estado, setEstado] = useState("");
  const [proveedorFiltro, setProveedorFiltro] = useState("");

  const proveedorNombre = (id) => proveedores.find((p) => p.id === id)?.nombre || "—";
  const materialNombre = (id) => materiales.find((m) => m.id === id)?.descripcion || "—";
  const proyectoNombre = (id) => (id ? proyectos.find((p) => p.id === id)?.nombre : null) || "Stock";

  const filtered = useMemo(() => {
    return pedidos.filter((p) => {
      if (!estado && !q && p.estado === "Recibido") return false;
      if (estado && p.estado !== estado) return false;
      if (proveedorFiltro && p.proveedorId !== proveedorFiltro) return false;
      if (!q) return true;
      const s = `${p.numero} ${proveedorNombre(p.proveedorId)} ${(p.lineas || []).map((l) => l.modo === "libre" ? l.referencia : materialNombre(l.materialId)).join(" ")}`.toLowerCase();
      return s.includes(q.toLowerCase());
    });
  }, [pedidos, q, estado, proveedorFiltro, proveedores, materiales]);

  if (view === "form") {
    const editing = pedidos.find((p) => p.id === editId) || null;
    return (
      <PedidoForm
        initial={editing}
        proveedores={proveedores}
        materiales={materiales}
        proyectos={proyectos}
        nextNumero={nextNumero}
        currentUser={currentUser}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={onUpsert}
        prefill={editing ? null : prefill}
        onClearPrefill={onClearPrefill}
      />
    );
  }

  if (view === "detail") {
    const pedido = pedidos.find((p) => p.id === detailId);
    if (!pedido) { setView("list"); return null; }
    return (
      <PedidoDetail
        pedido={pedido}
        proveedor={proveedores.find((p) => p.id === pedido.proveedorId)}
        materiales={materiales}
        proyectos={proyectos}
        currentUser={currentUser}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(pedido.id); setView("form"); }}
        onDelete={() => onDelete(pedido.id)}
        isAdmin={isAdmin}
        onRecibir={() => onRecibir(pedido.id)}
        onConfirmarAlbaran={(resultado) => onConfirmarAlbaran(pedido.id, resultado)}
        onMarcarEnviado={(metodo) => onMarcarEnviado(pedido.id, metodo)}
        onCrearPedidoFaltante={onCrearPedidoFaltante}
      />
    );
  }

  if (tabPrincipal === "solicitudes") {
    return (
      <SolicitudesPedidoModulo
        solicitudes={solicitudes}
        proyectos={proyectos}
        currentUser={currentUser}
        isAdmin={isAdmin}
        view={solicitudView}
        setView={setSolicitudView}
        editId={solicitudEditId}
        setEditId={setSolicitudEditId}
        detailId={solicitudDetailId}
        setDetailId={setSolicitudDetailId}
        onUpsert={onUpsertSolicitud}
        onDelete={onDeleteSolicitud}
        onAprobar={onAprobarSolicitud}
        onRechazar={onRechazarSolicitud}
        onComentar={onComentarSolicitud}
        onVolverPedidos={() => setTabPrincipal("pedidos")}
        prefill={solicitudPrefill}
        onClearPrefill={onClearSolicitudPrefill}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<ClipboardList size={20} className="text-[#2E8B57]" />}
        title="Pedidos"
        manualKey="pedidos"
        subtitle={`${pedidos.length} pedido${pedidos.length === 1 ? "" : "s"} registrado${pedidos.length === 1 ? "" : "s"}`}
      />

      <div className="flex gap-1 mb-6 border-b border-slate-200">
        <button onClick={() => setTabPrincipal("pedidos")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tabPrincipal === "pedidos" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Pedidos
        </button>
        <button onClick={() => setTabPrincipal("solicitudes")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition flex items-center gap-1.5 ${tabPrincipal === "solicitudes" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Solicitudes de empleados
          {solicitudes.filter((s) => s.estado === "Pendiente").length > 0 && (
            <Badge className="bg-rose-50 text-rose-700 ring-rose-200">{solicitudes.filter((s) => s.estado === "Pendiente").length}</Badge>
          )}
        </button>
      </div>

      {proveedores.length === 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
          ⚠ No puedes crear un pedido todavía: primero da de alta al menos un proveedor en la pestaña Proveedores.
        </div>
      )}
      <button
        onClick={() => { setEditId(null); setView("form"); }}
        disabled={proveedores.length === 0}
        title={proveedores.length === 0 ? "Necesitas dar de alta un proveedor primero" : ""}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg mb-3 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO PEDIDO
      </button>

      <button
        type="button"
        onClick={() => inputFotoRef.current?.click()}
        disabled={leyendoFoto}
        style={{ borderColor: "#2E8B57", color: "#2E8B57" }}
        className="w-full flex items-center justify-center gap-2 border-2 hover:bg-slate-50 disabled:opacity-50 text-sm font-semibold py-3 rounded-lg mb-6 cursor-pointer select-none"
      >
        <FileText size={16} /> {leyendoFoto ? "Leyendo..." : "Crear pedido a partir de una foto o PDF"}
      </button>
      <input
        ref={inputFotoRef}
        type="file"
        accept="image/*,application/pdf"
        className="hidden"
        onChange={(e) => { if (e.target.files?.[0]) leerFotoPedido(e.target.files[0]); e.target.value = ""; }}
      />
      {errorFoto && (
        <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold mb-6">⚠ {errorFoto}</div>
      )}

      {proveedores.length === 0 && (
        <div className="mb-4 px-4 py-3 rounded-md bg-amber-50 text-amber-800 text-sm border border-amber-200">
          Necesitas al menos un proveedor dado de alta antes de crear un pedido.
        </div>
      )}

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 min-w-[220px] max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nº, proveedor, material…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
        </div>
        <Select value={proveedorFiltro} onChange={(e) => setProveedorFiltro(e.target.value)} className="max-w-[200px]">
          <option value="">Todos los proveedores</option>
          {proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
        </Select>
        <Select value={estado} onChange={(e) => setEstado(e.target.value)} className="max-w-[180px]">
          <option value="">Todos los estados</option>
          {ESTADO_PEDIDO.map((t) => <option key={t}>{t}</option>)}
        </Select>
        <button
          onClick={() => descargarListaComoWord(
            "Pedidos",
            ["Nº", "Proveedor", "Líneas", "Estado", "Entrega prevista"],
            filtered.map((p) => [String(p.numero), proveedorNombre(p.proveedorId), String((p.lineas || []).length), p.estado, fmtDate(p.fechaEntregaPrevista)])
          )}
          className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
        >
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
      </div>
      {!estado && !q && pedidos.some((p) => p.estado === "Recibido") && (
        <p className="text-xs text-slate-400 mb-3">Los pedidos ya recibidos no se muestran aquí — para verlos, elige "Recibido" en el desplegable de estado, o búscalos por número.</p>
      )}

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nº Pedido</th>
              <th className="px-4 py-3 font-semibold">Proveedor</th>
              <th className="px-4 py-3 font-semibold">Materiales</th>
              <th className="px-4 py-3 font-semibold">Fecha compra</th>
              <th className="px-4 py-3 font-semibold">Entrega prevista</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay pedidos que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((p) => {
              const vencido = p.estado !== "Recibido" && p.estado !== "Cancelado" && p.fechaEntregaPrevista && new Date(p.fechaEntregaPrevista) < new Date();
              return (
                <tr key={p.id} onClick={() => { setDetailId(p.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                  <td className="px-4 py-3 font-mono-num text-slate-500">#{p.numero}</td>
                  <td className="px-4 py-3 font-medium text-slate-800">{proveedorNombre(p.proveedorId)}</td>
                  <td className="px-4 py-3 text-slate-600">{(p.lineas || []).length} línea{(p.lineas || []).length === 1 ? "" : "s"}</td>
                  <td className="px-4 py-3 text-slate-500">{fmtDate(p.fechaCompra)}</td>
                  <td className={`px-4 py-3 ${vencido ? "text-rose-600 font-semibold" : "text-slate-500"}`}>{fmtDate(p.fechaEntregaPrevista)}</td>
                  <td className="px-4 py-3"><Badge className={ESTADO_PEDIDO_STYLE[p.estado]}>{p.estado}</Badge></td>
                  <td className="px-4 py-3 text-right">
                    {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar el pedido #${p.numero}?`)) onDelete(p.id); }} className="text-slate-300 hover:text-rose-500 transition">
                      <Trash2 size={15} />
                    </button>)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PedidoForm({ initial, proveedores, materiales, proyectos, nextNumero, currentUser, onCancel, onSave, prefill, onClearPrefill }) {
  const blankLinea = () => ({ id: uid(), modo: "catalogo", materialId: materiales[0]?.id || "", referencia: "", ancho: "", alto: "", cantidad: "", precio: "", estado: "Solicitado" });
  const avisosPorDefecto = () => ({
    llegada: true, retraso: true,
    usuarioId: currentUser?.id || null,
    usuarioNombre: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "",
  });
  const [f, setF] = useState(() => {
    if (initial) return { avisos: avisosPorDefecto(), ...initial };
    if (prefill) {
      return {
        id: null, proveedorId: prefill.proveedorId || proveedores[0]?.id || "", proyectoId: prefill.proyectoId || "",
        fechaCompra: new Date().toISOString().slice(0, 10), fechaEntregaPrevista: "", estado: "Pendiente",
        comentarios: prefill.comentarios || "Generado automáticamente desde Stock (materiales por debajo del mínimo).",
        lineas: prefill.lineas && prefill.lineas.length ? prefill.lineas : [blankLinea()],
        avisos: avisosPorDefecto(),
      };
    }
    return {
      id: null, proveedorId: proveedores[0]?.id || "", proyectoId: "", fechaCompra: new Date().toISOString().slice(0, 10),
      fechaEntregaPrevista: "", estado: "Pendiente", comentarios: "",
      lineas: [blankLinea()],
      avisos: avisosPorDefecto(),
    };
  });
  const setAviso = (clave, valor) => {
    setF((prev) => ({
      ...prev,
      avisos: {
        ...(prev.avisos || {}),
        [clave]: valor,
        usuarioId: currentUser?.id || null,
        usuarioNombre: currentUser ? `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim() : "",
      },
    }));
  };
  useEffect(() => {
    if (!initial && prefill && onClearPrefill) onClearPrefill();
  }, []);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const [pasteText, setPasteText] = useState("");

  const setLinea = (id, patch) => setF({ ...f, lineas: f.lineas.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  const addLinea = (modo = "catalogo") => setF({ ...f, lineas: [...f.lineas, { ...blankLinea(), modo }] });
  const removeLinea = (id) => setF({ ...f, lineas: f.lineas.filter((l) => l.id !== id) });

  const convertirPegado = () => {
    if (!pasteText.trim()) return;
    const filas = pasteText.split("\n").map((r) => r.trim()).filter(Boolean);
    const nuevas = filas.map((fila) => {
      const cols = fila.includes("\t") ? fila.split("\t") : fila.split(",");
      const [referencia = "", ancho = "", alto = "", cantidad = "", precio = ""] = cols.map((c) => c.trim());
      return { id: uid(), modo: "libre", materialId: "", referencia, ancho, alto, cantidad, precio, estado: "Solicitado" };
    }).filter((l) => l.referencia || parseFloat(l.cantidad) > 0);
    if (nuevas.length === 0) return;
    setF({ ...f, lineas: [...f.lineas, ...nuevas] });
    setPasteText("");
  };

  const submit = (e) => {
    e.preventDefault();
    if (!f.proveedorId) { setErrorMsg("Falta seleccionar el proveedor."); return; }
    const lineas = f.lineas
      .filter((l) => (l.modo === "catalogo" ? l.materialId : l.referencia.trim()) && parseFloat(l.cantidad) > 0)
      .map((l) => ({ ...l, cantidad: parseFloat(l.cantidad) || 0, precio: parseFloat(l.precio) || 0, ancho: l.ancho ? parseFloat(l.ancho) || l.ancho : "", alto: l.alto ? parseFloat(l.alto) || l.alto : "" }));
    if (lineas.length === 0) { setErrorMsg("Añade al menos una línea válida (con cantidad mayor que 0)."); return; }
    setErrorMsg("");
    onSave({ ...f, lineas });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <div className="flex items-center gap-3 mb-6">
        <h1 className="font-display text-2xl font-extrabold text-slate-900">{initial ? `Editar pedido #${initial.numero}` : "Alta de pedido"}</h1>
        {!initial && <Badge className="bg-slate-100 text-slate-500 ring-slate-200 font-mono-num">Nº {nextNumero()} (automático)</Badge>}
      </div>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Proveedor" required>
            <Select value={f.proveedorId} onChange={set("proveedorId")} required>
              {proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </Select>
          </Field>
          <Field label="Estado">
            <Select value={f.estado} onChange={set("estado")}>
              {ESTADO_PEDIDO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
        </div>
        <div className="grid grid-cols-2 gap-4">
          <Field label="Fecha de compra"><TextInput type="date" value={f.fechaCompra} onChange={set("fechaCompra")} /></Field>
          <Field label="Fecha entrega prevista"><TextInput type="date" value={f.fechaEntregaPrevista} onChange={set("fechaEntregaPrevista")} /></Field>
        </div>

        <Field label="Proyecto / obra de este pedido">
          <Select value={f.proyectoId} onChange={set("proyectoId")}>
            <option value="">Stock (sin proyecto asociado)</option>
            {proyectos.map((pr) => <option key={pr.id} value={pr.id}>#{pr.numero} — {pr.nombre}</option>)}
          </Select>
          <p className="text-xs text-slate-400 mt-1">Todas las líneas de este pedido quedarán vinculadas a este proyecto.</p>
        </Field>

        {f.proyectoId && normalizarChecklist(proyectos.find((pr) => pr.id === f.proyectoId)?.checklistMateriales).some((c) => !c.estado) && (
          <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold">
            ⚠ El checklist "Qué lleva la obra" de este proyecto no está completo. No podrás guardar este pedido hasta rellenarlo en la ficha del proyecto (pestaña "Qué lleva la obra").
          </div>
        )}

        <div>
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-2">Líneas del pedido</span>
          <div className="space-y-2">
            {f.lineas.map((l) => (
              <div key={l.id} className="border border-slate-200 rounded-md p-2.5 bg-slate-50/50">
                <div className="flex items-center gap-2 mb-2">
                  <button type="button" onClick={() => setLinea(l.id, { modo: "catalogo" })}
                    className={`text-xs font-semibold px-2.5 py-1 rounded-full ${l.modo === "catalogo" ? "bg-[#2E8B57] text-white" : "bg-slate-200 text-slate-500"}`}>
                    Material de catálogo
                  </button>
                  <button type="button" onClick={() => setLinea(l.id, { modo: "libre" })}
                    className={`text-xs font-semibold px-2.5 py-1 rounded-full ${l.modo === "libre" ? "bg-[#2E8B57] text-white" : "bg-slate-200 text-slate-500"}`}>
                    Medida a medida
                  </button>
                  <button type="button" onClick={() => removeLinea(l.id)} className="ml-auto text-slate-300 hover:text-rose-500"><X size={15} /></button>
                </div>

                {l.modo === "catalogo" ? (
                  <div className="grid grid-cols-12 gap-2 items-center">
                    <div className="col-span-6">
                      <Select value={l.materialId} onChange={(e) => setLinea(l.id, { materialId: e.target.value })}>
                        <option value="">Selecciona material…</option>
                        {materiales.map((m) => <option key={m.id} value={m.id}>{m.codigo} — {m.descripcion}</option>)}
                      </Select>
                    </div>
                    <div className="col-span-3">
                      <TextInput type="number" placeholder="Cantidad" value={l.cantidad} onChange={(e) => setLinea(l.id, { cantidad: e.target.value })} />
                    </div>
                    <div className="col-span-3">
                      <TextInput type="number" step="0.01" placeholder="Precio ud. (€)" value={l.precio} onChange={(e) => setLinea(l.id, { precio: e.target.value })} />
                    </div>
                  </div>
                ) : (
                  <div className="grid grid-cols-12 gap-2 items-center">
                    <div className="col-span-4">
                      <TextInput placeholder="Referencia / descripción (ej. Cristal FL1)" value={l.referencia} onChange={(e) => setLinea(l.id, { referencia: e.target.value })} />
                    </div>
                    <div className="col-span-2">
                      <TextInput type="number" placeholder="Ancho" value={l.ancho} onChange={(e) => setLinea(l.id, { ancho: e.target.value })} />
                    </div>
                    <div className="col-span-2">
                      <TextInput type="number" placeholder="Alto" value={l.alto} onChange={(e) => setLinea(l.id, { alto: e.target.value })} />
                    </div>
                    <div className="col-span-2">
                      <TextInput type="number" placeholder="Cantidad" value={l.cantidad} onChange={(e) => setLinea(l.id, { cantidad: e.target.value })} />
                    </div>
                    <div className="col-span-2">
                      <TextInput type="number" step="0.01" placeholder="Precio ud. (€)" value={l.precio} onChange={(e) => setLinea(l.id, { precio: e.target.value })} />
                    </div>
                  </div>
                )}
              </div>
            ))}
          </div>
          <div className="flex flex-wrap gap-3 mt-2">
            <button type="button" onClick={() => addLinea("catalogo")} className="flex items-center gap-1 text-sm font-semibold text-[#2E8B57] hover:text-[#256E46]">
              <Plus size={14} /> Añadir material de catálogo
            </button>
            <button type="button" onClick={() => addLinea("libre")} className="flex items-center gap-1 text-sm font-semibold text-[#2E8B57] hover:text-[#256E46]">
              <Plus size={14} /> Añadir medida a medida
            </button>
          </div>
        </div>

        <div className="border border-dashed border-slate-300 rounded-md p-3 bg-slate-50/50">
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-1">Pegar varias medidas de golpe (desde Excel)</span>
          <p className="text-xs text-slate-400 mb-2">Copia y pega filas con columnas: Referencia, Ancho, Alto, Cantidad, Precio ud. (€) (opcional). Cada fila se convertirá en una línea "a medida".</p>
          <TextArea rows={3} placeholder={"FL1\t548\t863\t6\t12.50\nFL2\t526\t863\t2\t9.80"} value={pasteText} onChange={(e) => setPasteText(e.target.value)} />
          <button type="button" onClick={convertirPegado} className="mt-2 flex items-center gap-1.5 text-sm font-semibold text-white bg-slate-700 hover:bg-slate-800 px-3.5 py-2 rounded-md">
            <Plus size={14} /> Convertir en líneas
          </button>
        </div>

        <Field label="Comentarios"><TextArea rows={2} value={f.comentarios} onChange={set("comentarios")} /></Field>

        <Field label="Avisos">
          <div className="flex flex-col gap-2 pt-1">
            <label className="flex items-center gap-2 text-sm text-slate-600">
              <input type="checkbox" checked={!!f.avisos?.llegada} onChange={(e) => setAviso("llegada", e.target.checked)} />
              Avisarme a mí cuando llegue este pedido
            </label>
            <label className="flex items-center gap-2 text-sm text-slate-600">
              <input type="checkbox" checked={!!f.avisos?.retraso} onChange={(e) => setAviso("retraso", e.target.checked)} />
              Avisarme a mí si se pasa la fecha de entrega prevista y todavía no ha llegado
            </label>
            {(f.avisos?.llegada || f.avisos?.retraso) && (
              <p className="text-xs text-slate-400">El aviso te aparecerá a ti (usuario con la sesión iniciada ahora) en un banner al entrar al CRM.</p>
            )}
          </div>
        </Field>

        {errorMsg && (
          <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
            ⚠ {errorMsg}
          </div>
        )}

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md cursor-pointer select-none"><Save size={15} /> <span>Guardar pedido</span></button>
        </div>
      </form>
    </div>
  );
}

function PedidoDetail({ pedido, proveedor, materiales, proyectos, currentUser, onBack, onEdit, onDelete, onRecibir, onConfirmarAlbaran, onMarcarEnviado, onCrearPedidoFaltante, isAdmin }) {
  const materialInfo = (id) => materiales.find((m) => m.id === id);
  const proyecto = pedido.proyectoId ? proyectos.find((p) => p.id === pedido.proyectoId) : null;
  const puedeRecibir = pedido.estado !== "Recibido" && pedido.estado !== "Cancelado";
  const [confirmandoRecibir, setConfirmandoRecibir] = useState(false);
  const [leyendoAlbaran, setLeyendoAlbaran] = useState(false);
  const [resultadoAlbaran, setResultadoAlbaran] = useState(null);
  const [errorAlbaran, setErrorAlbaran] = useState("");
  const inputAlbaranRef = useRef(null);
  const [enviandoEmail, setEnviandoEmail] = useState(false);
  const [errorEnvioEmail, setErrorEnvioEmail] = useState("");
  const [emailEnviadoOk, setEmailEnviadoOk] = useState(false);

  const nombreLinea = (l) => {
    const esLibre = l.modo === "libre";
    const mat = esLibre ? null : materialInfo(l.materialId);
    return esLibre ? (l.referencia || "Sin referencia") : (mat ? mat.descripcion : "Material eliminado");
  };

  const leerAlbaran = async (file) => {
    setLeyendoAlbaran(true);
    setErrorAlbaran("");
    setResultadoAlbaran(null);
    try {
      const base64Data = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = () => rej(new Error("No se pudo leer el archivo"));
        r.readAsDataURL(file);
      });
      const esPdf = file.type === "application/pdf";
      const contentBlock = esPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } }
        : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64Data } };

      const prompt = 'Esto es un albarán de entrega de un proveedor. Puede tener varias líneas. Es MUY IMPORTANTE que revises el documento entero, de arriba a abajo, y devuelvas TODAS las líneas, sin saltarte ninguna ni resumir. Devuelve ÚNICAMENTE un JSON válido (sin texto adicional, sin backticks, sin explicación) con este formato exacto: [{"material":"nombre o referencia tal cual aparece en el albarán","cantidad":numero}]. Una línea por cada material o referencia distinta que aparezca. No omitas ninguna línea.';

      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 8000,
          messages: [{ role: "user", content: [contentBlock, { type: "text", text: prompt }] }],
        }),
      });
      if (!response.ok) {
        const errBody = await response.text();
        console.error("anthropic-proxy respuesta no válida:", response.status, errBody);
        throw new Error("Respuesta no válida de la API: " + response.status);
      }
      const data = await response.json();
      if (data.error) {
        console.error("Error devuelto por la API:", data.error);
        throw new Error(data.error.message || "Error de la API");
      }
      const textoRespuesta = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
      const limpio = textoRespuesta.replace(/```json|```/g, "").trim();
      const inicioA = limpio.indexOf("[");
      const finA = limpio.lastIndexOf("]");
      const jsonCandidatoA = inicioA !== -1 && finA !== -1 ? limpio.slice(inicioA, finA + 1) : limpio;
      let lineasAlbaran;
      try {
        lineasAlbaran = JSON.parse(jsonCandidatoA);
      } catch (parseErr) {
        console.error("No se pudo parsear el JSON del albarán. Texto recibido:", textoRespuesta);
        throw new Error("La respuesta de la IA no tenía formato válido");
      }

      const comparacion = pedido.lineas.map((l) => {
        const nombre = nombreLinea(l);
        const nombreLower = nombre.toLowerCase();
        const encontrado = lineasAlbaran.find((la) => {
          const laLower = String(la.material || "").toLowerCase();
          return laLower.includes(nombreLower.slice(0, 8)) || nombreLower.includes(laLower.slice(0, 8));
        });
        const cantidadCoincide = encontrado && parseFloat(encontrado.cantidad) === parseFloat(l.cantidad);
        return {
          nombre, cantidadPedida: l.cantidad,
          cantidadAlbaran: encontrado ? encontrado.cantidad : null,
          textoAlbaran: encontrado ? encontrado.material : null,
          coincide: !!encontrado && cantidadCoincide,
          encontrado: !!encontrado,
        };
      });
      const sobrantesAlbaran = lineasAlbaran.filter((la) =>
        !comparacion.some((c) => c.textoAlbaran === la.material)
      );

      setResultadoAlbaran({ comparacion, sobrantesAlbaran, fecha: new Date().toISOString(), archivo: file.name });
    } catch (err) {
      console.error("Error leyendo albarán:", err);
      setErrorAlbaran("No se pudo leer el albarán. Prueba con una foto más clara, con más luz, o inténtalo de nuevo. (" + err.message + ")");
    } finally {
      setLeyendoAlbaran(false);
    }
  };

  const enlaceEmailProveedor = () => {
    if (!proveedor?.email) return null;
    const lineasTexto = pedido.lineas.map((l) => {
      const medidas = l.modo === "libre" && (l.ancho || l.alto) ? ` (${l.ancho || "—"} x ${l.alto || "—"})` : "";
      return `- ${nombreLinea(l)}${medidas}: ${l.cantidad} ud.`;
    }).join("\n");
    const asunto = `Pedido ${pedido.numero} — ALUMAVEL`;
    const cuerpo = `Buenos días,\n\nLes hacemos el siguiente pedido:\n\n${lineasTexto}\n\nEntrega prevista: ${fmtDate(pedido.fechaEntregaPrevista) || "a concretar"}.\n${pedido.comentarios ? `\nComentarios: ${pedido.comentarios}\n` : ""}\nUn saludo,\nALUMAVEL`;
    return `mailto:${proveedor.email}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  };

  const enviarEmailAutomatico = async () => {
    if (!proveedor?.email) return;
    setEnviandoEmail(true);
    setErrorEnvioEmail("");
    setEmailEnviadoOk(false);
    try {
      const lineasTexto = pedido.lineas.map((l) => {
        const medidas = l.modo === "libre" && (l.ancho || l.alto) ? ` (${l.ancho || "—"} x ${l.alto || "—"})` : "";
        return `- ${nombreLinea(l)}${medidas}: ${l.cantidad} ud.`;
      }).join("\n");
      const cuerpo = `Buenos días,\n\nLes hacemos el siguiente pedido:\n\n${lineasTexto}\n\nEntrega prevista: ${fmtDate(pedido.fechaEntregaPrevista) || "a concretar"}.\n${pedido.comentarios ? `\nComentarios: ${pedido.comentarios}\n` : ""}\nUn saludo,\nALUMAVEL`;

      const response = await fetch("/.netlify/functions/enviar-email", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          destinatario: proveedor.email,
          asunto: `Pedido ${pedido.numero} — ALUMAVEL`,
          cuerpo,
          replyTo: currentUser?.email || "",
        }),
      });
      const data = await response.json();
      if (!response.ok || data.error) {
        throw new Error(data.error || "No se pudo enviar el correo.");
      }
      setEmailEnviadoOk(true);
      onMarcarEnviado("email");
    } catch (err) {
      console.error("Error enviando email automático:", err);
      setErrorEnvioEmail(err.message + " Puedes usar el botón de abajo para enviarlo desde tu propio correo mientras tanto.");
    } finally {
      setEnviandoEmail(false);
    }
  };

  const enlaceWhatsappProveedor = () => {
    if (!proveedor?.movil) return null;
    const lineasTexto = pedido.lineas.map((l) => {
      const medidas = l.modo === "libre" && (l.ancho || l.alto) ? ` (${l.ancho || "—"} x ${l.alto || "—"})` : "";
      return `- ${nombreLinea(l)}${medidas}: ${l.cantidad} ud.`;
    }).join("\n");
    const texto = `Pedido ${pedido.numero} — ALUMAVEL\n\n${lineasTexto}\n\nEntrega prevista: ${fmtDate(pedido.fechaEntregaPrevista) || "a concretar"}.${pedido.comentarios ? `\n\nComentarios: ${pedido.comentarios}` : ""}`;
    const tel = proveedor.movil.replace(/[^\d+]/g, "");
    return `https://wa.me/${tel}?text=${encodeURIComponent(texto)}`;
  };

  // Líneas que faltan o no cuadran según el albarán leído, listas para pasarlas
  // a un pedido nuevo (reponiendo modo/material/medidas de la línea original).
  const lineasFaltantesAlbaran = () => {
    if (!resultadoAlbaran) return [];
    return pedido.lineas
      .map((l) => {
        const comp = resultadoAlbaran.comparacion.find((c) => c.nombre === nombreLinea(l));
        if (!comp || comp.coincide) return null;
        const recibida = comp.encontrado ? parseFloat(comp.cantidadAlbaran) || 0 : 0;
        const pedida = parseFloat(l.cantidad) || 0;
        const falta = Math.round((pedida - recibida) * 100) / 100;
        if (falta <= 0) return null;
        return { ...l, id: uid(), cantidad: falta, estado: "Solicitado" };
      })
      .filter(Boolean);
  };

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">#{pedido.numero}</span>
            <h1 className="font-display text-2xl font-extrabold text-slate-900">{proveedor?.nombre || "Proveedor no encontrado"}</h1>
          </div>
          <p className="text-sm text-slate-500 mt-1">Compra {fmtDate(pedido.fechaCompra)} · Entrega prevista {fmtDate(pedido.fechaEntregaPrevista)}</p>
          <p className="text-sm text-slate-500 mt-0.5">Proyecto: <span className="font-semibold text-slate-700">{proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : "Stock (sin proyecto asociado)"}</span></p>
        </div>
        <div className="flex gap-2 shrink-0">
          {proveedor?.email ? (
            <div className="flex flex-col items-end gap-1">
              <button
                onClick={enviarEmailAutomatico}
                disabled={enviandoEmail}
                style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
                className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 disabled:opacity-60 px-3.5 py-2 rounded-md"
              >
                <Mail size={14} /> {enviandoEmail ? "Enviando..." : "Enviar por email"}
              </button>
              {emailEnviadoOk && <span className="text-xs text-emerald-600 font-semibold">✓ Correo enviado</span>}
              {errorEnvioEmail && (
                <div className="text-xs text-rose-600 font-semibold text-right max-w-[220px]">
                  ⚠ {errorEnvioEmail}
                  <a href={enlaceEmailProveedor()} onClick={() => onMarcarEnviado("email")} className="block underline mt-0.5">Abrir en mi correo</a>
                </div>
              )}
            </div>
          ) : (
            <span title="Este proveedor no tiene email guardado" className="flex items-center gap-1.5 text-sm font-semibold text-slate-300 border border-slate-200 px-3.5 py-2 rounded-md cursor-not-allowed">
              <Mail size={14} /> Sin email
            </span>
          )}
          {enlaceWhatsappProveedor() ? (
            <a href={enlaceWhatsappProveedor()} target="_blank" rel="noreferrer" onClick={() => onMarcarEnviado("whatsapp")} className="flex items-center gap-1.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 px-3.5 py-2 rounded-md">
              <MessageCircle size={14} /> Enviar por WhatsApp
            </a>
          ) : (
            <span title="Este proveedor no tiene móvil guardado" className="flex items-center gap-1.5 text-sm font-semibold text-slate-300 border border-slate-200 px-3.5 py-2 rounded-md cursor-not-allowed">
              <MessageCircle size={14} /> Sin móvil
            </span>
          )}
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      <div className="flex items-center gap-2 mb-6">
        <Badge className={ESTADO_PEDIDO_STYLE[pedido.estado]}>{pedido.estado}</Badge>
        {pedido.fechaRecibido && <span className="text-xs text-slate-400">Recibido el {fmtDate(pedido.fechaRecibido)}</span>}
        {(pedido.avisos?.llegada || pedido.avisos?.retraso) && (
          <span className="text-xs text-sky-600 font-semibold" title={`Avisos para ${pedido.avisos?.usuarioNombre || "un usuario"}`}>
            🔔 {[pedido.avisos?.llegada && "llegada", pedido.avisos?.retraso && "retraso"].filter(Boolean).join(" y ")} — {pedido.avisos?.usuarioNombre}
          </span>
        )}
      </div>

      {pedido.envioConfirmado ? (
        <div className="px-4 py-3 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm font-semibold mb-6">
          ✓ {pedido.envioMetodo === "email" ? "Enviado por email" : pedido.envioMetodo === "whatsapp" ? "Enviado por WhatsApp" : "Pedido realizado"} el {fmtDate(pedido.fechaEnvioConfirmado?.slice(0, 10))}.
        </div>
      ) : (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-6 space-y-2">
          <p>⚠ Este pedido todavía no consta como enviado al proveedor. No podrás marcarlo como recibido hasta que lo confirmes.</p>
          <p className="text-xs font-normal text-amber-700">
            Si le has llamado, o lo has pedido de otra forma que no sea el botón de email de arriba, confírmalo aquí:
          </p>
          <button
            onClick={() => onMarcarEnviado("manual")}
            className="text-xs font-semibold text-white bg-amber-600 hover:bg-amber-700 px-3 py-1.5 rounded-md"
          >
            He hecho este pedido de otra forma
          </button>
        </div>
      )}

      {puedeRecibir && !confirmandoRecibir && (
        <button
          onClick={() => { if (pedido.envioConfirmado) setConfirmandoRecibir(true); }}
          disabled={!pedido.envioConfirmado}
          title={!pedido.envioConfirmado ? "Confirma primero que el pedido se ha enviado al proveedor" : ""}
          className="mb-6 flex items-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-sm font-semibold px-4 py-2.5 rounded-md transition cursor-pointer select-none"
        >
          <CheckCircle2 size={16} /> Marcar pedido como recibido (actualiza stock)
        </button>
      )}
      {puedeRecibir && confirmandoRecibir && (
        <div className="mb-6 px-4 py-3 rounded-md bg-emerald-50 border border-emerald-300">
          <p className="text-sm text-emerald-800 font-medium mb-3">Al marcar el pedido como recibido se sumará automáticamente al stock de cada material. ¿Confirmas?</p>
          <div className="flex flex-wrap gap-2">
            <button
              type="button"
              onClick={() => { onRecibir(); setConfirmandoRecibir(false); }}
              className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-semibold px-4 py-2 rounded-md cursor-pointer select-none"
            >
              <CheckCircle2 size={15} /> Sí, marcar como recibido
            </button>
            <button type="button" onClick={() => setConfirmandoRecibir(false)} className="px-4 py-2 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          </div>
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-4">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Material / Referencia</th>
              <th className="px-4 py-2.5 font-semibold">Medidas</th>
              <th className="px-4 py-2.5 font-semibold text-right">Cantidad</th>
              <th className="px-4 py-2.5 font-semibold text-right">Precio ud.</th>
              <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
              <th className="px-4 py-2.5 font-semibold">Estado</th>
            </tr>
          </thead>
          <tbody>
            {pedido.lineas.map((l) => {
              const esLibre = l.modo === "libre";
              const mat = esLibre ? null : materialInfo(l.materialId);
              return (
                <tr key={l.id} className="border-b border-slate-100 last:border-0">
                  <td className="px-4 py-2.5 font-medium text-slate-800">
                    {esLibre ? (l.referencia || "Sin referencia") : (mat ? `${mat.codigo} — ${mat.descripcion}` : "Material eliminado")}
                    {esLibre && <Badge className="ml-2 bg-amber-50 text-amber-700 ring-amber-200">a medida</Badge>}
                  </td>
                  <td className="px-4 py-2.5 text-slate-600 font-mono-num">{esLibre && (l.ancho || l.alto) ? `${l.ancho || "—"} × ${l.alto || "—"}` : "—"}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num">{l.cantidad}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num text-slate-500">{l.precio ? money(l.precio) : "—"}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num font-semibold">{l.precio ? money((parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0)) : "—"}</td>
                  <td className="px-4 py-2.5"><Badge className={ESTADO_LINEA_PEDIDO_STYLE[l.estado] || ESTADO_LINEA_PEDIDO_STYLE["Solicitado"]}>{l.estado}</Badge></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {pedido.lineas.some((l) => l.precio) && (
        <div className="flex justify-end mb-4 px-2">
          <span className="text-sm text-slate-600">
            Total del pedido: <span className="font-bold text-slate-800 text-base">{money(pedido.lineas.reduce((s, l) => s + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0))}</span>
          </span>
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-6">
        <h2 className="font-display font-bold text-slate-800 mb-2">Comprobar albarán</h2>
        <p className="text-xs text-slate-400 mb-3">Sube la foto o el PDF del albarán que te ha llegado. Lo leo y lo comparo con lo que hay pedido — nada se guarda hasta que tú confirmes abajo.</p>

        <button
          type="button"
          onClick={() => inputAlbaranRef.current?.click()}
          disabled={leyendoAlbaran}
          style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
          className="flex items-center gap-2 text-sm font-semibold hover:opacity-90 disabled:bg-slate-300 px-4 py-2.5 rounded-md w-fit cursor-pointer"
        >
          {leyendoAlbaran ? "Leyendo..." : "Subir foto o PDF del albarán"}
        </button>
        <input
          ref={inputAlbaranRef}
          type="file"
          accept="image/*,application/pdf"
          className="hidden"
          onChange={(e) => { if (e.target.files?.[0]) leerAlbaran(e.target.files[0]); e.target.value = ""; }}
        />

        {errorAlbaran && (
          <div className="mt-3 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">⚠ {errorAlbaran}</div>
        )}

        {resultadoAlbaran && (
          <div className="mt-4">
            <div className="bg-slate-50 border border-slate-200 rounded-lg overflow-hidden mb-3">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-100 text-left text-[11px] uppercase tracking-wide text-slate-500">
                    <th className="px-3 py-2 font-semibold">Material (del pedido)</th>
                    <th className="px-3 py-2 font-semibold text-right">Pedido</th>
                    <th className="px-3 py-2 font-semibold text-right">En el albarán</th>
                    <th className="px-3 py-2 font-semibold">¿Coincide?</th>
                  </tr>
                </thead>
                <tbody>
                  {resultadoAlbaran.comparacion.map((c, i) => (
                    <tr key={i} className="border-t border-slate-200">
                      <td className="px-3 py-2 text-slate-700">{c.nombre}</td>
                      <td className="px-3 py-2 text-right font-mono-num">{c.cantidadPedida}</td>
                      <td className="px-3 py-2 text-right font-mono-num">{c.encontrado ? c.cantidadAlbaran : "No aparece"}</td>
                      <td className="px-3 py-2">
                        {c.coincide ? <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">✓ Coincide</Badge> : <Badge className="bg-rose-50 text-rose-700 ring-rose-200">⚠ Revisar</Badge>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            {resultadoAlbaran.sobrantesAlbaran.length > 0 && (
              <div className="px-3 py-2.5 rounded-md bg-amber-50 border border-amber-200 text-amber-800 text-sm mb-3">
                ⚠ En el albarán hay {resultadoAlbaran.sobrantesAlbaran.length} línea{resultadoAlbaran.sobrantesAlbaran.length === 1 ? "" : "s"} que no encuentro en el pedido: {resultadoAlbaran.sobrantesAlbaran.map((l) => l.material).join(", ")}.
              </div>
            )}
            <p className="text-xs text-slate-400 mb-3">La lectura del albarán es automática y puede tener algún error, sobre todo con fotos borrosas o letra manuscrita — revísalo tú antes de confirmar.</p>
            {resultadoAlbaran.comparacion.some((c) => !c.coincide) && (
              <div className="px-3 py-2.5 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
                ⚠ Hay líneas marcadas como "Revisar" arriba. Si confirmas igualmente, se recibirá el pedido completo tal y como está pedido.
              </div>
            )}
            {resultadoAlbaran.comparacion.some((c) => !c.coincide) && !pedido.albaranComprobado && (
              <button
                onClick={() => onCrearPedidoFaltante(pedido, lineasFaltantesAlbaran())}
                className="flex items-center gap-1.5 text-sm font-semibold text-amber-700 border border-amber-300 px-4 py-2.5 rounded-md hover:bg-amber-50 mb-3"
              >
                📦 Crear pedido al proveedor con lo que falta
              </button>
            )}
            {pedido.albaranComprobado ? (
              <div className="px-3 py-2.5 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm font-semibold">
                ✓ Comprobación confirmada el {fmtDate(pedido.albaranComprobado.fecha?.slice(0, 10))}.
              </div>
            ) : (
              <>
                <button
                  onClick={() => {
                    onConfirmarAlbaran(resultadoAlbaran);
                    if (puedeRecibir && pedido.envioConfirmado) onRecibir();
                  }}
                  disabled={puedeRecibir && !pedido.envioConfirmado}
                  title={puedeRecibir && !pedido.envioConfirmado ? "Confirma primero que el pedido se ha enviado al proveedor" : ""}
                  style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
                  className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 disabled:bg-slate-300 disabled:cursor-not-allowed px-4 py-2.5 rounded-md"
                >
                  <CheckCircle2 size={16} /> {puedeRecibir ? "Confirmar y recibir pedido (suma al stock)" : "Confirmar comprobación"}
                </button>
                {puedeRecibir && !pedido.envioConfirmado && (
                  <p className="text-xs text-amber-700 font-semibold mt-2">⚠ Confirma arriba que el pedido se ha enviado al proveedor antes de poder recibirlo.</p>
                )}
              </>
            )}
          </div>
        )}

        {!resultadoAlbaran && pedido.albaranComprobado && (
          <div className="mt-3 px-3 py-2.5 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm font-semibold">
            ✓ Ya hay una comprobación de albarán confirmada el {fmtDate(pedido.albaranComprobado.fecha?.slice(0, 10))} ({pedido.albaranComprobado.archivo}).
          </div>
        )}
      </div>

      {pedido.comentarios && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Comentarios</span>
          <p className="text-sm text-slate-600">{pedido.comentarios}</p>
        </CornerFrame>
      )}
    </div>
  );
}

/* ================= SOLICITUDES DE PEDIDO (empleados) ================= */

const URGENCIA_SOLICITUD = ["Normal", "Urgente"];
const TIPO_SOLICITUD = ["Material para proyecto", "Artículo nuevo para catálogo", "Otro"];
const ESTADO_SOLICITUD_STYLE = {
  Pendiente: "bg-amber-50 text-amber-700 ring-amber-200",
  Aprobada: "bg-emerald-50 text-emerald-700 ring-emerald-200",
  Rechazada: "bg-rose-50 text-rose-700 ring-rose-200",
};

function SolicitudesPedidoModulo({ solicitudes, proyectos, currentUser, isAdmin, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onAprobar, onRechazar, onComentar, onVolverPedidos, prefill, onClearPrefill }) {
  const [estadoFiltro, setEstadoFiltro] = useState("");
  const visibles = isAdmin ? solicitudes : solicitudes.filter((s) => s.solicitanteId === currentUser?.id);
  const filtered = estadoFiltro ? visibles.filter((s) => s.estado === estadoFiltro) : visibles;
  const pendientesN = solicitudes.filter((s) => s.estado === "Pendiente").length;

  if (view === "form") {
    const editing = solicitudes.find((s) => s.id === editId) || null;
    return (
      <SolicitudPedidoForm
        initial={editing || prefill}
        proyectos={proyectos}
        onCancel={() => { setView(editing ? "detail" : "list"); onClearPrefill && onClearPrefill(); }}
        onSave={(data) => { onUpsert(data); onClearPrefill && onClearPrefill(); }}
      />
    );
  }

  if (view === "detail") {
    const solicitud = solicitudes.find((s) => s.id === detailId);
    if (!solicitud) { setView("list"); return null; }
    return (
      <SolicitudPedidoDetail
        solicitud={solicitud}
        proyecto={proyectos.find((p) => p.id === solicitud.proyectoId)}
        currentUser={currentUser}
        isAdmin={isAdmin}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(solicitud.id); setView("form"); }}
        onDelete={() => onDelete(solicitud.id)}
        onAprobar={() => onAprobar(solicitud)}
        onRechazar={(motivo) => onRechazar(solicitud.id, motivo)}
        onComentar={(texto) => onComentar(solicitud.id, texto)}
      />
    );
  }

  return (
    <div className="p-8 max-w-5xl">
      <button onClick={onVolverPedidos} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver a Pedidos</button>

      <Header
        icon={<ClipboardList size={20} className="text-[#2E8B57]" />}
        title="Solicitudes de empleados"
        manualKey="solicitudesPedido"
        subtitle={isAdmin ? `${solicitudes.length} solicitud${solicitudes.length === 1 ? "" : "es"} · ${pendientesN} pendiente${pendientesN === 1 ? "" : "s"} de revisar` : `${visibles.length} solicitud${visibles.length === 1 ? "" : "es"} tuya${visibles.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> SOLICITAR UN PEDIDO
      </button>

      <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-6">
        Cualquier empleado puede pedir aquí que se compre algo — un material para una obra, un artículo nuevo de catálogo, o lo que sea. La persona encargada de pedidos la revisa, puede preguntarte dudas aquí mismo, y la aprueba o la rechaza.
      </div>

      <Select value={estadoFiltro} onChange={(e) => setEstadoFiltro(e.target.value)} className="max-w-[200px] mb-4">
        <option value="">Todos los estados</option>
        <option>Pendiente</option>
        <option>Aprobada</option>
        <option>Rechazada</option>
      </Select>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Fecha</th>
              {isAdmin && <th className="px-4 py-3 font-semibold">Solicitado por</th>}
              <th className="px-4 py-3 font-semibold">Qué se pide</th>
              <th className="px-4 py-3 font-semibold">Urgencia</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
              <th className="px-4 py-3 font-semibold text-right">Comentarios</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((s) => (
              <tr key={s.id} onClick={() => { setDetailId(s.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                <td className="px-4 py-3 text-slate-600">{fmtDate(s.fecha)}</td>
                {isAdmin && <td className="px-4 py-3 font-medium text-slate-800">{s.solicitanteNombre}</td>}
                <td className="px-4 py-3 text-slate-700">{s.descripcion?.slice(0, 60)}{s.descripcion?.length > 60 ? "…" : ""}</td>
                <td className="px-4 py-3">{s.urgencia === "Urgente" ? <Badge className="bg-rose-50 text-rose-700 ring-rose-200">Urgente</Badge> : <span className="text-slate-400 text-xs">Normal</span>}</td>
                <td className="px-4 py-3"><Badge className={ESTADO_SOLICITUD_STYLE[s.estado]}>{s.estado}</Badge></td>
                <td className="px-4 py-3 text-right text-slate-400">{(s.comentarios || []).length || "—"}</td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={isAdmin ? 6 : 5} className="px-4 py-10 text-center text-slate-400 text-sm">No hay solicitudes que coincidan.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SolicitudPedidoForm({ initial, proyectos, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || { id: null, proyectoId: "", tipo: "Material para proyecto", descripcion: "", urgencia: "Normal" }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.descripcion.trim()) { setErrorMsg("Falta describir qué necesitas pedir (es obligatorio)."); return; }
    setErrorMsg("");
    onSave(f);
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">{initial?.id ? "Editar solicitud" : "Solicitar un pedido"}</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Tipo de solicitud">
            <Select value={f.tipo} onChange={set("tipo")}>
              {TIPO_SOLICITUD.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Urgencia">
            <Select value={f.urgencia} onChange={set("urgencia")}>
              {URGENCIA_SOLICITUD.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
        </div>

        <Field label="Proyecto relacionado (si aplica)">
          <Select value={f.proyectoId} onChange={set("proyectoId")}>
            <option value="">General / Stock (sin proyecto)</option>
            {proyectos.map((p) => <option key={p.id} value={p.id}>#{p.numero} — {p.nombre}</option>)}
          </Select>
        </Field>

        <Field label="¿Qué necesitas pedir?" required>
          <TextArea rows={4} value={f.descripcion} onChange={set("descripcion")} placeholder="Ej: Necesitamos 10 tiradores modelo X para la obra de Antonio Vico" />
        </Field>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Enviar solicitud</button>
        </div>
      </form>
    </div>
  );
}

function SolicitudPedidoDetail({ solicitud, proyecto, currentUser, isAdmin, onBack, onEdit, onDelete, onAprobar, onRechazar, onComentar }) {
  const [comentario, setComentario] = useState("");
  const [rechazando, setRechazando] = useState(false);
  const [motivoRechazo, setMotivoRechazo] = useState("");
  const esPropia = solicitud.solicitanteId === currentUser?.id;

  const enviarComentario = (e) => {
    e.preventDefault();
    if (!comentario.trim()) return;
    onComentar(comentario);
    setComentario("");
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>

      <div className="flex items-start justify-between mb-4 gap-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <h1 className="font-display text-xl font-extrabold text-slate-900">{solicitud.tipo}</h1>
            <Badge className={ESTADO_SOLICITUD_STYLE[solicitud.estado]}>{solicitud.estado}</Badge>
            {solicitud.urgencia === "Urgente" && <Badge className="bg-rose-50 text-rose-700 ring-rose-200">Urgente</Badge>}
          </div>
          <p className="text-sm text-slate-500">Pedido por {solicitud.solicitanteNombre} el {fmtDate(solicitud.fecha)}{proyecto && ` · Proyecto #${proyecto.numero} — ${proyecto.nombre}`}</p>
        </div>
        {(esPropia || isAdmin) && solicitud.estado === "Pendiente" && (
          <div className="flex gap-2 shrink-0">
            <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          </div>
        )}
      </div>

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 text-sm text-slate-700 whitespace-pre-wrap">
        {solicitud.descripcion}
      </div>

      {solicitud.estado === "Rechazada" && solicitud.motivoRechazo && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-200 text-rose-700 text-sm">
          <span className="font-semibold">Motivo del rechazo:</span> {solicitud.motivoRechazo}
        </div>
      )}

      {isAdmin && solicitud.estado === "Pendiente" && !rechazando && (
        <div className="flex flex-wrap gap-2 mb-6">
          <button onClick={onAprobar} className="flex items-center gap-1.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 px-4 py-2.5 rounded-md">
            <CheckCircle2 size={16} /> Aprobar y crear pedido
          </button>
          <button onClick={() => setRechazando(true)} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-4 py-2.5 rounded-md hover:bg-rose-50">
            <X size={16} /> Rechazar
          </button>
        </div>
      )}
      {isAdmin && rechazando && (
        <div className="mb-6 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 space-y-2">
          <TextInput value={motivoRechazo} onChange={(e) => setMotivoRechazo(e.target.value)} placeholder="Motivo del rechazo (ej: ya tenemos stock suficiente)" />
          <div className="flex gap-2">
            <button onClick={() => { onRechazar(motivoRechazo || "Sin motivo especificado"); setRechazando(false); }} className="text-sm font-semibold text-white bg-rose-600 hover:bg-rose-700 px-4 py-2 rounded-md">Confirmar rechazo</button>
            <button onClick={() => setRechazando(false)} className="text-sm font-semibold text-slate-600 px-4 py-2 rounded-md hover:bg-slate-100">Cancelar</button>
          </div>
        </div>
      )}

      <h2 className="font-display font-bold text-slate-800 mb-3">Comentarios ({(solicitud.comentarios || []).length})</h2>
      <div className="space-y-3 mb-4">
        {(solicitud.comentarios || []).map((c) => (
          <div key={c.id} className="bg-slate-50 rounded-lg p-3 text-sm">
            <div className="flex items-center gap-2 mb-1">
              <span className="font-semibold text-slate-700">{c.autor}</span>
              <span className="text-xs text-slate-400">{new Date(c.fecha).toLocaleString("es-ES")}</span>
            </div>
            <p className="text-slate-600">{c.texto}</p>
          </div>
        ))}
        {(solicitud.comentarios || []).length === 0 && <p className="text-sm text-slate-400">Todavía no hay comentarios.</p>}
      </div>

      <form onSubmit={enviarComentario} className="flex gap-2">
        <TextInput value={comentario} onChange={(e) => setComentario(e.target.value)} placeholder="Escribe un comentario o pregunta..." className="flex-1" />
        <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-sm font-semibold hover:opacity-90 px-4 py-2 rounded-md">Enviar</button>
      </form>
    </div>
  );
}

/* ================= FÁBRICA (control independiente de almacén) ================= */

/* ================= CRISTALES (almacén de vidrio, dentro de Fábrica) ================= */

function CristalesModulo({ cristales, proyectos, proveedores, clientes, onAdd, onUpdate, onDelete, onUbicar, onLiberar }) {
  const [subTab, setSubTab] = useState("pendientes");
  const [q, setQ] = useState("");
  const [asignando, setAsignando] = useState(null); // cristal object being located right now
  const [verDetalle, setVerDetalle] = useState(null); // ubicación { zona, fila, hueco } to show contents of
  const [leyendoPacking, setLeyendoPacking] = useState(false);
  const [errorPacking, setErrorPacking] = useState("");
  const [mostrarNuevo, setMostrarNuevo] = useState(false);
  const inputPackingRef = useRef(null);

  const normalizar = (s) => (s || "").toString().toLowerCase().replace(/\s+/g, "").replace(/[×*]/g, "x");
  const filtered = useMemo(() => {
    if (!q) return cristales;
    const nq = normalizar(q);
    return cristales.filter((c) => normalizar(`${c.lote} ${c.secuencia} ${c.cliente} ${c.proveedor} ${c.expediente} ${c.medida}`).includes(nq));
  }, [cristales, q]);

  const pendientes = filtered.filter((c) => c.estado === "Pendiente");
  const colocados = cristales.filter((c) => c.estado === "Colocado");

  // Sugiere una ubicación: prioriza un hueco libre en la misma fila que otro caballete
  // del mismo expediente ya colocado, para mantenerlos juntos; si no, el primer hueco
  // libre de la zona que corresponda según el proveedor (dejando la fila reservada, si
  // la hay, para el final, como colchón cuando el resto esté lleno).
  const sugerirUbicacion = (cristal) => {
    const zonaSugerida = (cristal.proveedor || "").toLowerCase().includes("uxcar") ? "arriba" : "abajo";
    const ocupado = (zona, fila, hueco) => cristales.some((c) => c.ubicacion && c.ubicacion.zona === zona && c.ubicacion.fila === fila && c.ubicacion.hueco === hueco);

    if (cristal.expediente) {
      const mismos = cristales.filter((c) => c.id !== cristal.id && c.expediente === cristal.expediente && c.ubicacion);
      for (const m of mismos) {
        const { zona, fila } = m.ubicacion;
        for (let h = 1; h <= ZONAS_CRISTALES[zona].huecos; h++) {
          if (!ocupado(zona, fila, h)) return { zona, fila, hueco: h };
        }
      }
    }
    const cfg = ZONAS_CRISTALES[zonaSugerida];
    const filaReservada = FILA_RESERVA[zonaSugerida];
    const filasNormales = Array.from({ length: cfg.filas }, (_, i) => i + 1).filter((f) => f !== filaReservada);
    for (const f of filasNormales) {
      for (let h = 1; h <= cfg.huecos; h++) {
        if (!ocupado(zonaSugerida, f, h)) return { zona: zonaSugerida, fila: f, hueco: h };
      }
    }
    if (filaReservada) {
      for (let h = 1; h <= cfg.huecos; h++) {
        if (!ocupado(zonaSugerida, filaReservada, h)) return { zona: zonaSugerida, fila: filaReservada, hueco: h };
      }
    }
    return null;
  };

  const leerPackingList = async (file) => {
    setLeyendoPacking(true);
    setErrorPacking("");
    try {
      const base64Data = await new Promise((resolve, reject) => {
        const r = new FileReader();
        r.onload = () => resolve(r.result.split(",")[1]);
        r.onerror = reject;
        r.readAsDataURL(file);
      });
      const esPdf = file.type === "application/pdf";
      const contentBlock = esPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } }
        : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64Data } };
      const prompt = 'Esto es un packing list / albarán de entrega de caballetes de cristal. Puede tener muchas filas (a veces 10, 15 o más). Es MUY IMPORTANTE que revises el documento entero, de arriba a abajo, y devuelvas TODAS las filas, sin saltarte ninguna ni resumir. Antes de responder, cuenta cuántas filas de datos hay en el documento y asegúrate de que tu respuesta tiene exactamente ese número de elementos. Devuelve ÚNICAMENTE un JSON válido (sin texto adicional, sin backticks, sin explicación) como un array: [{"lote":"","secuencia":"","cliente":"","proveedor":"","expediente":"","medida":"","cantidad":numero}]. Una línea por cada caballete o referencia distinta que aparezca en el documento. Deja en blanco lo que no encuentres, pero no omitas ninguna fila.';
      // Usamos la función en segundo plano (sin límite de 26s) para evitar el 504.
      // La foto/PDF en base64 puede pesar varios MB, y Netlify rechaza con un
      // error 413 las peticiones grandes a sus funciones. Por eso guardamos el
      // archivo directamente en Firebase (sin ese límite) y a la función solo
      // le mandamos el jobId; ella misma va a buscar el archivo a Firebase.
      const jobId = uid();
      await fbSet(ref(fbDb, `packingListJobsInput/${jobId}`), {
        model: "claude-haiku-4-5-20251001",
        max_tokens: 8000,
        contentBlock,
        prompt,
      });
      await fetch("/.netlify/functions/anthropic-proxy-background", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ jobId }),
      });

      let resultado = null;
      for (let intento = 0; intento < 60; intento++) {
        await new Promise((r) => setTimeout(r, 3000));
        const snap = await fbGet(ref(fbDb, `packingListJobs/${jobId}`)).catch(() => null);
        const val = snap && snap.exists ? (snap.exists() ? snap.val() : null) : null;
        if (val && (val.status === "done" || val.status === "error")) {
          resultado = val;
          break;
        }
      }
      fbSet(ref(fbDb, `packingListJobs/${jobId}`), null).catch(() => {});
      fbSet(ref(fbDb, `packingListJobsInput/${jobId}`), null).catch(() => {});

      if (!resultado) {
        throw new Error("La lectura está tardando demasiado (más de 3 minutos). Prueba de nuevo o con un documento más corto.");
      }
      if (resultado.status === "error") {
        console.error("Error devuelto por la función en segundo plano:", resultado.error);
        throw new Error(resultado.error || "Error de la API");
      }
      const texto = resultado.texto || "";
      const limpio = texto.replace(/```json|```/g, "").trim();
      const inicio = limpio.indexOf("[");
      const fin = limpio.lastIndexOf("]");
      const jsonCandidato = inicio !== -1 && fin !== -1 ? limpio.slice(inicio, fin + 1) : limpio;
      let items;
      try {
        items = JSON.parse(jsonCandidato);
      } catch (parseErr) {
        console.error("No se pudo parsear el JSON de la IA. Texto recibido:", texto);
        throw new Error("La respuesta de la IA no tenía formato válido");
      }
      if (!Array.isArray(items) || items.length === 0) {
        setErrorPacking("No he podido leer ningún caballete claro en el documento. Prueba con una foto más nítida.");
        setLeyendoPacking(false);
        return;
      }
      // Simulamos la colocación de cada caballete nuevo uno a uno, usando la misma lógica
      // de sugerencia (agrupar por expediente en la misma fila, si no el primer hueco libre
      // de la zona). Trabajamos sobre una copia local para que los caballetes del mismo
      // packing list también se tengan en cuenta entre sí según se van "colocando".
      const ocupadosSimulado = cristales
        .filter((c) => c.ubicacion)
        .map((c) => ({ zona: c.ubicacion.zona, fila: c.ubicacion.fila, hueco: c.ubicacion.hueco, expediente: c.expediente }));

      const estaOcupado = (zona, fila, hueco) => ocupadosSimulado.some((o) => o.zona === zona && o.fila === fila && o.hueco === hueco);

      const calcularUbicacion = (item) => {
        const zonaSugerida = (item.proveedor || "").toLowerCase().includes("uxcar") ? "arriba" : "abajo";
        if (item.expediente) {
          const mismos = ocupadosSimulado.filter((o) => o.expediente === item.expediente);
          for (const m of mismos) {
            for (let h = 1; h <= ZONAS_CRISTALES[m.zona].huecos; h++) {
              if (!estaOcupado(m.zona, m.fila, h)) return { zona: m.zona, fila: m.fila, hueco: h };
            }
          }
        }
        const cfg = ZONAS_CRISTALES[zonaSugerida];
        const filaReservada = FILA_RESERVA[zonaSugerida];
        const filasNormales = Array.from({ length: cfg.filas }, (_, i) => i + 1).filter((f) => f !== filaReservada);
        for (const f of filasNormales) {
          for (let h = 1; h <= cfg.huecos; h++) {
            if (!estaOcupado(zonaSugerida, f, h)) return { zona: zonaSugerida, fila: f, hueco: h };
          }
        }
        if (filaReservada) {
          for (let h = 1; h <= cfg.huecos; h++) {
            if (!estaOcupado(zonaSugerida, filaReservada, h)) return { zona: zonaSugerida, fila: filaReservada, hueco: h };
          }
        }
        return null;
      };

      let sinHueco = 0;
      const hoy = new Date().toISOString().slice(0, 10);
      items.forEach((it) => {
        const datosBase = {
          lote: it.lote || "", secuencia: it.secuencia || "", cliente: it.cliente || "",
          proveedor: it.proveedor || "", expediente: it.expediente || "", medida: it.medida || "",
          cantidad: it.cantidad || 1,
        };
        const ubicacion = calcularUbicacion(datosBase);
        if (ubicacion) {
          ocupadosSimulado.push({ ...ubicacion, expediente: datosBase.expediente });
          onAdd({ ...datosBase, ubicacion, estado: "Colocado", fechaColocado: hoy });
        } else {
          sinHueco++;
          onAdd(datosBase);
        }
      });
      if (sinHueco > 0) {
        setErrorPacking(`Aviso: el almacén está lleno y ${sinHueco} caballete(s) se han guardado sin ubicar. Colócalos a mano cuando haya sitio.`);
      }
    } catch (e) {
      console.error("Error leyendo packing list:", e);
      setErrorPacking("No se pudo leer el archivo. Prueba de nuevo con otra foto o PDF. (" + e.message + ")");
    } finally {
      setLeyendoPacking(false);
    }
  };

  return (
    <div>
      <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-4">
        La ubicación se guarda por <b>caballete completo</b> (no por cristal individual). Arriba van los caballetes de Uxcar (3 filas), abajo los de ALUMAVEL (2 filas), 15 huecos por fila.
      </div>

      <div className="flex flex-wrap items-center gap-2 mb-4">
        <div className="relative flex-1 min-w-[200px]">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por lote, secuencia, cliente, medida..." className={inputCls + " pl-9"} />
        </div>
        <button type="button" onClick={() => inputPackingRef.current?.click()} disabled={leyendoPacking}
          style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
          className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 disabled:opacity-50 px-3.5 py-2 rounded-md">
          <ImageIcon size={14} /> {leyendoPacking ? "Leyendo..." : "Importar packing list (foto/PDF)"}
        </button>
        <input ref={inputPackingRef} type="file" accept="image/*,application/pdf" className="hidden"
          onChange={(e) => { if (e.target.files?.[0]) leerPackingList(e.target.files[0]); e.target.value = ""; }} />
        <button type="button" onClick={() => setMostrarNuevo(true)} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
          <Plus size={14} /> Añadir a mano
        </button>
      </div>
      {errorPacking && (
        <div className="flex items-start justify-between gap-3 mb-4 px-3 py-2 rounded-md bg-rose-50 border border-rose-200">
          <p className="text-xs text-rose-600 font-semibold">⚠ {errorPacking}</p>
          <button onClick={() => setErrorPacking("")} className="text-rose-400 hover:text-rose-600 text-xs font-bold shrink-0">✕</button>
        </div>
      )}

      {mostrarNuevo && (
        <NuevoCristalForm onCancel={() => setMostrarNuevo(false)} onSave={(data) => { onAdd(data); setMostrarNuevo(false); }} />
      )}

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        <button onClick={() => setSubTab("pendientes")}
          className={`px-4 py-2 text-sm font-semibold border-b-2 -mb-px transition flex items-center gap-1.5 ${subTab === "pendientes" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Pendientes de ubicar {pendientes.length > 0 && <Badge className="bg-amber-50 text-amber-700 ring-amber-200">{pendientes.length}</Badge>}
        </button>
        <button onClick={() => setSubTab("mapa")}
          className={`px-4 py-2 text-sm font-semibold border-b-2 -mb-px transition ${subTab === "mapa" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Mapa del almacén
        </button>
        <button onClick={() => setSubTab("estadisticas")}
          className={`px-4 py-2 text-sm font-semibold border-b-2 -mb-px transition ${subTab === "estadisticas" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Estadísticas
        </button>
      </div>

      {subTab === "pendientes" && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          {pendientes.length > 0 && (
            <div className="flex justify-end px-4 pt-3">
              <button
                onClick={() => {
                  if (window.confirm(`¿Borrar los ${pendientes.length} caballete(s) pendientes que ves ahora en la lista (según el buscador)? Esta acción no se puede deshacer.`)) {
                    pendientes.forEach((c) => onDelete(c.id));
                  }
                }}
                className="text-xs font-semibold text-rose-600 border border-rose-200 px-3 py-1.5 rounded-md hover:bg-rose-50"
              >
                Vaciar {q ? "esta lista filtrada" : "todos los pendientes"} ({pendientes.length})
              </button>
            </div>
          )}
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                <th className="px-4 py-2.5 font-semibold">Lote / Secuencia</th>
                <th className="px-4 py-2.5 font-semibold">Cliente</th>
                <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                <th className="px-4 py-2.5 font-semibold">Medida</th>
                <th className="px-4 py-2.5 font-semibold">Llegada</th>
                <th className="px-4 py-2.5 font-semibold text-right">Acción</th>
              </tr>
            </thead>
            <tbody>
              {pendientes.map((c) => (
                <tr key={c.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50">
                  <td className="px-4 py-2.5 font-medium text-slate-800">{c.lote || "—"} {c.secuencia && `/ ${c.secuencia}`}</td>
                  <td className="px-4 py-2.5 text-slate-600">{c.cliente || "—"}</td>
                  <td className="px-4 py-2.5 text-slate-600">{c.proveedor || "—"}</td>
                  <td className="px-4 py-2.5 text-slate-600 font-mono-num">{c.medida || "—"}</td>
                  <td className="px-4 py-2.5 text-slate-500">{fmtDate(c.fechaLlegada)}</td>
                  <td className="px-4 py-2.5 text-right">
                    <div className="flex gap-1.5 justify-end">
                      <button onClick={() => setAsignando(c)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold hover:opacity-90 px-3 py-1.5 rounded-md">
                        Asignar ubicación
                      </button>
                      <button
                        onClick={() => { if (window.confirm("¿Borrar este caballete pendiente?")) onDelete(c.id); }}
                        className="text-xs font-semibold text-rose-600 border border-rose-200 px-2.5 py-1.5 rounded-md hover:bg-rose-50"
                      >
                        Borrar
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
              {pendientes.length === 0 && (
                <tr><td colSpan={6} className="px-4 py-10 text-center text-slate-400 text-sm">No hay caballetes pendientes de ubicar.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {subTab === "mapa" && (
        <MapaAlmacenCristales cristales={cristales} q={q} onVerHueco={setVerDetalle} onAsignarDesdeMapa={setAsignando} />
      )}

      {subTab === "estadisticas" && (
        <EstadisticasCristales cristales={cristales} />
      )}

      {asignando && (
        <UbicacionPicker
          cristal={asignando}
          cristales={cristales}
          sugerencia={sugerirUbicacion(asignando)}
          onClose={() => setAsignando(null)}
          onConfirmar={(ubicacion) => { const ok = onUbicar(asignando.id, ubicacion); if (ok) setAsignando(null); }}
        />
      )}

      {verDetalle && (
        <DetalleHuecoModal
          ubicacion={verDetalle}
          cristalesEnHueco={cristales.filter((c) => c.ubicacion && c.ubicacion.zona === verDetalle.zona && c.ubicacion.fila === verDetalle.fila && c.ubicacion.hueco === verDetalle.hueco)}
          onClose={() => setVerDetalle(null)}
          onLiberar={(id) => onLiberar(id)}
          onEliminar={(id) => onDelete(id)}
          onMover={(cristal) => { setVerDetalle(null); setAsignando(cristal); }}
        />
      )}
    </div>
  );
}

function NuevoCristalForm({ onCancel, onSave }) {
  const [f, setF] = useState({ lote: "", secuencia: "", cliente: "", proveedor: "", expediente: "", medida: "", cantidad: 1, fechaLlegada: new Date().toISOString().slice(0, 10) });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  return (
    <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4">
      <h3 className="font-display font-bold text-slate-800 mb-3">Nuevo caballete</h3>
      <div className="grid grid-cols-3 gap-3 mb-3">
        <Field label="Lote"><TextInput value={f.lote} onChange={set("lote")} /></Field>
        <Field label="Secuencia"><TextInput value={f.secuencia} onChange={set("secuencia")} /></Field>
        <Field label="Expediente"><TextInput value={f.expediente} onChange={set("expediente")} /></Field>
        <Field label="Cliente"><TextInput value={f.cliente} onChange={set("cliente")} /></Field>
        <Field label="Proveedor"><TextInput value={f.proveedor} onChange={set("proveedor")} placeholder="Ej: Uxcar" /></Field>
        <Field label="Medida"><TextInput value={f.medida} onChange={set("medida")} placeholder="Ej: 1200x1500" /></Field>
      </div>
      <div className="flex justify-end gap-2">
        <button onClick={onCancel} className="px-4 py-2 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
        <button onClick={() => onSave(f)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-sm font-semibold hover:opacity-90 px-4 py-2 rounded-md">Guardar</button>
      </div>
    </div>
  );
}

function MapaAlmacenCristales({ cristales, q, onVerHueco, onAsignarDesdeMapa }) {
  const normalizarMapa = (s) => (s || "").toString().toLowerCase().replace(/\s+/g, "").replace(/[×*]/g, "x");
  const nq = normalizarMapa(q);
  const ocupantes = (zona, fila, hueco) => cristales.filter((c) => c.ubicacion && c.ubicacion.zona === zona && c.ubicacion.fila === fila && c.ubicacion.hueco === hueco);
  return (
    <div className="space-y-6">
      {q && (
        <div className="px-4 py-2.5 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold">
          Los huecos resaltados en naranja coinciden con "{q}".
        </div>
      )}
      {Object.entries(ZONAS_CRISTALES).map(([zonaId, cfg]) => (
        <div key={zonaId} className="bg-white border border-slate-200 rounded-lg p-4">
          <h3 className="font-display font-bold text-slate-800 mb-3">{cfg.label}</h3>
          <div className="space-y-1.5">
            {Array.from({ length: cfg.filas }, (_, i) => i + 1).map((fila) => (
              <div key={fila} className="flex items-center gap-2">
                <span className="text-xs text-slate-400 w-14 shrink-0">Fila {fila}</span>
                <div className="flex gap-1.5 flex-1">
                  {Array.from({ length: cfg.huecos }, (_, i) => i + 1).map((hueco) => {
                    const cs = ocupantes(zonaId, fila, hueco);
                    const c = cs[0];
                    const coincide = q && cs.some((x) => normalizarMapa(`${x.lote} ${x.secuencia} ${x.cliente} ${x.proveedor} ${x.expediente} ${x.medida}`).includes(nq));
                    const clase = coincide
                      ? "bg-amber-500 text-white ring-2 ring-amber-300 cursor-pointer hover:opacity-80"
                      : c
                      ? "bg-[#2E8B57] text-white cursor-pointer hover:opacity-80"
                      : "bg-slate-100 text-slate-300";
                    return (
                      <button
                        key={hueco}
                        onClick={() => c && onVerHueco({ zona: zonaId, fila, hueco })}
                        title={c ? cs.map((x) => `${x.lote || ""} ${x.secuencia || ""} — ${x.cliente || ""}`).join(" | ") : "Libre"}
                        className={`relative flex-1 h-10 rounded-md text-[11px] font-semibold flex items-center justify-center px-1 truncate ${clase}`}
                      >
                        {c ? (c.lote || c.secuencia || "•") : "—"}
                        {cs.length > 1 && (
                          <span className="absolute -top-1.5 -right-1.5 bg-slate-800 text-white text-[9px] w-4 h-4 rounded-full flex items-center justify-center">
                            {cs.length}
                          </span>
                        )}
                      </button>
                    );
                  })}
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

function UbicacionPicker({ cristal, cristales, sugerencia, onClose, onConfirmar }) {
  const [zona, setZona] = useState(sugerencia?.zona || "arriba");
  const ocupantesEn = (z, fila, hueco) => cristales.filter((c) => c.id !== cristal.id && c.ubicacion && c.ubicacion.zona === z && c.ubicacion.fila === fila && c.ubicacion.hueco === hueco);
  const cfg = ZONAS_CRISTALES[zona];

  const elegir = (z, fila, hueco) => {
    const ocupantes = ocupantesEn(z, fila, hueco);
    if (ocupantes.length > 0) {
      const nombres = ocupantes.map((o) => o.expediente || o.lote || o.secuencia || "un expediente").join(", ");
      if (!window.confirm(`Ese hueco ya tiene ${ocupantes.length} expediente(s) (${nombres}). ¿Añadir también este caballete ahí, juntos?`)) return;
    }
    onConfirmar({ zona: z, fila, hueco });
  };

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-lg p-5 max-w-lg w-full max-h-[85vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-display font-bold text-slate-800 mb-1">Ubicar: {cristal.lote || cristal.secuencia || "Caballete"}</h3>
        <p className="text-xs text-slate-500 mb-3">{cristal.cliente} · {cristal.proveedor}</p>
        <p className="text-xs text-slate-400 mb-3">Los huecos "Ocupado" no están bloqueados: puedes pulsarlos igualmente si quieres juntar varios expedientes en el mismo caballete.</p>
        {sugerencia && (
          <button
            onClick={() => elegir(sugerencia.zona, sugerencia.fila, sugerencia.hueco)}
            className="w-full flex items-center justify-center gap-2 text-sm font-semibold text-white bg-amber-600 hover:bg-amber-700 px-4 py-2.5 rounded-md mb-3"
          >
            ✨ Usar sugerencia: {ubicacionTexto(sugerencia)}
          </button>
        )}
        <div className="flex gap-2 mb-3">
          {Object.entries(ZONAS_CRISTALES).map(([id, c]) => (
            <button key={id} onClick={() => setZona(id)} className={`flex-1 px-3 py-2 rounded-md text-sm font-semibold border ${zona === id ? "bg-[#2E8B57] text-white border-[#2E8B57]" : "border-slate-300 text-slate-600"}`}>
              {c.label}
            </button>
          ))}
        </div>
        <div className="space-y-1.5 mb-4">
          {Array.from({ length: cfg.filas }, (_, i) => i + 1).map((fila) => (
            <div key={fila} className="flex items-center gap-2">
              <span className="text-xs text-slate-400 w-12 shrink-0">F{fila}</span>
              <div className="flex gap-1.5 flex-1">
                {Array.from({ length: cfg.huecos }, (_, i) => i + 1).map((hueco) => {
                  const ocupantes = ocupantesEn(zona, fila, hueco);
                  const libre = ocupantes.length === 0;
                  return (
                    <button
                      key={hueco}
                      onClick={() => elegir(zona, fila, hueco)}
                      className={`flex-1 h-9 rounded-md text-xs font-semibold ${libre ? "bg-emerald-50 text-emerald-700 hover:bg-emerald-100 cursor-pointer" : "bg-amber-50 text-amber-700 hover:bg-amber-100 cursor-pointer"}`}
                    >
                      {libre ? "Libre" : `Ocupado (${ocupantes.length})`}
                    </button>
                  );
                })}
              </div>
            </div>
          ))}
        </div>
        <button onClick={onClose} className="w-full text-sm font-semibold text-slate-600 px-4 py-2 rounded-md hover:bg-slate-100">Cancelar</button>
      </div>
    </div>
  );
}

function DetalleHuecoModal({ ubicacion, cristalesEnHueco, onClose, onLiberar, onEliminar, onMover }) {
  useEffect(() => {
    if (cristalesEnHueco.length === 0) onClose();
  }, [cristalesEnHueco.length]);

  if (cristalesEnHueco.length === 0) return null;
  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-lg p-5 max-w-md w-full max-h-[85vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-display font-bold text-slate-800 mb-1">{ubicacionTexto(ubicacion)}</h3>
        <p className="text-xs text-slate-500 mb-3">
          {cristalesEnHueco.length === 1 ? "1 expediente en este caballete." : `${cristalesEnHueco.length} expedientes juntos en este mismo caballete.`}
        </p>
        <div className="space-y-3">
          {cristalesEnHueco.map((cristal) => (
            <div key={cristal.id} className="border border-slate-200 rounded-md p-3">
              <div className="text-sm text-slate-600 space-y-1 mb-2">
                <p><b>Lote:</b> {cristal.lote || "—"}</p>
                <p><b>Secuencia:</b> {cristal.secuencia || "—"}</p>
                <p><b>Cliente:</b> {cristal.cliente || "—"}</p>
                <p><b>Proveedor:</b> {cristal.proveedor || "—"}</p>
                <p><b>Expediente:</b> {cristal.expediente || "—"}</p>
                <p><b>Medida:</b> {cristal.medida || "—"}</p>
                <p><b>Colocado el:</b> {fmtDate(cristal.fechaColocado)}</p>
              </div>
              <div className="flex flex-wrap gap-2">
                <button onClick={() => onMover(cristal)} className="flex-1 text-xs font-semibold text-sky-700 border border-sky-300 px-3 py-1.5 rounded-md hover:bg-sky-50">
                  Mover a otro hueco
                </button>
                <button onClick={() => onLiberar(cristal.id)} className="flex-1 text-xs font-semibold text-amber-700 border border-amber-300 px-3 py-1.5 rounded-md hover:bg-amber-50">
                  Liberar (vuelve a pendiente)
                </button>
                <button onClick={() => onEliminar(cristal.id)} className="flex-1 text-xs font-semibold text-rose-600 border border-rose-200 px-3 py-1.5 rounded-md hover:bg-rose-50">
                  Dar de baja este expediente
                </button>
              </div>
            </div>
          ))}
        </div>
        <button onClick={onClose} className="w-full mt-4 text-sm font-semibold text-slate-600 px-4 py-2 rounded-md hover:bg-slate-100">Cerrar</button>
      </div>
    </div>
  );
}

function EstadisticasCristales({ cristales }) {
  const porDia = useMemo(() => {
    const map = {};
    cristales.forEach((c) => {
      if (c.fechaLlegada) map[c.fechaLlegada] = { ...(map[c.fechaLlegada] || { fecha: c.fechaLlegada, llegados: 0, colocados: 0 }), llegados: (map[c.fechaLlegada]?.llegados || 0) + 1 };
      if (c.fechaColocado) map[c.fechaColocado] = { ...(map[c.fechaColocado] || { fecha: c.fechaColocado, llegados: 0, colocados: 0 }), colocados: (map[c.fechaColocado]?.colocados || 0) + 1 };
    });
    return Object.values(map).sort((a, b) => a.fecha.localeCompare(b.fecha)).slice(-30);
  }, [cristales]);

  const porProveedor = useMemo(() => {
    const map = {};
    cristales.forEach((c) => { const k = c.proveedor || "Sin proveedor"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([proveedor, n]) => ({ proveedor, n }));
  }, [cristales]);

  const porCliente = useMemo(() => {
    const map = {};
    cristales.forEach((c) => { const k = c.cliente || "Sin cliente"; map[k] = (map[k] || 0) + 1; });
    return Object.entries(map).map(([cliente, n]) => ({ cliente, n })).sort((a, b) => b.n - a.n).slice(0, 10);
  }, [cristales]);

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-3 gap-3">
        <Kpi label="Total caballetes" value={String(cristales.length)} />
        <Kpi label="Colocados" value={String(cristales.filter((c) => c.estado === "Colocado").length)} />
        <Kpi label="Pendientes" value={String(cristales.filter((c) => c.estado === "Pendiente").length)} />
      </div>

      <div className="bg-white border border-slate-200 rounded-lg p-4">
        <h3 className="font-display font-bold text-slate-800 mb-3">Llegadas vs. colocados por día</h3>
        {porDia.length > 0 ? (
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={porDia}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="fecha" tick={{ fontSize: 11 }} />
              <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="llegados" name="Llegados" stroke="#2E8B57" strokeWidth={2} />
              <Line type="monotone" dataKey="colocados" name="Colocados" stroke="#D97706" strokeWidth={2} />
            </LineChart>
          </ResponsiveContainer>
        ) : <p className="text-sm text-slate-400 text-center py-8">Todavía no hay datos suficientes.</p>}
      </div>

      <div className="grid grid-cols-2 gap-4">
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <h3 className="font-display font-bold text-slate-800 mb-3">Por proveedor</h3>
          {porProveedor.length > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={porProveedor}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="proveedor" tick={{ fontSize: 11 }} />
                <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                <Tooltip />
                <Bar dataKey="n" fill="#2E8B57" />
              </BarChart>
            </ResponsiveContainer>
          ) : <p className="text-sm text-slate-400 text-center py-8">Sin datos.</p>}
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <h3 className="font-display font-bold text-slate-800 mb-3">Top 10 clientes</h3>
          {porCliente.length > 0 ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={porCliente} layout="vertical">
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" tick={{ fontSize: 11 }} allowDecimals={false} />
                <YAxis dataKey="cliente" type="category" tick={{ fontSize: 10 }} width={90} />
                <Tooltip />
                <Bar dataKey="n" fill="#D97706" />
              </BarChart>
            </ResponsiveContainer>
          ) : <p className="text-sm text-slate-400 text-center py-8">Sin datos.</p>}
        </div>
      </div>
    </div>
  );
}

function FabricaModulo({ proyectos, pedidos, proveedores, materiales, clientes, onConfirmarLinea, onIniciarFabricacion, cristales, onAddCristal, onUpdateCristal, onDeleteCristal, onUbicarCristal, onLiberarCristal }) {
  const [q, setQ] = useState("");
  const [tab, setTab] = useState("listo");
  const proveedorNombre = (id) => proveedores.find((p) => p.id === id)?.nombre || "—";
  const materialInfo = (id) => materiales.find((m) => m.id === id);

  const proyectosEnFabricacion = proyectos.filter((p) => ["En proceso", "Albarán de carga firmado", "Listo para reparto/recogida"].includes(p.estadoTrabajo));

  const pedidosEnCurso = pedidos.filter((p) => p.estado !== "Cancelado" && (p.lineas || []).some((l) => !l.confirmadoFabrica));

  const filtered = useMemo(() => {
    if (!q) return pedidosEnCurso;
    return pedidosEnCurso.filter((p) => {
      const s = `${p.numero} ${proveedorNombre(p.proveedorId)} ${(p.lineas || []).map((l) => l.modo === "libre" ? l.referencia : materialInfo(l.materialId)?.descripcion).join(" ")}`.toLowerCase();
      return s.includes(q.toLowerCase());
    });
  }, [pedidosEnCurso, q]);

  // Proyectos donde TODO coincide: oficina ha recibido todos los pedidos, fábrica ha
  // confirmado la llegada de todas las líneas, y el checklist "Qué lleva la obra" está
  // completo. Solo cuando las tres cosas cuadran, el proyecto está listo para fabricar.
  const listoParaFabricar = useMemo(() => {
    return proyectos.filter((p) => {
      if (["Entregado", "Cancelado"].includes(p.estadoTrabajo) || p.estadoTrabajo === "En proceso" || p.estadoTrabajo === "Albarán de carga firmado" || p.estadoTrabajo === "Listo para reparto/recogida") return false;
      const misPedidos = pedidos.filter((pd) => pd.proyectoId === p.id && pd.estado !== "Cancelado");
      if (misPedidos.length === 0) return false;
      const oficinaOk = misPedidos.every((pd) => pd.estado === "Recibido");
      const fabricaOk = misPedidos.every((pd) => (pd.lineas || []).every((l) => l.confirmadoFabrica));
      const checklist = normalizarChecklist(p.checklistMateriales);
      const checklistOk = checklist.length > 0 && checklist.every((c) => c.estado);
      return oficinaOk && fabricaOk && checklistOk;
    });
  }, [proyectos, pedidos]);

  const descargarWord = () => {
    const titulos = { listo: "Listo para fabricar", materiales: "Materiales pendientes", enfab: "En fabricación" };
    const filaTabla = (cols) => `<tr>${cols.map((c) => `<td style="border:1px solid #ccc;padding:6px 10px;">${c}</td>`).join("")}</tr>`;
    const cabeceraTabla = (cols) => `<tr>${cols.map((c) => `<th style="border:1px solid #ccc;padding:6px 10px;background:#f1f1f1;text-align:left;">${c}</th>`).join("")}</tr>`;

    let cuerpo = "";
    if (tab === "listo") {
      cuerpo = `
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Proyecto", "Entrega prevista", "Fabricación prevista"])}
          ${listoParaFabricar.map((p) => filaTabla([`#${p.numero} — ${p.nombre}`, fmtDate(p.fechaEntregaPrevista) || "—", fmtDate(p.fechaFabricacion) || "—"])).join("")}
        </table>`;
    } else if (tab === "materiales") {
      cuerpo = filtered.map((pedido) => {
        const proyecto = pedido.proyectoId ? proyectos.find((p) => p.id === pedido.proyectoId) : null;
        const filasLineas = pedido.lineas.map((l) => {
          const esLibre = l.modo === "libre";
          const mat = esLibre ? null : materialInfo(l.materialId);
          const nombre = esLibre ? (l.referencia || "Sin referencia") : (mat ? `${mat.codigo} — ${mat.descripcion}` : "Material eliminado");
          return filaTabla([nombre, l.cantidad, l.confirmadoFabrica ? `Ha llegado (${fmtDate(l.fechaConfirmadoFabrica)})` : "Pendiente"]);
        }).join("");
        return `<h3>Pedido #${pedido.numero} — ${proveedorNombre(pedido.proveedorId)}${proyecto ? ` · Proyecto #${proyecto.numero} ${proyecto.nombre}` : ""}</h3>
          <table style="border-collapse:collapse;width:100%;margin-bottom:16px;">
            ${cabeceraTabla(["Material", "Cantidad", "Estado en fábrica"])}
            ${filasLineas}
          </table>`;
      }).join("");
    } else {
      cuerpo = `
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Proyecto", "Estado", "Entrega prevista"])}
          ${proyectosEnFabricacion.map((p) => filaTabla([`#${p.numero} — ${p.nombre}`, p.estadoTrabajo, fmtDate(p.fechaEntregaPrevista) || "—"])).join("")}
        </table>`;
    }

    const htmlDoc = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="utf-8"><title>Fábrica</title></head>
      <body style="font-family: Calibri, Arial, sans-serif; color:#1a1a1a;">
        <h1 style="color:#2E8B57;">Fábrica — ${titulos[tab]}</h1>
        <p style="color:#666;">Generado el ${fmtDate(new Date().toISOString().slice(0, 10))}</p>
        ${cuerpo}
      </body>
      </html>
    `;
    const blob = new Blob(["\ufeff", htmlDoc], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `fabrica_${tab}_${new Date().toISOString().slice(0, 10)}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="p-8">
      <Header icon={<Factory size={20} className="text-[#2E8B57]" />} title="Fábrica" manualKey="fabrica" subtitle="Control desde almacén — independiente del control de oficina" />

      <div className="flex flex-wrap items-center gap-1 mb-6 border-b border-slate-200">
        <button onClick={() => setTab("listo")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition flex items-center gap-1.5 ${tab === "listo" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Listo para fabricar
          {listoParaFabricar.length > 0 && <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">{listoParaFabricar.length}</Badge>}
        </button>
        <button onClick={() => setTab("materiales")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === "materiales" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Materiales pendientes
        </button>
        <button onClick={() => setTab("enfab")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === "enfab" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          En fabricación
        </button>
        <button onClick={() => setTab("cristales")}
          className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition flex items-center gap-1.5 ${tab === "cristales" ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
          Cristales
          {cristales.filter((c) => c.estado === "Pendiente").length > 0 && <Badge className="bg-amber-50 text-amber-700 ring-amber-200">{cristales.filter((c) => c.estado === "Pendiente").length}</Badge>}
        </button>
        {tab !== "cristales" && (
        <button onClick={descargarWord} className="ml-auto mb-1 flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
        )}
      </div>

      {tab === "listo" && (
        <>
          <div className="px-4 py-3 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-800 text-sm mb-6">
            Aquí solo aparecen los proyectos donde coincide todo: oficina ha recibido los pedidos, fábrica ha confirmado que ha llegado el material, y el checklist "Qué lleva la obra" está completo. Nada más entrar aquí, ya se puede empezar a fabricar.
          </div>
          <div className="space-y-3">
            {listoParaFabricar.map((p) => (
              <div key={p.id} className="bg-white border border-emerald-200 rounded-lg p-4 flex flex-wrap items-center justify-between gap-3">
                <div>
                  <div className="font-semibold text-slate-800">#{p.numero} — {p.nombre}</div>
                  <div className="text-xs text-slate-500 mt-0.5">Entrega prevista {fmtDate(p.fechaEntregaPrevista)}{p.fechaFabricacion ? ` · Fabricación prevista ${fmtDate(p.fechaFabricacion)}` : ""}</div>
                </div>
                <button
                  onClick={() => onIniciarFabricacion(p.id)}
                  style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
                  className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 px-4 py-2.5 rounded-md"
                >
                  <Factory size={15} /> Empezar a fabricar
                </button>
              </div>
            ))}
            {listoParaFabricar.length === 0 && (
              <div className="bg-white border border-slate-200 rounded-lg px-4 py-10 text-center text-slate-400 text-sm">
                Todavía no hay ningún proyecto con todo listo (oficina + fábrica + checklist completos).
              </div>
            )}
          </div>
        </>
      )}

      {tab === "enfab" && (
        <>
          <h2 className="font-display font-bold text-slate-800 mb-3">Proyectos en fabricación ({proyectosEnFabricacion.length})</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Proyecto</th>
                  <th className="px-4 py-2.5 font-semibold">Estado</th>
                  <th className="px-4 py-2.5 font-semibold">Entrega prevista</th>
                </tr>
              </thead>
              <tbody>
                {proyectosEnFabricacion.map((p) => (
                  <tr key={p.id} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-medium text-slate-800">#{p.numero} — {p.nombre}</td>
                    <td className="px-4 py-2.5"><Badge className={ESTADO_TRABAJO_STYLE[p.estadoTrabajo]}>{p.estadoTrabajo}</Badge></td>
                    <td className="px-4 py-2.5 text-slate-500">{fmtDate(p.fechaEntregaPrevista)}</td>
                  </tr>
                ))}
                {proyectosEnFabricacion.length === 0 && (
                  <tr><td colSpan={3} className="px-4 py-8 text-center text-slate-400 text-sm">No hay proyectos en fabricación ahora mismo.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "materiales" && (
        <>
          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-6">
            Cuando llegue un material (cristal, persiana, aluminio, PVC...), márcalo aquí como "Ha llegado". Este control es tuyo, no depende de lo que oficina haya marcado — así luego se puede comparar y ver si todo cuadra.
          </div>

          <div className="relative mb-4 max-w-sm">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por proveedor, material, referencia..." className={inputCls + " pl-9"} />
          </div>

          <div className="space-y-4">
            {filtered.map((pedido) => {
              const proyecto = pedido.proyectoId ? proyectos.find((p) => p.id === pedido.proyectoId) : null;
              return (
                <div key={pedido.id} className="bg-white border border-slate-200 rounded-lg overflow-hidden">
                  <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-2">
                    <div>
                      <span className="font-mono-num text-sm text-[#2E8B57] font-bold">#{pedido.numero}</span>
                      <span className="ml-2 font-semibold text-slate-700">{proveedorNombre(pedido.proveedorId)}</span>
                      {proyecto && <span className="ml-2 text-xs text-slate-500">Proyecto #{proyecto.numero} — {proyecto.nombre}</span>}
                    </div>
                    <span className="text-xs text-slate-400">Entrega prevista {fmtDate(pedido.fechaEntregaPrevista)}</span>
                  </div>
                  <div className="divide-y divide-slate-100">
                    {pedido.lineas.map((l) => {
                      const esLibre = l.modo === "libre";
                      const mat = esLibre ? null : materialInfo(l.materialId);
                      const nombre = esLibre ? (l.referencia || "Sin referencia") : (mat ? `${mat.codigo} — ${mat.descripcion}` : "Material eliminado");
                      return (
                        <label key={l.id} className="flex items-center gap-3 px-4 py-3 cursor-pointer hover:bg-slate-50">
                          <input
                            type="checkbox"
                            checked={!!l.confirmadoFabrica}
                            onChange={(e) => onConfirmarLinea(pedido.id, l.id, e.target.checked)}
                            className="w-5 h-5 rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]"
                          />
                          <div className="flex-1">
                            <div className={`text-sm font-medium ${l.confirmadoFabrica ? "text-slate-400 line-through" : "text-slate-800"}`}>
                              {nombre}{esLibre && (l.ancho || l.alto) ? ` (${l.ancho || "—"} x ${l.alto || "—"})` : ""} — {l.cantidad} ud.
                            </div>
                            {l.confirmadoFabrica && <div className="text-xs text-emerald-600">Ha llegado · {fmtDate(l.fechaConfirmadoFabrica)} · {l.confirmadoPorFabrica}</div>}
                          </div>
                        </label>
                      );
                    })}
                  </div>
                </div>
              );
            })}
            {filtered.length === 0 && (
              <div className="bg-white border border-slate-200 rounded-lg px-4 py-10 text-center text-slate-400 text-sm">
                No hay materiales pendientes de confirmar llegada.
              </div>
            )}
          </div>
        </>
      )}

      {tab === "cristales" && (
        <CristalesModulo
          cristales={cristales}
          proyectos={proyectos}
          proveedores={proveedores}
          clientes={clientes}
          onAdd={onAddCristal}
          onUpdate={onUpdateCristal}
          onDelete={onDeleteCristal}
          onUbicar={onUbicarCristal}
          onLiberar={onLiberarCristal}
        />
      )}
    </div>
  );
}

/* ================= INSTALACIONES ================= */

const ESTADO_INSTALACION = ["Pendiente de instalación", "En curso", "Finalizada"];
const ESTADO_INSTALACION_STYLE = {
  "Pendiente de instalación": "bg-amber-50 text-amber-700 ring-amber-200",
  "En curso": "bg-sky-50 text-sky-700 ring-sky-200",
  "Finalizada": "bg-emerald-50 text-emerald-700 ring-emerald-200",
};

function InstalacionesModulo({ instalaciones, proyectos, clientes, vehiculos, onUpsertVehiculo, onDeleteVehiculo, view, setView, detailId, setDetailId, onUpdate, onCrearManual, onAddHora, onDeleteHora, onAddGasto, onDeleteGasto, onAddMaterialFurgoneta, onCicloMaterialFurgoneta, onDeleteMaterialFurgoneta, isAdmin, incidencias, onUpsertIncidencia, materiales, usuarios, onCrearTarea, currentUser }) {
  const [tabPrincipal, setTabPrincipal] = useState("lista");
  const [q, setQ] = useState("");
  const [estadoFiltro, setEstadoFiltro] = useState("");

  const datos = (inst) => {
    const proyecto = inst.proyectoId ? proyectos.find((p) => p.id === inst.proyectoId) : null;
    const cliente = proyecto ? clientes.find((c) => c.id === proyecto.clienteId) : null;
    const nombreMostrar = proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : (inst.nombre || "Instalación sin proyecto");
    const clienteMostrar = cliente?.nombre || inst.clienteNombre || "—";
    const totalHoras = (inst.registroHoras || []).reduce((s, r) => s + (parseFloat(r.horas) || 0), 0);
    const costeManoObra = totalHoras * (parseFloat(inst.costeHora) || 0);
    const totalGastos = (inst.gastos || []).reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
    const costeTotal = costeManoObra + totalGastos;
    const presupuesto = parseFloat(inst.presupuestoInstalacion) || 0;
    const diferencia = presupuesto - costeTotal;
    return { proyecto, cliente, nombreMostrar, clienteMostrar, totalHoras, costeManoObra, totalGastos, costeTotal, presupuesto, diferencia };
  };

  // Si no eres admin, solo ves las instalaciones donde estás asignado como instalador.
  const instalacionesVisibles = isAdmin
    ? instalaciones
    : instalaciones.filter((i) => (i.instaladoresAsignados || []).includes(currentUser?.id));

  const filtered = useMemo(() => {
    return instalacionesVisibles.filter((i) => {
      if (estadoFiltro && i.estado !== estadoFiltro) return false;
      if (!q) return true;
      const { nombreMostrar, clienteMostrar } = datos(i);
      return `${nombreMostrar} ${clienteMostrar}`.toLowerCase().includes(q.toLowerCase());
    });
  }, [instalacionesVisibles, q, estadoFiltro, proyectos, clientes]);

  if (view === "form") {
    return (
      <InstalacionForm
        proyectos={proyectos}
        clientes={clientes}
        onCancel={() => setView("list")}
        onSave={onCrearManual}
      />
    );
  }

  if (view === "detail") {
    const instalacion = instalacionesVisibles.find((i) => i.id === detailId);
    if (!instalacion) { setView("list"); return null; }
    const proyecto = instalacion.proyectoId ? proyectos.find((p) => p.id === instalacion.proyectoId) : null;
    return (
      <InstalacionDetail
        instalacion={instalacion}
        proyecto={proyecto}
        cliente={proyecto ? clientes.find((c) => c.id === proyecto.clienteId) : null}
        onBack={() => setView("list")}
        onUpdate={(patch) => onUpdate(instalacion.id, patch)}
        onAddHora={(r) => onAddHora(instalacion.id, r)}
        onDeleteHora={(rid) => onDeleteHora(instalacion.id, rid)}
        onAddGasto={(g) => onAddGasto(instalacion.id, g)}
        onDeleteGasto={(gid) => onDeleteGasto(instalacion.id, gid)}
        onAddMaterialFurgoneta={(nombre) => onAddMaterialFurgoneta(instalacion.id, nombre)}
        onCicloMaterialFurgoneta={(itemId) => onCicloMaterialFurgoneta(instalacion.id, itemId)}
        onDeleteMaterialFurgoneta={(itemId) => onDeleteMaterialFurgoneta(instalacion.id, itemId)}
        vehiculos={vehiculos}
        otrasInstalaciones={instalaciones.filter((i) => i.id !== instalacion.id)}
        proyectos={proyectos}
        isAdmin={isAdmin}
        incidencias={incidencias}
        onUpsertIncidencia={onUpsertIncidencia}
        materiales={materiales}
        usuarios={usuarios}
        onCrearTarea={onCrearTarea}
      />
    );
  }

  if (tabPrincipal === "vehiculos") {
    return (
      <VehiculosModulo
        vehiculos={vehiculos}
        instalaciones={instalacionesVisibles}
        proyectos={proyectos}
        onUpsert={onUpsertVehiculo}
        onDelete={onDeleteVehiculo}
        onVolver={() => setTabPrincipal("lista")}
      />
    );
  }

  if (tabPrincipal === "furgoneta") {
    return (
      <FurgonetaModulo
        instalaciones={instalacionesVisibles}
        proyectos={proyectos}
        onCiclo={onCicloMaterialFurgoneta}
        onVerInstalacion={(id) => { setDetailId(id); setView("detail"); }}
        onVolver={() => setTabPrincipal("lista")}
      />
    );
  }

  return (
    <div className="p-8">
      <Header
        icon={<Wrench size={20} className="text-[#2E8B57]" />}
        title="Instalaciones"
        manualKey="instalaciones"
        subtitle={`${instalacionesVisibles.length} instalación${instalacionesVisibles.length === 1 ? "" : "es"} registrada${instalacionesVisibles.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => setView("form")}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-3 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVA INSTALACIÓN (SIN PROYECTO)
      </button>

      <button
        onClick={() => setTabPrincipal("furgoneta")}
        style={{ borderColor: "#2E8B57", color: "#2E8B57" }}
        className="w-full flex items-center justify-center gap-2 border-2 hover:bg-slate-50 text-sm font-semibold py-3 rounded-lg mb-3 cursor-pointer select-none"
      >
        <Truck size={16} /> Ver furgoneta — qué falta cargar / qué está en obra
      </button>

      <button
        onClick={() => setTabPrincipal("vehiculos")}
        className="w-full flex items-center justify-center gap-2 border-2 border-slate-300 text-slate-600 hover:bg-slate-50 text-sm font-semibold py-3 rounded-lg mb-6 cursor-pointer select-none"
      >
        <Truck size={16} /> Gestionar vehículos de la flota
      </button>

      <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-6">
        Las instalaciones de un proyecto se crean solas al marcar "Este proyecto lleva instalación" en su ficha. Si necesitas una instalación suelta, sin proyecto en el CRM, créala aquí arriba. Lleva el presupuesto, el parte diario de horas de la montadora y los gastos (dietas, etc.) para comparar.
      </div>

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 min-w-[200px]">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por proyecto o cliente..." className={inputCls + " pl-9"} />
        </div>
        <Select value={estadoFiltro} onChange={(e) => setEstadoFiltro(e.target.value)} className="max-w-[220px]">
          <option value="">Todos los estados</option>
          {ESTADO_INSTALACION.map((t) => <option key={t}>{t}</option>)}
        </Select>
        <button
          onClick={() => descargarListaComoWord(
            "Instalaciones",
            ["Instalación", "Cliente", "Estado", "Presupuesto", "Coste actual", "Diferencia"],
            filtered.map((i) => {
              const d = datos(i);
              return [d.nombreMostrar, d.clienteMostrar, i.estado, d.presupuesto ? money(d.presupuesto) : "—", money(d.costeTotal), d.presupuesto ? money(d.diferencia) : "—"];
            })
          )}
          className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
        >
          <FileText size={14} /> Descargar esta vista (Word)
        </button>
      </div>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Proyecto</th>
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
              <th className="px-4 py-3 font-semibold text-right">Presupuesto</th>
              <th className="px-4 py-3 font-semibold text-right">Coste actual</th>
              <th className="px-4 py-3 font-semibold text-right">Diferencia</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((i) => {
              const d = datos(i);
              return (
                <tr key={i.id} onClick={() => { setDetailId(i.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                  <td className="px-4 py-3 font-medium text-slate-800">{d.nombreMostrar}{!d.proyecto && <Badge className="ml-2 bg-violet-50 text-violet-700 ring-violet-200">Sin proyecto</Badge>}</td>
                  <td className="px-4 py-3 text-slate-600">{d.clienteMostrar}</td>
                  <td className="px-4 py-3"><Badge className={ESTADO_INSTALACION_STYLE[i.estado]}>{i.estado}</Badge></td>
                  <td className="px-4 py-3 text-right font-mono-num text-slate-500">{d.presupuesto ? money(d.presupuesto) : "—"}</td>
                  <td className="px-4 py-3 text-right font-mono-num text-slate-600">{money(d.costeTotal)}</td>
                  <td className={`px-4 py-3 text-right font-mono-num font-semibold ${d.diferencia < 0 ? "text-rose-600" : "text-emerald-600"}`}>{d.presupuesto ? money(d.diferencia) : "—"}</td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={6} className="px-4 py-10 text-center text-slate-400 text-sm">No hay instalaciones que coincidan.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function VehiculosModulo({ vehiculos, instalaciones, proyectos, onUpsert, onDelete, onVolver }) {
  const [nombre, setNombre] = useState("");
  const [tipo, setTipo] = useState("Furgoneta pequeña");
  const [editandoId, setEditandoId] = useState(null);

  const nombreInstalacion = (inst) => {
    const p = inst.proyectoId ? proyectos.find((pr) => pr.id === inst.proyectoId) : null;
    return p ? `#${p.numero} — ${p.nombre}` : (inst.nombre || "Instalación sin proyecto");
  };

  const asignacionesProximas = instalaciones
    .filter((i) => i.vehiculoId && i.fechaMontaje && i.fechaMontaje >= new Date().toISOString().slice(0, 10))
    .sort((a, b) => a.fechaMontaje.localeCompare(b.fechaMontaje));

  const submit = (e) => {
    e.preventDefault();
    if (!nombre.trim()) return;
    onUpsert({ id: editandoId, nombre, tipo });
    setNombre(""); setTipo("Furgoneta pequeña"); setEditandoId(null);
  };

  const editar = (v) => { setEditandoId(v.id); setNombre(v.nombre); setTipo(v.tipo); };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onVolver} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver a Instalaciones</button>

      <Header icon={<Truck size={20} className="text-[#2E8B57]" />} title="Vehículos de la flota" manualKey="vehiculos" subtitle="Furgonetas y camiones disponibles para asignar a instalaciones" />

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-4 flex flex-wrap gap-2 items-end mb-4">
        <Field label="Nombre del vehículo">
          <TextInput value={nombre} onChange={(e) => setNombre(e.target.value)} placeholder="Ej: Furgoneta pequeña 1" />
        </Field>
        <Field label="Tipo">
          <Select value={tipo} onChange={(e) => setTipo(e.target.value)}>
            <option>Furgoneta pequeña</option>
            <option>Furgoneta grande</option>
            <option>Camión</option>
          </Select>
        </Field>
        <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1 text-sm font-semibold px-4 py-2 rounded-md h-[38px]">
          <Plus size={15} /> {editandoId ? "Guardar cambios" : "Añadir vehículo"}
        </button>
      </form>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-8">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Vehículo</th>
              <th className="px-4 py-2.5 font-semibold">Tipo</th>
              <th className="px-4 py-2.5 font-semibold"></th>
            </tr>
          </thead>
          <tbody>
            {vehiculos.map((v) => (
              <tr key={v.id} className="border-b border-slate-100 last:border-0">
                <td className="px-4 py-2.5 font-medium text-slate-800">{v.nombre}</td>
                <td className="px-4 py-2.5 text-slate-500">{v.tipo}</td>
                <td className="px-4 py-2.5 flex gap-2 justify-end">
                  <button onClick={() => editar(v)} className="text-slate-400 hover:text-slate-700"><Pencil size={14} /></button>
                  <button onClick={() => onDelete(v.id)} className="text-slate-400 hover:text-rose-600"><Trash2 size={14} /></button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h2 className="font-display font-bold text-slate-800 mb-3">Próximas asignaciones (para ver de un vistazo si algo se solapa)</h2>
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Fecha</th>
              <th className="px-4 py-2.5 font-semibold">Vehículo</th>
              <th className="px-4 py-2.5 font-semibold">Instalación</th>
            </tr>
          </thead>
          <tbody>
            {asignacionesProximas.map((i) => {
              const v = vehiculos.find((veh) => veh.id === i.vehiculoId);
              const solapa = asignacionesProximas.some((otra) => otra.id !== i.id && otra.vehiculoId === i.vehiculoId && otra.fechaMontaje === i.fechaMontaje);
              return (
                <tr key={i.id} className={`border-b border-slate-100 last:border-0 ${solapa ? "bg-rose-50" : ""}`}>
                  <td className="px-4 py-2.5 text-slate-600">{fmtDate(i.fechaMontaje)}{solapa && " ⚠"}</td>
                  <td className="px-4 py-2.5 font-medium text-slate-800">{v?.nombre || "—"}</td>
                  <td className="px-4 py-2.5 text-slate-600">{nombreInstalacion(i)}</td>
                </tr>
              );
            })}
            {asignacionesProximas.length === 0 && (
              <tr><td colSpan={3} className="px-4 py-8 text-center text-slate-400 text-sm">Todavía no hay vehículos asignados a ninguna fecha.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FurgonetaModulo({ instalaciones, proyectos, onCiclo, onVerInstalacion, onVolver }) {
  const nombreInstalacion = (inst) => {
    const proyecto = inst.proyectoId ? proyectos.find((p) => p.id === inst.proyectoId) : null;
    return proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : (inst.nombre || "Instalación sin proyecto");
  };

  const activas = instalaciones.filter((i) => i.estado !== "Finalizada" && (i.materialFurgoneta || []).length > 0);
  const pendientesCargar = [];
  const enFurgoneta = [];
  const enObra = [];
  activas.forEach((inst) => {
    (inst.materialFurgoneta || []).forEach((m) => {
      const fila = { inst, item: m };
      if (m.estado === "Pendiente de cargar") pendientesCargar.push(fila);
      else if (m.estado === "En la furgoneta") enFurgoneta.push(fila);
      else enObra.push(fila);
    });
  });

  const Columna = ({ titulo, filas, color }) => (
    <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
      <div className={`px-4 py-2.5 text-sm font-bold ${color}`}>{titulo} ({filas.length})</div>
      <div className="divide-y divide-slate-100">
        {filas.length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-6">Nada aquí.</p>
        ) : (
          filas.map(({ inst, item }) => (
            <div key={item.id} className="px-4 py-2.5">
              <button onClick={() => onCiclo(inst.id, item.id)} className="text-sm font-medium text-slate-800 hover:text-[#2E8B57] block">{item.nombre}</button>
              <button onClick={() => onVerInstalacion(inst.id)} className="text-xs text-slate-400 hover:text-slate-600">{nombreInstalacion(inst)}</button>
            </div>
          ))
        )}
      </div>
    </div>
  );

  return (
    <div className="p-8">
      <button onClick={onVolver} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver a Instalaciones</button>

      <Header icon={<Truck size={20} className="text-[#2E8B57]" />} title="Furgoneta" manualKey="furgoneta" subtitle="Qué falta cargar, qué va ya en la furgoneta, y qué está en la obra — de todas las instalaciones activas" />

      <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-6">
        Toca un material para pasarlo a la siguiente columna (Pendiente de cargar → En la furgoneta → En la obra). Toca el nombre de la instalación para abrir su ficha.
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <Columna titulo="Pendiente de cargar" filas={pendientesCargar} color="bg-amber-50 text-amber-700" />
        <Columna titulo="En la furgoneta" filas={enFurgoneta} color="bg-sky-50 text-sky-700" />
        <Columna titulo="En la obra" filas={enObra} color="bg-emerald-50 text-emerald-700" />
      </div>
    </div>
  );
}

function InstalacionForm({ proyectos, clientes, onCancel, onSave }) {
  const [f, setF] = useState({ proyectoId: "", nombre: "", clienteNombre: "", presupuestoInstalacion: "", notas: "" });
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const clientesDisponibles = useMemo(() => clientes.map((c) => c.nombre).sort(), [clientes]);

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.proyectoId && !f.nombre.trim()) {
      setErrorMsg("Si no eliges un proyecto, tienes que ponerle un nombre a la instalación (es obligatorio).");
      return;
    }
    setErrorMsg("");
    onSave(f);
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">Nueva instalación</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <Field label="Proyecto (opcional)">
          <Select value={f.proyectoId} onChange={set("proyectoId")}>
            <option value="">Sin proyecto — instalación suelta</option>
            {proyectos.map((p) => <option key={p.id} value={p.id}>#{p.numero} — {p.nombre}</option>)}
          </Select>
          <p className="text-xs text-slate-400 mt-1">Si la eliges de un proyecto, se usa su nombre y su cliente automáticamente.</p>
        </Field>

        {!f.proyectoId && (
          <div className="grid grid-cols-2 gap-4">
            <Field label="Nombre de la instalación" required>
              <TextInput value={f.nombre} onChange={set("nombre")} placeholder="Ej: Instalación puerta garaje Sr. López" />
            </Field>
            <Field label="Cliente">
              <TextInput value={f.clienteNombre} onChange={set("clienteNombre")} list="instalacion-clientes-datalist" placeholder="Nombre del cliente (opcional)" />
              <datalist id="instalacion-clientes-datalist">
                {clientesDisponibles.map((n) => <option key={n} value={n} />)}
              </datalist>
            </Field>
          </div>
        )}

        <Field label="Presupuesto de instalación (€)">
          <TextInput type="number" step="0.01" value={f.presupuestoInstalacion} onChange={set("presupuestoInstalacion")} />
        </Field>

        <Field label="Notas">
          <TextArea rows={2} value={f.notas} onChange={set("notas")} />
        </Field>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Crear instalación</button>
        </div>
      </form>
    </div>
  );
}

function InstalacionDetail({ instalacion, proyecto, cliente, onBack, onUpdate, onAddHora, onDeleteHora, onAddGasto, onDeleteGasto, onAddMaterialFurgoneta, onCicloMaterialFurgoneta, onDeleteMaterialFurgoneta, vehiculos, otrasInstalaciones, proyectos, isAdmin, incidencias, onUpsertIncidencia, materiales, usuarios, onCrearTarea }) {
  const [nuevoMaterial, setNuevoMaterial] = useState("");
  const [errorMaterial, setErrorMaterial] = useState("");
  const [fechaMontajeInput, setFechaMontajeInput] = useState(instalacion.fechaMontaje || "");

  const agregarMaterial = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!nuevoMaterial.trim()) {
      setErrorMaterial("Escribe primero qué material hace falta llevar.");
      return;
    }
    setErrorMaterial("");
    onAddMaterialFurgoneta(nuevoMaterial.trim());
    setNuevoMaterial("");
  };

  const nombreOtraInstalacion = (inst) => {
    const p = inst.proyectoId ? proyectos.find((pr) => pr.id === inst.proyectoId) : null;
    return p ? `#${p.numero} — ${p.nombre}` : (inst.nombre || "Instalación sin proyecto");
  };

  const conflicto = instalacion.vehiculoId && instalacion.fechaMontaje
    ? otrasInstalaciones.find((i) => i.vehiculoId === instalacion.vehiculoId && i.fechaMontaje === instalacion.fechaMontaje)
    : null;

  const guardarFechaMontaje = () => onUpdate({ fechaMontaje: fechaMontajeInput });

  // ---- Fecha de instalación propuesta: fabricación (manual) + stock cubierto + primer hueco libre del vehículo ----
  const [fechaFabricacionInput, setFechaFabricacionInput] = useState(instalacion.fechaFabricacionEstimada || "");
  const [avisandoTarea, setAvisandoTarea] = useState(false);
  const [usuarioAvisoId, setUsuarioAvisoId] = useState("");
  const guardarFechaFabricacion = () => onUpdate({ fechaFabricacionEstimada: fechaFabricacionInput });

  const despieceConStock = proyecto ? compararDespieceConStock(calcularDespieceConjuntoProyecto(proyecto.techos), materiales) : [];
  const hayTechos = proyecto && (proyecto.techos || []).length > 0;
  const stockCubierto = hayTechos && despieceConStock.every((d) => d.falta <= 0.01);

  const primerHuecoLibre = (desde, vehiculoId) => {
    if (!vehiculoId || !desde) return null;
    let fecha = new Date(desde + "T00:00:00");
    for (let i = 0; i < 90; i++) {
      const iso = fecha.toISOString().slice(0, 10);
      const ocupado = otrasInstalaciones.some((i2) => i2.vehiculoId === vehiculoId && i2.fechaMontaje === iso);
      if (!ocupado) return iso;
      fecha.setDate(fecha.getDate() + 1);
    }
    return null;
  };
  const fechaPropuesta = stockCubierto ? primerHuecoLibre(instalacion.fechaFabricacionEstimada, instalacion.vehiculoId) : null;

  const avisarFabricacionTerminada = () => {
    if (!usuarioAvisoId) { alert("Elige a quién avisar."); return; }
    const usuario = usuarios.find((u) => u.id === usuarioAvisoId);
    onCrearTarea({
      titulo: `Fábrica terminada: ${proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : "instalación"}`,
      descripcion: `La fabricación de esta obra ya está lista. ${instalacion.fechaMontaje ? `Día de montaje: ${instalacion.fechaMontaje}.` : "Falta fijar día de montaje."}`,
      asignadoA: usuarioAvisoId,
      asignadoANombre: usuario ? `${usuario.nombre} ${usuario.apellidos || ""}`.trim() : "",
      requiereConfirmacion: true,
    });
    setAvisandoTarea(false);
    setUsuarioAvisoId("");
    alert("Aviso enviado a Tareas.");
  };

  const [presupuestoInput, setPresupuestoInput] = useState(instalacion.presupuestoInstalacion || "");
  const [costeHoraInput, setCosteHoraInput] = useState(instalacion.costeHora || "");
  const [hForm, setHForm] = useState({ fecha: new Date().toISOString().slice(0, 10), instalador: "", horas: "" });
  const [gForm, setGForm] = useState({ fecha: new Date().toISOString().slice(0, 10), concepto: "Dietas", importe: "" });

  const totalHoras = (instalacion.registroHoras || []).reduce((s, r) => s + (parseFloat(r.horas) || 0), 0);
  const costeManoObra = totalHoras * (parseFloat(instalacion.costeHora) || 0);
  const totalGastos = (instalacion.gastos || []).reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
  const costeTotal = costeManoObra + totalGastos;
  const presupuesto = parseFloat(instalacion.presupuestoInstalacion) || 0;
  const diferencia = presupuesto - costeTotal;

  const guardarPresupuesto = () => onUpdate({ presupuestoInstalacion: parseFloat(presupuestoInput) || 0 });
  const guardarCosteHora = () => onUpdate({ costeHora: parseFloat(costeHoraInput) || 0 });

  const [errorHora, setErrorHora] = useState("");
  const submitHora = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!hForm.instalador.trim()) { setErrorHora("Falta el nombre del instalador."); return; }
    if (!(parseFloat(hForm.horas) > 0)) { setErrorHora("Pon un número de horas mayor que 0."); return; }
    setErrorHora("");
    onAddHora({ ...hForm, horas: parseFloat(hForm.horas) || 0 });
    setHForm({ fecha: new Date().toISOString().slice(0, 10), instalador: "", horas: "" });
  };

  const [errorGasto, setErrorGasto] = useState("");
  const submitGasto = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!gForm.concepto.trim()) { setErrorGasto("Falta el concepto del gasto."); return; }
    if (!(parseFloat(gForm.importe) > 0)) { setErrorGasto("Pon un importe mayor que 0."); return; }
    setErrorGasto("");
    onAddGasto({ ...gForm, importe: parseFloat(gForm.importe) || 0 });
    setGForm({ fecha: new Date().toISOString().slice(0, 10), concepto: "Dietas", importe: "" });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>

      <div className="mb-6">
        <div className="flex items-center gap-2 mb-1">
          <h1 className="font-display text-2xl font-extrabold text-slate-900">{proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : (instalacion.nombre || "Instalación sin proyecto")}</h1>
          {!proyecto && <Badge className="bg-violet-50 text-violet-700 ring-violet-200">Sin proyecto</Badge>}
        </div>
        <p className="text-sm text-slate-500">{cliente?.nombre || instalacion.clienteNombre || "—"}</p>
      </div>

      <div className="flex gap-2 mb-6">
        {ESTADO_INSTALACION.map((e) => (
          <button key={e} onClick={() => onUpdate({ estado: e })}
            className={`px-3 py-2 rounded-md text-sm font-semibold transition ${instalacion.estado === e ? ESTADO_INSTALACION_STYLE[e] + " ring-1" : "bg-white border border-slate-300 text-slate-500 hover:bg-slate-50"}`}>
            {e}
          </button>
        ))}
      </div>

      <div className="grid grid-cols-2 gap-4 mb-6">
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-1">Presupuesto de instalación (€)</label>
          <div className="flex gap-2">
            <TextInput type="number" step="0.01" value={presupuestoInput} onChange={(e) => setPresupuestoInput(e.target.value)} />
            <button onClick={guardarPresupuesto} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold hover:opacity-90 px-3 rounded-md">Guardar</button>
          </div>
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-1">Coste por hora de instalador (€/h, opcional)</label>
          <div className="flex gap-2">
            <TextInput type="number" step="0.01" value={costeHoraInput} onChange={(e) => setCosteHoraInput(e.target.value)} />
            <button onClick={guardarCosteHora} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold hover:opacity-90 px-3 rounded-md">Guardar</button>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-2 gap-4 mb-2">
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-1">Día de montaje</label>
          <div className="flex gap-2">
            <TextInput type="date" value={fechaMontajeInput} onChange={(e) => setFechaMontajeInput(e.target.value)} />
            <button onClick={guardarFechaMontaje} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold hover:opacity-90 px-3 rounded-md">Guardar</button>
          </div>
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-1">Vehículo asignado</label>
          <Select value={instalacion.vehiculoId || ""} onChange={(e) => onUpdate({ vehiculoId: e.target.value || null })}>
            <option value="">Sin asignar</option>
            {vehiculos.map((v) => <option key={v.id} value={v.id}>{v.nombre} ({v.tipo})</option>)}
          </Select>
        </div>
      </div>

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-6">
        <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-2">Instaladores asignados a esta obra</label>
        {usuarios.length === 0 ? (
          <p className="text-sm text-slate-400">No hay usuarios dados de alta todavía.</p>
        ) : (
          <div className="flex flex-wrap gap-2">
            {usuarios.map((u) => {
              const asignado = (instalacion.instaladoresAsignados || []).includes(u.id);
              return (
                <button
                  key={u.id}
                  type="button"
                  onClick={() => {
                    const actuales = instalacion.instaladoresAsignados || [];
                    const next = asignado ? actuales.filter((id) => id !== u.id) : [...actuales, u.id];
                    onUpdate({ instaladoresAsignados: next });
                  }}
                  className={`text-xs font-semibold px-3 py-1.5 rounded-full border ${asignado ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-300 text-slate-500 hover:bg-slate-50"}`}
                >
                  {u.nombre} {u.apellidos || ""}
                </button>
              );
            })}
          </div>
        )}
        <p className="text-xs text-slate-400 mt-2">Si no eres administrador, solo ves las obras donde estás marcado aquí.</p>
      </div>

      {conflicto && (
        <div className="mb-6 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ Ese vehículo ya está asignado ese día a: {nombreOtraInstalacion(conflicto)}. Se solapan.
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-6">
        <h3 className="font-display font-bold text-slate-800 mb-3">Fecha de instalación</h3>

        {!hayTechos ? (
          <p className="text-sm text-slate-400">Esta obra no tiene techos calculados desde Mediciones, así que no se puede proponer fecha automáticamente. Pon el día de montaje a mano arriba.</p>
        ) : (
          <>
            <div className="mb-3">
              <label className="text-[11px] uppercase tracking-wide text-slate-400 block mb-1">Fecha estimada de fin de fabricación</label>
              <div className="flex gap-2 max-w-xs">
                <TextInput type="date" value={fechaFabricacionInput} onChange={(e) => setFechaFabricacionInput(e.target.value)} />
                <button onClick={guardarFechaFabricacion} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold hover:opacity-90 px-3 rounded-md">Guardar</button>
              </div>
            </div>

            <div className={`text-sm px-3 py-2 rounded-md mb-3 ${stockCubierto ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
              {stockCubierto ? "✓ El despiece de esta obra está cubierto en Stock." : "El despiece todavía tiene material pendiente (revisa la pestaña \"Despiece de techos\" en el Proyecto) — no se propone fecha hasta que esté cubierto."}
            </div>

            {stockCubierto && !instalacion.fechaFabricacionEstimada && (
              <p className="text-sm text-slate-400">Pon la fecha estimada de fin de fabricación arriba para que se pueda proponer día de montaje.</p>
            )}

            {stockCubierto && instalacion.fechaFabricacionEstimada && !instalacion.vehiculoId && (
              <p className="text-sm text-slate-400">Asigna un vehículo abajo para poder buscar el primer hueco libre.</p>
            )}

            {stockCubierto && instalacion.fechaFabricacionEstimada && instalacion.vehiculoId && (
              fechaPropuesta ? (
                <div className="flex items-center gap-3">
                  <p className="text-sm text-slate-700">Fecha propuesta: <span className="font-semibold">{fmtDate(fechaPropuesta)}</span> (primer hueco libre de ese vehículo desde fin de fabricación)</p>
                  <button
                    onClick={() => { onUpdate({ fechaMontaje: fechaPropuesta }); setFechaMontajeInput(fechaPropuesta); }}
                    className="text-xs font-semibold text-emerald-700 border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 rounded-md hover:bg-emerald-100"
                  >
                    Usar esta fecha
                  </button>
                </div>
              ) : (
                <p className="text-sm text-slate-400">No se ha encontrado hueco libre en los próximos 90 días para ese vehículo.</p>
              )
            )}
          </>
        )}

        <div className="mt-4 pt-4 border-t border-slate-100">
          {!avisandoTarea ? (
            <button onClick={() => setAvisandoTarea(true)} className="text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
              Fábrica terminada → avisar a instalaciones
            </button>
          ) : (
            <div className="flex items-center gap-2">
              <Select value={usuarioAvisoId} onChange={(e) => setUsuarioAvisoId(e.target.value)} className="max-w-xs">
                <option value="">¿A quién avisamos?</option>
                {usuarios.map((u) => <option key={u.id} value={u.id}>{u.nombre} {u.apellidos || ""}</option>)}
              </Select>
              <button onClick={avisarFabricacionTerminada} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold px-3 py-2 rounded-md">Enviar aviso</button>
              <button onClick={() => setAvisandoTarea(false)} className="text-xs font-semibold text-slate-400 px-2">Cancelar</button>
            </div>
          )}
          <p className="text-xs text-slate-400 mt-1.5">Crea una tarea con confirmación para que quede constancia de que se avisó.</p>
        </div>
      </div>


      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-8">
        <Kpi label="Horas totales" value={`${totalHoras.toFixed(1)} h`} />
        <Kpi label="Coste mano de obra" value={money(costeManoObra)} />
        <Kpi label="Gastos (dietas, etc.)" value={money(totalGastos)} />
        <Kpi label="Diferencia vs presupuesto" value={presupuesto ? money(diferencia) : "—"} tone={presupuesto ? (diferencia < 0 ? "bad" : "good") : "neutral"} />
      </div>

      <h2 className="font-display font-bold text-slate-800 mb-3">Material para la furgoneta ({(instalacion.materialFurgoneta || []).length})</h2>
      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 space-y-2">
        {(instalacion.materialFurgoneta || []).length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-2">Todavía no hay materiales apuntados para esta instalación.</p>
        ) : (
          (instalacion.materialFurgoneta || []).map((m) => {
            const estilo = m.estado === "En la obra" ? "bg-emerald-50 text-emerald-700 ring-emerald-200" : m.estado === "En la furgoneta" ? "bg-sky-50 text-sky-700 ring-sky-200" : "bg-amber-50 text-amber-700 ring-amber-200";
            return (
              <div key={m.id} className="flex items-center justify-between gap-2 border-b border-slate-100 last:border-0 pb-2 last:pb-0">
                <span className="text-sm text-slate-700">{m.nombre}</span>
                <div className="flex items-center gap-2">
                  <button onClick={() => onCicloMaterialFurgoneta(m.id)} className={`text-xs font-semibold px-2.5 py-1 rounded-md ring-1 ${estilo}`} title="Toca para cambiar de estado">
                    {m.estado}
                  </button>
                  <button onClick={() => onDeleteMaterialFurgoneta(m.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
                </div>
              </div>
            );
          })
        )}
      </div>
      <form onSubmit={agregarMaterial} className="flex gap-2 mb-2">
        <TextInput value={nuevoMaterial} onChange={(e) => setNuevoMaterial(e.target.value)} placeholder="Ej: Caja de tiradores, escalera, silicona..." className="flex-1" />
        <button type="submit" onClick={agregarMaterial} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1 text-sm font-semibold px-4 py-2 rounded-md"><Plus size={15} /> Añadir</button>
      </form>
      {errorMaterial && (
        <p className="text-xs text-rose-600 font-semibold -mt-1 mb-2">⚠ {errorMaterial}</p>
      )}
      <p className="text-xs text-slate-400 mb-8">Toca la etiqueta de estado para ir pasando de Pendiente de cargar → En la furgoneta → En la obra.</p>

      <h2 className="font-display font-bold text-slate-800 mb-3">Parte diario de horas ({(instalacion.registroHoras || []).length})</h2>
      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 space-y-2">
        {(instalacion.registroHoras || []).length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-2">Todavía no hay horas registradas.</p>
        ) : (
          (instalacion.registroHoras || []).map((r) => (
            <div key={r.id} className="flex items-center justify-between border-b border-slate-100 last:border-0 pb-2 last:pb-0 text-sm">
              <span>{fmtDate(r.fecha)} · <span className="font-medium text-slate-700">{r.instalador}</span> · {r.horas}h</span>
              <button onClick={() => onDeleteHora(r.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
            </div>
          ))
        )}
      </div>
      <form onSubmit={submitHora} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-4 gap-2 items-end mb-2">
        <Field label="Fecha"><TextInput type="date" value={hForm.fecha} onChange={(e) => setHForm({ ...hForm, fecha: e.target.value })} /></Field>
        <Field label="Instalador"><TextInput value={hForm.instalador} onChange={(e) => setHForm({ ...hForm, instalador: e.target.value })} placeholder="Nombre" /></Field>
        <Field label="Horas"><TextInput type="number" step="0.5" value={hForm.horas} onChange={(e) => setHForm({ ...hForm, horas: e.target.value })} /></Field>
        <button type="submit" onClick={submitHora} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Añadir</button>
      </form>
      {errorHora && <p className="text-xs text-rose-600 font-semibold mb-8">⚠ {errorHora}</p>}
      {!errorHora && <div className="mb-8" />}

      <h2 className="font-display font-bold text-slate-800 mb-3">Gastos (dietas, desplazamientos...) ({(instalacion.gastos || []).length})</h2>
      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 space-y-2">
        {(instalacion.gastos || []).length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-2">Todavía no hay gastos registrados.</p>
        ) : (
          (instalacion.gastos || []).map((g) => (
            <div key={g.id} className="flex items-center justify-between border-b border-slate-100 last:border-0 pb-2 last:pb-0 text-sm">
              <span>{fmtDate(g.fecha)} · <span className="font-medium text-slate-700">{g.concepto}</span> · {money(g.importe)}</span>
              <button onClick={() => onDeleteGasto(g.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button>
            </div>
          ))
        )}
      </div>
      <form onSubmit={submitGasto} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-4 gap-2 items-end">
        <Field label="Fecha"><TextInput type="date" value={gForm.fecha} onChange={(e) => setGForm({ ...gForm, fecha: e.target.value })} /></Field>
        <Field label="Concepto"><TextInput value={gForm.concepto} onChange={(e) => setGForm({ ...gForm, concepto: e.target.value })} /></Field>
        <Field label="Importe (€)"><TextInput type="number" step="0.01" value={gForm.importe} onChange={(e) => setGForm({ ...gForm, importe: e.target.value })} /></Field>
        <button type="submit" onClick={submitGasto} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Añadir</button>
      </form>
      {errorGasto && <p className="text-xs text-rose-600 font-semibold mt-2">⚠ {errorGasto}</p>}

      <div className="mt-10 pt-8 border-t border-slate-200">
        <ControlMontajeVivienda
          instalacionId={instalacion.id}
          proyectoId={instalacion.proyectoId}
          incidencias={incidencias}
          onUpsertIncidencia={onUpsertIncidencia}
          instalacion={instalacion}
          proyecto={proyecto}
          cliente={cliente}
        />
      </div>
    </div>
  );
}

/* ================= CONTROL DE MONTAJE POR VIVIENDA ================= */

// Convierte un texto tipo "CAR BL1 PB A0-1" en piezas de vivienda, y a partir
// de ahí lee las filas de elementos (puertas/ventanas) que le siguen.
function parsearExcelMontaje(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const viviendas = [];
  const hojasBloque = workbook.SheetNames.filter((n) => n.toUpperCase().startsWith("BLOQUE"));

  hojasBloque.forEach((nombreHoja) => {
    const hoja = workbook.Sheets[nombreHoja];
    const filas = XLSX.utils.sheet_to_json(hoja, { header: 1, defval: null });
    let viviendaActual = null;

    const cerrarVivienda = () => {
      if (viviendaActual && viviendaActual.elementos.length > 0) viviendas.push(viviendaActual);
      viviendaActual = null;
    };

    filas.forEach((fila) => {
      // Buscamos la primera celda con contenido en la fila, sea cual sea su columna
      // (algunos lectores de ODS no dejan la columna A vacía como en el Excel original)
      let colEtiqueta = -1;
      for (let c = 0; c < fila.length; c++) {
        if (fila[c] != null && String(fila[c]).trim() !== "") { colEtiqueta = c; break; }
      }
      if (colEtiqueta === -1) return;
      const etiqueta = String(fila[colEtiqueta]).trim();
      const up = etiqueta.toUpperCase();

      if (up.startsWith("BLOQUE")) return;
      if (up.startsWith("TOTAL")) { cerrarVivienda(); return; }

      if (up.startsWith("CAR ")) {
        cerrarVivienda();
        const resto = etiqueta.substring(4).trim();
        const partes = resto.split(/\s+/);
        const bloqueCode = partes[0] || "";
        const planta = partes[1] || "";
        const codigoVivienda = partes.slice(2).join(" ") || "";
        const esZonaComun = /HUECO|ESCALERA|ZONA/i.test(codigoVivienda);
        const codigoLimpio = codigoVivienda.replace(/[()]/g, "").trim();
        viviendaActual = {
          id: uid(),
          claveImportacion: `${bloqueCode}_${planta}_${codigoLimpio}`.replace(/\s+/g, "_"),
          bloque: bloqueCode,
          planta,
          codigoVivienda: codigoLimpio,
          esZonaComun,
          elementos: [],
          extras: [],
          fotos: [],
          documentos: [],
          incidenciasVinculadas: [],
        };
        return;
      }

      if (viviendaActual) {
        const partesCodigo = etiqueta.split(/\s+/);
        const codigoBase = partesCodigo[0] || etiqueta;
        const sufijo = partesCodigo[1] || null;
        const esVentana = codigoBase.toUpperCase().startsWith("V");
        const esPuerta = codigoBase.toUpperCase().startsWith("P");
        viviendaActual.elementos.push({
          id: uid(),
          codigo: codigoBase,
          tipo: esPuerta ? "puerta" : esVentana ? "ventana" : "otro",
          ladoApertura: esPuerta ? sufijo : null,
          tipoVentana: esVentana ? sufijo : null,
          medida: fila[colEtiqueta + 1] != null ? String(fila[colEtiqueta + 1]).trim() : null,
          cantidad: fila[colEtiqueta + 3] != null ? Number(fila[colEtiqueta + 3]) : 1,
          instalado: false,
          tapajuntas: { izquierda: false, derecha: false, arriba: false },
        });
      }
    });

    cerrarVivienda();
  });

  return viviendas;
}

// Lista de componentes que se controlan por defecto en cada puerta/ventana.
// El usuario puede añadir más componentes personalizados con el botón "+".
const COMPONENTES_DEFECTO_ELEMENTO = ["Marco", "Hoja izquierda", "Hoja derecha", "Persiana", "Mosquitera", "Cajón de obra", "Montaje", "Tapajuntas", "Postigo", "Silicona"];

// Si el elemento todavía no tiene su propia lista de componentes guardada
// (por ejemplo, viene de una importación antigua), se genera la lista por
// defecto al vuelo. No se guarda hasta que el usuario marca/añade algo.
function componentesDeElemento(el) {
  if (el.componentes && el.componentes.length) return el.componentes;
  return COMPONENTES_DEFECTO_ELEMENTO.map((nombre) => ({ id: uid(), nombre, hecho: false }));
}

// Categorías para clasificar los documentos guardados por vivienda (planos, etc.)
const CATEGORIAS_DOCUMENTO_MONTAJE = ["Planos", "Contratos", "Fichas técnicas", "Fotos de obra", "Otros"];

function agruparDocumentosPorCategoria(documentos) {
  const grupos = {};
  (documentos || []).forEach((d) => {
    const cat = d.categoria || "Otros";
    if (!grupos[cat]) grupos[cat] = [];
    grupos[cat].push(d);
  });
  return grupos;
}

function nombreElementoMontaje(el) {
  if (el.nombre) return el.nombre;
  const base = el.tipo === "puerta" ? "Puerta" : el.tipo === "ventana" ? "Ventana" : el.codigo;
  const sufijo =
    el.tipo === "puerta" && el.ladoApertura
      ? ` (apertura ${el.ladoApertura === "I" ? "izquierda" : "derecha"})`
      : el.tipo === "ventana" && el.tipoVentana
      ? ` (${el.tipoVentana})`
      : "";
  return `${base} ${el.codigo}${sufijo}`;
}

function comprimirFotoMontaje(file, maxAncho = 1200, calidad = 0.7) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const escala = Math.min(1, maxAncho / img.width);
        const canvas = document.createElement("canvas");
        canvas.width = img.width * escala;
        canvas.height = img.height * escala;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", calidad));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function leerDocumentoMontaje(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// Guarda todo bajo la clave "controlMontaje/<instalacionId>" en Firebase.
// No usa las funciones persist()/uid() de tiempo real del resto de la app
// porque este control vive solo dentro de la ficha de instalación — pero
// reutiliza fbDb, uid y toArray que ya están definidos arriba en el archivo.
// Para las incidencias sí usa las funciones reales de la app (onUpsertIncidencia,
// y el array "incidencias" ya cargado), para no crear un sistema paralelo.
// Componente genérico: la lista de "viviendas" con su checklist de elementos,
// fotos, documentos por categoría, extras e incidencias. Lo usa tanto
// "Control de montaje" (dentro de cada instalación, basePath=controlMontaje/<id>)
// como "Mediciones" (basePath=mediciones/<id>/detalle), para que ambos
// tengan exactamente las mismas funciones.
// Envoltorio usado dentro de cada instalación: calcula la ruta de Firebase a
// partir del id de la instalación y delega en el componente genérico de
// arriba. Esto es exactamente el mismo comportamiento que tenía antes.
function ControlMontajeVivienda({ instalacionId, proyectoId, incidencias, onUpsertIncidencia, instalacion, proyecto, cliente }) {
  return (
    <ControlElementosPorVivienda
      basePath={`controlMontaje/${instalacionId}`}
      proyectoId={proyectoId}
      incidencias={incidencias}
      onUpsertIncidencia={onUpsertIncidencia}
      instalacion={instalacion}
      proyecto={proyecto}
      cliente={cliente}
    />
  );
}

// Escapa texto antes de insertarlo en el HTML del informe imprimible, para
// que direcciones/nombres con caracteres especiales no rompan el documento.
function escaparHtmlInforme(texto) {
  return String(texto ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
  }[c]));
}

// Genera el HTML del "Informe del día de montaje" para una vivienda concreta
// (datos de obra/contacto, plano/croquis si hay uno subido, y los materiales
// —elementos— que hay que llevar ese día) y lo abre en una pestaña nueva con
// un botón de imprimir. No incluye el checklist de instalado/pendiente: eso
// se sigue gestionando dentro del propio Control de Montaje.
function abrirInformeDiaMontaje(v, { proyecto, cliente, instalacion, jefeDeObra }) {
  const e = escaparHtmlInforme;
  const direccionObra = [cliente?.direccion, cliente?.pueblo, cliente?.provincia === "Otra ciudad..." ? cliente?.provinciaManual : cliente?.provincia, cliente?.cp]
    .filter(Boolean).join(", ");
  const nombreObra = proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : (instalacion?.nombre || "Instalación sin proyecto");
  const planos = (v.documentos || []).filter((d) => d.categoria === "Planos");
  const elementos = [...(v.elementos || [])].sort((a, b) => (a.codigo || "").localeCompare(b.codigo || ""));

  const filasMateriales = elementos.length
    ? elementos.map((el) => `<tr><td>${e(nombreElementoMontaje(el))}</td><td>${e(el.medida || "—")}</td></tr>`).join("")
    : `<tr><td colspan="2" style="color:#94a3b8;">Sin elementos registrados en esta vivienda.</td></tr>`;

  const bloquePlanos = planos.length
    ? planos.map((p) => {
        const esImagen = /\.(png|jpe?g|gif|webp)$/i.test(p.nombre || "") || (p.url || "").startsWith("data:image");
        return esImagen
          ? `<div class="plano"><img src="${p.url}" alt="${e(p.nombre)}" /><p class="pie">${e(p.nombre)}</p></div>`
          : `<div class="plano"><a href="${p.url}" target="_blank">${e(p.nombre)} (abrir documento)</a></div>`;
      }).join("")
    : `<p style="color:#94a3b8;">No hay ningún plano/croquis subido en la categoría "Planos" de esta vivienda.</p>`;

  const html = `<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8" />
<title>Informe de montaje — ${e(v.codigoVivienda || v.bloque)}</title>
<style>
  body { font-family: -apple-system, Arial, sans-serif; color: #1e293b; margin: 24px; max-width: 800px; }
  h1 { font-size: 18px; margin-bottom: 2px; }
  h2 { font-size: 14px; margin: 20px 0 8px; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px; }
  .sub { color: #64748b; font-size: 13px; margin-bottom: 16px; }
  table { width: 100%; border-collapse: collapse; font-size: 13px; }
  td { padding: 6px 4px; border-bottom: 1px solid #f1f5f9; }
  .plano img { max-width: 100%; border: 1px solid #e2e8f0; border-radius: 6px; }
  .plano .pie { color: #94a3b8; font-size: 12px; margin-top: 2px; }
  .btn-print { background: #2E8B57; color: #fff; border: none; padding: 8px 16px; border-radius: 6px; font-weight: 600; cursor: pointer; margin-bottom: 16px; }
  @media print { .btn-print { display: none; } }
</style></head>
<body>
  <button class="btn-print" onclick="window.print()">Imprimir</button>
  <h1>Informe del día de montaje</h1>
  <p class="sub">${e(nombreObra)} · ${v.bloque || ""} ${v.planta || ""} ${v.codigoVivienda || ""}</p>

  <h2>Datos de contacto / obra</h2>
  <table>
    <tr><td style="width:140px;color:#64748b;">Obra</td><td>${e(nombreObra)}</td></tr>
    <tr><td style="color:#64748b;">Cliente</td><td>${e(cliente?.nombre || instalacion?.clienteNombre || "—")}</td></tr>
    <tr><td style="color:#64748b;">Dirección</td><td>${e(direccionObra || "—")}</td></tr>
    <tr><td style="color:#64748b;">Jefe de obra</td><td>${e(jefeDeObra?.email || "—")} ${jefeDeObra?.telefono ? "· " + e(jefeDeObra.telefono) : ""}</td></tr>
  </table>

  <h2>Plano / croquis de la vivienda</h2>
  ${bloquePlanos}

  <h2>Materiales que se llevan ese día</h2>
  <table>
    <tr><td style="color:#64748b;font-weight:600;">Elemento</td><td style="color:#64748b;font-weight:600;">Medida</td></tr>
    ${filasMateriales}
  </table>
</body></html>`;

  const ventana = window.open("", "_blank");
  if (!ventana) {
    alert("El navegador ha bloqueado la ventana emergente. Permite las ventanas emergentes para este sitio e inténtalo de nuevo.");
    return;
  }
  ventana.document.write(html);
  ventana.document.close();
}

function ControlElementosPorVivienda({ basePath, proyectoId, incidencias, onUpsertIncidencia, accionPrincipal, instalacion, proyecto, cliente }) {
  const [viviendas, setViviendas] = useState([]);
  const [jefeDeObra, setJefeDeObra] = useState(null);
  const [cargando, setCargando] = useState(true);
  const [importando, setImportando] = useState(false);
  const [expandida, setExpandida] = useState(null);
  const [filtroBloque, setFiltroBloque] = useState("todos");
  const [extraTexto, setExtraTexto] = useState({});
  const [editandoJefe, setEditandoJefe] = useState(false);
  const [jefeEmail, setJefeEmail] = useState("");
  const [jefeTelefono, setJefeTelefono] = useState("");
  const [nuevaIncidenciaTexto, setNuevaIncidenciaTexto] = useState({});
  const [vinculandoEn, setVinculandoEn] = useState(null);
  const [mostrarManual, setMostrarManual] = useState(false);
  const [anadiendoComponente, setAnadiendoComponente] = useState(null);
  const [nuevoComponenteTexto, setNuevoComponenteTexto] = useState({});
  const [docCategoria, setDocCategoria] = useState({});
  const [nuevaViviendaTexto, setNuevaViviendaTexto] = useState("");
  const [nuevoElementoNombre, setNuevoElementoNombre] = useState({});
  const [nuevoElementoMedida, setNuevoElementoMedida] = useState({});
  const fileInputRef = useRef(null);

  const descargarPlantilla = () => {
    const datosLeeme = [
      ["CÓMO RELLENAR ESTE EXCEL PARA EL CONTROL DE MONTAJE POR VIVIENDA"],
      [""],
      ["Esta hoja LEEME es solo para ti — el CRM no la lee. Rellena las hojas BLOQUE_1, BLOQUE_2, etc."],
      [""],
      ["1. Cada hoja que quieras que el CRM lea debe empezar su nombre por BLOQUE (BLOQUE_1, BLOQUE_2, BLOQUE_3...)."],
      ["2. Cada vivienda empieza con una fila en la COLUMNA B con este formato exacto: CAR <bloque> <planta> <vivienda>"],
      ["   Ejemplo: CAR BL1 PB A0-1        Zona común: CAR BL1 PB (HUECO ESCALERA)"],
      ["   Esta fila es IMPRESCINDIBLE — así el CRM sabe que empieza una vivienda nueva."],
      ["3. Debajo de cada CAR, una fila por cada puerta o ventana de esa vivienda:"],
      ["   Columna B = código. IMPRESCINDIBLE. Empieza por P (puerta) o V (ventana), ej: P01 I / V01 CA"],
      ["   Columna E = cantidad. IMPRESCINDIBLE (si se deja vacío, el CRM pone 1)."],
      ["   Columnas C y D y F (medida, valor, total) son opcionales, solo para tu referencia."],
      ["4. Deja una fila completamente en blanco entre una vivienda y la siguiente."],
      ["5. Puedes cerrar cada bloque con una fila que empiece por TOTAL (ej: TOTAL BLOQUE 1). Es opcional."],
      [""],
      ["En la hoja BLOQUE_1 tienes un ejemplo ya relleno — cópialo y sustituye por los datos de tu obra."],
    ];
    const datosBloque1 = [
      ["", "Código / vivienda", "Medida (opcional)", "Valor (opcional)", "Cantidad (OBLIGATORIO)", "Total (opcional)"],
      ["", "BLOQUE 1"],
      ["", "CAR BL1 PB A0-1"],
      ["", "P01 I", "0,90 x 2,20", 108, 1, 108],
      ["", "V01 CA", "2,60 x 2,20", 204, 2, 408],
      ["", "V02 CA", "1,80 x 1,25", 108, 1, 108],
      [],
      ["", "CAR BL1 PB B0-2"],
      ["", "P01 D", "0,90 x 2,20", 108, 1, 108],
      ["", "V02 CA", "1,80 x 1,25", 108, 1, 108],
      [],
      ["", "CAR BL1 PB (HUECO ESCALERA)"],
      ["", "V07", "0,80 x 1,25", 84, 1, 84],
      [],
      ["", "TOTAL BLOQUE 1"],
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(datosLeeme), "LEEME");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(datosBloque1), "BLOQUE_1");
    XLSX.writeFile(wb, "Plantilla_Control_Montaje_Vivienda.xlsx");
  };

  useEffect(() => {
    if (!basePath) return;
    (async () => {
      setCargando(true);
      try {
        const snap = await fbGet(ref(fbDb, basePath));
        const datos = snap.val() || {};
        setViviendas(toArray(datos.viviendas));
        setJefeDeObra(datos.jefeDeObra || null);
        setJefeEmail(datos.jefeDeObra?.email || "");
        setJefeTelefono(datos.jefeDeObra?.telefono || "");
      } catch (e) {
        console.error(e);
      } finally {
        setCargando(false);
      }
    })();
  }, [basePath]);

  const guardarViviendas = async (next) => {
    setViviendas(next);
    try {
      await fbSet(ref(fbDb, `${basePath}/viviendas`), next);
    } catch (e) {
      console.error(e);
    }
  };

  const guardarJefeDeObra = async () => {
    const datos = { email: jefeEmail.trim(), telefono: jefeTelefono.trim() };
    setJefeDeObra(datos);
    setEditandoJefe(false);
    try {
      await fbSet(ref(fbDb, `${basePath}/jefeDeObra`), datos);
    } catch (e) {
      console.error(e);
    }
  };

  const listaOrdenada = [...viviendas].sort((a, b) => {
    if (a.bloque !== b.bloque) return String(a.bloque).localeCompare(String(b.bloque));
    if (a.planta !== b.planta) return String(a.planta).localeCompare(String(b.planta));
    return String(a.codigoVivienda).localeCompare(String(b.codigoVivienda));
  });
  const bloquesDisponibles = [...new Set(listaOrdenada.map((v) => v.bloque))].sort();
  const listaFiltrada = filtroBloque === "todos" ? listaOrdenada : listaOrdenada.filter((v) => v.bloque === filtroBloque);

  const handleImportar = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setImportando(true);
    try {
      const arrayBuffer = await file.arrayBuffer();
      const nuevas = parsearExcelMontaje(arrayBuffer);
      const clavesExistentes = new Set(viviendas.map((v) => v.claveImportacion));
      const aAgregar = nuevas.filter((v) => !clavesExistentes.has(v.claveImportacion));
      // si la vivienda ya existía, añadimos solo los elementos nuevos sin tocar el estado marcado
      const actualizadas = viviendas.map((v) => {
        const importada = nuevas.find((n) => n.claveImportacion === v.claveImportacion);
        if (!importada) return v;
        const codigosExistentes = new Set(v.elementos.map((el) => el.codigo + (el.ladoApertura || "") + (el.tipoVentana || "")));
        const nuevosElementos = importada.elementos.filter(
          (el) => !codigosExistentes.has(el.codigo + (el.ladoApertura || "") + (el.tipoVentana || ""))
        );
        return nuevosElementos.length ? { ...v, elementos: [...v.elementos, ...nuevosElementos] } : v;
      });
      await guardarViviendas([...actualizadas, ...aAgregar]);
      alert(`Importación completa: ${nuevas.length} viviendas leídas del archivo.`);
    } catch (err) {
      console.error(err);
      alert("No se pudo leer el archivo. Comprueba que es el Excel/ODS de cálculo de montaje.");
    } finally {
      setImportando(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  };

  const actualizarVivienda = (viviendaId, cambios) => {
    guardarViviendas(viviendas.map((v) => (v.id === viviendaId ? { ...v, ...cambios } : v)));
  };

  const agregarViviendaAMano = (nombre) => {
    const limpio = (nombre || "").trim();
    if (!limpio) return;
    const nueva = { id: uid(), bloque: limpio, planta: "", puerta: "", elementos: [], extras: [], fotos: [], documentos: [], incidenciasVinculadas: [] };
    guardarViviendas([...viviendas, nueva]);
  };

  const agregarElementoAMano = (v, nombre, medida) => {
    const limpio = (nombre || "").trim();
    if (!limpio) return;
    const nuevo = { id: uid(), nombre: limpio, medida: (medida || "").trim(), instalado: false, componentes: [], fotos: [] };
    actualizarVivienda(v.id, { elementos: [...v.elementos, nuevo] });
  };

  const toggleInstalado = (v, elId) => {
    actualizarVivienda(v.id, {
      elementos: v.elementos.map((el) => (el.id === elId ? { ...el, instalado: !el.instalado, instaladoEn: !el.instalado ? Date.now() : null } : el)),
    });
  };

  const toggleTapajuntas = (v, elId, lado) => {
    actualizarVivienda(v.id, {
      elementos: v.elementos.map((el) =>
        el.id === elId ? { ...el, tapajuntas: { ...el.tapajuntas, [lado]: !el.tapajuntas[lado] } } : el
      ),
    });
  };

  const toggleComponente = (v, elId, compId) => {
    actualizarVivienda(v.id, {
      elementos: v.elementos.map((el) => {
        if (el.id !== elId) return el;
        const lista = componentesDeElemento(el);
        return { ...el, componentes: lista.map((c) => (c.id === compId ? { ...c, hecho: !c.hecho } : c)) };
      }),
    });
  };

  const agregarComponente = (v, elId, nombre) => {
    const limpio = (nombre || "").trim();
    if (!limpio) return;
    actualizarVivienda(v.id, {
      elementos: v.elementos.map((el) => {
        if (el.id !== elId) return el;
        const lista = componentesDeElemento(el);
        return { ...el, componentes: [...lista, { id: uid(), nombre: limpio, hecho: false }] };
      }),
    });
  };

  const subirFotoElemento = async (v, elId, file) => {
    if (!file) return;
    const dataUrl = await comprimirFotoMontaje(file);
    actualizarVivienda(v.id, {
      elementos: v.elementos.map((el) =>
        el.id === elId ? { ...el, fotos: [...(el.fotos || []), { id: uid(), url: dataUrl, subidaEn: Date.now() }] } : el
      ),
    });
  };

  const agregarExtra = (v) => {
    const texto = (extraTexto[v.id] || "").trim();
    if (!texto) return;
    actualizarVivienda(v.id, { extras: [...(v.extras || []), { id: uid(), descripcion: texto, creadoEn: Date.now() }] });
    setExtraTexto((prev) => ({ ...prev, [v.id]: "" }));
  };

  const subirFoto = async (v, file) => {
    if (!file) return;
    const dataUrl = await comprimirFotoMontaje(file);
    actualizarVivienda(v.id, { fotos: [...(v.fotos || []), { id: uid(), url: dataUrl, subidaEn: Date.now() }] });
  };

  const subirDocumento = async (v, file, categoria) => {
    if (!file) return;
    if (file.size > 8 * 1024 * 1024) {
      alert("El archivo pesa más de 8 MB. Prueba a comprimirlo o súbelo a Dropbox y enlázalo aparte.");
      return;
    }
    const dataUrl = await leerDocumentoMontaje(file);
    actualizarVivienda(v.id, {
      documentos: [...(v.documentos || []), { id: uid(), nombre: file.name, categoria: categoria || "Otros", url: dataUrl, subidoEn: Date.now() }],
    });
  };

  const crearIncidencia = (v) => {
    const texto = (nuevaIncidenciaTexto[v.id] || "").trim();
    if (!texto) return;
    onUpsertIncidencia({
      id: null,
      proyectoId: proyectoId || "",
      fecha: new Date().toISOString().slice(0, 10),
      especificaciones: texto,
      observaciones: `Vivienda ${v.bloque} ${v.planta} ${v.codigoVivienda} (control de montaje)`,
      responsableInicial: "",
      comercialAsociado: "",
      responsableActual: "",
      estadoTrabajo: "Pendiente revisión",
      estadoIncidencia: "Pendiente revisión",
      fechaEntregaPrevista: "",
      fechaEntregado: "",
    });
    actualizarVivienda(v.id, {
      incidenciasVinculadas: [...(v.incidenciasVinculadas || []), { id: uid(), descripcion: texto, creadaEn: Date.now() }],
    });
    setNuevaIncidenciaTexto((prev) => ({ ...prev, [v.id]: "" }));
  };

  const vincularIncidenciaExistente = (v, incidencia) => {
    actualizarVivienda(v.id, {
      incidenciasVinculadas: [
        ...(v.incidenciasVinculadas || []),
        { id: incidencia.id, numero: incidencia.numero, descripcion: incidencia.especificaciones, vinculadaEn: Date.now() },
      ],
    });
    setVinculandoEn(null);
  };

  const resumenVivienda = (v) => {
    const instalados = v.elementos.filter((el) => el.instalado).length;
    const lineas = v.elementos.map((el) => {
      const comps = componentesDeElemento(el);
      const estado = el.instalado ? "Instalado" : "Pendiente";
      const detalleComp = comps.map((c) => `${c.nombre}:${c.hecho ? "SI" : "NO"}`).join(" | ");
      return `- ${nombreElementoMontaje(el)}: ${estado} (${detalleComp})`;
    });
    const incs = v.incidenciasVinculadas || [];
    const lineasInc = incs.length ? `\nIncidencias vinculadas:\n${incs.map((i) => `- ${i.descripcion}`).join("\n")}` : "";
    return `Control montaje ${v.bloque} ${v.planta} ${v.codigoVivienda} (${instalados}/${v.elementos.length})\n${lineas.join("\n")}${lineasInc}`;
  };

  const avisarWhatsapp = (v) => {
    const texto = encodeURIComponent(resumenVivienda(v));
    const telefono = jefeDeObra?.telefono ? jefeDeObra.telefono.replace(/[^\d+]/g, "") : "";
    window.open(`https://wa.me/${telefono}?text=${texto}`, "_blank");
  };

  const avisarCorreo = (v) => {
    const asunto = encodeURIComponent(`Control montaje ${v.bloque} ${v.codigoVivienda}`);
    const cuerpo = encodeURIComponent(resumenVivienda(v));
    window.location.href = `mailto:${jefeDeObra?.email || ""}?subject=${asunto}&body=${cuerpo}`;
  };

  if (cargando) return <p className="text-sm text-slate-400">Cargando control de montaje...</p>;

  return (
    <div>
      <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
        <h2 className="font-display font-bold text-slate-800">Control de montaje por vivienda</h2>
        <div className="flex gap-2">
          <button onClick={descargarPlantilla} className="px-3 py-2 rounded-md text-sm font-semibold text-slate-600 border border-slate-300 hover:bg-slate-50">
            Descargar plantilla vacía
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.ods" className="hidden" onChange={handleImportar} />
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={importando}
            style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
            className="px-3 py-2 rounded-md text-sm font-semibold"
          >
            {importando ? "Importando..." : "Importar Excel/ODS"}
          </button>
        </div>
      </div>

      <div className="flex gap-2 mb-4">
        <TextInput
          value={nuevaViviendaTexto}
          onChange={(e) => setNuevaViviendaTexto(e.target.value)}
          placeholder="Ej: Vivienda única, Piso 1A, Local..."
          className="max-w-xs"
        />
        <button
          onClick={() => { agregarViviendaAMano(nuevaViviendaTexto); setNuevaViviendaTexto(""); }}
          className="px-3 py-2 rounded-md text-sm font-semibold text-slate-600 border border-slate-300 hover:bg-slate-50 shrink-0"
        >
          + Añadir a mano (sin Excel)
        </button>
      </div>

      <button onClick={() => setMostrarManual((v) => !v)} className="text-xs font-semibold text-emerald-700 mb-4">
        {mostrarManual ? "Ocultar instrucciones ▲" : "¿Cómo relleno el Excel? ▼"}
      </button>
      {mostrarManual && (
        <div className="bg-slate-50 border border-slate-200 rounded-lg p-4 mb-4 text-sm text-slate-600 space-y-1.5">
          <p className="font-semibold text-slate-700">Cómo tiene que estar el Excel para que el CRM lo reconozca:</p>
          <p>1. Cada hoja que quieras que se lea debe empezar su nombre por <b>BLOQUE</b> (BLOQUE_1, BLOQUE_2, BLOQUE_3...).</p>
          <p>2. Cada vivienda empieza con una fila en la columna B: <b>CAR &lt;bloque&gt; &lt;planta&gt; &lt;vivienda&gt;</b>, ej: <span className="font-mono">CAR BL1 PB A0-1</span>. Para zonas comunes: <span className="font-mono">CAR BL1 PB (HUECO ESCALERA)</span>.</p>
          <p>3. Debajo, una fila por cada puerta/ventana: columna B = código (empieza por P o V), columna E = cantidad. Ambas son imprescindibles.</p>
          <p>4. Deja una fila en blanco entre cada vivienda.</p>
          <p>5. Puedes cerrar cada bloque con una fila "TOTAL BLOQUE N" (opcional).</p>
          <p className="text-slate-400">Descarga la plantilla de arriba para ver un ejemplo ya relleno.</p>
        </div>
      )}


      <div className="bg-slate-50 border border-slate-200 rounded-lg p-3 mb-4">
        <div className="flex items-center justify-between">
          <div className="text-sm">
            <span className="font-semibold">Jefe de obra: </span>
            {jefeDeObra?.email || jefeDeObra?.telefono ? (
              <span className="text-slate-600">{jefeDeObra.email} {jefeDeObra.telefono && `· ${jefeDeObra.telefono}`}</span>
            ) : (
              <span className="text-slate-400">sin configurar</span>
            )}
          </div>
          <button onClick={() => setEditandoJefe((v) => !v)} className="text-xs font-semibold text-emerald-700">
            {editandoJefe ? "Cerrar" : "Configurar"}
          </button>
        </div>
        {editandoJefe && (
          <div className="flex flex-col gap-2 mt-2">
            <TextInput type="email" placeholder="Correo del jefe de obra" value={jefeEmail} onChange={(e) => setJefeEmail(e.target.value)} />
            <TextInput type="tel" placeholder="WhatsApp (con prefijo, ej: +34600000000)" value={jefeTelefono} onChange={(e) => setJefeTelefono(e.target.value)} />
            <button onClick={guardarJefeDeObra} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="px-3 py-1.5 rounded-md text-sm font-semibold self-start">
              Guardar
            </button>
          </div>
        )}
      </div>

      {bloquesDisponibles.length > 1 && (
        <div className="flex gap-2 mb-4 overflow-x-auto">
          <button onClick={() => setFiltroBloque("todos")} className={`px-3 py-1 rounded-md text-sm whitespace-nowrap ${filtroBloque === "todos" ? "bg-emerald-600 text-white" : "bg-slate-100 text-slate-600"}`}>Todos</button>
          {bloquesDisponibles.map((b) => (
            <button key={b} onClick={() => setFiltroBloque(b)} className={`px-3 py-1 rounded-md text-sm whitespace-nowrap ${filtroBloque === b ? "bg-emerald-600 text-white" : "bg-slate-100 text-slate-600"}`}>{b}</button>
          ))}
        </div>
      )}

      {accionPrincipal && listaOrdenada.length > 0 && (
        <button
          onClick={() => accionPrincipal.onClick(listaOrdenada.map(resumenVivienda).join("\n\n"), listaOrdenada)}
          style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
          className="w-full flex items-center justify-center gap-2 hover:opacity-90 text-sm font-bold py-3 rounded-lg mb-4 shadow-md"
        >
          {accionPrincipal.etiqueta}
        </button>
      )}

      {listaFiltrada.length === 0 && (
        <p className="text-sm text-slate-400">Todavía no hay viviendas importadas. Pulsa "Importar Excel/ODS" y sube el archivo de cálculo de montaje.</p>
      )}

      <div className="space-y-3">
        {listaFiltrada.map((v) => {
          const elementos = [...v.elementos].sort((a, b) => a.codigo.localeCompare(b.codigo));
          const instalados = elementos.filter((el) => el.instalado).length;
          const abierta = expandida === v.id;
          return (
            <div key={v.id} className="border border-slate-200 rounded-lg overflow-hidden bg-white">
              <button onClick={() => setExpandida(abierta ? null : v.id)} className="w-full flex items-center justify-between px-4 py-3 text-left">
                <div className="font-semibold text-slate-800 text-sm">
                  {v.bloque} · {v.planta} · {v.codigoVivienda}
                  {v.esZonaComun && <span className="ml-2 text-xs bg-slate-100 text-slate-500 px-2 py-0.5 rounded-md">Zona común</span>}
                </div>
                <span className={`text-xs px-2 py-1 rounded-md font-semibold ${instalados === elementos.length && elementos.length > 0 ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
                  {instalados}/{elementos.length} instalado
                </span>
              </button>

              {abierta && (
                <div className="px-4 pb-4 border-t border-slate-100">
                  <button
                    type="button"
                    onClick={() => abrirInformeDiaMontaje(v, { proyecto, cliente, instalacion, jefeDeObra })}
                    className="mt-3 flex items-center gap-1.5 text-xs font-semibold text-emerald-700 border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 rounded-md hover:bg-emerald-100"
                    title="Abre un informe imprimible con datos de la obra, el plano y los materiales del día"
                  >
                    <Printer size={13} /> Informe del día de montaje
                  </button>
                  <div className="space-y-2 mt-3">
                    {elementos.map((el) => {
                      const comps = componentesDeElemento(el);
                      const keyAdd = `${v.id}_${el.id}`;
                      return (
                        <div key={el.id} className="border border-slate-200 rounded-md p-2">
                          <label className="flex items-center gap-2 text-sm font-semibold text-slate-700">
                            <input type="checkbox" checked={!!el.instalado} onChange={() => toggleInstalado(v, el.id)} />
                            {nombreElementoMontaje(el)}
                            {el.medida && <span className="text-xs text-slate-400 font-normal">({el.medida})</span>}
                          </label>

                          <div className="flex flex-wrap gap-3 mt-2 ml-6 text-xs text-slate-600 items-center">
                            {comps.map((c) => (
                              <label key={c.id} className="flex items-center gap-1">
                                <input type="checkbox" checked={!!c.hecho} onChange={() => toggleComponente(v, el.id, c.id)} />
                                {c.nombre}
                              </label>
                            ))}
                            <button
                              type="button"
                              onClick={() => setAnadiendoComponente(anadiendoComponente === keyAdd ? null : keyAdd)}
                              title="Añadir otro elemento a controlar (por ejemplo: cerradura, mosquitera...)"
                              className="w-5 h-5 flex items-center justify-center rounded-full border border-slate-300 text-slate-500 hover:bg-slate-100 font-bold shrink-0"
                            >
                              +
                            </button>
                          </div>

                          {anadiendoComponente === keyAdd && (
                            <div className="flex gap-2 mt-2 ml-6">
                              <TextInput
                                value={nuevoComponenteTexto[keyAdd] || ""}
                                onChange={(e) => setNuevoComponenteTexto((prev) => ({ ...prev, [keyAdd]: e.target.value }))}
                                placeholder="Ej: cerradura, mosquitera..."
                                className="flex-1 !text-xs !py-1"
                              />
                              <button
                                onClick={() => {
                                  agregarComponente(v, el.id, nuevoComponenteTexto[keyAdd]);
                                  setNuevoComponenteTexto((prev) => ({ ...prev, [keyAdd]: "" }));
                                  setAnadiendoComponente(null);
                                }}
                                className="px-2.5 py-1 bg-slate-100 rounded-md text-xs font-semibold shrink-0"
                              >
                                Añadir
                              </button>
                            </div>
                          )}

                          <div className="flex gap-2 flex-wrap mt-2 ml-6">
                            {(el.fotos || []).map((f) => (
                              <img key={f.id} src={f.url} alt="" className="w-12 h-12 object-cover rounded-md border border-slate-200" />
                            ))}
                            <label
                              title="Añadir foto de este elemento"
                              className="w-12 h-12 flex items-center justify-center border border-dashed border-slate-300 rounded-md text-slate-400 cursor-pointer hover:bg-slate-50"
                            >
                              <ImageIcon size={16} />
                              <input
                                type="file"
                                accept="image/*"
                                capture="environment"
                                className="hidden"
                                onChange={(e) => subirFotoElemento(v, el.id, e.target.files[0])}
                              />
                            </label>
                          </div>
                        </div>
                      );
                    })}
                  </div>

                  <div className="flex gap-2 mt-3">
                    <TextInput
                      value={nuevoElementoNombre[v.id] || ""}
                      onChange={(e) => setNuevoElementoNombre((prev) => ({ ...prev, [v.id]: e.target.value }))}
                      placeholder="Nombre (ej: Cocina, Salón, Dormitorio 1...)"
                      className="flex-1 !text-xs !py-1.5"
                    />
                    <TextInput
                      value={nuevoElementoMedida[v.id] || ""}
                      onChange={(e) => setNuevoElementoMedida((prev) => ({ ...prev, [v.id]: e.target.value }))}
                      placeholder="Medida (ej: 120x150)"
                      className="max-w-[140px] !text-xs !py-1.5"
                    />
                    <button
                      onClick={() => {
                        agregarElementoAMano(v, nuevoElementoNombre[v.id], nuevoElementoMedida[v.id]);
                        setNuevoElementoNombre((prev) => ({ ...prev, [v.id]: "" }));
                        setNuevoElementoMedida((prev) => ({ ...prev, [v.id]: "" }));
                      }}
                      className="px-3 py-1.5 rounded-md text-xs font-semibold text-slate-600 border border-slate-300 hover:bg-slate-50 shrink-0"
                    >
                      + Añadir habitación/elemento
                    </button>
                  </div>

                  <div className="mt-4">
                    <div className="text-sm font-semibold text-slate-700 mb-1">Extras añadidos ({(v.extras || []).length})</div>
                    {(v.extras || []).map((ex) => (
                      <div key={ex.id} className="text-sm text-slate-600 ml-2">· {ex.descripcion}</div>
                    ))}
                    <div className="flex gap-2 mt-2">
                      <TextInput value={extraTexto[v.id] || ""} onChange={(e) => setExtraTexto((prev) => ({ ...prev, [v.id]: e.target.value }))} placeholder="Describe el extra" className="flex-1" />
                      <button onClick={() => agregarExtra(v)} className="px-3 py-1.5 bg-slate-100 rounded-md text-sm font-semibold">Añadir</button>
                    </div>
                  </div>

                  <div className="mt-4">
                    <div className="text-sm font-semibold text-slate-700 mb-1">Fotos ({(v.fotos || []).length})</div>
                    <div className="flex gap-2 flex-wrap mb-2">
                      {(v.fotos || []).map((f) => <img key={f.id} src={f.url} alt="" className="w-16 h-16 object-cover rounded-md" />)}
                    </div>
                    <input type="file" accept="image/*" capture="environment" onChange={(e) => subirFoto(v, e.target.files[0])} className="text-sm" />
                  </div>

                  <div className="mt-4">
                    <div className="text-sm font-semibold text-slate-700 mb-1">Documentos ({(v.documentos || []).length})</div>
                    {Object.entries(agruparDocumentosPorCategoria(v.documentos)).map(([cat, docs]) => (
                      <div key={cat} className="mb-2">
                        <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wide mt-2 ml-2">{cat}</div>
                        {docs.map((d) => (
                          <a key={d.id} href={d.url} download={d.nombre} className="block text-sm text-emerald-700 underline ml-2">{d.nombre}</a>
                        ))}
                      </div>
                    ))}
                    <div className="flex gap-2 mt-2 items-center">
                      <Select
                        value={docCategoria[v.id] || "Planos"}
                        onChange={(e) => setDocCategoria((prev) => ({ ...prev, [v.id]: e.target.value }))}
                        className="max-w-[150px] !py-1 !text-xs shrink-0"
                      >
                        {CATEGORIAS_DOCUMENTO_MONTAJE.map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
                      <input
                        type="file"
                        accept="application/pdf,image/*,.xlsx,.xls,.ods,.doc,.docx"
                        onChange={(e) => subirDocumento(v, e.target.files[0], docCategoria[v.id] || "Planos")}
                        className="text-sm flex-1"
                      />
                    </div>
                  </div>

                  <div className="mt-4">
                    <div className="text-sm font-semibold text-slate-700 mb-1">Incidencias vinculadas ({(v.incidenciasVinculadas || []).length})</div>
                    {(v.incidenciasVinculadas || []).map((inc) => (
                      <div key={inc.id} className="text-sm text-slate-600 ml-2">· {inc.descripcion}</div>
                    ))}
                    <div className="flex gap-2 mt-2">
                      <TextInput value={nuevaIncidenciaTexto[v.id] || ""} onChange={(e) => setNuevaIncidenciaTexto((prev) => ({ ...prev, [v.id]: e.target.value }))} placeholder="Describe la incidencia nueva" className="flex-1" />
                      <button onClick={() => crearIncidencia(v)} className="px-3 py-1.5 bg-slate-100 rounded-md text-sm font-semibold">Crear</button>
                    </div>
                    <button onClick={() => setVinculandoEn(vinculandoEn === v.id ? null : v.id)} className="text-xs font-semibold text-emerald-700 mt-2">
                      {vinculandoEn === v.id ? "Cerrar" : "Vincular incidencia ya existente"}
                    </button>
                    {vinculandoEn === v.id && (
                      <div className="mt-2 border border-slate-200 rounded-md max-h-40 overflow-y-auto">
                        {(incidencias || []).length === 0 && <div className="text-sm text-slate-400 p-2">No hay incidencias registradas.</div>}
                        {(incidencias || []).map((inc) => (
                          <button key={inc.id} onClick={() => vincularIncidenciaExistente(v, inc)} className="block w-full text-left text-sm px-2 py-1 hover:bg-slate-50">
                            #{inc.numero} — {inc.especificaciones}
                          </button>
                        ))}
                      </div>
                    )}
                  </div>

                  <div className="flex gap-2 mt-4">
                    <button onClick={() => avisarWhatsapp(v)} className="flex-1 px-3 py-2 bg-emerald-50 text-emerald-700 rounded-md text-sm font-semibold">Avisar por WhatsApp</button>
                    <button onClick={() => avisarCorreo(v)} className="flex-1 px-3 py-2 bg-sky-50 text-sky-700 rounded-md text-sm font-semibold">Avisar por correo</button>
                  </div>
                  {!jefeDeObra?.email && !jefeDeObra?.telefono && (
                    <p className="text-xs text-slate-400 mt-1">Configura el jefe de obra arriba para que se rellene el destinatario automáticamente.</p>
                  )}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

/* ================= MEDICIONES ================= */

/* ================= PRESUPUESTOS Y DESPIECE DE TECHOS (calculado en el propio CRM) =================
   Réplica exacta, en JavaScript, de las plantillas Excel reales que se usaban a mano:
   - Techo abatible: marco cerrado + con puerta + sin barandilla (caso más habitual).
   - Techo corredero: solo el despiece de corte (la parte de precios de ese Excel usa
     cantidades manuales por obra, así que de momento no se automatiza el coste). */

// Techo abatible — marco cerrado, con puerta, sin barandilla.
// Fórmulas portadas 1:1 desde la plantilla real (validado: 900x2400 -> 2200,57 € de coste).
function calcularPresupuestoTechoAbatible(anchoIn, largoIn, nTechosIn) {
  const ancho = parseFloat(anchoIn) || 0;
  const largo = parseFloat(largoIn) || 0;
  const nTechos = parseFloat(nTechosIn) || 1;
  const perimetro = ancho * 2 + largo * 2;

  const despiece = [
    { perfil: "120x40 (premarco)", cantidad: 2, medida: ancho, nota: "" },
    { perfil: "120x40 (premarco)", cantidad: 2, medida: largo, nota: "" },
    { perfil: "Ángulo 80x40", cantidad: 1, medida: ancho, nota: "para tapar Pladur" },
    { perfil: "Ángulo 80x40", cantidad: 2, medida: largo, nota: "para tapar Pladur" },
    { perfil: "Marco liso 40x20", cantidad: 1, medida: ancho, nota: "" },
    { perfil: "Marco liso 40x20", cantidad: 2, medida: largo - 5, nota: "(-5 del tapón)" },
    { perfil: "120x40 (marco)", cantidad: 1, medida: ancho, nota: "" },
    { perfil: "120x40 (marco)", cantidad: 2, medida: largo - 5, nota: "(-5 del tapón)" },
    { perfil: "T ventana 40x20", cantidad: 2, medida: ancho + 70, nota: "" },
    { perfil: "T ventana 40x20", cantidad: 2, medida: largo + 81, nota: "" },
    { perfil: "100x40 rajado", cantidad: 2, medida: ancho + 112, nota: "" },
    { perfil: "100x40 rajado", cantidad: 2, medida: largo + 122, nota: "" },
    { perfil: "35x35 bruto", cantidad: 8, medida: 180, nota: "" },
    { perfil: "35x35 bruto", cantidad: 8, medida: 130, nota: "" },
  ];

  // [precio unitario, cantidad] — cantidad ya en la unidad correcta (m, ud...), escala con nTechos.
  const materiales = [
    ["Ángulo 80x40", 3.5, (ancho * 1 + largo * 2) / 1000],
    ["120x40", 11.2, (ancho * 1 + (largo - 5) * 2) / 1000],
    ["100x40", 9.2, ((ancho + 112) * 2 + (largo + 122) * 2) / 1000],
    ["Marco liso 40x20", 3.7, (ancho * 1 + (largo - 5) * 2) / 1000],
    ["T 40x20", 4.6, ((ancho + 70) * 2 + (largo + 81) * 2) / 1000],
    ["Bisagras", 1.1, 16],
    ["Escuadra ventana", 0.5, 6],
    ["35x35", 3.5, 2.7],
    ["Premarco techo", 50, 1],
    ["Goma marco y hoja", 0.3, perimetro / 1000],
    ["Polímero y silicona", 5, 4],
    ["Panel sandwich (Polipanel)", 180, 1],
    ["Loseta (m2)", 9.4, (1.5 * largo) / 1000],
    ["Danopol (m2)", 6, (1.8 * largo) / 1000],
    ["Cola (litro)", 2.9, 4],
    ["Pistones", 19.4, 2],
    ["Pletinas", 3, 3],
    ["Placa anclaje", 10, 1],
    ["Pintura", 5, 1],
    ["Motor", 284, 1],
    ["Batería", 82, 1],
    ["Tapón 120x40", 0.9, 2],
    ["Tapón 40x20", 0.3, 2],
    ["Tornillo 4,2x19", 0.02, 20],
    ["Tornillo 4,8x25 (perímetro panel)", 0.02, perimetro / 10 / 8],
    ["Tornillo 4,2x22 (interior)", 0.02, largo / 10 / 5],
    ["Tornillo 4,8x25 (tubo-marco)", 0.02, perimetro / 10 / 9],
  ];
  // Estos 4 son costes fijos por proyecto: NO escalan con el nº de techos (así está en la plantilla real).
  const fijos = [
    ["Fabricación", 300, 1],
    ["Montaje", 720, 1],
    ["Dietas y hotel", 90, 1],
    ["Km", 100, 1],
  ];

  let costeTotal = 0;
  materiales.forEach(([, precio, cantidad]) => { costeTotal += precio * cantidad * nTechos; });
  fijos.forEach(([, precio, cantidad]) => { costeTotal += precio * cantidad; });

  const mas20 = costeTotal + costeTotal * 0.2;
  const mas20_5 = mas20 + mas20 * 0.05;
  const mas30 = costeTotal + costeTotal * 0.3;

  return {
    medidaPanel: { ancho: ancho + 102, largo: largo + 115 },
    despiece,
    costeTotal, mas20, mas20_5, mas30,
  };
}

// Techo corredero — despiece de corte (fórmulas de la plantilla real, validado contra el
// ejemplo 840x2780x860 -> panel 980x2950). La parte de PRECIOS de este tipo de techo usa
// cantidades manuales por obra en el Excel original, así que no se automatiza aquí todavía.
function calcularDespieceTechoCorredero(anchoIn, largoIn, alturaIn) {
  const ancho = parseFloat(anchoIn) || 0;
  const largo = parseFloat(largoIn) || 0;
  const altura = parseFloat(alturaIn) || 0;
  // Deducido del único ejemplo disponible (840→980, es decir +140). A confirmar con más casos reales.
  const panelAncho = ancho + 140;
  const panelLargo = largo + 170;

  const items = [
    { perfil: "Cobija", cantidad: 1, medida: panelAncho < 1001 ? 2000 : panelAncho * 2, nota: "" },
    { perfil: "Larguero largo (marco)", cantidad: 1, medida: largo + 185, nota: "dibujo 1" },
    { perfil: "Larguero corto (marco)", cantidad: 1, medida: altura + 185, nota: "dibujo 2" },
    { perfil: "Larguero largo (marco)", cantidad: 1, medida: largo + 57, nota: "dibujo 3" },
    { perfil: "Larguero corto (marco)", cantidad: 1, medida: altura + 57, nota: "dibujo 2" },
    { perfil: "Larguero largo (hoja) — rectángulo 80x40x3", cantidad: 2, medida: largo + 50, nota: "" },
    { perfil: "Larguero corto (hoja) — rectángulo 120x40x3", cantidad: 2, medida: altura + 46, nota: "" },
    { perfil: "Ancho (hoja) — rectángulo 80x40x3", cantidad: 4, medida: ancho, nota: "" },
    { perfil: "Ángulo largo — 100x40x3", cantidad: 2, medida: largo + 170, nota: "" },
    { perfil: "Ángulo corto — 100x40x3", cantidad: 2, medida: altura + 96, nota: "" },
    { perfil: "Rodamiento", cantidad: 1, medida: ancho + 80, nota: "" },
  ];

  return { panel: { ancho: panelAncho, largo: panelLargo }, items };
}

// Agrupa el despiece de TODOS los techos de un proyecto (abatibles + correderos) en una
// sola lista por tipo de perfil, sumando los metros lineales totales que hacen falta.
function calcularDespieceConjuntoProyecto(techos) {
  const grupos = {};
  (techos || []).forEach((t) => {
    const items = t.tipo === "abatible" ? t.resultado.despiece : t.despiece.items;
    items.forEach((it) => {
      const clave = it.perfil;
      if (!grupos[clave]) grupos[clave] = { perfil: clave, piezas: 0, metros: 0 };
      grupos[clave].piezas += it.cantidad;
      grupos[clave].metros += (it.cantidad * it.medida) / 1000;
    });
  });
  return Object.values(grupos).sort((a, b) => b.metros - a.metros);
}

// Compara el despiece agrupado con el catálogo de Stock (materiales) por coincidencia de
// nombre, para saber qué hay y qué falta. No sabe de almacenes/ubicaciones (eso solo
// existe hoy para Cristales) — solo compara cantidad total disponible vs necesaria.
function compararDespieceConStock(despieceAgrupado, materiales) {
  return despieceAgrupado.map((d) => {
    const nombreBuscado = d.perfil.toLowerCase();
    const material = (materiales || []).find((m) => {
      const n = (m.descripcion || "").toLowerCase();
      return n.includes(nombreBuscado) || nombreBuscado.includes(n);
    });
    const disponible = material ? parseFloat(material.stockReal) || 0 : null;
    const falta = disponible === null ? d.metros : Math.max(0, d.metros - disponible);
    return { ...d, material, disponible, falta };
  });
}

// Croquis guía fijo del techo corredero, con el mismo estilo de los croquis a mano
// (ancho arriba, código de vivienda a la izquierda, largo a la derecha, altura abajo en V).
function CroquisTechoCorredero() {
  return (
    <svg viewBox="0 0 280 220" className="w-full max-w-xs mx-auto" role="img" aria-label="Croquis de premarco: ancho arriba, largo lateral, altura abajo">
      <line x1="50" y1="30" x2="200" y2="30" stroke="#334155" strokeWidth="2" />
      <text x="105" y="22" fontSize="13" fill="#334155">ancho</text>
      <line x1="50" y1="30" x2="50" y2="130" stroke="#334155" strokeWidth="2" />
      <line x1="130" y1="30" x2="130" y2="130" stroke="#334155" strokeWidth="2" />
      <text x="10" y="84" fontSize="13" fill="#334155">vivienda</text>
      <text x="145" y="84" fontSize="13" fill="#334155">largo</text>
      <line x1="50" y1="130" x2="80" y2="170" stroke="#334155" strokeWidth="2" />
      <line x1="130" y1="130" x2="100" y2="170" stroke="#334155" strokeWidth="2" />
      <text x="65" y="195" fontSize="13" fill="#334155">altura</text>
    </svg>
  );
}

// Lienzo táctil sencillo para anotar particularidades de la medición a mano,
// igual que se hace hoy en papel. Guarda el trazo como imagen (dataURL).
function LienzoDibujo({ onGuardar }) {
  const canvasRef = useRef(null);
  const dibujando = useRef(false);

  const getPos = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return {
      x: ((clientX - rect.left) / rect.width) * canvas.width,
      y: ((clientY - rect.top) / rect.height) * canvas.height,
    };
  };

  const empezar = (e) => {
    e.preventDefault();
    dibujando.current = true;
    const ctx = canvasRef.current.getContext("2d");
    const p = getPos(e);
    ctx.beginPath();
    ctx.moveTo(p.x, p.y);
  };
  const dibujar = (e) => {
    if (!dibujando.current) return;
    e.preventDefault();
    const ctx = canvasRef.current.getContext("2d");
    const p = getPos(e);
    ctx.lineTo(p.x, p.y);
    ctx.strokeStyle = "#1e293b";
    ctx.lineWidth = 2;
    ctx.lineCap = "round";
    ctx.stroke();
  };
  const terminar = () => { dibujando.current = false; };

  const limpiar = () => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  };

  useEffect(() => { limpiar(); }, []);

  return (
    <div>
      <canvas
        ref={canvasRef}
        width={300}
        height={180}
        className="w-full border border-slate-300 rounded-md touch-none bg-white"
        onMouseDown={empezar} onMouseMove={dibujar} onMouseUp={terminar} onMouseLeave={terminar}
        onTouchStart={empezar} onTouchMove={dibujar} onTouchEnd={terminar}
      />
      <div className="flex gap-2 mt-2">
        <button type="button" onClick={limpiar} className="text-xs font-semibold text-slate-500 border border-slate-300 px-2.5 py-1.5 rounded-md hover:bg-slate-50">Limpiar</button>
        <button
          type="button"
          onClick={() => onGuardar(canvasRef.current.toDataURL("image/png"))}
          className="text-xs font-semibold text-emerald-700 border border-emerald-200 bg-emerald-50 px-2.5 py-1.5 rounded-md hover:bg-emerald-100"
        >
          Guardar dibujo en esta medición
        </button>
      </div>
    </div>
  );
}

// Módulo "Techos" dentro de cada medición: añadir techos abatibles (con presupuesto
// automático) o correderos (con croquis guía + dibujo + despiece de corte), y ver el
// historial de los que ya se han calculado para esa medición.
function TechosMedicion({ medicionId }) {
  const basePath = `mediciones/${medicionId}/techos`;
  const [techos, setTechos] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [nuevoTipo, setNuevoTipo] = useState(null);
  const [expandidoId, setExpandidoId] = useState(null);
  const [form, setForm] = useState({ vivienda: "", ancho: "", largo: "", altura: "", nTechos: "1", abre: "Izquierda" });
  const [dibujoPendiente, setDibujoPendiente] = useState(null);

  useEffect(() => {
    if (!medicionId) return;
    (async () => {
      setCargando(true);
      try {
        const snap = await fbGet(ref(fbDb, basePath));
        setTechos(toArray(snap.val()));
      } catch (e) {
        console.error(e);
      } finally {
        setCargando(false);
      }
    })();
  }, [medicionId]);

  const guardarTechos = async (next) => {
    setTechos(next);
    try {
      await fbSet(ref(fbDb, basePath), next);
    } catch (e) {
      console.error(e);
    }
  };

  const resetForm = () => {
    setForm({ vivienda: "", ancho: "", largo: "", altura: "", nTechos: "1", abre: "Izquierda" });
    setDibujoPendiente(null);
    setNuevoTipo(null);
  };

  const anadirAbatible = () => {
    if (!form.ancho || !form.largo) { alert("Faltan el ancho y el largo."); return; }
    const resultado = calcularPresupuestoTechoAbatible(form.ancho, form.largo, form.nTechos);
    const nuevo = {
      id: uid(),
      tipo: "abatible",
      vivienda: form.vivienda || "Sin nombre",
      ancho: form.ancho, largo: form.largo, nTechos: form.nTechos || "1",
      resultado,
      creadoEn: Date.now(),
    };
    guardarTechos([nuevo, ...techos]);
    resetForm();
  };

  const anadirCorredero = () => {
    if (!form.ancho || !form.largo || !form.altura) { alert("Faltan ancho, largo o altura."); return; }
    const despiece = calcularDespieceTechoCorredero(form.ancho, form.largo, form.altura);
    const nuevo = {
      id: uid(),
      tipo: "corredero",
      vivienda: form.vivienda || "Sin nombre",
      ancho: form.ancho, largo: form.largo, altura: form.altura, abre: form.abre,
      despiece,
      dibujo: dibujoPendiente || null,
      creadoEn: Date.now(),
    };
    guardarTechos([nuevo, ...techos]);
    resetForm();
  };

  const borrarTecho = (id) => {
    if (!window.confirm("¿Borrar este techo de la medición?")) return;
    guardarTechos(techos.filter((t) => t.id !== id));
  };

  return (
    <div className="mt-8 pt-8 border-t border-slate-200">
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-display font-bold text-slate-800">Techos ({techos.length})</h2>
        {!nuevoTipo && (
          <div className="flex gap-2">
            <button onClick={() => setNuevoTipo("abatible")} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold px-3 py-2 rounded-md">
              + Techo abatible
            </button>
            <button onClick={() => setNuevoTipo("corredero")} className="text-xs font-semibold text-slate-600 border border-slate-300 px-3 py-2 rounded-md hover:bg-slate-50">
              + Techo corredero
            </button>
          </div>
        )}
      </div>

      {nuevoTipo && (
        <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4">
          <p className="text-sm font-semibold text-slate-700 mb-3">
            Nuevo techo {nuevoTipo === "abatible" ? "abatible (marco cerrado, con puerta, sin barandilla)" : "corredero"}
          </p>

          {nuevoTipo === "corredero" && (
            <div className="mb-4">
              <CroquisTechoCorredero />
            </div>
          )}

          <div className="grid grid-cols-2 gap-3 mb-3">
            <Field label="Vivienda / referencia">
              <TextInput value={form.vivienda} onChange={(e) => setForm({ ...form, vivienda: e.target.value })} placeholder="Ej: B9-V130" />
            </Field>
            {nuevoTipo === "abatible" && (
              <Field label="Nº de techos iguales">
                <TextInput type="number" value={form.nTechos} onChange={(e) => setForm({ ...form, nTechos: e.target.value })} />
              </Field>
            )}
            {nuevoTipo === "corredero" && (
              <Field label="Abre visto desde fuera">
                <select className="w-full border border-slate-300 rounded-md px-3 py-2 text-sm" value={form.abre} onChange={(e) => setForm({ ...form, abre: e.target.value })}>
                  <option>Izquierda</option>
                  <option>Derecha</option>
                </select>
              </Field>
            )}
            <Field label="Ancho (mm)">
              <TextInput type="number" value={form.ancho} onChange={(e) => setForm({ ...form, ancho: e.target.value })} placeholder="900" />
            </Field>
            <Field label="Largo (mm)">
              <TextInput type="number" value={form.largo} onChange={(e) => setForm({ ...form, largo: e.target.value })} placeholder="2400" />
            </Field>
            {nuevoTipo === "corredero" && (
              <Field label="Altura (mm)">
                <TextInput type="number" value={form.altura} onChange={(e) => setForm({ ...form, altura: e.target.value })} placeholder="860" />
              </Field>
            )}
          </div>

          {nuevoTipo === "corredero" && (
            <div className="mb-4">
              <p className="text-xs font-semibold text-slate-600 mb-1.5">Dibuja aquí particularidades (opcional)</p>
              <LienzoDibujo onGuardar={(dataUrl) => { setDibujoPendiente(dataUrl); alert("Dibujo guardado, se adjuntará al pulsar \"Calcular y guardar\"."); }} />
            </div>
          )}

          <div className="flex gap-2">
            <button
              onClick={nuevoTipo === "abatible" ? anadirAbatible : anadirCorredero}
              style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
              className="text-sm font-semibold px-4 py-2 rounded-md"
            >
              Calcular y guardar
            </button>
            <button onClick={resetForm} className="text-sm font-semibold text-slate-500 px-4 py-2 rounded-md hover:bg-slate-50">Cancelar</button>
          </div>
        </div>
      )}

      {cargando ? (
        <p className="text-sm text-slate-400">Cargando…</p>
      ) : techos.length === 0 && !nuevoTipo ? (
        <p className="text-sm text-slate-400">Todavía no hay techos calculados en esta medición.</p>
      ) : (
        <div className="space-y-2">
          {techos.map((t) => {
            const abierto = expandidoId === t.id;
            return (
              <div key={t.id} className="border border-slate-200 rounded-lg overflow-hidden bg-white">
                <button onClick={() => setExpandidoId(abierto ? null : t.id)} className="w-full flex items-center justify-between px-4 py-3 text-left">
                  <div className="text-sm">
                    <span className="font-semibold text-slate-800">{t.vivienda}</span>
                    <span className="text-slate-400"> · {t.tipo === "abatible" ? "Abatible" : "Corredero"} · {t.ancho}×{t.largo}{t.tipo === "corredero" ? `×${t.altura}` : ""} mm</span>
                  </div>
                  <div className="flex items-center gap-3">
                    {t.tipo === "abatible" && <span className="text-xs font-semibold text-emerald-700">{money(t.resultado.costeTotal)}</span>}
                    <Trash2 size={14} className="text-slate-300 hover:text-rose-500" onClick={(e) => { e.stopPropagation(); borrarTecho(t.id); }} />
                  </div>
                </button>

                {abierto && (
                  <div className="px-4 pb-4 border-t border-slate-100 pt-3">
                    {t.tipo === "abatible" ? (
                      <>
                        <p className="text-xs text-slate-500 mb-2">
                          Medida de fabricación del panel: {Math.round(t.resultado.medidaPanel.ancho)} x {Math.round(t.resultado.medidaPanel.largo)} mm
                        </p>
                        <table className="w-full text-xs mb-3">
                          <thead><tr className="text-slate-400 text-left"><th className="pb-1">Perfil</th><th className="pb-1">Cant.</th><th className="pb-1">Medida (mm)</th><th className="pb-1">Nota</th></tr></thead>
                          <tbody>
                            {t.resultado.despiece.map((d, i) => (
                              <tr key={i} className="border-t border-slate-100">
                                <td className="py-1 pr-2">{d.perfil}</td>
                                <td className="py-1 pr-2">{d.cantidad}</td>
                                <td className="py-1 pr-2">{Math.round(d.medida)}</td>
                                <td className="py-1 text-slate-400">{d.nota}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <div className="grid grid-cols-4 gap-2 text-xs">
                          <div className="bg-slate-50 rounded-md p-2"><div className="text-slate-400">Coste</div><div className="font-semibold">{money(t.resultado.costeTotal)}</div></div>
                          <div className="bg-slate-50 rounded-md p-2"><div className="text-slate-400">+20%</div><div className="font-semibold">{money(t.resultado.mas20)}</div></div>
                          <div className="bg-slate-50 rounded-md p-2"><div className="text-slate-400">+20%+5%</div><div className="font-semibold">{money(t.resultado.mas20_5)}</div></div>
                          <div className="bg-slate-50 rounded-md p-2"><div className="text-slate-400">+30%</div><div className="font-semibold">{money(t.resultado.mas30)}</div></div>
                        </div>
                      </>
                    ) : (
                      <>
                        <p className="text-xs text-slate-500 mb-2">Abre: {t.abre} · Medida panel: {Math.round(t.despiece.panel.ancho)} x {Math.round(t.despiece.panel.largo)} mm</p>
                        {t.dibujo && <img src={t.dibujo} alt="Dibujo de la medición" className="border border-slate-200 rounded-md mb-3 max-w-xs" />}
                        <table className="w-full text-xs">
                          <thead><tr className="text-slate-400 text-left"><th className="pb-1">Pieza</th><th className="pb-1">Cant.</th><th className="pb-1">Medida (mm)</th><th className="pb-1">Nota</th></tr></thead>
                          <tbody>
                            {t.despiece.items.map((d, i) => (
                              <tr key={i} className="border-t border-slate-100">
                                <td className="py-1 pr-2">{d.perfil}</td>
                                <td className="py-1 pr-2">{d.cantidad}</td>
                                <td className="py-1 pr-2">{Math.round(d.medida)}</td>
                                <td className="py-1 text-slate-400">{d.nota}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        <p className="text-xs text-amber-600 mt-2">El coste de este tipo de techo todavía no se calcula solo — se sigue presupuestando a mano.</p>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function MedicionesModulo({ mediciones, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, incidencias, onUpsertIncidencia, onPasarAPresupuesto }) {
  const [q, setQ] = useState("");

  const filtered = mediciones.filter((m) => {
    if (!q) return true;
    return `${m.clienteNombre} ${m.direccion || ""}`.toLowerCase().includes(q.toLowerCase());
  });

  if (view === "form") {
    const editing = mediciones.find((m) => m.id === editId) || null;
    return (
      <MedicionForm
        initial={editing}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={(data) => { onUpsert(data); setDetailId(data.id); setView("detail"); }}
      />
    );
  }

  if (view === "detail") {
    const medicion = mediciones.find((m) => m.id === detailId);
    if (!medicion) { setView("list"); return null; }
    return (
      <MedicionDetail
        medicion={medicion}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(medicion.id); setView("form"); }}
        onDelete={() => { onDelete(medicion.id); setView("list"); }}
        incidencias={incidencias}
        onUpsertIncidencia={onUpsertIncidencia}
        onPasarAPresupuesto={onPasarAPresupuesto}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<Ruler size={20} className="text-[#2E8B57]" />}
        title="Mediciones"
        manualKey="mediciones"
        subtitle={`${mediciones.length} medición${mediciones.length === 1 ? "" : "es"} registrada${mediciones.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVA MEDICIÓN
      </button>

      <div className="relative flex-1 max-w-sm mb-4">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder="Buscar por cliente o dirección..."
          className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]"
        />
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Dirección</th>
              <th className="px-4 py-3 font-semibold">Fecha</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={4} className="px-4 py-10 text-center text-slate-400 text-sm">No hay mediciones que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((m) => (
              <tr
                key={m.id}
                onClick={() => { setDetailId(m.id); setView("detail"); }}
                className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition"
              >
                <td className="px-4 py-3 font-medium text-slate-800">{m.clienteNombre}</td>
                <td className="px-4 py-3 text-slate-500">{m.direccion || "—"}</td>
                <td className="px-4 py-3 text-slate-500">{m.fecha || "—"}</td>
                <td className="px-4 py-3">
                  {m.presupuestoCreado
                    ? <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Pasada a presupuesto</Badge>
                    : <Badge className="bg-amber-50 text-amber-700 ring-amber-200">En curso</Badge>}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function MedicionForm({ initial, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || { id: null, clienteNombre: "", direccion: "", fecha: new Date().toISOString().slice(0, 10), notas: "" }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.clienteNombre.trim()) { setErrorMsg("Falta el nombre del cliente (es obligatorio)."); return; }
    if (!f.fecha) { setErrorMsg("Falta la fecha (es obligatoria)."); return; }
    setErrorMsg("");
    onSave({ ...f, id: f.id || uid() });
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">{initial?.id ? "Editar medición" : "Nueva medición"}</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <p className="text-xs text-slate-400">
          El cliente no tiene por qué existir todavía en el CRM — puedes escribir su nombre aunque aún no lo hayas dado de alta como cliente. Cuando pases esta medición a presupuesto, si el cliente no existe, se te avisará para darlo de alta primero.
        </p>
        <Field label="Cliente" required>
          <TextInput value={f.clienteNombre} onChange={set("clienteNombre")} placeholder="Nombre del cliente o de la obra" />
        </Field>
        <Field label="Dirección">
          <TextInput value={f.direccion} onChange={set("direccion")} placeholder="Dirección donde se mide" />
        </Field>
        <Field label="Fecha" required>
          <TextInput type="date" value={f.fecha} onChange={set("fecha")} />
        </Field>
        <Field label="Notas">
          <TextArea value={f.notas} onChange={set("notas")} rows={3} placeholder="Cualquier apunte general de la visita" />
        </Field>
        <div className="flex gap-3 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2 rounded-md border border-slate-300 text-sm font-semibold text-slate-600">Cancelar</button>
          <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="px-4 py-2 rounded-md text-sm font-semibold">Guardar</button>
        </div>
      </form>
    </div>
  );
}

function MedicionDetail({ medicion, onBack, onEdit, onDelete, incidencias, onUpsertIncidencia, onPasarAPresupuesto }) {
  return (
    <div className="p-8 max-w-5xl overflow-x-hidden">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>

      <div className="flex items-start justify-between flex-wrap gap-3 mb-6">
        <div>
          <h1 className="font-display text-2xl font-extrabold text-slate-900">{medicion.clienteNombre}</h1>
          <p className="text-sm text-slate-500 mt-1">{medicion.direccion || "Sin dirección"} · {medicion.fecha || "—"}</p>
          {medicion.notas && <p className="text-sm text-slate-500 mt-1">{medicion.notas}</p>}
        </div>
        <div className="flex gap-2">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          <button onClick={() => { if (window.confirm("¿Borrar esta medición? Esta acción no se puede deshacer.")) onDelete(); }} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Borrar</button>
        </div>
      </div>

      {medicion.presupuestoCreado && (
        <div className="px-4 py-3 rounded-md bg-emerald-50 border border-emerald-300 text-emerald-800 text-sm font-semibold mb-4">
          ✓ Esta medición ya se pasó a presupuesto. Si mides algo más, puedes volver a pasarla a presupuesto cuando quieras.
        </div>
      )}

      <ControlElementosPorVivienda
        basePath={`mediciones/${medicion.id}/detalle`}
        proyectoId={null}
        incidencias={incidencias}
        onUpsertIncidencia={onUpsertIncidencia}
        accionPrincipal={{
          etiqueta: "Pasar a presupuesto →",
          onClick: (resumenGlobal) => onPasarAPresupuesto(medicion, resumenGlobal),
        }}
      />

      <TechosMedicion medicionId={medicion.id} />
    </div>
  );
}

/* ================= CHAT INTERNO ================= */

function ChatModulo({ usuarios, currentUser, proyectos, onCrearTarea }) {
  const salas = useMemo(() => [
    { id: "general", nombre: "General" },
    ...proyectos.map((p) => ({ id: `proyecto_${p.id}`, nombre: p.nombre || `Proyecto ${p.numero || ""}` })),
  ], [proyectos]);

  const [salaId, setSalaId] = useState("general");
  const sala = salas.find((s) => s.id === salaId) || salas[0];
  const [mensajes, setMensajes] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [texto, setTexto] = useState("");
  const [seleccionando, setSeleccionando] = useState(false);
  const [seleccionados, setSeleccionados] = useState({});
  const [creandoTareaDesde, setCreandoTareaDesde] = useState(null);
  const [busquedaSala, setBusquedaSala] = useState("");
  const finRef = useRef(null);

  useEffect(() => {
    let activo = true;
    const cargar = async () => {
      try {
        const snap = await fbGet(ref(fbDb, `chatMensajes/${salaId}`));
        const val = snap.exists() ? snap.val() : {};
        const lista = toArray(val).sort((a, b) => (a.fecha || 0) - (b.fecha || 0));
        if (activo) setMensajes(lista);
      } catch (e) { /* sin conexión momentánea, se reintenta en el siguiente sondeo */ }
      if (activo) setCargando(false);
    };
    setCargando(true);
    cargar();
    const intervalo = setInterval(cargar, 8000);
    return () => { activo = false; clearInterval(intervalo); };
  }, [salaId]);

  useEffect(() => {
    finRef.current?.scrollIntoView({ block: "end" });
  }, [mensajes.length, salaId]);

  const enviar = async () => {
    const limpio = texto.trim();
    if (!limpio || !currentUser) return;
    const msgId = uid();
    const mensaje = {
      id: msgId,
      autorId: currentUser.id,
      autorNombre: `${currentUser.nombre} ${currentUser.apellidos || ""}`.trim(),
      texto: limpio,
      fecha: Date.now(),
    };
    setMensajes((prev) => [...prev, mensaje]);
    setTexto("");
    await fbSet(ref(fbDb, `chatMensajes/${salaId}/${msgId}`), mensaje).catch(() => {});
  };

  const toggleSeleccion = (id) => {
    setSeleccionados((prev) => {
      const next = { ...prev };
      if (next[id]) delete next[id]; else next[id] = true;
      return next;
    });
  };

  const exportarPDF = () => {
    const elegidos = mensajes.filter((m) => seleccionados[m.id]).sort((a, b) => (a.fecha || 0) - (b.fecha || 0));
    if (elegidos.length === 0) { alert("Selecciona al menos un mensaje para exportar."); return; }
    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Conversación - ${escapeHtml(sala.nombre)}</title>
      <style>
        body { font-family: Arial, Helvetica, sans-serif; padding: 28px; color: #1e293b; }
        h1 { font-size: 18px; margin-bottom: 2px; }
        .meta { color: #64748b; font-size: 12px; margin-bottom: 24px; }
        .msg { margin-bottom: 14px; padding-bottom: 10px; border-bottom: 1px solid #e2e8f0; }
        .autor { font-weight: bold; }
        .hora { color: #94a3b8; font-size: 11px; margin-left: 8px; }
        .texto { margin-top: 4px; white-space: pre-wrap; }
      </style></head><body>
      <h1>Chat ALUMAVEL — ${escapeHtml(sala.nombre)}</h1>
      <div class="meta">Exportado el ${new Date().toLocaleString("es-ES")} · ${elegidos.length} mensaje(s)</div>
      ${elegidos.map((m) => `<div class="msg"><span class="autor">${escapeHtml(m.autorNombre)}</span><span class="hora">${new Date(m.fecha).toLocaleString("es-ES")}</span><div class="texto">${escapeHtml(m.texto)}</div></div>`).join("")}
      </body></html>`;
    const ventana = window.open("", "_blank");
    if (!ventana) { alert("El navegador ha bloqueado la ventana emergente. Permítela para poder exportar."); return; }
    ventana.document.write(html);
    ventana.document.close();
    ventana.focus();
    setTimeout(() => ventana.print(), 350);
  };

  const salasFiltradas = busquedaSala ? salas.filter((s) => s.nombre.toLowerCase().includes(busquedaSala.toLowerCase())) : salas;

  return (
    <div className="p-8 max-w-6xl h-[calc(100vh-2rem)] flex flex-col overflow-hidden">
      <Header icon={<MessageCircle size={20} className="text-[#2E8B57]" />} title="Chat interno" manualKey="chat" subtitle="General y por proyecto — todo el equipo puede escribir y ver el historial" />

      <div className="flex-1 flex gap-4 min-h-0">
        <div className="w-56 shrink-0 bg-white border border-slate-200 rounded-lg overflow-hidden flex flex-col">
          <div className="p-2 border-b border-slate-100">
            <input value={busquedaSala} onChange={(e) => setBusquedaSala(e.target.value)} placeholder="Buscar proyecto..." className="w-full text-xs px-2 py-1.5 rounded-md border border-slate-200" />
          </div>
          <div className="overflow-y-auto flex-1">
            {salasFiltradas.map((s) => (
              <button
                key={s.id}
                onClick={() => { setSalaId(s.id); setSeleccionando(false); setSeleccionados({}); }}
                className={`w-full text-left px-3 py-2.5 text-sm border-b border-slate-50 truncate ${salaId === s.id ? "bg-[#2E8B57]/10 text-[#2E8B57] font-semibold" : "text-slate-600 hover:bg-slate-50"}`}
              >
                {s.id === "general" ? "💬 " : "📁 "}{s.nombre}
              </button>
            ))}
          </div>
        </div>

        <div className="flex-1 bg-white border border-slate-200 rounded-lg flex flex-col min-w-0">
          <div className="px-4 py-3 border-b border-slate-100 flex items-center justify-between">
            <h2 className="font-semibold text-slate-800 text-sm truncate">{sala.nombre}</h2>
            <div className="flex gap-2 shrink-0">
              {seleccionando ? (
                <>
                  <button onClick={exportarPDF} className="text-xs font-semibold text-white bg-[#2E8B57] px-3 py-1.5 rounded-md">Exportar ({Object.keys(seleccionados).length})</button>
                  <button onClick={() => { setSeleccionando(false); setSeleccionados({}); }} className="text-xs font-semibold text-slate-500 border border-slate-300 px-3 py-1.5 rounded-md">Cancelar</button>
                </>
              ) : (
                <button onClick={() => setSeleccionando(true)} className="text-xs font-semibold text-slate-600 border border-slate-300 px-3 py-1.5 rounded-md hover:bg-slate-50">
                  <Download size={12} className="inline -mt-0.5 mr-1" /> Exportar a PDF
                </button>
              )}
            </div>
          </div>

          <div className="flex-1 overflow-y-auto px-4 py-3 space-y-3">
            {cargando && mensajes.length === 0 && <p className="text-center text-slate-400 text-sm py-8">Cargando conversación...</p>}
            {!cargando && mensajes.length === 0 && <p className="text-center text-slate-400 text-sm py-8">Todavía no hay mensajes en esta sala. ¡Escribe el primero!</p>}
            {mensajes.map((m) => {
              const esMio = m.autorId === currentUser?.id;
              return (
                <div key={m.id} className={`flex gap-2 ${seleccionando ? "items-center" : ""}`}>
                  {seleccionando && (
                    <input type="checkbox" checked={!!seleccionados[m.id]} onChange={() => toggleSeleccion(m.id)} className="mt-1 shrink-0" />
                  )}
                  <div className={`group max-w-[75%] ${esMio ? "ml-auto" : ""}`}>
                    <div className={`rounded-lg px-3 py-2 text-sm ${esMio ? "bg-[#2E8B57] text-white" : "bg-slate-100 text-slate-800"}`}>
                      {!esMio && <p className="text-[11px] font-bold opacity-70 mb-0.5">{m.autorNombre}</p>}
                      <p className="whitespace-pre-wrap">{m.texto}</p>
                    </div>
                    <div className={`flex items-center gap-2 mt-1 ${esMio ? "justify-end" : ""}`}>
                      <span className="text-[10px] text-slate-400">{new Date(m.fecha).toLocaleString("es-ES", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" })}</span>
                      <button onClick={() => setCreandoTareaDesde(m)} className="text-[10px] font-semibold text-sky-600 opacity-0 group-hover:opacity-100 transition">→ Crear tarea</button>
                    </div>
                  </div>
                </div>
              );
            })}
            <div ref={finRef} />
          </div>

          <div className="p-3 border-t border-slate-100 flex gap-2">
            <TextArea
              value={texto}
              onChange={(e) => setTexto(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); enviar(); } }}
              rows={1}
              placeholder="Escribe un mensaje... (Enter para enviar, Shift+Enter para salto de línea)"
              className="flex-1 resize-none"
            />
            <button onClick={enviar} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="px-4 rounded-md flex items-center justify-center shrink-0">
              <Send size={16} />
            </button>
          </div>
        </div>
      </div>

      {creandoTareaDesde && (
        <CrearTareaModal
          mensaje={creandoTareaDesde}
          usuarios={usuarios}
          currentUser={currentUser}
          salaId={salaId}
          salaNombre={sala.nombre}
          onClose={() => setCreandoTareaDesde(null)}
          onCrear={(data) => { onCrearTarea(data); setCreandoTareaDesde(null); }}
        />
      )}
    </div>
  );
}

function CrearTareaModal({ mensaje, usuarios, currentUser, salaId, salaNombre, onClose, onCrear }) {
  const [asignadoA, setAsignadoA] = useState("");
  const [titulo, setTitulo] = useState(mensaje.texto.slice(0, 80));
  const [descripcion, setDescripcion] = useState(mensaje.texto);
  const [fechaLimite, setFechaLimite] = useState("");
  const [requiereConfirmacion, setRequiereConfirmacion] = useState(true);
  const [error, setError] = useState("");

  const candidatos = usuarios.filter((u) => u.id !== currentUser?.id);

  const submit = () => {
    if (!asignadoA) { setError("Elige a quién se la asignas."); return; }
    if (!titulo.trim()) { setError("Ponle un título a la tarea."); return; }
    const usuario = usuarios.find((u) => u.id === asignadoA);
    onCrear({
      titulo: titulo.trim(),
      descripcion: descripcion.trim(),
      asignadoA,
      asignadoANombre: usuario ? `${usuario.nombre} ${usuario.apellidos || ""}`.trim() : "",
      requiereConfirmacion,
      fechaLimite,
      salaId,
      salaNombre,
    });
  };

  return (
    <div className="fixed inset-0 bg-black/40 z-50 flex items-center justify-center p-4" onClick={onClose}>
      <div className="bg-white rounded-lg p-5 max-w-md w-full max-h-[85vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <h3 className="font-display font-bold text-slate-800 mb-3">Crear tarea desde este mensaje</h3>
        {error && <p className="text-xs text-rose-600 font-semibold mb-3">⚠ {error}</p>}
        <div className="space-y-3">
          <Field label="Asignar a" required>
            <Select value={asignadoA} onChange={(e) => setAsignadoA(e.target.value)}>
              <option value="">Elige un compañero...</option>
              {candidatos.map((u) => <option key={u.id} value={u.id}>{u.nombre} {u.apellidos || ""}</option>)}
            </Select>
          </Field>
          <Field label="Título" required>
            <TextInput value={titulo} onChange={(e) => setTitulo(e.target.value)} />
          </Field>
          <Field label="Descripción">
            <TextArea value={descripcion} onChange={(e) => setDescripcion(e.target.value)} rows={3} />
          </Field>
          <Field label="Fecha límite (opcional)">
            <TextInput type="date" value={fechaLimite} onChange={(e) => setFechaLimite(e.target.value)} />
          </Field>
          <label className="flex items-center gap-2 text-sm text-slate-600">
            <input type="checkbox" checked={requiereConfirmacion} onChange={(e) => setRequiereConfirmacion(e.target.checked)} />
            Pedir confirmación cuando la marque como hecha (si no, se da por completada directamente)
          </label>
        </div>
        <div className="flex gap-2 mt-5">
          <button onClick={onClose} className="flex-1 text-sm font-semibold text-slate-600 border border-slate-300 px-4 py-2 rounded-md">Cancelar</button>
          <button onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex-1 text-sm font-semibold px-4 py-2 rounded-md">Crear tarea</button>
        </div>
      </div>
    </div>
  );
}

/* ================= TAREAS ================= */

function TareasModulo({ tareas, usuarios, currentUser, onMarcarHecha, onConfirmar, onDelete }) {
  const [filtro, setFiltro] = useState("asignadas");

  const asignadas = tareas.filter((t) => t.asignadoA === currentUser?.id).sort((a, b) => (b.fechaCreacion || 0) - (a.fechaCreacion || 0));
  const creadas = tareas.filter((t) => t.asignadoPor === currentUser?.id).sort((a, b) => (b.fechaCreacion || 0) - (a.fechaCreacion || 0));
  const lista = filtro === "asignadas" ? asignadas : creadas;

  const badgeEstado = (t) => {
    if (t.estado === "Hecha") return <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Hecha</Badge>;
    if (t.estado === "Pendiente de confirmar") return <Badge className="bg-amber-50 text-amber-700 ring-amber-200">Pendiente de confirmar</Badge>;
    return <Badge className="bg-slate-100 text-slate-600 ring-slate-200">Pendiente</Badge>;
  };

  return (
    <div className="p-8 max-w-4xl">
      <Header icon={<ClipboardList size={20} className="text-[#2E8B57]" />} title="Tareas" manualKey="tareas" subtitle={`${asignadas.filter((t) => t.estado === "Pendiente").length} pendiente(s) para ti`} />

      <div className="flex gap-2 mb-5">
        <button onClick={() => setFiltro("asignadas")} className={`px-4 py-2 rounded-md text-sm font-semibold ${filtro === "asignadas" ? "bg-[#2E8B57] text-white" : "bg-white border border-slate-300 text-slate-600"}`}>
          Asignadas a mí ({asignadas.length})
        </button>
        <button onClick={() => setFiltro("creadas")} className={`px-4 py-2 rounded-md text-sm font-semibold ${filtro === "creadas" ? "bg-[#2E8B57] text-white" : "bg-white border border-slate-300 text-slate-600"}`}>
          Creadas por mí ({creadas.length})
        </button>
      </div>

      <div className="space-y-3">
        {lista.length === 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-8 text-center text-slate-400 text-sm">
            {filtro === "asignadas" ? "No tienes tareas asignadas." : "No has creado ninguna tarea todavía. Puedes crear una desde cualquier mensaje del Chat."}
          </div>
        )}
        {lista.map((t) => (
          <div key={t.id} className="bg-white border border-slate-200 rounded-lg p-4">
            <div className="flex items-start justify-between gap-3 mb-1.5">
              <h3 className="font-semibold text-slate-800 text-sm">{t.titulo}</h3>
              {badgeEstado(t)}
            </div>
            {t.descripcion && <p className="text-sm text-slate-600 mb-2 whitespace-pre-wrap">{t.descripcion}</p>}
            <div className="text-xs text-slate-400 flex flex-wrap gap-x-3 gap-y-1 mb-3">
              {filtro === "asignadas" ? <span>De: {t.asignadoPorNombre}</span> : <span>Para: {t.asignadoANombre}</span>}
              {t.salaNombre && <span>Sala: {t.salaNombre}</span>}
              {t.fechaLimite && <span>Fecha límite: {fmtDate(t.fechaLimite)}</span>}
              <span>Creada: {new Date(t.fechaCreacion).toLocaleDateString("es-ES")}</span>
            </div>
            <div className="flex gap-2">
              {filtro === "asignadas" && t.estado === "Pendiente" && (
                <button onClick={() => onMarcarHecha(t.id)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold px-3 py-1.5 rounded-md">
                  Marcar como hecha
                </button>
              )}
              {filtro === "creadas" && t.estado === "Pendiente de confirmar" && (
                <>
                  <button onClick={() => onConfirmar(t.id, true)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-xs font-semibold px-3 py-1.5 rounded-md">
                    Confirmar hecha
                  </button>
                  <button onClick={() => onConfirmar(t.id, false)} className="text-xs font-semibold text-amber-700 border border-amber-300 px-3 py-1.5 rounded-md">
                    No está hecha, devolver
                  </button>
                </>
              )}
              {filtro === "creadas" && (
                <button onClick={() => { if (window.confirm("¿Borrar esta tarea?")) onDelete(t.id); }} className="text-xs font-semibold text-rose-600 border border-rose-200 px-3 py-1.5 rounded-md ml-auto">
                  Borrar
                </button>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ================= ARCHIVOS DE EMPRESA ================= */

const CATEGORIAS_ARCHIVO_EMPRESA = ["Tarifas", "Contactos", "Plantillas", "Manuales", "Otros"];

function ArchivosModulo({ archivos, onSubir, onDelete }) {
  const [categoria, setCategoria] = useState("Tarifas");
  const [filtro, setFiltro] = useState("todas");
  const [q, setQ] = useState("");
  const inputRef = useRef(null);

  const visibles = archivos.filter((a) => {
    if (filtro !== "todas" && (a.categoria || "Otros") !== filtro) return false;
    if (q && !a.nombre.toLowerCase().includes(q.toLowerCase())) return false;
    return true;
  });

  const grupos = {};
  visibles.forEach((a) => {
    const cat = a.categoria || "Otros";
    if (!grupos[cat]) grupos[cat] = [];
    grupos[cat].push(a);
  });

  return (
    <div className="p-8 max-w-4xl">
      <Header
        icon={<FileText size={20} className="text-[#2E8B57]" />}
        title="Archivos"
        manualKey="archivos"
        subtitle={`${archivos.length} archivo(s) guardado(s) — tarifas, contactos y otros documentos de referencia`}
      />

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-5 flex flex-wrap gap-2 items-center">
        <Select value={categoria} onChange={(e) => setCategoria(e.target.value)} className="max-w-[160px]">
          {CATEGORIAS_ARCHIVO_EMPRESA.map((c) => <option key={c} value={c}>{c}</option>)}
        </Select>
        <input
          ref={inputRef}
          type="file"
          onChange={(e) => { onSubir(e.target.files[0], categoria); e.target.value = ""; }}
          className="hidden"
        />
        <button onClick={() => inputRef.current?.click()} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="px-4 py-2 rounded-md text-sm font-semibold">
          Subir archivo
        </button>
        <p className="text-xs text-slate-400">Excel, PDF, Word, imágenes... cualquier tipo de archivo, hasta 8 MB.</p>
      </div>

      <div className="flex flex-wrap gap-2 mb-4">
        <button onClick={() => setFiltro("todas")} className={`px-3 py-1.5 rounded-md text-xs font-semibold ${filtro === "todas" ? "bg-[#2E8B57] text-white" : "bg-white border border-slate-300 text-slate-600"}`}>Todas</button>
        {CATEGORIAS_ARCHIVO_EMPRESA.map((c) => (
          <button key={c} onClick={() => setFiltro(c)} className={`px-3 py-1.5 rounded-md text-xs font-semibold ${filtro === c ? "bg-[#2E8B57] text-white" : "bg-white border border-slate-300 text-slate-600"}`}>{c}</button>
        ))}
        <div className="relative flex-1 min-w-[180px]">
          <Search size={13} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nombre..." className="w-full pl-8 pr-3 py-1.5 rounded-md border border-slate-300 text-xs" />
        </div>
      </div>

      {Object.keys(grupos).length === 0 && (
        <div className="bg-white border border-slate-200 rounded-lg p-10 text-center text-slate-400 text-sm">No hay archivos guardados todavía.</div>
      )}

      {Object.entries(grupos).map(([cat, lista]) => (
        <div key={cat} className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-4">
          <h3 className="font-display font-bold text-slate-700 text-sm px-4 py-2.5 bg-slate-50 border-b border-slate-200">{cat}</h3>
          <div className="divide-y divide-slate-100">
            {lista.map((a) => (
              <div key={a.id} className="flex items-center justify-between gap-3 px-4 py-2.5">
                <a href={a.url} download={a.nombre} className="text-sm text-emerald-700 underline truncate">{a.nombre}</a>
                <div className="flex items-center gap-3 shrink-0">
                  <span className="text-xs text-slate-400">{new Date(a.subidoEn).toLocaleDateString("es-ES")}{a.subidoPor ? ` · ${a.subidoPor}` : ""}</span>
                  <button onClick={() => { if (window.confirm("¿Borrar este archivo?")) onDelete(a.id); }} className="text-xs font-semibold text-rose-600 hover:underline">Borrar</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
    </div>
  );
}

/* ================= INCIDENCIAS ================= */

function IncidenciasModulo({ incidencias, proyectos, clientes, pedidos, proveedores, openPedido, onPedirMateriales, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onInlineUpdate, nextNumero, onImportarMasivo, isAdmin }) {
  const [q, setQ] = useState("");
  const [estadoIncidencia, setEstadoIncidencia] = useState("");
  const [estadoTrabajo, setEstadoTrabajo] = useState("");
  const inputIncidenciasRef = useRef(null);

  const manejarImportarIncidencias = async (file) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const filasExcel = leerFilasExcel(buffer);
    const filas = filasExcel.map((fila) => ({
      clienteTexto: String(valorPorCabeceras(fila, ["clientereferencia", "cliente", "referencia", "nombre"])).trim(),
      telefono: String(valorPorCabeceras(fila, ["telefono", "tel", "movil"])).trim(),
      tareas: String(valorPorCabeceras(fila, ["tareasdereparacion", "tareas", "reparacion", "descripcion"])).trim(),
      materiales: String(valorPorCabeceras(fila, ["materialesapedir", "materiales"])).trim(),
      estado: String(valorPorCabeceras(fila, ["estado"])).trim(),
      dudas: String(valorPorCabeceras(fila, ["dudasaconfirmarconmiguel", "dudas"])).trim(),
      notas: String(valorPorCabeceras(fila, ["notas", "observaciones"])).trim(),
    })).filter((f) => f.clienteTexto);
    if (filas.length === 0) {
      alert("No se ha encontrado ninguna fila con cliente/referencia. Revisa las cabeceras del Excel.");
      return;
    }
    onImportarMasivo(filas);
  };

  const proyecto = (id) => proyectos.find((p) => p.id === id);
  const clienteNombre = (proyectoId) => {
    const p = proyecto(proyectoId);
    return p ? clientes.find((c) => c.id === p.clienteId)?.nombre : null;
  };

  const filtered = useMemo(() => {
    return incidencias.filter((i) => {
      if (estadoIncidencia && i.estadoIncidencia !== estadoIncidencia) return false;
      if (estadoTrabajo && i.estadoTrabajo !== estadoTrabajo) return false;
      if (!q) return true;
      const p = proyecto(i.proyectoId);
      const s = `${i.numero} ${p?.numero || ""} ${p?.nombre || ""} ${clienteNombre(i.proyectoId) || ""} ${i.especificaciones}`.toLowerCase();
      return s.includes(q.toLowerCase());
    });
  }, [incidencias, q, estadoIncidencia, estadoTrabajo, proyectos, clientes]);

  if (view === "form") {
    const editing = incidencias.find((i) => i.id === editId) || null;
    return (
      <IncidenciaForm
        initial={editing}
        proyectos={proyectos}
        clientes={clientes}
        nextNumero={nextNumero}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={onUpsert}
      />
    );
  }

  if (view === "detail") {
    const incidencia = incidencias.find((i) => i.id === detailId);
    if (!incidencia) { setView("list"); return null; }
    return (
      <IncidenciaDetail
        incidencia={incidencia}
        proyecto={proyecto(incidencia.proyectoId)}
        cliente={clientes.find((c) => c.id === proyecto(incidencia.proyectoId)?.clienteId)}
        pedidos={pedidos.filter((pd) => pd.proyectoId === incidencia.proyectoId)}
        proveedores={proveedores}
        openPedido={openPedido}
        onPedirMateriales={() => onPedirMateriales(incidencia)}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(incidencia.id); setView("form"); }}
        onDelete={() => onDelete(incidencia.id)}
        isAdmin={isAdmin}
        onInlineUpdate={onInlineUpdate}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<AlertOctagon size={20} className="text-[#2E8B57]" />}
        title="Incidencias"
        manualKey="incidencias"
        subtitle={`${incidencias.length} incidencia${incidencias.length === 1 ? "" : "s"} registrada${incidencias.length === 1 ? "" : "s"}`}
      />

      {proyectos.length === 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
          ⚠ No puedes crear una incidencia todavía: primero da de alta al menos un proyecto en la pestaña Proyectos.
        </div>
      )}
      <button
        onClick={() => { setEditId(null); setView("form"); }}
        disabled={proyectos.length === 0}
        title={proyectos.length === 0 ? "Necesitas un proyecto dado de alta" : ""}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVA INCIDENCIA
      </button>

      <input
        ref={inputIncidenciasRef}
        type="file"
        accept=".xlsx,.xls,.ods,.csv"
        className="hidden"
        onChange={(e) => { manejarImportarIncidencias(e.target.files[0]); e.target.value = ""; }}
      />
      <button
        onClick={() => inputIncidenciasRef.current?.click()}
        className="w-full flex items-center justify-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 py-2.5 rounded-md mb-6 hover:bg-slate-50"
      >
        Importar incidencias desde Excel (masivo) — ej: hoja de control de reparaciones
      </button>

      {proyectos.length === 0 && (
        <div className="mb-4 px-4 py-3 rounded-md bg-amber-50 text-amber-800 text-sm border border-amber-200">
          Necesitas al menos un proyecto dado de alta para registrar una incidencia.
        </div>
      )}

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 min-w-[220px] max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nº, proyecto, cliente, descripción…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
        </div>
        <Select value={estadoTrabajo} onChange={(e) => setEstadoTrabajo(e.target.value)} className="max-w-[190px]">
          <option value="">Estado del trabajo</option>
          {ESTADO_TRABAJO_INCIDENCIA.map((t) => <option key={t}>{t}</option>)}
        </Select>
        <Select value={estadoIncidencia} onChange={(e) => setEstadoIncidencia(e.target.value)} className="max-w-[190px]">
          <option value="">Estado de la incidencia</option>
          {ESTADO_INCIDENCIA.map((t) => <option key={t}>{t}</option>)}
        </Select>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nº Incidencia</th>
              <th className="px-4 py-3 font-semibold">Proyecto</th>
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Fecha</th>
              <th className="px-4 py-3 font-semibold">Estado trabajo</th>
              <th className="px-4 py-3 font-semibold">Estado incidencia</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay incidencias que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((i) => {
              const p = proyecto(i.proyectoId);
              return (
                <tr key={i.id} onClick={() => { setDetailId(i.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                  <td className="px-4 py-3 font-mono-num text-slate-500">#{i.numero}</td>
                  <td className="px-4 py-3 font-medium text-slate-800">{p ? `#${p.numero} — ${p.nombre}` : "Proyecto eliminado"}</td>
                  <td className="px-4 py-3 text-slate-600">{clienteNombre(i.proyectoId) || "—"}</td>
                  <td className="px-4 py-3 text-slate-500">{fmtDate(i.fecha)}</td>
                  <td className="px-4 py-3"><Badge className={ESTADO_TRABAJO_INCIDENCIA_STYLE[i.estadoTrabajo]}>{i.estadoTrabajo}</Badge></td>
                  <td className="px-4 py-3"><Badge className={ESTADO_INCIDENCIA_STYLE[i.estadoIncidencia]}>{i.estadoIncidencia}</Badge></td>
                  <td className="px-4 py-3 text-right">
                    {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar la incidencia #${i.numero}?`)) onDelete(i.id); }} className="text-slate-300 hover:text-rose-500 transition">
                      <Trash2 size={15} />
                    </button>)}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function IncidenciaForm({ initial, proyectos, clientes, nextNumero, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, proyectoId: proyectos[0]?.id || "", fecha: new Date().toISOString().slice(0, 10),
      especificaciones: "", observaciones: "", responsableInicial: "", comercialAsociado: "",
      responsableActual: "", estadoTrabajo: "Pendiente revisión", estadoIncidencia: "Pendiente revisión",
      fechaEntregaPrevista: "", fechaEntregado: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const proyectoSel = proyectos.find((p) => p.id === f.proyectoId);
  const clienteSel = proyectoSel ? clientes.find((c) => c.id === proyectoSel.clienteId) : null;

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.proyectoId) {
      setErrorMsg("Falta seleccionar el proyecto asociado (es obligatorio).");
      return;
    }
    if (!f.especificaciones.trim()) {
      setErrorMsg("Falta el campo Especificaciones de la incidencia (es obligatorio).");
      return;
    }
    setErrorMsg("");
    onSave(f);
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <div className="flex items-center gap-3 mb-6">
        <h1 className="font-display text-2xl font-extrabold text-slate-900">{initial ? `Editar incidencia #${initial.numero}` : "Alta de incidencia"}</h1>
        {!initial && <Badge className="bg-slate-100 text-slate-500 ring-slate-200 font-mono-num">Nº {nextNumero()} (automático)</Badge>}
      </div>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <Field label="Proyecto asociado" required>
          <Select value={f.proyectoId} onChange={set("proyectoId")} required>
            {proyectos.map((p) => <option key={p.id} value={p.id}>#{p.numero} — {p.nombre}</option>)}
          </Select>
        </Field>

        {proyectoSel && (
          <div className="grid grid-cols-3 gap-4 text-sm bg-slate-50 border border-slate-200 rounded-md p-3">
            <InfoRow icon={<Users size={13} />} label="Cliente" value={clienteSel?.nombre || "—"} />
            <InfoRow icon={<Euro size={13} />} label="Importe presupuesto" value={money(proyectoSel.importePresupuesto)} />
            <InfoRow icon={<MapPin size={13} />} label="Dirección" value={proyectoSel.ubicacion || "—"} />
          </div>
        )}

        <div className="grid grid-cols-3 gap-4">
          <Field label="Fecha"><TextInput type="date" value={f.fecha} onChange={set("fecha")} /></Field>
          <Field label="Responsable inicial del proyecto"><TextInput value={f.responsableInicial} onChange={set("responsableInicial")} /></Field>
          <Field label="Responsable actual"><TextInput value={f.responsableActual} onChange={set("responsableActual")} /></Field>
        </div>
        <Field label="Comercial asociado"><TextInput value={f.comercialAsociado} onChange={set("comercialAsociado")} /></Field>

        <Field label="Especificaciones de la incidencia" required>
          <TextArea rows={3} value={f.especificaciones} onChange={set("especificaciones")} required />
        </Field>
        <Field label="Observaciones"><TextArea rows={2} value={f.observaciones} onChange={set("observaciones")} /></Field>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Estado del trabajo">
            <Select value={f.estadoTrabajo} onChange={set("estadoTrabajo")}>
              {ESTADO_TRABAJO_INCIDENCIA.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Estado de la incidencia">
            <Select value={f.estadoIncidencia} onChange={set("estadoIncidencia")}>
              {ESTADO_INCIDENCIA.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Fecha entrega prevista"><TextInput type="date" value={f.fechaEntregaPrevista} onChange={set("fechaEntregaPrevista")} /></Field>
          <Field label="Fecha entregado"><TextInput type="date" value={f.fechaEntregado} onChange={set("fechaEntregado")} /></Field>
        </div>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar incidencia</button>
        </div>
      </form>
    </div>
  );
}

function IncidenciaDetail({ incidencia, proyecto, cliente, pedidos, proveedores, openPedido, onPedirMateriales, onBack, onEdit, onDelete, onInlineUpdate, isAdmin }) {
  const [tab, setTab] = useState("datos");
  const gastos = incidencia.gastos || [];
  const totalGastos = gastos.reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
  const checklistIncidencia = normalizarChecklist(proyecto?.checklistMateriales);
  const checklistIncompleto = checklistIncidencia.some((c) => !c.estado);
  const [gForm, setGForm] = useState({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  const archivos = incidencia.archivos || [];
  const [subiendoArchivo, setSubiendoArchivo] = useState(false);
  const [errorArchivo, setErrorArchivo] = useState("");
  const inputArchivoRef = useRef(null);

  // Reduce el tamaño de una foto antes de guardarla (máx. 1000px de lado, calidad 0.7)
  // para no llenar el límite de guardado compartido de Incidencias.
  const comprimirImagen = (file) => new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const maxLado = 1000;
        let { width, height } = img;
        if (width > height && width > maxLado) { height = Math.round(height * (maxLado / width)); width = maxLado; }
        else if (height > maxLado) { width = Math.round(width * (maxLado / height)); height = maxLado; }
        const canvas = document.createElement("canvas");
        canvas.width = width; canvas.height = height;
        canvas.getContext("2d").drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", 0.7));
      };
      img.onerror = reject;
      img.src = e.target.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });

  const subirArchivo = async (file) => {
    setSubiendoArchivo(true);
    setErrorArchivo("");
    try {
      const esImagen = file.type.startsWith("image/");
      let dataUrl;
      if (esImagen) {
        dataUrl = await comprimirImagen(file);
      } else {
        dataUrl = await new Promise((resolve, reject) => {
          const r = new FileReader();
          r.onload = () => resolve(r.result);
          r.onerror = reject;
          r.readAsDataURL(file);
        });
      }
      if (dataUrl.length > 900000) {
        setErrorArchivo("Este archivo sigue siendo demasiado grande incluso comprimido. Prueba con una foto más sencilla, o un PDF más corto.");
        setSubiendoArchivo(false);
        return;
      }
      const nuevo = { id: uid(), nombre: file.name, tipo: esImagen ? "imagen" : "documento", dataUrl, fecha: new Date().toISOString() };
      onInlineUpdate(incidencia.id, { archivos: [nuevo, ...archivos] });
    } catch (err) {
      setErrorArchivo("No se pudo subir el archivo. Prueba de nuevo.");
    } finally {
      setSubiendoArchivo(false);
    }
  };

  const eliminarArchivo = (id) => onInlineUpdate(incidencia.id, { archivos: archivos.filter((a) => a.id !== id) });

  const addGasto = (e) => {
    e.preventDefault();
    if (!gForm.proveedor.trim()) return;
    const next = [...gastos, { ...gForm, id: uid(), importe: parseFloat(gForm.importe) || 0 }];
    onInlineUpdate(incidencia.id, { gastos: next });
    setGForm({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  };
  const removeGasto = (id) => onInlineUpdate(incidencia.id, { gastos: gastos.filter((g) => g.id !== id) });

  const enlaceEmailCliente = () => {
    if (!cliente?.email) return null;
    const asunto = `Incidencia #${incidencia.numero} — ALUMAVEL`;
    const cuerpo = `Hola ${cliente.nombre},\n\nLe escribimos respecto a la incidencia registrada${proyecto ? ` sobre el proyecto #${proyecto.numero} — ${proyecto.nombre}` : ""}.\n\n${incidencia.descripcion || ""}\n\nQuedamos a su disposición para cualquier duda.\n\nUn saludo,\nALUMAVEL`;
    return `mailto:${cliente.email}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  };

  const enlaceWhatsappCliente = () => {
    const tel = (cliente?.movil || "").replace(/[^\d+]/g, "");
    if (!tel) return null;
    const telConPrefijo = tel.startsWith("+") ? tel.replace("+", "") : (tel.startsWith("34") ? tel : `34${tel}`);
    const mensaje = `Hola ${cliente?.nombre || ""}, le escribimos de ALUMAVEL respecto a la incidencia #${incidencia.numero}${proyecto ? ` del proyecto #${proyecto.numero} — ${proyecto.nombre}` : ""}. Quedamos a su disposición para cualquier duda. Un saludo.`;
    return `https://wa.me/${telConPrefijo}?text=${encodeURIComponent(mensaje)}`;
  };

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">Incidencia #{incidencia.numero}</span>
          </div>
          <h1 className="font-display text-2xl font-extrabold text-slate-900 mt-0.5">
            {proyecto ? `#${proyecto.numero} — ${proyecto.nombre}` : "Proyecto eliminado"}
          </h1>
          <p className="text-sm text-slate-500 mt-1">{cliente?.nombre || "—"} · {fmtDate(incidencia.fecha)}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          {enlaceWhatsappCliente() ? (
            <a href={enlaceWhatsappCliente()} target="_blank" rel="noopener noreferrer" className="flex items-center gap-1.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 px-3.5 py-2 rounded-md">
              <MessageCircle size={14} /> WhatsApp
            </a>
          ) : (
            <span title="Este cliente no tiene móvil guardado" className="flex items-center gap-1.5 text-sm font-semibold text-slate-300 border border-slate-200 px-3.5 py-2 rounded-md cursor-not-allowed">
              <MessageCircle size={14} /> Sin móvil
            </span>
          )}
          {enlaceEmailCliente() ? (
            <a href={enlaceEmailCliente()} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 px-3.5 py-2 rounded-md">
              <Mail size={14} /> Email
            </a>
          ) : (
            <span title="Este cliente no tiene email guardado" className="flex items-center gap-1.5 text-sm font-semibold text-slate-300 border border-slate-200 px-3.5 py-2 rounded-md cursor-not-allowed">
              <Mail size={14} /> Sin email
            </span>
          )}
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      <div className="flex flex-wrap gap-2 mb-6 mt-3">
        <Badge className={ESTADO_TRABAJO_INCIDENCIA_STYLE[incidencia.estadoTrabajo]}>Trabajo: {incidencia.estadoTrabajo}</Badge>
        <Badge className={ESTADO_INCIDENCIA_STYLE[incidencia.estadoIncidencia]}>Incidencia: {incidencia.estadoIncidencia}</Badge>
      </div>

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        {[
          { id: "datos", label: "Datos", icon: FileText },
          { id: "archivos", label: `Fotos y documentos (${archivos.length})`, icon: ImageIcon },
          { id: "pedidos", label: `Pedidos de materiales (${(pedidos || []).length})`, icon: ClipboardList },
          { id: "gastos", label: `Gastos asociados (${gastos.length})`, icon: Receipt },
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "datos" && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
            <InfoRow icon={<Euro size={14} />} label="Importe presupuesto" value={money(proyecto?.importePresupuesto)} />
            <InfoRow icon={<MapPin size={14} />} label="Dirección" value={proyecto?.ubicacion || "—"} />
            <InfoRow icon={<Users size={14} />} label="Responsable inicial" value={incidencia.responsableInicial || "—"} />
            <InfoRow icon={<Users size={14} />} label="Responsable actual" value={incidencia.responsableActual || "—"} />
            <InfoRow icon={<Users size={14} />} label="Comercial asociado" value={incidencia.comercialAsociado || "—"} />
            <InfoRow icon={<FileText size={14} />} label="Entrega prevista / entregado" value={`${fmtDate(incidencia.fechaEntregaPrevista)} / ${fmtDate(incidencia.fechaEntregado)}`} />
          </div>
          <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Especificaciones de la incidencia</span>
            {incidencia.especificaciones}
          </div>
          {incidencia.observaciones && (
            <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Observaciones</span>
              {incidencia.observaciones}
            </div>
          )}
        </CornerFrame>
      )}

      {tab === "archivos" && (
        <div>
          <button
            type="button"
            onClick={() => inputArchivoRef.current?.click()}
            disabled={subiendoArchivo}
            style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
            className="flex items-center gap-2 text-sm font-semibold hover:opacity-90 disabled:opacity-50 px-4 py-2.5 rounded-md mb-3"
          >
            <ImageIcon size={15} /> {subiendoArchivo ? "Subiendo..." : "Subir foto o documento"}
          </button>
          <input
            ref={inputArchivoRef}
            type="file"
            accept="image/*,application/pdf"
            className="hidden"
            onChange={(e) => { if (e.target.files?.[0]) subirArchivo(e.target.files[0]); e.target.value = ""; }}
          />
          {errorArchivo && (
            <p className="text-xs text-rose-600 font-semibold mb-3">⚠ {errorArchivo}</p>
          )}
          <p className="text-xs text-slate-400 mb-4">Las fotos se comprimen automáticamente al subirlas para no ocupar demasiado espacio.</p>

          {archivos.length === 0 ? (
            <div className="bg-white border border-slate-200 rounded-lg px-4 py-10 text-center text-slate-400 text-sm">
              Todavía no hay fotos ni documentos en esta incidencia.
            </div>
          ) : (
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {archivos.map((a) => (
                <div key={a.id} className="bg-white border border-slate-200 rounded-lg overflow-hidden">
                  {a.tipo === "imagen" ? (
                    <a href={a.dataUrl} target="_blank" rel="noopener noreferrer">
                      <img src={a.dataUrl} alt={a.nombre} className="w-full h-28 object-cover" />
                    </a>
                  ) : (
                    <a href={a.dataUrl} target="_blank" rel="noopener noreferrer" download={a.nombre} className="flex items-center justify-center h-28 bg-slate-50">
                      <FileText size={28} className="text-slate-400" />
                    </a>
                  )}
                  <div className="p-2">
                    <p className="text-xs text-slate-600 truncate" title={a.nombre}>{a.nombre}</p>
                    <div className="flex items-center justify-between mt-1">
                      <span className="text-[10px] text-slate-400">{fmtDate(a.fecha?.slice(0, 10))}</span>
                      <button onClick={() => eliminarArchivo(a.id)} className="text-slate-300 hover:text-rose-500"><X size={12} /></button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {tab === "pedidos" && (
        <div className="space-y-4">
          <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm">
            Pedidos de materiales del proyecto de esta incidencia (cristales, persianas, etc.). Puedes pedir uno nuevo, incluyendo medidas pegadas desde Excel.
          </div>
          {checklistIncompleto && (
            <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold">
              ⚠ El checklist "Qué lleva la obra" del proyecto no está completo todavía. No podrás crear un pedido nuevo hasta rellenarlo en la ficha del proyecto.
            </div>
          )}
          <button
            onClick={onPedirMateriales}
            disabled={!proyecto || checklistIncompleto}
            title={!proyecto ? "Esta incidencia no tiene proyecto asociado" : checklistIncompleto ? "Completa antes el checklist \"Qué lleva la obra\" del proyecto" : ""}
            className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg shadow-md cursor-pointer select-none"
          >
            <Plus size={22} /> PEDIR MATERIALES (CRISTAL, ETC.)
          </button>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {(pedidos || []).length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin pedidos de materiales vinculados a este proyecto todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Nº Pedido</th>
                    <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                    <th className="px-4 py-2.5 font-semibold">Líneas</th>
                    <th className="px-4 py-2.5 font-semibold">Entrega prevista</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {(pedidos || []).map((p) => (
                    <tr key={p.id} onClick={() => openPedido(p.id)} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                      <td className="px-4 py-2.5 font-mono-num text-slate-500">#{p.numero}</td>
                      <td className="px-4 py-2.5 font-medium text-slate-800">{proveedores.find((pr) => pr.id === p.proveedorId)?.nombre || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{(p.lineas || []).length} línea{(p.lineas || []).length === 1 ? "" : "s"}</td>
                      <td className="px-4 py-2.5 text-slate-500">{fmtDate(p.fechaEntregaPrevista)}</td>
                      <td className="px-4 py-2.5"><Badge className={ESTADO_PEDIDO_STYLE[p.estado]}>{p.estado}</Badge></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {tab === "gastos" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {gastos.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin gastos registrados todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                    <th className="px-4 py-2.5 font-semibold">Producto / Tarea</th>
                    <th className="px-4 py-2.5 font-semibold">Factura</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {gastos.map((g) => (
                    <tr key={g.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-medium text-slate-800">{g.proveedor}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.producto || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.facturaAsociada || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.estadoFactura}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money(g.importe)}</td>
                      <td className="px-4 py-2.5 text-right"><button onClick={() => removeGasto(g.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button></td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5" colSpan={4}>Total gastos asociados</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(totalGastos)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>
          <form onSubmit={addGasto} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-6 gap-2 items-end">
            <Field label="Proveedor"><TextInput value={gForm.proveedor} onChange={(e) => setGForm({ ...gForm, proveedor: e.target.value })} /></Field>
            <Field label="Producto / Tarea"><TextInput value={gForm.producto} onChange={(e) => setGForm({ ...gForm, producto: e.target.value })} /></Field>
            <Field label="Factura asociada"><TextInput value={gForm.facturaAsociada} onChange={(e) => setGForm({ ...gForm, facturaAsociada: e.target.value })} /></Field>
            <Field label="Estado factura">
              <Select value={gForm.estadoFactura} onChange={(e) => setGForm({ ...gForm, estadoFactura: e.target.value })}>
                <option>Pendiente</option><option>Pagada</option>
              </Select>
            </Field>
            <Field label="Importe (€)"><TextInput type="number" step="0.01" value={gForm.importe} onChange={(e) => setGForm({ ...gForm, importe: e.target.value })} /></Field>
            <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Añadir</button>
          </form>
        </div>
      )}
    </div>
  );
}

/* ================= CALENDARIO ================= */

const DIAS_SEMANA = ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"];
const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

function toKey(d) {
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return null;
  return d;
}

function buildMonthMatrix(year, month) {
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7; // lunes=0
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < startOffset; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(new Date(year, month, d));
  while (cells.length % 7 !== 0) cells.push(null);
  const weeks = [];
  for (let i = 0; i < cells.length; i += 7) weeks.push(cells.slice(i, i + 7));
  return weeks;
}

const EVENT_DOT = {
  proyecto_entrega: (estado) => ESTADO_TRABAJO_STYLE[estado] || "bg-slate-100 text-slate-600 ring-slate-200",
  proyecto_solicitud: () => "bg-rose-50 text-rose-700 ring-rose-200",
  proyecto_fabricacion: () => "bg-orange-50 text-orange-700 ring-orange-200",
  proyecto_montaje: () => "bg-fuchsia-50 text-fuchsia-700 ring-fuchsia-200",
  pedido: () => "bg-sky-50 text-sky-700 ring-sky-200",
  incidencia: () => "bg-violet-50 text-violet-700 ring-violet-200",
};

const CALENDARIO_TABS = [
  { id: "todo", label: "Todo" },
  { id: "proyecto_fabricacion", label: "Fabricación" },
  { id: "proyecto_montaje", label: "Montaje" },
  { id: "proyecto_entrega", label: "Entregas" },
  { id: "pedido", label: "Pedidos" },
  { id: "incidencia", label: "Incidencias" },
];

function CalendarioModulo({ proyectos, clientes, pedidos, incidencias, openProyecto, openPedido, openIncidencia, onCambiarFechaProyecto, onCambiarFechaPedido, onCambiarFechaIncidencia }) {
  const today = new Date();
  const [cursor, setCursor] = useState(new Date(today.getFullYear(), today.getMonth(), 1));
  const [selectedDay, setSelectedDay] = useState(toKey(today.toISOString().slice(0, 10)));
  const [filtroEstado, setFiltroEstado] = useState("");
  const [tab, setTab] = useState("todo");
  const [editandoFechaId, setEditandoFechaId] = useState(null);

  const clienteNombre = (proyectoId) => {
    const p = proyectos.find((pr) => pr.id === proyectoId);
    return p ? clientes.find((c) => c.id === p.clienteId)?.nombre : null;
  };

  const eventsByDay = useMemo(() => {
    const map = {};
    const push = (dateStr, ev) => {
      const key = toKey(dateStr);
      if (!key) return;
      if (tab !== "todo" && ev.type !== tab) return;
      if (!map[key]) map[key] = [];
      map[key].push(ev);
    };

    proyectos.forEach((p) => {
      if (filtroEstado && p.estadoTrabajo !== filtroEstado) return;
      if (p.fechaSolicitud) {
        push(p.fechaSolicitud, {
          type: "proyecto_solicitud", id: p.id, campo: "fechaSolicitud",
          label: `Presupuesto enviado — #${p.numero} ${p.nombre}`,
          badge: EVENT_DOT.proyecto_solicitud(),
        });
      }
      if (p.fechaFabricacion) {
        push(p.fechaFabricacion, {
          type: "proyecto_fabricacion", id: p.id, campo: "fechaFabricacion",
          label: `Fabricación — #${p.numero} ${p.nombre}`,
          badge: EVENT_DOT.proyecto_fabricacion(),
        });
      }
      if (p.fechaMontaje) {
        push(p.fechaMontaje, {
          type: "proyecto_montaje", id: p.id, campo: "fechaMontaje",
          label: `Montaje — #${p.numero} ${p.nombre}`,
          badge: EVENT_DOT.proyecto_montaje(),
        });
      }
      if (p.fechaEntregaPrevista) {
        push(p.fechaEntregaPrevista, {
          type: "proyecto_entrega", id: p.id, campo: "fechaEntregaPrevista",
          label: `#${p.numero} ${p.nombre} · ${p.estadoTrabajo}`,
          badge: EVENT_DOT.proyecto_entrega(p.estadoTrabajo),
        });
      }
    });

    pedidos.forEach((pd) => {
      if (pd.estado === "Recibido" || pd.estado === "Cancelado") return;
      if (pd.fechaEntregaPrevista) {
        push(pd.fechaEntregaPrevista, {
          type: "pedido", id: pd.id, campo: "fechaEntregaPrevista",
          label: `Pedido #${pd.numero} — entrega prevista`,
          badge: EVENT_DOT.pedido(),
        });
      }
    });

    incidencias.forEach((inc) => {
      if (inc.fechaEntregaPrevista) {
        const p = proyectos.find((pr) => pr.id === inc.proyectoId);
        push(inc.fechaEntregaPrevista, {
          type: "incidencia", id: inc.id, campo: "fechaEntregaPrevista",
          label: `Incidencia #${inc.numero}${p ? ` — #${p.numero} ${p.nombre}` : ""}`,
          badge: EVENT_DOT.incidencia(),
        });
      }
    });

    return map;
  }, [proyectos, pedidos, incidencias, filtroEstado, tab]);

  const weeks = useMemo(() => buildMonthMatrix(cursor.getFullYear(), cursor.getMonth()), [cursor]);
  const todayKey = today.toISOString().slice(0, 10);
  const selectedEvents = eventsByDay[selectedDay] || [];

  const openEvent = (ev) => {
    if (ev.type.startsWith("proyecto_")) openProyecto(ev.id);
    else if (ev.type === "pedido") openPedido(ev.id);
    else if (ev.type === "incidencia") openIncidencia(ev.id);
  };

  const cambiarFecha = (ev, nuevaFecha) => {
    if (ev.type.startsWith("proyecto_")) onCambiarFechaProyecto(ev.id, ev.campo, nuevaFecha);
    else if (ev.type === "pedido") onCambiarFechaPedido(ev.id, ev.campo, nuevaFecha);
    else if (ev.type === "incidencia") onCambiarFechaIncidencia(ev.id, ev.campo, nuevaFecha);
    setEditandoFechaId(null);
  };

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <Header
        icon={<CalendarDays size={20} className="text-[#2E8B57]" />}
        title="Calendario de trabajos"
        manualKey="calendario"
        subtitle="Vista mensual automática a partir de proyectos, pedidos e incidencias"
      />

      <div className="flex gap-1 mb-5 border-b border-slate-200 overflow-x-auto">
        {CALENDARIO_TABS.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition whitespace-nowrap ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            {t.label}
          </button>
        ))}
      </div>

      <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
        <div className="flex items-center gap-2">
          <button onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1))} className="p-2 rounded-md border border-slate-300 bg-white hover:bg-slate-50">
            <ChevronLeft size={16} />
          </button>
          <span className="font-display font-bold text-lg text-slate-800 w-44 text-center">{MESES[cursor.getMonth()]} {cursor.getFullYear()}</span>
          <button onClick={() => setCursor(new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1))} className="p-2 rounded-md border border-slate-300 bg-white hover:bg-slate-50">
            <ChevronRight size={16} />
          </button>
          <button
            onClick={() => { const t = new Date(); setCursor(new Date(t.getFullYear(), t.getMonth(), 1)); setSelectedDay(t.toISOString().slice(0, 10)); }}
            className="ml-1 text-sm font-semibold text-[#2E8B57] hover:text-[#256E46]"
          >
            Hoy
          </button>
        </div>
        <Select value={filtroEstado} onChange={(e) => setFiltroEstado(e.target.value)} className="max-w-[210px]">
          <option value="">Todos los estados de trabajo</option>
          {ESTADO_TRABAJO.map((t) => <option key={t}>{t}</option>)}
        </Select>
      </div>

      <div className="flex flex-wrap gap-3 mb-4 text-xs text-slate-500">
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-emerald-400" /> Entregado</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-amber-400" /> En proceso</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-rose-400" /> Pendiente aceptación / Presupuesto enviado</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-sky-400" /> Pedido pendiente</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-violet-400" /> Incidencia</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-orange-400" /> Fabricación</span>
        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-fuchsia-400" /> Montaje</span>
      </div>

      <div className="grid grid-cols-[1fr_300px] gap-5">
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="grid grid-cols-7 bg-slate-50 border-b border-slate-200 text-[11px] uppercase tracking-wide text-slate-500 font-semibold">
            {DIAS_SEMANA.map((d) => <div key={d} className="px-2 py-2 text-center">{d}</div>)}
          </div>
          {weeks.map((week, wi) => (
            <div key={wi} className="grid grid-cols-7 border-b border-slate-100 last:border-0">
              {week.map((day, di) => {
                if (!day) return <div key={di} className="min-h-[92px] border-r border-slate-100 last:border-0 bg-slate-50/40" />;
                const key = day.toISOString().slice(0, 10);
                const evs = eventsByDay[key] || [];
                const isToday = key === todayKey;
                const isSelected = key === selectedDay;
                return (
                  <button
                    key={di}
                    onClick={() => setSelectedDay(key)}
                    className={`min-h-[92px] border-r border-slate-100 last:border-0 p-1.5 text-left align-top flex flex-col gap-1 transition ${
                      isSelected ? "bg-[#EAF1F8]" : "hover:bg-slate-50"
                    }`}
                  >
                    <span className={`inline-flex items-center justify-center w-5 h-5 rounded-full text-xs font-mono-num ${isToday ? "bg-[#2E8B57] text-white font-bold" : "text-slate-500"}`}>
                      {day.getDate()}
                    </span>
                    <div className="flex flex-col gap-0.5">
                      {evs.slice(0, 2).map((ev, i) => (
                        <span key={i} className={`text-[10px] px-1 py-0.5 rounded truncate ring-1 ${ev.badge}`} title={ev.label}>
                          {ev.label}
                        </span>
                      ))}
                      {evs.length > 2 && <span className="text-[10px] text-slate-400 px-1">+{evs.length - 2} más</span>}
                    </div>
                  </button>
                );
              })}
            </div>
          ))}
        </div>

        <div className="bg-white border border-slate-200 rounded-lg p-4 h-fit sticky top-4">
          <h3 className="font-display font-bold text-slate-800 mb-1">
            {new Date(selectedDay + "T00:00:00").toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" })}
          </h3>
          <p className="text-xs text-slate-400 mb-3">{selectedEvents.length} evento{selectedEvents.length === 1 ? "" : "s"}</p>
          {selectedEvents.length === 0 ? (
            <p className="text-sm text-slate-400 py-4 text-center">Sin eventos este día.</p>
          ) : (
            <div className="space-y-2">
              {selectedEvents.map((ev, i) => (
                <div key={i} className={`rounded-md ring-1 text-sm ${ev.badge}`}>
                  <div className="flex items-center gap-1">
                    <button onClick={() => openEvent(ev)} className="flex-1 text-left px-3 py-2.5 hover:opacity-80 transition">
                      {ev.label}
                    </button>
                    <button onClick={() => setEditandoFechaId(editandoFechaId === i ? null : i)} title="Cambiar fecha" className="px-2 py-2.5 hover:opacity-70">
                      <CalendarDays size={14} />
                    </button>
                  </div>
                  {editandoFechaId === i && (
                    <div className="flex items-center gap-2 px-3 pb-2.5">
                      <input
                        type="date"
                        defaultValue={selectedDay}
                        onChange={(e) => { if (e.target.value) cambiarFecha(ev, e.target.value); }}
                        className="text-xs border border-white/60 rounded px-2 py-1 bg-white/70 flex-1"
                      />
                    </div>
                  )}
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ================= ARTÍCULOS (Fase II) ================= */

function ArticulosModulo({ articulos, proveedores, materiales, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onInlineUpdate, nextNumero, isAdmin, onSolicitarArticulo, onImportarTarifas }) {
  const [q, setQ] = useState("");
  const inputTarifasArticulosRef = useRef(null);
  const proveedorNombre = (id) => proveedores.find((p) => p.id === id)?.nombre || "—";

  const manejarImportarTarifasArticulos = async (file) => {
    if (!file) return;
    const buffer = await file.arrayBuffer();
    const filasExcel = leerFilasExcel(buffer);
    const filas = filasExcel.map((fila) => ({
      nombre: String(valorPorCabeceras(fila, ["nombre", "articulo", "producto"])).trim(),
      precioVenta: valorPorCabeceras(fila, ["precioventa", "pventa", "pvp", "precio", "tarifa"]),
    })).filter((f) => f.nombre);
    if (filas.length === 0) {
      alert("No se ha encontrado ninguna fila con nombre. Revisa las cabeceras del Excel.");
      return;
    }
    onImportarTarifas(filas);
  };

  const precioMateriales = (a) => (a.materiales || []).reduce((s, l) => s + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0);

  const filtered = useMemo(() => {
    if (!q) return articulos;
    const s = q.toLowerCase();
    return articulos.filter((a) => `${a.numero} ${a.nombre} ${a.descripcion} ${proveedorNombre(a.proveedorId)}`.toLowerCase().includes(s));
  }, [articulos, q, proveedores]);

  if (view === "form") {
    const editing = articulos.find((a) => a.id === editId) || null;
    return (
      <ArticuloForm
        initial={editing}
        proveedores={proveedores}
        materiales={materiales}
        nextNumero={nextNumero}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={onUpsert}
      />
    );
  }

  if (view === "detail") {
    const articulo = articulos.find((a) => a.id === detailId);
    if (!articulo) { setView("list"); return null; }
    return (
      <ArticuloDetail
        articulo={articulo}
        proveedor={proveedores.find((p) => p.id === articulo.proveedorId)}
        materiales={materiales}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(articulo.id); setView("form"); }}
        onDelete={() => onDelete(articulo.id)}
        isAdmin={isAdmin}
        onInlineUpdate={onInlineUpdate}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <div className="flex items-center gap-2 mb-1">
        <Badge className="bg-violet-50 text-violet-700 ring-violet-200">Fase II</Badge>
        <span className="text-xs text-slate-400">Este módulo se está probando de forma anticipada, según el sistema actual de HETMO</span>
      </div>
      <Header
        icon={<Layers size={20} className="text-[#2E8B57]" />}
        title="Artículos"
        manualKey="articulos"
        subtitle={`${articulos.length} artículo${articulos.length === 1 ? "" : "s"} en catálogo`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-3 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO ARTÍCULO
      </button>

      <input
        ref={inputTarifasArticulosRef}
        type="file"
        accept=".xlsx,.xls,.ods,.csv"
        className="hidden"
        onChange={(e) => { manejarImportarTarifasArticulos(e.target.files[0]); e.target.value = ""; }}
      />
      <button
        onClick={() => inputTarifasArticulosRef.current?.click()}
        className="w-full flex items-center justify-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 py-2.5 rounded-md mb-6 hover:bg-slate-50"
      >
        Importar tarifas desde Excel (masivo) — actualiza precios existentes o da de alta artículos nuevos
      </button>

      <button
        onClick={() => onSolicitarArticulo()}
        style={{ borderColor: "#2E8B57", color: "#2E8B57" }}
        className="w-full flex items-center justify-center gap-2 border-2 hover:bg-slate-50 text-sm font-semibold py-3 rounded-lg mb-6 cursor-pointer select-none"
      >
        <ClipboardList size={16} /> Solicitar un artículo nuevo al encargado de pedidos
      </button>

      <div className="relative flex-1 max-w-sm mb-4">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nº, nombre, descripción, proveedor…"
          className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nº</th>
              <th className="px-4 py-3 font-semibold">Nombre</th>
              <th className="px-4 py-3 font-semibold">Descripción</th>
              <th className="px-4 py-3 font-semibold">Proveedor</th>
              <th className="px-4 py-3 font-semibold text-right">Precio materiales</th>
              <th className="px-4 py-3 font-semibold text-right">Precio venta</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay artículos que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((a) => (
              <tr key={a.id} onClick={() => { setDetailId(a.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                <td className="px-4 py-3 font-mono-num text-slate-500">{a.numero}</td>
                <td className="px-4 py-3 font-medium text-slate-800">{a.nombre}</td>
                <td className="px-4 py-3 text-slate-600 max-w-xs truncate">{a.descripcion || "—"}</td>
                <td className="px-4 py-3 text-slate-600">{proveedorNombre(a.proveedorId)}</td>
                <td className="px-4 py-3 text-right font-mono-num text-slate-500">{money(precioMateriales(a))}</td>
                <td className="px-4 py-3 text-right font-mono-num font-semibold">{money(a.precioVenta)}</td>
                <td className="px-4 py-3 text-right">
                  {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar el artículo ${a.nombre}?`)) onDelete(a.id); }} className="text-slate-300 hover:text-rose-500 transition">
                    <Trash2 size={15} />
                  </button>)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function ArticuloForm({ initial, proveedores, materiales, nextNumero, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, nombre: "", descripcion: "", proveedorId: proveedores[0]?.id || "", precioVenta: "",
      materiales: [], fases: [{ id: uid(), nombre: "", tiempoEstimado: "", materialesNecesarios: "" }],
      tamano: "", medidas: "", volumen: "", importeEnvio: "", importeMontaje: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });

  const setMaterialLinea = (id, patch) => setF({ ...f, materiales: f.materiales.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  const addMaterialLinea = () => setF({ ...f, materiales: [...f.materiales, { id: uid(), materialId: materiales[0]?.id || "", cantidad: "", precio: "" }] });
  const removeMaterialLinea = (id) => setF({ ...f, materiales: f.materiales.filter((l) => l.id !== id) });
  const pickMaterialPrecio = (id, materialId) => {
    const mat = materiales.find((m) => m.id === materialId);
    setMaterialLinea(id, { materialId, precio: mat ? mat.precioCompra : "" });
  };

  const setFase = (id, patch) => setF({ ...f, fases: f.fases.map((l) => (l.id === id ? { ...l, ...patch } : l)) });
  const addFase = () => setF({ ...f, fases: [...f.fases, { id: uid(), nombre: "", tiempoEstimado: "", materialesNecesarios: "" }] });
  const removeFase = (id) => setF({ ...f, fases: f.fases.filter((l) => l.id !== id) });

  const precioMateriales = f.materiales.reduce((s, l) => s + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0);
  const tiempoTotal = f.fases.reduce((s, l) => s + (parseFloat(l.tiempoEstimado) || 0), 0);

  const [errorMsg, setErrorMsg] = useState("");
  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.nombre.trim()) {
      setErrorMsg("Falta el campo Nombre del artículo (es obligatorio).");
      return;
    }
    setErrorMsg("");
    const mats = f.materiales.filter((l) => l.materialId && parseFloat(l.cantidad) > 0).map((l) => ({ ...l, cantidad: parseFloat(l.cantidad) || 0, precio: parseFloat(l.precio) || 0 }));
    const fases = f.fases.filter((l) => l.nombre.trim()).map((l) => ({ ...l, tiempoEstimado: parseFloat(l.tiempoEstimado) || 0 }));
    onSave({ ...f, materiales: mats, fases, precioVenta: parseFloat(f.precioVenta) || 0 });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <div className="flex items-center gap-3 mb-6">
        <h1 className="font-display text-2xl font-extrabold text-slate-900">{initial ? `Editar artículo ${initial.numero}` : "Alta de artículo"}</h1>
        {!initial && <Badge className="bg-slate-100 text-slate-500 ring-slate-200 font-mono-num">Nº {nextNumero()} (automático)</Badge>}
      </div>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-6">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Nombre" required><TextInput value={f.nombre} onChange={set("nombre")} required /></Field>
          <Field label="Proveedor">
            <Select value={f.proveedorId} onChange={set("proveedorId")}>
              <option value="">Sin proveedor específico</option>
              {proveedores.map((p) => <option key={p.id} value={p.id}>{p.nombre}</option>)}
            </Select>
          </Field>
        </div>
        <Field label="Descripción"><TextArea rows={2} value={f.descripcion} onChange={set("descripcion")} /></Field>

        {/* Materiales */}
        <div>
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-2">Materiales que lo componen</span>
          <div className="space-y-2">
            {f.materiales.map((l) => (
              <div key={l.id} className="grid grid-cols-12 gap-2 items-center">
                <div className="col-span-5">
                  <Select value={l.materialId} onChange={(e) => pickMaterialPrecio(l.id, e.target.value)}>
                    {materiales.map((m) => <option key={m.id} value={m.id}>{m.codigo} — {m.descripcion}</option>)}
                  </Select>
                </div>
                <div className="col-span-3"><TextInput type="number" placeholder="Cantidad" value={l.cantidad} onChange={(e) => setMaterialLinea(l.id, { cantidad: e.target.value })} /></div>
                <div className="col-span-3"><TextInput type="number" step="0.01" placeholder="Precio unitario (€)" value={l.precio} onChange={(e) => setMaterialLinea(l.id, { precio: e.target.value })} /></div>
                <div className="col-span-1 text-right"><button type="button" onClick={() => removeMaterialLinea(l.id)} className="text-slate-300 hover:text-rose-500"><X size={15} /></button></div>
              </div>
            ))}
            {f.materiales.length === 0 && <p className="text-sm text-slate-400">Sin materiales añadidos.</p>}
          </div>
          <div className="flex items-center justify-between mt-2">
            <button type="button" onClick={addMaterialLinea} className="flex items-center gap-1 text-sm font-semibold text-[#2E8B57] hover:text-[#256E46]"><Plus size={14} /> Añadir material</button>
            <span className="text-sm font-mono-num text-slate-500">Total materiales: <span className="font-semibold text-slate-800">{money(precioMateriales)}</span></span>
          </div>
        </div>

        {/* Fases */}
        <div>
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-2">Fases de composición</span>
          <div className="space-y-2">
            {f.fases.map((l) => (
              <div key={l.id} className="grid grid-cols-12 gap-2 items-center">
                <div className="col-span-4"><TextInput placeholder="Nombre de la fase" value={l.nombre} onChange={(e) => setFase(l.id, { nombre: e.target.value })} /></div>
                <div className="col-span-5"><TextInput placeholder="Materiales necesarios en esta fase" value={l.materialesNecesarios} onChange={(e) => setFase(l.id, { materialesNecesarios: e.target.value })} /></div>
                <div className="col-span-2"><TextInput type="number" step="0.25" placeholder="Horas" value={l.tiempoEstimado} onChange={(e) => setFase(l.id, { tiempoEstimado: e.target.value })} /></div>
                <div className="col-span-1 text-right"><button type="button" onClick={() => removeFase(l.id)} className="text-slate-300 hover:text-rose-500"><X size={15} /></button></div>
              </div>
            ))}
          </div>
          <div className="flex items-center justify-between mt-2">
            <button type="button" onClick={addFase} className="flex items-center gap-1 text-sm font-semibold text-[#2E8B57] hover:text-[#256E46]"><Plus size={14} /> Añadir fase</button>
            <span className="text-sm font-mono-num text-slate-500">Tiempo total estimado: <span className="font-semibold text-slate-800">{tiempoTotal.toFixed(2)} h</span></span>
          </div>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Tamaño"><TextInput value={f.tamano} onChange={set("tamano")} /></Field>
          <Field label="Medidas"><TextInput value={f.medidas} onChange={set("medidas")} placeholder="Por defecto, ajustables por proyecto" /></Field>
          <Field label="Volumen"><TextInput value={f.volumen} onChange={set("volumen")} /></Field>
        </div>

        <div className="grid grid-cols-3 gap-4">
          <Field label="Importe de envío (€)"><TextInput type="number" step="0.01" value={f.importeEnvio} onChange={set("importeEnvio")} /></Field>
          <Field label="Importe de montaje (€)"><TextInput type="number" step="0.01" value={f.importeMontaje} onChange={set("importeMontaje")} /></Field>
          <Field label="Precio de venta (€)"><TextInput type="number" step="0.01" value={f.precioVenta} onChange={set("precioVenta")} /></Field>
        </div>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar artículo</button>
        </div>
      </form>
    </div>
  );
}

function ArticuloDetail({ articulo, proveedor, materiales, onBack, onEdit, onDelete, onInlineUpdate, isAdmin }) {
  const [tab, setTab] = useState("datos");
  const gastos = articulo.gastos || [];
  const totalGastos = gastos.reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
  const precioMateriales = (articulo.materiales || []).reduce((s, l) => s + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0);
  const tiempoTotal = (articulo.fases || []).reduce((s, l) => s + (parseFloat(l.tiempoEstimado) || 0), 0);
  const materialNombre = (id) => materiales.find((m) => m.id === id)?.descripcion || "—";

  const [gForm, setGForm] = useState({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  const addGasto = (e) => {
    e.preventDefault();
    if (!gForm.proveedor.trim()) return;
    const next = [...gastos, { ...gForm, id: uid(), importe: parseFloat(gForm.importe) || 0 }];
    onInlineUpdate(articulo.id, { gastos: next });
    setGForm({ proveedor: "", producto: "", importe: "", facturaAsociada: "", estadoFactura: "Pendiente" });
  };
  const removeGasto = (id) => onInlineUpdate(articulo.id, { gastos: gastos.filter((g) => g.id !== id) });

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">Artículo {articulo.numero}</span>
            <Badge className="bg-violet-50 text-violet-700 ring-violet-200">Fase II</Badge>
          </div>
          <h1 className="font-display text-2xl font-extrabold text-slate-900 mt-0.5">{articulo.nombre}</h1>
          <p className="text-sm text-slate-500 mt-1">{proveedor?.nombre || "Sin proveedor"}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-4 gap-3 my-6">
        <Kpi label="Precio materiales" value={money(precioMateriales)} />
        <Kpi label="Tiempo total estimado" value={`${tiempoTotal.toFixed(2)} h`} />
        <Kpi label="Precio venta" value={money(articulo.precioVenta)} />
        <Kpi label="Gastos asociados" value={money(totalGastos)} />
      </div>

      <div className="flex gap-1 mb-4 border-b border-slate-200">
        {[
          { id: "datos", label: "Datos", icon: FileText },
          { id: "materiales", label: `Materiales (${(articulo.materiales || []).length})`, icon: Package },
          { id: "fases", label: `Fases de composición (${(articulo.fases || []).length})`, icon: Ruler },
          { id: "gastos", label: `Gastos asociados (${gastos.length})`, icon: Receipt },
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "datos" && (
        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
            <InfoRow icon={<Ruler size={14} />} label="Tamaño" value={articulo.tamano || "—"} />
            <InfoRow icon={<Ruler size={14} />} label="Medidas" value={articulo.medidas || "—"} />
            <InfoRow icon={<Package size={14} />} label="Volumen" value={articulo.volumen || "—"} />
            <InfoRow icon={<Euro size={14} />} label="Importe envío" value={money(articulo.importeEnvio)} />
            <InfoRow icon={<Euro size={14} />} label="Importe montaje" value={money(articulo.importeMontaje)} />
          </div>
          {articulo.descripcion && (
            <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
              <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Descripción</span>
              {articulo.descripcion}
            </div>
          )}
        </CornerFrame>
      )}

      {tab === "materiales" && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          {(articulo.materiales || []).length === 0 ? (
            <p className="px-4 py-8 text-center text-sm text-slate-400">Sin materiales asociados.</p>
          ) : (
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Material</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Cantidad</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Precio unitario</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Subtotal</th>
                </tr>
              </thead>
              <tbody>
                {articulo.materiales.map((l) => (
                  <tr key={l.id} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-medium text-slate-800">{materialNombre(l.materialId)}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{l.cantidad}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(l.precio)}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(l.cantidad * l.precio)}</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="bg-slate-50 font-semibold">
                  <td className="px-4 py-2.5" colSpan={3}>Total precio materiales</td>
                  <td className="px-4 py-2.5 text-right font-mono-num">{money(precioMateriales)}</td>
                </tr>
              </tfoot>
            </table>
          )}
        </div>
      )}

      {tab === "fases" && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          {(articulo.fases || []).length === 0 ? (
            <p className="px-4 py-8 text-center text-sm text-slate-400">Sin fases de composición definidas.</p>
          ) : (
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Fase</th>
                  <th className="px-4 py-2.5 font-semibold">Materiales necesarios</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Tiempo estimado</th>
                </tr>
              </thead>
              <tbody>
                {articulo.fases.map((l) => (
                  <tr key={l.id} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-medium text-slate-800">{l.nombre}</td>
                    <td className="px-4 py-2.5 text-slate-600">{l.materialesNecesarios || "—"}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{Number(l.tiempoEstimado).toFixed(2)} h</td>
                  </tr>
                ))}
              </tbody>
              <tfoot>
                <tr className="bg-slate-50 font-semibold">
                  <td className="px-4 py-2.5" colSpan={2}>Tiempo total estimado</td>
                  <td className="px-4 py-2.5 text-right font-mono-num">{tiempoTotal.toFixed(2)} h</td>
                </tr>
              </tfoot>
            </table>
          )}
        </div>
      )}

      {tab === "gastos" && (
        <div className="space-y-4">
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {gastos.length === 0 ? (
              <p className="px-4 py-8 text-center text-sm text-slate-400">Sin gastos registrados todavía.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Proveedor</th>
                    <th className="px-4 py-2.5 font-semibold">Producto / Tarea</th>
                    <th className="px-4 py-2.5 font-semibold">Factura</th>
                    <th className="px-4 py-2.5 font-semibold">Estado</th>
                    <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                    <th className="px-4 py-2.5"></th>
                  </tr>
                </thead>
                <tbody>
                  {gastos.map((g) => (
                    <tr key={g.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-medium text-slate-800">{g.proveedor}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.producto || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.facturaAsociada || "—"}</td>
                      <td className="px-4 py-2.5 text-slate-600">{g.estadoFactura}</td>
                      <td className="px-4 py-2.5 text-right font-mono-num">{money(g.importe)}</td>
                      <td className="px-4 py-2.5 text-right"><button onClick={() => removeGasto(g.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button></td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-50 font-semibold">
                    <td className="px-4 py-2.5" colSpan={4}>Total gastos asociados</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(totalGastos)}</td>
                    <td></td>
                  </tr>
                </tfoot>
              </table>
            )}
          </div>
          <form onSubmit={addGasto} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-6 gap-2 items-end">
            <Field label="Proveedor"><TextInput value={gForm.proveedor} onChange={(e) => setGForm({ ...gForm, proveedor: e.target.value })} /></Field>
            <Field label="Producto / Tarea"><TextInput value={gForm.producto} onChange={(e) => setGForm({ ...gForm, producto: e.target.value })} /></Field>
            <Field label="Factura asociada"><TextInput value={gForm.facturaAsociada} onChange={(e) => setGForm({ ...gForm, facturaAsociada: e.target.value })} /></Field>
            <Field label="Estado factura">
              <Select value={gForm.estadoFactura} onChange={(e) => setGForm({ ...gForm, estadoFactura: e.target.value })}>
                <option>Pendiente</option><option>Pagada</option>
              </Select>
            </Field>
            <Field label="Importe (€)"><TextInput type="number" step="0.01" value={gForm.importe} onChange={(e) => setGForm({ ...gForm, importe: e.target.value })} /></Field>
            <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Añadir</button>
          </form>
        </div>
      )}
    </div>
  );
}

/* ================= FACTURAS ================= */

function FacturasModulo({ facturas, clientes, proyectos, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onAddPago, onRemovePago, nextNumero, isAdmin }) {
  const [q, setQ] = useState("");
  const [tipo, setTipo] = useState("");
  const [estado, setEstado] = useState("");

  const clienteNombre = (id) => clientes.find((c) => c.id === id)?.nombre || "—";

  const filtered = useMemo(() => {
    return facturas.filter((f) => {
      if (tipo && f.tipo !== tipo) return false;
      if (estado && estadoFacturaCalc(f) !== estado) return false;
      if (!q) return true;
      const s = `${f.numero} ${clienteNombre(f.clienteId)}`.toLowerCase();
      return s.includes(q.toLowerCase());
    });
  }, [facturas, q, tipo, estado, clientes]);

  if (view === "form") {
    const editing = facturas.find((f) => f.id === editId) || null;
    return (
      <FacturaForm
        initial={editing}
        clientes={clientes}
        proyectos={proyectos}
        nextNumero={nextNumero}
        onCancel={() => setView(editId ? "detail" : "list")}
        onSave={onUpsert}
      />
    );
  }

  if (view === "detail") {
    const factura = facturas.find((f) => f.id === detailId);
    if (!factura) { setView("list"); return null; }
    return (
      <FacturaDetail
        factura={factura}
        cliente={clientes.find((c) => c.id === factura.clienteId)}
        proyectos={proyectos}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(factura.id); setView("form"); }}
        onDelete={() => onDelete(factura.id)}
        isAdmin={isAdmin}
        onAddPago={(pago) => onAddPago(factura.id, pago)}
        onRemovePago={(pagoId) => onRemovePago(factura.id, pagoId)}
      />
    );
  }

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <div className="flex items-center gap-2 mb-1">
        <Badge className="bg-slate-100 text-slate-500 ring-slate-200">Fase posterior</Badge>
        <span className="text-xs text-slate-400">A integrar con el presupuesto real de HETMO más adelante</span>
      </div>
      <Header
        icon={<Receipt size={20} className="text-[#2E8B57]" />}
        title="Facturas"
        manualKey="facturas"
        subtitle={`${facturas.length} factura${facturas.length === 1 ? "" : "s"} emitida${facturas.length === 1 ? "" : "s"}`}
      />

      {clientes.length === 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-3">
          ⚠ No puedes crear una factura todavía: primero da de alta al menos un cliente en la pestaña Clientes.
        </div>
      )}
      <button
        onClick={() => { setEditId(null); setView("form"); }}
        disabled={clientes.length === 0}
        title={clientes.length === 0 ? "Necesitas un cliente dado de alta" : ""}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 disabled:bg-slate-300 disabled:cursor-not-allowed text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVA FACTURA
      </button>

      <div className="flex flex-wrap gap-3 mb-4">
        <div className="relative flex-1 min-w-[220px] max-w-sm">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por nº o cliente…"
            className="w-full pl-9 pr-3 py-2 rounded-md border border-slate-300 bg-white text-sm focus:outline-none focus:ring-2 focus:ring-[#2E8B57]/40 focus:border-[#2E8B57]" />
        </div>
        <Select value={tipo} onChange={(e) => setTipo(e.target.value)} className="max-w-[160px]">
          <option value="">Todos los tipos</option>
          {TIPO_FACTURA.map((t) => <option key={t}>{t}</option>)}
        </Select>
        <Select value={estado} onChange={(e) => setEstado(e.target.value)} className="max-w-[160px]">
          <option value="">Todos los estados</option>
          <option>Pendiente</option><option>Parcial</option><option>Pagada</option>
        </Select>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nº Factura</th>
              <th className="px-4 py-3 font-semibold">Tipo</th>
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Fecha</th>
              <th className="px-4 py-3 font-semibold">Trabajos</th>
              <th className="px-4 py-3 font-semibold text-right">Total</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={8} className="px-4 py-10 text-center text-slate-400 text-sm">No hay facturas que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((f) => (
              <tr key={f.id} onClick={() => { setDetailId(f.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                <td className="px-4 py-3 font-mono-num text-slate-500">{f.numero}</td>
                <td className="px-4 py-3 text-slate-600">{f.tipo}</td>
                <td className="px-4 py-3 font-medium text-slate-800">{clienteNombre(f.clienteId)}</td>
                <td className="px-4 py-3 text-slate-500">{fmtDate(f.fecha)}</td>
                <td className="px-4 py-3 text-slate-600">{(f.proyectosIds || []).length}</td>
                <td className="px-4 py-3 text-right font-mono-num font-semibold">{money(f.total)}</td>
                <td className="px-4 py-3"><Badge className={ESTADO_FACTURA_STYLE[estadoFacturaCalc(f)]}>{estadoFacturaCalc(f)}</Badge></td>
                <td className="px-4 py-3 text-right">
                  {isAdmin && (<button onClick={(e) => { e.stopPropagation(); if (confirm(`¿Eliminar la factura ${f.numero}?`)) onDelete(f.id); }} className="text-slate-300 hover:text-rose-500 transition">
                    <Trash2 size={15} />
                  </button>)}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function FacturaForm({ initial, clientes, proyectos, nextNumero, onCancel, onSave }) {
  const [clienteId, setClienteId] = useState(initial?.clienteId || clientes[0]?.id || "");
  const [tipo, setTipo] = useState(initial?.tipo || "Definitiva");
  const [fecha, setFecha] = useState(initial?.fecha || new Date().toISOString().slice(0, 10));
  const [proyectosIds, setProyectosIds] = useState(initial?.proyectosIds || []);
  const [observaciones, setObservaciones] = useState(initial?.observaciones || "");
  const [errorMsg, setErrorMsg] = useState("");

  const trabajosCliente = proyectos.filter((p) => p.clienteId === clienteId);
  const total = proyectosIds.reduce((s, pid) => s + (proyectos.find((p) => p.id === pid)?.importePresupuesto || 0), 0);

  const toggleProyecto = (pid) => {
    setProyectosIds((ids) => (ids.includes(pid) ? ids.filter((x) => x !== pid) : [...ids, pid]));
  };

  const exportarFacturaExcel = (numeroFactura) => {
    const clienteNombre = clientes.find((c) => c.id === clienteId)?.nombre || "";
    const filaFactura = [{
      "Nº Factura": numeroFactura, Cliente: clienteNombre, Tipo: tipo, Fecha: fecha,
      Total: total, Observaciones: observaciones || "",
    }];
    const filasTrabajos = proyectosIds.map((pid) => {
      const p = proyectos.find((pr) => pr.id === pid);
      return { "Nº Proyecto": p?.numero || "", Proyecto: p?.nombre || "", Importe: p?.importePresupuesto || 0 };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filaFactura), "Factura");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filasTrabajos), "Trabajos");
    XLSX.writeFile(wb, `factura_${numeroFactura || "nueva"}_${fecha}.xlsx`);
  };

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!clienteId) {
      setErrorMsg("Falta seleccionar un cliente.");
      return;
    }
    if (proyectosIds.length === 0) {
      setErrorMsg("Selecciona al menos un trabajo para facturar (marca la casilla de un proyecto en la lista de abajo).");
      return;
    }
    setErrorMsg("");
    const numeroFactura = initial?.numero || nextNumero();
    onSave({ id: initial?.id || null, clienteId, tipo, fecha, proyectosIds, total, observaciones });
    try {
      exportarFacturaExcel(numeroFactura);
    } catch (err) {
      console.error("Error exportando factura a Excel:", err);
    }
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <div className="flex items-center gap-3 mb-6">
        <h1 className="font-display text-2xl font-extrabold text-slate-900">{initial ? `Editar factura ${initial.numero}` : "Emitir factura"}</h1>
        {!initial && <Badge className="bg-slate-100 text-slate-500 ring-slate-200 font-mono-num">Nº {nextNumero()} (automático)</Badge>}
      </div>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-3 gap-4">
          <Field label="Cliente" required>
            <Select value={clienteId} onChange={(e) => { setClienteId(e.target.value); setProyectosIds([]); }} required>
              {clientes.map((c) => <option key={c.id} value={c.id}>{c.nombre}</option>)}
            </Select>
          </Field>
          <Field label="Tipo">
            <Select value={tipo} onChange={(e) => setTipo(e.target.value)}>
              {TIPO_FACTURA.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <Field label="Fecha"><TextInput type="date" value={fecha} onChange={(e) => setFecha(e.target.value)} /></Field>
        </div>

        <div>
          <span className="block text-[11px] font-semibold tracking-wide uppercase text-slate-500 mb-2">Trabajos a facturar (del cliente seleccionado)</span>
          {trabajosCliente.length === 0 ? (
            <p className="text-sm text-slate-400">Este cliente no tiene proyectos disponibles.</p>
          ) : (
            <div className="border border-slate-200 rounded-md divide-y divide-slate-100 max-h-64 overflow-y-auto">
              {trabajosCliente.map((p) => (
                <label key={p.id} className="flex items-center gap-3 px-3 py-2.5 text-sm hover:bg-slate-50 cursor-pointer">
                  <input type="checkbox" checked={proyectosIds.includes(p.id)} onChange={() => toggleProyecto(p.id)} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
                  <span className="font-mono-num text-slate-400">#{p.numero}</span>
                  <span className="flex-1 text-slate-700">{p.nombre}</span>
                  <span className="font-mono-num text-slate-600">{money(p.importePresupuesto)}</span>
                </label>
              ))}
            </div>
          )}
          <div className="flex justify-end mt-2 text-sm">
            <span className="font-mono-num text-slate-500">Total factura: <span className="font-bold text-slate-800 text-base">{money(total)}</span></span>
          </div>
        </div>

        <Field label="Observaciones">
          <TextArea rows={3} value={observaciones} onChange={(e) => setObservaciones(e.target.value)} placeholder="Cualquier nota sobre esta factura (opcional)" />
        </Field>

        <div className="flex flex-wrap items-center justify-end gap-2 pt-2">
          <span className="text-xs text-slate-400 mr-auto">Al guardar, se descargará también un Excel con los datos de esta factura.</span>
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar factura</button>
        </div>
      </form>
    </div>
  );
}

function FacturaDetail({ factura, cliente, proyectos, onBack, onEdit, onDelete, onAddPago, onRemovePago, isAdmin }) {
  const trabajos = proyectos.filter((p) => (factura.proyectosIds || []).includes(p.id));
  const pagos = factura.pagos || [];
  const totalPagado = pagos.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
  const saldo = (parseFloat(factura.total) || 0) - totalPagado;
  const estado = estadoFacturaCalc(factura);

  const exportarFacturaDetalleExcel = (fac, cli, trabajosLista) => {
    const filaFactura = [{
      "Nº Factura": fac.numero, Cliente: cli?.nombre || "", Tipo: fac.tipo, Fecha: fac.fecha,
      Total: fac.total, Observaciones: fac.observaciones || "",
    }];
    const filasTrabajos = trabajosLista.map((t) => ({ "Nº Proyecto": t.numero, Proyecto: t.nombre, Importe: t.importePresupuesto || 0 }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filaFactura), "Factura");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filasTrabajos), "Trabajos");
    XLSX.writeFile(wb, `factura_${fac.numero}_${fac.fecha}.xlsx`);
  };

  const [pForm, setPForm] = useState({ fecha: new Date().toISOString().slice(0, 10), importe: "", formaPago: "Transferencia" });
  const addPago = (e) => {
    e.preventDefault();
    const importe = parseFloat(pForm.importe) || 0;
    if (importe <= 0) return;
    onAddPago({ ...pForm, id: uid(), importe });
    setPForm({ fecha: new Date().toISOString().slice(0, 10), importe: "", formaPago: "Transferencia" });
  };

  return (
    <div className="p-8 max-w-4xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver al listado</button>

      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-mono-num text-sm text-[#2E8B57] font-bold">Factura {factura.numero}</span>
            <Badge className="bg-slate-100 text-slate-600 ring-slate-200">{factura.tipo}</Badge>
          </div>
          <h1 className="font-display text-2xl font-extrabold text-slate-900 mt-0.5">{cliente?.nombre || "Cliente no encontrado"}</h1>
          <p className="text-sm text-slate-500 mt-1">{cliente?.cif} · {fmtDate(factura.fecha)}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          <button onClick={() => exportarFacturaDetalleExcel(factura, cliente, trabajos)} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><FileSpreadsheet size={14} /> Exportar a Excel</button>
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-4 gap-3 my-6">
        <Kpi label="Total factura" value={money(factura.total)} />
        <Kpi label="Total pagado" value={money(totalPagado)} tone="good" />
        <Kpi label="Saldo pendiente" value={money(Math.max(saldo, 0))} tone={saldo > 0 ? "bad" : "neutral"} />
        <div className="bg-white border border-slate-200 rounded-lg px-4 py-3 flex flex-col justify-center">
          <div className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 mb-1">Estado</div>
          <Badge className={ESTADO_FACTURA_STYLE[estado] + " w-fit"}>{estado}</Badge>
        </div>
      </div>

      {factura.observaciones && (
        <div className="bg-white border border-slate-200 rounded-lg px-4 py-3 mb-6 text-sm text-slate-600">
          <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Observaciones</span>
          {factura.observaciones}
        </div>
      )}

      <h2 className="font-display font-bold text-slate-800 mb-2">Trabajos incluidos</h2>
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-6">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Nº</th>
              <th className="px-4 py-2.5 font-semibold">Proyecto</th>
              <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
            </tr>
          </thead>
          <tbody>
            {trabajos.map((t) => (
              <tr key={t.id} className="border-b border-slate-100 last:border-0">
                <td className="px-4 py-2.5 font-mono-num text-slate-500">#{t.numero}</td>
                <td className="px-4 py-2.5 font-medium text-slate-800">{t.nombre}</td>
                <td className="px-4 py-2.5 text-right font-mono-num">{money(t.importePresupuesto)}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <h2 className="font-display font-bold text-slate-800 mb-2 flex items-center gap-2"><Wallet size={16} className="text-[#2E8B57]" /> Histórico de pagos</h2>
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-4">
        {pagos.length === 0 ? (
          <p className="px-4 py-8 text-center text-sm text-slate-400">Sin pagos registrados todavía.</p>
        ) : (
          <table className="w-full text-sm">
            <thead>
              <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                <th className="px-4 py-2.5 font-semibold">Fecha</th>
                <th className="px-4 py-2.5 font-semibold">Forma de pago</th>
                <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                <th className="px-4 py-2.5"></th>
              </tr>
            </thead>
            <tbody>
              {pagos.map((p) => (
                <tr key={p.id} className="border-b border-slate-100 last:border-0">
                  <td className="px-4 py-2.5 text-slate-500">{fmtDate(p.fecha)}</td>
                  <td className="px-4 py-2.5 text-slate-600">{p.formaPago}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num">{money(p.importe)}</td>
                  <td className="px-4 py-2.5 text-right"><button onClick={() => onRemovePago(p.id)} className="text-slate-300 hover:text-rose-500"><X size={14} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {saldo > 0 && (
        <form onSubmit={addPago} className="bg-white border border-slate-200 rounded-lg p-4 grid grid-cols-4 gap-2 items-end">
          <Field label="Fecha"><TextInput type="date" value={pForm.fecha} onChange={(e) => setPForm({ ...pForm, fecha: e.target.value })} /></Field>
          <Field label="Importe (€)"><TextInput type="number" step="0.01" value={pForm.importe} onChange={(e) => setPForm({ ...pForm, importe: e.target.value })} placeholder={saldo.toFixed(2)} /></Field>
          <Field label="Forma de pago">
            <Select value={pForm.formaPago} onChange={(e) => setPForm({ ...pForm, formaPago: e.target.value })}>
              {FORMA_PAGO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center justify-center gap-1 text-sm font-semibold px-3 py-2 rounded-md h-[38px]"><Plus size={15} /> Registrar pago</button>
        </form>
      )}
    </div>
  );
}

/* ================= PRESUPUESTOS ================= */

const ESTADO_PRESUPUESTO_TRACKER = ["Pendiente", "Aceptado", "Rechazado", "En espera"];
const MOTIVO_RECHAZO = ["Precio", "Competencia", "Plazo", "Otro"];
const ESTADO_PRESUPUESTO_TRACKER_STYLE = {
  Pendiente: "bg-amber-50 text-amber-700 ring-amber-200",
  Aceptado: "bg-emerald-50 text-emerald-700 ring-emerald-200",
  Rechazado: "bg-rose-50 text-rose-700 ring-rose-200",
  "En espera": "bg-slate-100 text-slate-600 ring-slate-200",
};

const diasEntre = (a, b) => {
  if (!a || !b) return null;
  const ms = new Date(b).setHours(0, 0, 0, 0) - new Date(a).setHours(0, 0, 0, 0);
  return Math.round(ms / 86400000);
};

const diasSinRespuestaDe = (p) => {
  if (p.estado !== "Pendiente" || !p.fechaEnvio) return null;
  return diasEntre(p.fechaEnvio, new Date().toISOString().slice(0, 10));
};

const mensajeWhatsappPresupuesto = (p) => {
  return `Hola ${p.clienteNombre}, le escribimos de ALUMAVEL para saber si ha podido revisar el presupuesto ${p.numero}${p.descripcion ? ` (${p.descripcion})` : ""}. Quedamos a su disposición para cualquier duda. Un saludo.`;
};

const enlaceWhatsapp = (p) => {
  const tel = (p.telefono || "").replace(/[^\d+]/g, "");
  if (!tel) return null;
  const telConPrefijo = tel.startsWith("+") ? tel.replace("+", "") : (tel.startsWith("34") ? tel : `34${tel}`);
  return `https://wa.me/${telConPrefijo}?text=${encodeURIComponent(mensajeWhatsappPresupuesto(p))}`;
};

const enlaceLlamar = (p) => {
  const tel = (p.telefono || "").replace(/[^\d+]/g, "");
  if (!tel) return null;
  return `tel:${tel.startsWith("+") ? tel : `+${tel.startsWith("34") ? tel : `34${tel}`}`}`;
};

const semanaISO = (fechaStr) => {
  const d = new Date(fechaStr);
  d.setHours(0, 0, 0, 0);
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
  return `${d.getFullYear()}-S${String(weekNo).padStart(2, "0")}`;
};

function PresupuestosModulo({ presupuestos, clientes, onCrearClienteRapido, view, setView, editId, setEditId, detailId, setDetailId, onUpsert, onDelete, onAddLlamada, onDeleteLlamada, onCrearProyecto, onDuplicar, isAdmin, prefill, onClearPrefill }) {
  const [tab, setTab] = useState("lista");
  const [q, setQ] = useState("");
  const [estadoFiltro, setEstadoFiltro] = useState("");
  const [zonaFiltro, setZonaFiltro] = useState("");
  const [mostrarEmailPanel, setMostrarEmailPanel] = useState(false);
  const [prefillPresupuesto, setPrefillPresupuesto] = useState(null);
  const [leyendoFoto, setLeyendoFoto] = useState(false);
  const [errorFoto, setErrorFoto] = useState("");
  const inputFotoRef = useRef(null);

  // Cuando llega un "prefill" desde fuera (por ejemplo, al pulsar "Pasar a
  // presupuesto" en una medición), lo convertimos al mismo formato que ya usa
  // el prefill interno de "leer foto", y abrimos el formulario directamente.
  useEffect(() => {
    if (!prefill) return;
    setPrefillPresupuesto({
      id: null, numero: "", fechaEnvio: new Date().toISOString().slice(0, 10),
      clienteNombre: prefill.clienteNombre || "", telefono: "",
      descripcion: prefill.descripcion || "", importe: "", estado: "Pendiente",
      motivoRechazo: "", fechaRespuesta: "", comentarios: "Creado a partir de una medición. Revisa los datos y añade el importe antes de guardar.",
      fechaPrevistaConfirmacion: "", envio: false, direccionEnvio: prefill.direccionEnvio || "", montaje: false, zona: "",
      techos: prefill.techos || [],
    });
    setEditId(null);
    setView("form");
    if (onClearPrefill) onClearPrefill();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [prefill]);

  const leerDatosDesdeArchivo = async (file) => {
    const base64Data = await new Promise((res, rej) => {
      const r = new FileReader();
      r.onload = () => res(r.result.split(",")[1]);
      r.onerror = () => rej(new Error("No se pudo leer el archivo"));
      r.readAsDataURL(file);
    });
    const esPdf = file.type === "application/pdf";
    const contentBlock = esPdf
      ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64Data } }
      : { type: "image", source: { type: "base64", media_type: file.type || "image/jpeg", data: base64Data } };

    const prompt = 'Esto es un presupuesto o una nota con datos de un presupuesto para un cliente (puede ser una foto de algo escrito a mano, un documento impreso de un programa de presupuestos, etc). Es MUY IMPORTANTE que revises el documento entero, de arriba a abajo, y devuelvas TODAS las medidas/piezas, sin saltarte ninguna ni resumir. Devuelve ÚNICAMENTE un JSON válido (sin texto adicional, sin backticks) con este formato exacto: {"clienteNombre":"","telefono":"","importe":numero_o_vacio,"descripcionGeneral":"","zona":"","medidas":[{"referencia":"","ancho":"","alto":"","cantidad":""}]}. En "medidas" incluye una línea por cada pieza, ventana, puerta, etc. que tenga ancho y alto (en la unidad que aparezca, normalmente mm), con su referencia o nombre y la cantidad. No omitas ninguna pieza. Si no hay medidas, deja el array vacío. Deja en blanco lo que no encuentres.';

    const response = await fetch("/.netlify/functions/anthropic-proxy", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 8000,
        messages: [{ role: "user", content: [contentBlock, { type: "text", text: prompt }] }],
      }),
    });
    if (!response.ok) {
      const errBody = await response.text();
      console.error("anthropic-proxy respuesta no válida:", response.status, errBody);
      throw new Error("Respuesta no válida de la API: " + response.status);
    }
    const data = await response.json();
    if (data.error) {
      console.error("Error devuelto por la API:", data.error);
      throw new Error(data.error.message || "Error de la API");
    }
    const textoRespuesta = (data.content || []).filter((c) => c.type === "text").map((c) => c.text).join("");
    const limpio = textoRespuesta.replace(/```json|```/g, "").trim();
    const inicioP = limpio.indexOf("{");
    const finP = limpio.lastIndexOf("}");
    const jsonCandidatoP = inicioP !== -1 && finP !== -1 ? limpio.slice(inicioP, finP + 1) : limpio;
    let info;
    try {
      info = JSON.parse(jsonCandidatoP);
    } catch (parseErr) {
      console.error("No se pudo parsear el JSON del presupuesto. Texto recibido:", textoRespuesta);
      throw new Error("La respuesta de la IA no tenía formato válido");
    }

    const medidas = Array.isArray(info.medidas) ? info.medidas.filter((m) => m && (m.ancho || m.alto)) : [];
    const textoMedidas = medidas.length > 0
      ? "\n\nMedidas:\n" + medidas.map((m) => `- ${m.referencia || "Pieza"}: ${m.ancho || "—"} x ${m.alto || "—"} mm${m.cantidad ? ` (x${m.cantidad})` : ""}`).join("\n")
      : "";
    const descripcionCompleta = `${info.descripcionGeneral || ""}${textoMedidas}`.trim();

    if (!info.clienteNombre && !descripcionCompleta && !info.importe) {
      throw new Error("No he podido leer datos claros en la imagen. Prueba con una foto más nítida.");
    }

    return {
      clienteNombre: info.clienteNombre || "",
      telefono: info.telefono || "",
      descripcion: descripcionCompleta,
      importe: info.importe || "",
      zona: info.zona || "",
      nombreArchivo: file.name,
    };
  };

  const leerFotoPresupuesto = async (file) => {
    setLeyendoFoto(true);
    setErrorFoto("");
    try {
      const datos = await leerDatosDesdeArchivo(file);
      setPrefillPresupuesto({
        id: null, numero: "", fechaEnvio: new Date().toISOString().slice(0, 10),
        clienteNombre: datos.clienteNombre, telefono: datos.telefono,
        descripcion: datos.descripcion, importe: datos.importe, estado: "Pendiente",
        motivoRechazo: "", fechaRespuesta: "", comentarios: `Creado a partir de una foto/PDF subida (${datos.nombreArchivo}). Revisa los datos antes de guardar.`,
        fechaPrevistaConfirmacion: "", envio: false, direccionEnvio: "", montaje: false, zona: datos.zona,
      });
      setEditId(null);
      setView("form");
    } catch (err) {
      console.error("Error leyendo foto de presupuesto:", err);
      setErrorFoto("No se pudo leer el archivo. Prueba con una foto más clara, con más luz, o inténtalo de nuevo. (" + err.message + ")");
    } finally {
      setLeyendoFoto(false);
    }
  };

  const zonasDisponibles = useMemo(() => {
    const uniq = new Set(presupuestos.map((p) => (p.zona || "").trim()).filter(Boolean));
    return Array.from(uniq).sort();
  }, [presupuestos]);

  // Una réplica es un presupuesto cuyo número es "base-N" y ese "base" coincide con
  // el número de otro presupuesto ya existente (el original).
  const esReplica = (p) => {
    const partes = String(p.numero).split("-");
    if (partes.length < 2) return false;
    const base = partes[0].trim();
    const sufijo = partes[partes.length - 1];
    return /^\d+$/.test(sufijo) && presupuestos.some((other) => other.id !== p.id && String(other.numero).trim() === base);
  };
  const replicasDe = (p) => presupuestos.filter((other) => other.id !== p.id && String(other.numero).split("-")[0].trim() === String(p.numero).trim() && esReplica(other));

  const filtered = useMemo(() => {
    return presupuestos.filter((p) => {
      if (!q && esReplica(p) && (p.estado || "Pendiente") === "Pendiente") return false;
      if (estadoFiltro && p.estado !== estadoFiltro) return false;
      if (zonaFiltro && p.zona !== zonaFiltro) return false;
      if (!q) return true;
      const hay = `${p.numero} ${p.clienteNombre} ${p.descripcion} ${p.zona}`.toLowerCase();
      return hay.includes(q.toLowerCase());
    });
  }, [presupuestos, q, estadoFiltro, zonaFiltro]);

  const llamarHoy = presupuestos.filter((p) => {
    const d = diasSinRespuestaDe(p);
    return d !== null && d > 7;
  });
  const contactarVencidos = presupuestos.filter((p) => {
    if (p.estado !== "En espera" || !p.fechaPrevistaConfirmacion) return false;
    return new Date(p.fechaPrevistaConfirmacion) < new Date(new Date().toISOString().slice(0, 10));
  });

  const exportarExcel = () => {
    const filas = presupuestos.map((p) => ({
      "Nº Presupuesto": p.numero, "Fecha envío": p.fechaEnvio, Cliente: p.clienteNombre,
      "Descripción / Obra": p.descripcion, "Importe (€)": p.importe, Estado: p.estado,
      "Motivo rechazo": p.motivoRechazo || "", "Fecha respuesta": p.fechaRespuesta || "",
      "Días respuesta": diasEntre(p.fechaEnvio, p.fechaRespuesta), Comentarios: p.comentarios || "",
      "Días sin respuesta": diasSinRespuestaDe(p) ?? "", "¿Envío?": p.envio ? "sí" : "no",
      "Dirección de envío": p.direccionEnvio || "", "¿Montaje?": p.montaje ? "sí" : "no", Zona: p.zona || "",
    }));

    const total = presupuestos.length;
    const porEstado = ESTADO_PRESUPUESTO_TRACKER.map((estado) => {
      const lista = presupuestos.filter((p) => p.estado === estado);
      return { estado, n: lista.length, pct: total ? (lista.length / total) * 100 : 0, importe: lista.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0) };
    });
    const rechazados = presupuestos.filter((p) => p.estado === "Rechazado");
    const porMotivo = MOTIVO_RECHAZO.map((m) => {
      const n = rechazados.filter((p) => p.motivoRechazo === m).length;
      return { motivo: m, n, pct: rechazados.length ? (n / rechazados.length) * 100 : 0 };
    });
    const zonas = Array.from(new Set(presupuestos.map((p) => (p.zona || "").trim()).filter(Boolean))).sort();
    const zonasStats = zonas.map((zona) => {
      const lista = presupuestos.filter((p) => p.zona === zona);
      const aceptados = lista.filter((p) => p.estado === "Aceptado");
      const rechazadosZ = lista.filter((p) => p.estado === "Rechazado");
      const pct = (aceptados.length + rechazadosZ.length) ? (aceptados.length / (aceptados.length + rechazadosZ.length)) * 100 : null;
      return { zona, enviados: lista.length, aceptados: aceptados.length, rechazados: rechazadosZ.length, pct, importeAceptado: aceptados.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0) };
    });

    const agrupar = (keyFn) => {
      const map = new Map();
      presupuestos.forEach((p) => {
        if (!p.fechaEnvio) return;
        const key = keyFn(p.fechaEnvio);
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(p);
      });
      return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([periodo, lista]) => {
        const aceptados = lista.filter((p) => p.estado === "Aceptado");
        const rechazadosP = lista.filter((p) => p.estado === "Rechazado");
        const pendientes = lista.filter((p) => p.estado === "Pendiente");
        const pct = (aceptados.length + rechazadosP.length) ? (aceptados.length / (aceptados.length + rechazadosP.length)) * 100 : null;
        return {
          periodo, enviados: lista.length, aceptados: aceptados.length, rechazados: rechazadosP.length, pendientes: pendientes.length, pct,
          importeTotal: lista.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0),
          importeAceptado: aceptados.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0),
        };
      });
    };
    const porAno = agrupar((f) => String(new Date(f).getFullYear()));
    const porMes = agrupar((f) => f.slice(0, 7));
    const porSemana = agrupar(semanaISO);

    const aoaStats = [
      ["Resumen general"],
      ["Estado", "Nº presupuestos", "% del total", "Importe total (€)"],
      ...porEstado.map((e) => [e.estado, e.n, `${e.pct.toFixed(1)}%`, e.importe]),
      [],
      ["Motivos de rechazo"],
      ["Motivo", "Nº presupuestos", "% sobre rechazados"],
      ...porMotivo.map((m) => [m.motivo, m.n, `${m.pct.toFixed(1)}%`]),
      [],
      ["Por Zona"],
      ["Zona", "Enviados", "Aceptados", "Rechazados", "% Aceptación", "Importe aceptado (€)"],
      ...zonasStats.map((z) => [z.zona, z.enviados, z.aceptados, z.rechazados, z.pct !== null ? `${z.pct.toFixed(1)}%` : "", z.importeAceptado]),
    ];

    const cabeceraPeriodoAoa = ["Periodo", "Enviados", "Aceptados", "Rechazados", "Pendientes", "% Aceptación", "Importe total (€)", "Importe aceptado (€)"];
    const filaPeriodoAoa = (r) => [r.periodo, r.enviados, r.aceptados, r.rechazados, r.pendientes, r.pct !== null ? `${r.pct.toFixed(1)}%` : "", r.importeTotal, r.importeAceptado];
    const aoaPeriodo = [
      ["Por Año"], cabeceraPeriodoAoa, ...porAno.map(filaPeriodoAoa), [],
      ["Por Mes"], cabeceraPeriodoAoa, ...porMes.map(filaPeriodoAoa), [],
      ["Por Semana (lunes-domingo)"], cabeceraPeriodoAoa, ...porSemana.map(filaPeriodoAoa),
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), "Presupuestos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaStats), "Estadísticas");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoaPeriodo), "Por Periodo");
    XLSX.writeFile(wb, `presupuestos_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  if (view === "form") {
    const editing = presupuestos.find((p) => p.id === editId) || null;
    return (
      <PresupuestoForm
        initial={editing ? { ...editing, estado: editing.estado || "Pendiente" } : prefillPresupuesto}
        clientes={clientes}
        presupuestosExistentes={presupuestos}
        onCrearClienteRapido={onCrearClienteRapido}
        onLeerDatos={leerDatosDesdeArchivo}
        onCancel={() => { setView(editId ? "detail" : "list"); setPrefillPresupuesto(null); }}
        onSave={(data) => { onUpsert(data); setPrefillPresupuesto(null); }}
      />
    );
  }

  if (view === "detail") {
    const presupuesto = presupuestos.find((p) => p.id === detailId);
    if (!presupuesto) { setView("list"); return null; }
    return (
      <PresupuestoDetail
        presupuesto={presupuesto}
        onBack={() => setView("list")}
        onEdit={() => { setEditId(presupuesto.id); setView("form"); }}
        onDelete={() => onDelete(presupuesto.id)}
        onAddLlamada={(llamada) => onAddLlamada(presupuesto.id, llamada)}
        onDeleteLlamada={(llamadaId) => onDeleteLlamada(presupuesto.id, llamadaId)}
        onCrearProyecto={() => onCrearProyecto(presupuesto)}
        onDuplicar={() => onDuplicar(presupuesto)}
        replicas={replicasDe(presupuesto)}
        onAbrirReplica={(id) => { setDetailId(id); setView("detail"); }}
        isAdmin={isAdmin}
      />
    );
  }

  return (
    <div className="p-8">
      <Header
        icon={<FileSpreadsheet size={20} className="text-[#2E8B57]" />}
        title="Presupuestos"
        manualKey="presupuestos"
        subtitle={`${presupuestos.length} presupuesto${presupuestos.length === 1 ? "" : "s"} registrado${presupuestos.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-3 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO PRESUPUESTO
      </button>

      <button
        type="button"
        onClick={() => inputFotoRef.current?.click()}
        disabled={leyendoFoto}
        style={{ borderColor: "#2E8B57", color: "#2E8B57" }}
        className="w-full flex items-center justify-center gap-2 border-2 hover:bg-slate-50 disabled:opacity-50 text-sm font-semibold py-3 rounded-lg mb-6 cursor-pointer select-none"
      >
        <FileText size={16} /> {leyendoFoto ? "Leyendo..." : "Crear presupuesto a partir de una foto o PDF"}
      </button>
      <input
        ref={inputFotoRef}
        type="file"
        accept="image/*,application/pdf"
        className="hidden"
        onChange={(e) => { if (e.target.files?.[0]) leerFotoPresupuesto(e.target.files[0]); e.target.value = ""; }}
      />
      {errorFoto && (
        <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold mb-6">⚠ {errorFoto}</div>
      )}

      {(llamarHoy.length > 0 || contactarVencidos.length > 0) && (
        <div className="mb-6 space-y-3">
          {llamarHoy.length > 0 && (
            <div className="px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm">
              <div className="font-semibold mb-2">⚠ Llamar hoy (más de 7 días sin respuesta): {llamarHoy.length}</div>
              <div className="space-y-1.5">
                {llamarHoy.map((p) => (
                  <div key={p.id} className="flex items-center justify-between gap-2 bg-white/60 rounded px-3 py-1.5">
                    <span>{p.numero} — {p.clienteNombre} ({diasSinRespuestaDe(p)}d)</span>
                    {enlaceWhatsapp(p) ? (
                      <a href={enlaceWhatsapp(p)} target="_blank" rel="noopener noreferrer" className="flex items-center gap-1 text-xs font-semibold bg-emerald-600 text-white px-2.5 py-1 rounded">
                        <MessageCircle size={13} /> WhatsApp
                      </a>
                    ) : (
                      <span className="text-xs text-slate-400">Sin teléfono</span>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
          {contactarVencidos.length > 0 && (
            <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold">
              📅 Contactar (en espera vencida): {contactarVencidos.length}
            </div>
          )}
          <button
            onClick={() => setMostrarEmailPanel((v) => !v)}
            className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"
          >
            <Mail size={14} /> Enviar aviso a mi correo
          </button>
          {mostrarEmailPanel && (
            <EnviarAvisoEmailPanel llamarHoy={llamarHoy} contactarVencidos={contactarVencidos} />
          )}
        </div>
      )}

      <div className="flex gap-1 mb-5 border-b border-slate-200">
        {[{ id: "lista", label: "Lista" }, { id: "stats", label: "Estadísticas" }, { id: "presentacion", label: "Presentación" }].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)}
            className={`px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            {t.label}
          </button>
        ))}
      </div>

      {tab === "lista" && (
        <>
          <div className="flex flex-wrap gap-3 mb-4">
            <div className="relative flex-1 min-w-[200px]">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por número, cliente, obra..." className={inputCls + " pl-9"} />
            </div>
            <Select value={estadoFiltro} onChange={(e) => setEstadoFiltro(e.target.value)} className="max-w-[180px]">
              <option value="">Todos los estados</option>
              {ESTADO_PRESUPUESTO_TRACKER.map((t) => <option key={t}>{t}</option>)}
            </Select>
            {zonasDisponibles.length > 0 && (
              <Select value={zonaFiltro} onChange={(e) => setZonaFiltro(e.target.value)} className="max-w-[180px]">
                <option value="">Todas las zonas</option>
                {zonasDisponibles.map((z) => <option key={z}>{z}</option>)}
              </Select>
            )}
            <button onClick={exportarExcel} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><FileSpreadsheet size={14} /> Exportar a Excel</button>
          </div>
          {presupuestos.some((p) => esReplica(p) && (p.estado || "Pendiente") === "Pendiente") && !q && (
            <p className="text-xs text-slate-400 mb-3">Las réplicas pendientes (números con guion, ej. 4192-1) están ocultas aquí — se ven dentro de la ficha del presupuesto original, o buscando su número. En cuanto una réplica se acepta o se rechaza, vuelve a aparecer aquí como cualquier otro presupuesto.</p>
          )}

          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-3 font-semibold">Nº</th>
                  <th className="px-4 py-3 font-semibold">Fecha envío</th>
                  <th className="px-4 py-3 font-semibold">Cliente</th>
                  <th className="px-4 py-3 font-semibold">Descripción / Obra</th>
                  <th className="px-4 py-3 font-semibold text-right">Importe</th>
                  <th className="px-4 py-3 font-semibold">Estado</th>
                  <th className="px-4 py-3 font-semibold">Zona</th>
                  <th className="px-4 py-3 font-semibold">Días s/resp.</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((p) => {
                  const dias = diasSinRespuestaDe(p);
                  return (
                    <tr key={p.id} onClick={() => { setDetailId(p.id); setView("detail"); }} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                      <td className="px-4 py-3 font-mono-num text-slate-500">{p.numero}</td>
                      <td className="px-4 py-3 text-slate-600">{fmtDate(p.fechaEnvio)}</td>
                      <td className="px-4 py-3 font-medium text-slate-800">{p.clienteNombre}</td>
                      <td className="px-4 py-3 text-slate-600">{p.descripcion}</td>
                      <td className="px-4 py-3 text-right font-mono-num">{money(p.importe)}</td>
                      <td className="px-4 py-3">
                        <div className="flex items-center gap-1.5 flex-wrap">
                          <Badge className={ESTADO_PRESUPUESTO_TRACKER_STYLE[p.estado || "Pendiente"]}>{p.estado || "Pendiente"}</Badge>
                          {p.proyectoCreadoId && <Badge className="bg-violet-50 text-violet-700 ring-violet-200">✓ Ya tiene proyecto</Badge>}
                        </div>
                      </td>
                      <td className="px-4 py-3 text-slate-500">{p.zona || "—"}</td>
                      <td className="px-4 py-3">{dias !== null ? <span className={dias > 7 ? "text-rose-600 font-semibold" : "text-slate-500"}>{dias}d</span> : "—"}</td>
                    </tr>
                  );
                })}
                {filtered.length === 0 && (
                  <tr><td colSpan={8} className="px-4 py-10 text-center text-slate-400 text-sm">No hay presupuestos que coincidan con la búsqueda.</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </>
      )}

      {tab === "stats" && <PresupuestosEstadisticas presupuestos={presupuestos} />}
      {tab === "presentacion" && <PresupuestosPresentacion presupuestos={presupuestos} />}
    </div>
  );
}

function PresupuestosEstadisticas({ presupuestos }) {
  const COLOR_ESTADO = { Pendiente: "#f59e0b", Aceptado: "#10b981", Rechazado: "#f43f5e", "En espera": "#94a3b8" };
  const total = presupuestos.length;
  const porEstado = ESTADO_PRESUPUESTO_TRACKER.map((estado) => {
    const lista = presupuestos.filter((p) => p.estado === estado);
    const importe = lista.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
    return { estado, n: lista.length, pct: total ? (lista.length / total) * 100 : 0, importe };
  });
  const datosPie = porEstado.filter((e) => e.n > 0).map((e) => ({ name: e.estado, value: e.n }));

  const rechazados = presupuestos.filter((p) => p.estado === "Rechazado");
  const porMotivo = MOTIVO_RECHAZO.map((m) => {
    const n = rechazados.filter((p) => p.motivoRechazo === m).length;
    return { motivo: m, n, pct: rechazados.length ? (n / rechazados.length) * 100 : 0 };
  });

  const aceptadosN = presupuestos.filter((p) => p.estado === "Aceptado").length;
  const pctAceptacion = (aceptadosN + rechazados.length) ? (aceptadosN / (aceptadosN + rechazados.length)) * 100 : null;

  const zonasStats = useMemo(() => {
    const zonas = Array.from(new Set(presupuestos.map((p) => (p.zona || "").trim()).filter(Boolean))).sort();
    return zonas.map((zona) => {
      const lista = presupuestos.filter((p) => p.zona === zona);
      const aceptados = lista.filter((p) => p.estado === "Aceptado");
      const rechazadosZ = lista.filter((p) => p.estado === "Rechazado");
      const pct = (aceptados.length + rechazadosZ.length) ? (aceptados.length / (aceptados.length + rechazadosZ.length)) * 100 : null;
      const importeAceptado = aceptados.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
      return { zona, enviados: lista.length, aceptados: aceptados.length, rechazados: rechazadosZ.length, pct, importeAceptado };
    });
  }, [presupuestos]);

  const periodos = useMemo(() => {
    const agrupar = (keyFn) => {
      const map = new Map();
      presupuestos.forEach((p) => {
        if (!p.fechaEnvio) return;
        const key = keyFn(p.fechaEnvio);
        if (!map.has(key)) map.set(key, []);
        map.get(key).push(p);
      });
      return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([periodo, lista]) => {
        const aceptados = lista.filter((p) => p.estado === "Aceptado");
        const rechazadosP = lista.filter((p) => p.estado === "Rechazado");
        const pendientes = lista.filter((p) => p.estado === "Pendiente");
        const pct = (aceptados.length + rechazadosP.length) ? (aceptados.length / (aceptados.length + rechazadosP.length)) * 100 : null;
        const importeTotal = lista.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
        const importeAceptado = aceptados.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
        return { periodo, enviados: lista.length, aceptados: aceptados.length, rechazados: rechazadosP.length, pendientes: pendientes.length, pct, importeTotal, importeAceptado };
      });
    };
    return {
      porAno: agrupar((f) => String(new Date(f).getFullYear())),
      porMes: agrupar((f) => f.slice(0, 7)),
      porSemana: agrupar(semanaISO),
    };
  }, [presupuestos]);

  const filaPeriodo = (r) => (
    <tr key={r.periodo} className="border-b border-slate-100 last:border-0">
      <td className="px-4 py-2.5 font-medium text-slate-700">{r.periodo}</td>
      <td className="px-4 py-2.5 text-right font-mono-num">{r.enviados}</td>
      <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{r.aceptados}</td>
      <td className="px-4 py-2.5 text-right font-mono-num text-rose-600">{r.rechazados}</td>
      <td className="px-4 py-2.5 text-right font-mono-num text-amber-600">{r.pendientes}</td>
      <td className="px-4 py-2.5 text-right font-mono-num">{r.pct !== null ? `${r.pct.toFixed(1)}%` : "—"}</td>
      <td className="px-4 py-2.5 text-right font-mono-num">{money(r.importeTotal)}</td>
      <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{money(r.importeAceptado)}</td>
    </tr>
  );

  const cabeceraPeriodo = (
    <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
      <th className="px-4 py-2.5 font-semibold">Periodo</th>
      <th className="px-4 py-2.5 font-semibold text-right">Enviados</th>
      <th className="px-4 py-2.5 font-semibold text-right">Aceptados</th>
      <th className="px-4 py-2.5 font-semibold text-right">Rechazados</th>
      <th className="px-4 py-2.5 font-semibold text-right">Pendientes</th>
      <th className="px-4 py-2.5 font-semibold text-right">% Aceptación</th>
      <th className="px-4 py-2.5 font-semibold text-right">Importe total</th>
      <th className="px-4 py-2.5 font-semibold text-right">Importe aceptado</th>
    </tr>
  );

  return (
    <div className="space-y-8">
      <div>
        <h2 className="font-display font-bold text-slate-800 mb-3">Resumen general</h2>
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-3">
          {porEstado.map((e) => (
            <div key={e.estado} className="bg-white border border-slate-200 rounded-lg p-4">
              <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">{e.estado}</div>
              <div className="text-2xl font-bold text-slate-800">{e.n}</div>
              <div className="text-xs text-slate-500">{e.pct.toFixed(1)}% · {money(e.importe)}</div>
            </div>
          ))}
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4 text-sm text-slate-600 mb-3">
          Total enviados: <span className="font-semibold text-slate-800">{total}</span>
          {pctAceptacion !== null && <> · % Aceptación (aceptados / aceptados+rechazados): <span className="font-semibold text-slate-800">{pctAceptacion.toFixed(1)}%</span></>}
        </div>
        {datosPie.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-4">
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={datosPie} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={90} label={(e) => `${e.name}: ${e.value}`}>
                  {datosPie.map((e) => <Cell key={e.name} fill={COLOR_ESTADO[e.name]} />)}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {rechazados.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-800 mb-3">Motivos de rechazo</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Motivo</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Nº</th>
                  <th className="px-4 py-2.5 font-semibold text-right">% sobre rechazados</th>
                </tr>
              </thead>
              <tbody>
                {porMotivo.map((m) => (
                  <tr key={m.motivo} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5">{m.motivo}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{m.n}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{m.pct.toFixed(1)}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {zonasStats.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-800 mb-3">Por zona</h2>
          <div className="bg-white border border-slate-200 rounded-lg p-4 mb-3">
            <ResponsiveContainer width="100%" height={Math.max(220, zonasStats.length * 45)}>
              <BarChart data={zonasStats} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="zona" width={110} />
                <Tooltip />
                <Legend />
                <Bar dataKey="aceptados" name="Aceptados" fill={COLOR_ESTADO.Aceptado} />
                <Bar dataKey="rechazados" name="Rechazados" fill={COLOR_ESTADO.Rechazado} />
                <Bar dataKey="enviados" name="Enviados" fill="#94a3b8" />
              </BarChart>
            </ResponsiveContainer>
          </div>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Zona</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Enviados</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Aceptados</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Rechazados</th>
                  <th className="px-4 py-2.5 font-semibold text-right">% Aceptación</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Importe aceptado</th>
                </tr>
              </thead>
              <tbody>
                {zonasStats.map((z) => (
                  <tr key={z.zona} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-medium text-slate-700">{z.zona}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{z.enviados}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{z.aceptados}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num text-rose-600">{z.rechazados}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{z.pct !== null ? `${z.pct.toFixed(1)}%` : "—"}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{money(z.importeAceptado)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {periodos.porAno.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-800 mb-3">Por año</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>{cabeceraPeriodo}</thead>
              <tbody>{periodos.porAno.map(filaPeriodo)}</tbody>
            </table>
          </div>
        </div>
      )}

      {periodos.porMes.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-800 mb-3">Por mes</h2>
          <div className="bg-white border border-slate-200 rounded-lg p-4 mb-3">
            <ResponsiveContainer width="100%" height={260}>
              <LineChart data={periodos.porMes}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="periodo" />
                <YAxis allowDecimals={false} />
                <Tooltip />
                <Legend />
                <Line type="monotone" dataKey="enviados" name="Enviados" stroke="#94a3b8" strokeWidth={2} />
                <Line type="monotone" dataKey="aceptados" name="Aceptados" stroke={COLOR_ESTADO.Aceptado} strokeWidth={2} />
                <Line type="monotone" dataKey="rechazados" name="Rechazados" stroke={COLOR_ESTADO.Rechazado} strokeWidth={2} />
              </LineChart>
            </ResponsiveContainer>
          </div>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>{cabeceraPeriodo}</thead>
              <tbody>{periodos.porMes.map(filaPeriodo)}</tbody>
            </table>
          </div>
        </div>
      )}

      {periodos.porSemana.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-800 mb-3">Por semana (lunes-domingo)</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>{cabeceraPeriodo}</thead>
              <tbody>{periodos.porSemana.map(filaPeriodo)}</tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

function PresupuestosPresentacion({ presupuestos }) {
  const COLOR_ESTADO = { Pendiente: "#f59e0b", Aceptado: "#10b981", Rechazado: "#f43f5e", "En espera": "#94a3b8" };
  const total = presupuestos.length;
  const porEstado = ESTADO_PRESUPUESTO_TRACKER.map((estado) => {
    const lista = presupuestos.filter((p) => p.estado === estado);
    return { estado, n: lista.length, importe: lista.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0) };
  });
  const datosPie = porEstado.filter((e) => e.n > 0).map((e) => ({ name: e.estado, value: e.n }));
  const aceptadosN = presupuestos.filter((p) => p.estado === "Aceptado").length;
  const rechazadosN = presupuestos.filter((p) => p.estado === "Rechazado").length;
  const pctAceptacion = (aceptadosN + rechazadosN) ? (aceptadosN / (aceptadosN + rechazadosN)) * 100 : null;
  const importeTotal = presupuestos.reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
  const importeAceptado = presupuestos.filter((p) => p.estado === "Aceptado").reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);

  const porMes = useMemo(() => {
    const map = new Map();
    presupuestos.forEach((p) => {
      if (!p.fechaEnvio) return;
      const key = p.fechaEnvio.slice(0, 7);
      if (!map.has(key)) map.set(key, []);
      map.get(key).push(p);
    });
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0])).map(([periodo, lista]) => ({
      periodo,
      enviados: lista.length,
      aceptados: lista.filter((p) => p.estado === "Aceptado").length,
      rechazados: lista.filter((p) => p.estado === "Rechazado").length,
    }));
  }, [presupuestos]);

  const zonasStats = useMemo(() => {
    const zonas = Array.from(new Set(presupuestos.map((p) => (p.zona || "").trim()).filter(Boolean))).sort();
    return zonas.map((zona) => {
      const lista = presupuestos.filter((p) => p.zona === zona);
      return {
        zona,
        aceptados: lista.filter((p) => p.estado === "Aceptado").length,
        rechazados: lista.filter((p) => p.estado === "Rechazado").length,
        enviados: lista.length,
      };
    });
  }, [presupuestos]);

  return (
    <div className="bg-white rounded-xl border border-slate-200 p-10 space-y-10">
      <div className="text-center">
        <h1 className="font-display text-3xl font-extrabold text-slate-900 mb-1">Control de Presupuestos — ALUMAVEL</h1>
        <p className="text-slate-400 text-sm">Generado el {fmtDate(new Date().toISOString().slice(0, 10))}</p>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-5 gap-4">
        <div className="text-center bg-slate-50 rounded-xl py-6">
          <div className="text-4xl font-extrabold text-slate-800">{total}</div>
          <div className="text-xs uppercase tracking-wide text-slate-400 mt-1">Enviados</div>
        </div>
        <div className="text-center bg-emerald-50 rounded-xl py-6">
          <div className="text-4xl font-extrabold text-emerald-600">{aceptadosN}</div>
          <div className="text-xs uppercase tracking-wide text-emerald-500 mt-1">Aceptados</div>
        </div>
        <div className="text-center bg-rose-50 rounded-xl py-6">
          <div className="text-4xl font-extrabold text-rose-600">{rechazadosN}</div>
          <div className="text-xs uppercase tracking-wide text-rose-500 mt-1">Rechazados</div>
        </div>
        <div className="text-center bg-amber-50 rounded-xl py-6">
          <div className="text-4xl font-extrabold text-amber-600">{pctAceptacion !== null ? `${pctAceptacion.toFixed(0)}%` : "—"}</div>
          <div className="text-xs uppercase tracking-wide text-amber-500 mt-1">% Aceptación</div>
        </div>
        <div className="text-center bg-sky-50 rounded-xl py-6">
          <div className="text-3xl font-extrabold text-sky-700">{money(importeAceptado)}</div>
          <div className="text-xs uppercase tracking-wide text-sky-500 mt-1">Importe aceptado</div>
          <div className="text-[11px] text-sky-400 mt-0.5">de {money(importeTotal)} enviado</div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
        {datosPie.length > 0 && (
          <div>
            <h2 className="font-display font-bold text-slate-700 text-center mb-2">Distribución por estado</h2>
            <ResponsiveContainer width="100%" height={300}>
              <PieChart>
                <Pie data={datosPie} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={110} label={(e) => `${e.name}: ${e.value}`}>
                  {datosPie.map((e) => <Cell key={e.name} fill={COLOR_ESTADO[e.name]} />)}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
        {zonasStats.length > 0 && (
          <div>
            <h2 className="font-display font-bold text-slate-700 text-center mb-2">Por zona</h2>
            <ResponsiveContainer width="100%" height={300}>
              <BarChart data={zonasStats} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="zona" width={110} />
                <Tooltip />
                <Legend />
                <Bar dataKey="aceptados" name="Aceptados" fill={COLOR_ESTADO.Aceptado} />
                <Bar dataKey="rechazados" name="Rechazados" fill={COLOR_ESTADO.Rechazado} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {porMes.length > 0 && (
        <div>
          <h2 className="font-display font-bold text-slate-700 text-center mb-2">Evolución mensual</h2>
          <ResponsiveContainer width="100%" height={320}>
            <LineChart data={porMes}>
              <CartesianGrid strokeDasharray="3 3" />
              <XAxis dataKey="periodo" />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Legend />
              <Line type="monotone" dataKey="enviados" name="Enviados" stroke="#94a3b8" strokeWidth={3} dot={{ r: 4 }} />
              <Line type="monotone" dataKey="aceptados" name="Aceptados" stroke={COLOR_ESTADO.Aceptado} strokeWidth={3} dot={{ r: 4 }} />
              <Line type="monotone" dataKey="rechazados" name="Rechazados" stroke={COLOR_ESTADO.Rechazado} strokeWidth={3} dot={{ r: 4 }} />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

function IngresosModulo({ ingresos, clientes, proyectos, view, setView, editId, setEditId, onUpsert, onDelete, onVincular, onRegistrarIngreso, prefill, onClearPrefill, isAdmin }) {
  const [q, setQ] = useState("");
  const sinProyecto = ingresos.filter((i) => !i.proyectoId);
  const totalSinProyecto = sinProyecto.reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);

  const proyectosConSaldo = useMemo(() => {
    return proyectos
      .map((p) => {
        const importePresupuesto = parseFloat(p.importePresupuesto) || 0;
        const recibido = ingresos.filter((i) => i.proyectoId === p.id).reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);
        return { proyecto: p, importePresupuesto, recibido, pendiente: importePresupuesto - recibido };
      })
      .filter((r) => r.importePresupuesto > 0 && r.pendiente > 0.01)
      .sort((a, b) => b.pendiente - a.pendiente);
  }, [proyectos, ingresos]);
  const totalPendienteProyectos = proyectosConSaldo.reduce((s, r) => s + r.pendiente, 0);

  const filtered = useMemo(() => {
    return ingresos.filter((i) => {
      if (!q) return true;
      const hay = `${i.clienteNombre} ${i.concepto}`.toLowerCase();
      return hay.includes(q.toLowerCase());
    });
  }, [ingresos, q]);

  if (view === "form") {
    const editing = ingresos.find((i) => i.id === editId) || null;
    return (
      <IngresoForm
        initial={editing || prefill}
        clientes={clientes}
        proyectos={proyectos}
        onCancel={() => { setView("list"); onClearPrefill && onClearPrefill(); }}
        onSave={(data) => { onUpsert(data); onClearPrefill && onClearPrefill(); }}
      />
    );
  }

  return (
    <div className="p-8">
      <Header
        icon={<Wallet size={20} className="text-[#2E8B57]" />}
        title="Entrada de dinero"
        manualKey="ingresos"
        subtitle={`${ingresos.length} entrada${ingresos.length === 1 ? "" : "s"} registrada${ingresos.length === 1 ? "" : "s"}`}
      />

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVA ENTRADA DE DINERO
      </button>

      <div className="px-4 py-3 rounded-md bg-sky-50 border border-sky-200 text-sky-800 text-sm mb-6">
        Usa esto cuando un cliente te dé dinero (señal, anticipo) <strong>antes</strong> de tener el proyecto dado de alta en el CRM. En cuanto crees el proyecto, vuelve aquí y vincúlalo — se generará la factura automáticamente por ese importe.
      </div>

      {sinProyecto.length > 0 && (
        <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 text-amber-800 text-sm font-semibold mb-6">
          ⚠ {sinProyecto.length} entrada{sinProyecto.length === 1 ? "" : "s"} sin proyecto asignado todavía — {money(totalSinProyecto)} pendientes de vincular.
        </div>
      )}

      {proyectosConSaldo.length > 0 && (
        <div className="mb-6">
          <h2 className="font-display font-bold text-slate-800 mb-2">Proyectos con dinero pendiente de cobrar</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Proyecto</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Presupuesto</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Recibido</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Pendiente</th>
                  <th className="px-4 py-2.5 font-semibold"></th>
                </tr>
              </thead>
              <tbody>
                {proyectosConSaldo.map((r) => (
                  <tr key={r.proyecto.id} className="border-b border-slate-100 last:border-0">
                    <td className="px-4 py-2.5 font-medium text-slate-800">#{r.proyecto.numero} — {r.proyecto.nombre}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num text-slate-500">{money(r.importePresupuesto)}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{money(r.recibido)}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num font-bold text-amber-700">{money(r.pendiente)}</td>
                    <td className="px-4 py-2.5 text-right">
                      <button
                        onClick={() => onRegistrarIngreso(r.proyecto, clientes.find((c) => c.id === r.proyecto.clienteId)?.nombre, r.pendiente)}
                        style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
                        className="text-xs font-semibold hover:opacity-90 px-3 py-1.5 rounded-md"
                      >
                        Registrar este importe
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-slate-400 mt-1.5">Total pendiente de cobrar en proyectos: <span className="font-semibold text-slate-600">{money(totalPendienteProyectos)}</span></p>
        </div>
      )}

      <div className="relative mb-4 max-w-sm">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Buscar por cliente o concepto..." className={inputCls + " pl-9"} />
      </div>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Fecha</th>
              <th className="px-4 py-3 font-semibold">Cliente</th>
              <th className="px-4 py-3 font-semibold">Concepto</th>
              <th className="px-4 py-3 font-semibold text-right">Importe</th>
              <th className="px-4 py-3 font-semibold">Forma de pago</th>
              <th className="px-4 py-3 font-semibold">Proyecto</th>
              <th className="px-4 py-3 font-semibold"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((i) => {
              const proyecto = proyectos.find((p) => p.id === i.proyectoId);
              return (
                <tr key={i.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 transition">
                  <td className="px-4 py-3 text-slate-600">{fmtDate(i.fecha)}</td>
                  <td className="px-4 py-3 font-medium text-slate-800">{i.clienteNombre}</td>
                  <td className="px-4 py-3 text-slate-600">{i.concepto || "—"}</td>
                  <td className="px-4 py-3 text-right font-mono-num">{money(i.importe)}</td>
                  <td className="px-4 py-3 text-slate-500">{i.formaPago || "—"}</td>
                  <td className="px-4 py-3">
                    {proyecto ? (
                      <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">#{proyecto.numero} — {proyecto.nombre}</Badge>
                    ) : (
                      <VincularProyectoSelect proyectos={proyectos} onVincular={(proyectoId) => onVincular(i.id, proyectoId)} />
                    )}
                  </td>
                  <td className="px-4 py-3 flex gap-2 justify-end">
                    <button onClick={() => { setEditId(i.id); setView("form"); }} className="text-slate-400 hover:text-slate-700"><Pencil size={15} /></button>
                    {isAdmin && <button onClick={() => onDelete(i.id)} className="text-slate-400 hover:text-rose-600"><Trash2 size={15} /></button>}
                  </td>
                </tr>
              );
            })}
            {filtered.length === 0 && (
              <tr><td colSpan={7} className="px-4 py-10 text-center text-slate-400 text-sm">No hay entradas de dinero registradas.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function VincularProyectoSelect({ proyectos, onVincular }) {
  const [sel, setSel] = useState("");
  return (
    <div className="flex items-center gap-1.5">
      <Select value={sel} onChange={(e) => setSel(e.target.value)} className="max-w-[160px] !py-1 !text-xs">
        <option value="">Vincular a...</option>
        {proyectos.map((p) => <option key={p.id} value={p.id}>#{p.numero} — {p.nombre}</option>)}
      </Select>
      <button
        onClick={() => sel && onVincular(sel)}
        disabled={!sel}
        title={!sel ? "Elige primero un proyecto de la lista" : ""}
        style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
        className="text-xs font-semibold disabled:bg-slate-300 px-2.5 py-1.5 rounded-md"
      >
        Vincular
      </button>
    </div>
  );
}

function IngresoForm({ initial, clientes, proyectos, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, fecha: new Date().toISOString().slice(0, 10), clienteNombre: "", importe: "",
      concepto: "Señal / Anticipo", formaPago: "Transferencia", proyectoId: "", notas: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const clientesDisponibles = useMemo(() => clientes.map((c) => c.nombre).sort(), [clientes]);

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.clienteNombre.trim()) { setErrorMsg("Falta el nombre del cliente (es obligatorio)."); return; }
    if (!f.importe || parseFloat(f.importe) <= 0) { setErrorMsg("El importe tiene que ser mayor que 0."); return; }
    if (!f.fecha) { setErrorMsg("Falta la fecha (es obligatoria)."); return; }
    setErrorMsg("");
    onSave({ ...f, importe: parseFloat(f.importe) || 0, proyectoId: f.proyectoId || null });
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">{initial?.id ? "Editar entrada de dinero" : "Nueva entrada de dinero"}</h1>

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Fecha" required>
            <TextInput type="date" value={f.fecha} onChange={set("fecha")} />
          </Field>
          <Field label="Importe (€)" required>
            <TextInput type="number" step="0.01" value={f.importe} onChange={set("importe")} />
          </Field>
        </div>

        <Field label="Cliente" required>
          <TextInput value={f.clienteNombre} onChange={set("clienteNombre")} list="ingreso-clientes-datalist" placeholder="Nombre del cliente" />
          <datalist id="ingreso-clientes-datalist">
            {clientesDisponibles.map((n) => <option key={n} value={n} />)}
          </datalist>
        </Field>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Concepto">
            <TextInput value={f.concepto} onChange={set("concepto")} placeholder="Ej: Señal / Anticipo" />
          </Field>
          <Field label="Forma de pago">
            <Select value={f.formaPago} onChange={set("formaPago")}>
              {FORMA_PAGO.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
        </div>

        <Field label="Notas">
          <TextArea rows={2} value={f.notas} onChange={set("notas")} />
        </Field>

        <Field label="Proyecto (déjalo en blanco si todavía no existe)">
          <Select value={f.proyectoId || ""} onChange={set("proyectoId")}>
            <option value="">Sin proyecto todavía</option>
            {proyectos.map((p) => <option key={p.id} value={p.id}>#{p.numero} — {p.nombre}</option>)}
          </Select>
          {!initial?.id && <p className="text-xs text-slate-400 mt-1">Si eliges un proyecto ahora, se generará la factura al guardar.</p>}
          {initial?.vinculado && <p className="text-xs text-amber-600 mt-1">Esta entrada ya generó una factura. Si cambias el proyecto aquí, solo se actualiza el vínculo — no se genera ni se mueve ninguna factura nueva.</p>}
        </Field>

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar entrada</button>
        </div>
      </form>
    </div>
  );
}

const PERIODOS_INFORME = [
  { id: "mes", label: "Este mes" },
  { id: "trimestre", label: "Este trimestre" },
  { id: "ano", label: "Este año" },
  { id: "todo", label: "Todo" },
];

const rangoPeriodo = (periodoId, offsetPeriodos = 0) => {
  const hoy = new Date();
  if (periodoId === "todo") return { desde: null, hasta: null };
  if (periodoId === "mes") {
    const mes = hoy.getMonth() + offsetPeriodos;
    const desde = new Date(hoy.getFullYear(), mes, 1);
    const hasta = new Date(hoy.getFullYear(), mes + 1, 0);
    return { desde, hasta };
  }
  if (periodoId === "trimestre") {
    const trimActual = Math.floor(hoy.getMonth() / 3);
    const mesInicio = (trimActual + offsetPeriodos) * 3;
    const desde = new Date(hoy.getFullYear(), mesInicio, 1);
    const hasta = new Date(hoy.getFullYear(), mesInicio + 3, 0);
    return { desde, hasta };
  }
  if (periodoId === "ano") {
    const desde = new Date(hoy.getFullYear() + offsetPeriodos, 0, 1);
    const hasta = new Date(hoy.getFullYear() + offsetPeriodos, 11, 31);
    return { desde, hasta };
  }
  return { desde: null, hasta: null };
};

const enRango = (fechaStr, rango) => {
  if (!fechaStr) return false;
  if (!rango.desde) return true;
  const f = new Date(fechaStr);
  return f >= rango.desde && f <= rango.hasta;
};

function InformesModulo({ proyectos, presupuestos, ingresos, facturas, incidencias, pedidos, clientes, materiales, instalaciones, usuarios, sesionesUsuario, isAdmin }) {
  const COLOR_ESTADO = { Pendiente: "#f59e0b", Aceptado: "#10b981", Rechazado: "#f43f5e", "En espera": "#94a3b8" };
  const [periodo, setPeriodo] = useState("mes");
  const rango = rangoPeriodo(periodo, 0);
  const rangoAnterior = rangoPeriodo(periodo, -1);

  // Datos de "Control de montaje por vivienda" (una entrada por instalación),
  // se cargan una vez para poder cruzar horas trabajadas con viviendas completadas.
  const [controlMontajeData, setControlMontajeData] = useState({});
  useEffect(() => {
    (async () => {
      try {
        const snap = await fbGet(ref(fbDb, "controlMontaje"));
        setControlMontajeData(snap.val() || {});
      } catch (e) {
        console.error(e);
      }
    })();
  }, []);

  const presupActual = presupuestos.filter((p) => enRango(p.fechaEnvio, rango));
  const presupAnterior = presupuestos.filter((p) => enRango(p.fechaEnvio, rangoAnterior));
  const ingresosActual = ingresos.filter((i) => enRango(i.fecha, rango));
  const ingresosAnterior = ingresos.filter((i) => enRango(i.fecha, rangoAnterior));
  const proyectosActual = proyectos.filter((p) => enRango(p.fechaSolicitud, rango));

  const aceptadosN = presupActual.filter((p) => p.estado === "Aceptado").length;
  const sinAceptarN = presupActual.filter((p) => p.estado === "Pendiente" || p.estado === "En espera").length;
  const rechazadosN = presupActual.filter((p) => p.estado === "Rechazado").length;

  const totalIngresado = ingresosActual.reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);
  const totalIngresadoAnterior = ingresosAnterior.reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);
  const deltaIngresado = totalIngresadoAnterior ? ((totalIngresado - totalIngresadoAnterior) / totalIngresadoAnterior) * 100 : null;
  const sinVincular = ingresos.filter((i) => !i.proyectoId);
  const totalSinVincular = sinVincular.reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);

  const facturasActual = facturas.filter((f) => enRango(f.fecha, rango));
  const totalFacturado = facturasActual.reduce((s, f) => s + (parseFloat(f.total) || 0), 0);
  const totalCobrado = facturasActual.reduce((s, f) => s + (f.pagos || []).reduce((s2, p) => s2 + (parseFloat(p.importe) || 0), 0), 0);
  const totalPendienteCobro = totalFacturado - totalCobrado;

  const datosPieProyectos = [
    { name: "Aceptado", value: aceptadosN },
    { name: "Sin aceptar", value: sinAceptarN },
    { name: "Rechazado", value: rechazadosN },
  ].filter((e) => e.value > 0);

  const rechazadosMotivo = presupActual.filter((p) => p.estado === "Rechazado");
  const porMotivo = MOTIVO_RECHAZO.map((m) => ({ motivo: m, n: rechazadosMotivo.filter((p) => p.motivoRechazo === m).length })).filter((m) => m.n > 0);

  const porZona = useMemo(() => {
    const zonas = Array.from(new Set(presupActual.map((p) => (p.zona || "").trim()).filter(Boolean))).sort();
    return zonas.map((zona) => {
      const lista = presupActual.filter((p) => p.zona === zona);
      return { zona, aceptados: lista.filter((p) => p.estado === "Aceptado").length, rechazados: lista.filter((p) => p.estado === "Rechazado").length, enviados: lista.length };
    });
  }, [presupActual]);

  const evolucionMensual = useMemo(() => {
    const map = new Map();
    ingresos.forEach((i) => {
      if (!i.fecha) return;
      const key = i.fecha.slice(0, 7);
      map.set(key, (map.get(key) || 0) + (parseFloat(i.importe) || 0));
    });
    return Array.from(map.entries()).sort((a, b) => a[0].localeCompare(b[0])).slice(-6).map(([mes, importe]) => ({ mes, importe }));
  }, [ingresos]);

  const porEstadoTrabajo = ESTADO_TRABAJO.map((estado) => ({ estado, n: proyectosActual.filter((p) => p.estadoTrabajo === estado).length })).filter((e) => e.n > 0);

  const incidenciasActual = incidencias.filter((i) => enRango(i.fecha, rango));
  const porEstadoIncidencia = ESTADO_INCIDENCIA.map((estado) => ({ estado, n: incidenciasActual.filter((i) => i.estadoIncidencia === estado).length })).filter((e) => e.n > 0);
  const gastosIncidenciasTotal = incidenciasActual.reduce((s, i) => s + (i.gastos || []).reduce((s2, g) => s2 + (parseFloat(g.importe) || 0), 0), 0);

  const pedidosActual = pedidos.filter((p) => enRango(p.fechaCompra, rango));
  const porEstadoPedido = ESTADO_PEDIDO.map((estado) => ({ estado, n: pedidosActual.filter((p) => p.estado === estado).length })).filter((e) => e.n > 0);
  const importePedidosActual = pedidosActual.reduce((s, p) => s + (p.lineas || []).reduce((s2, l) => s2 + (parseFloat(l.cantidad) || 0) * (parseFloat(l.precio) || 0), 0), 0);

  const ultimosMovimientos = [...ingresos].sort((a, b) => (b.fecha || "").localeCompare(a.fecha || "")).slice(0, 8);

  const formatoHorasMin = (ms) => {
    if (!ms || ms <= 0) return "0 min";
    const totalMin = Math.round(ms / 60000);
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    if (h === 0) return `${m} min`;
    return `${h} h ${m} min`;
  };

  const resumenConexion = useMemo(() => {
    const inicioHoy = new Date();
    inicioHoy.setHours(0, 0, 0, 0);
    const tsInicioHoy = inicioHoy.getTime();
    const inicioSemana = new Date(inicioHoy);
    inicioSemana.setDate(inicioSemana.getDate() - 6);
    const tsInicioSemana = inicioSemana.getTime();

    return (usuarios || [])
      .map((u) => {
        const sesiones = Object.values((sesionesUsuario && sesionesUsuario[u.id]) || {});
        let total = 0, hoy = 0, ultimos7 = 0, ultimaConexion = null;
        sesiones.forEach((s) => {
          const inicio = s.inicio || 0;
          const fin = Math.max(s.fin || s.inicio || 0, inicio);
          total += fin - inicio;
          if (fin >= tsInicioHoy) hoy += Math.max(fin - Math.max(inicio, tsInicioHoy), 0);
          if (fin >= tsInicioSemana) ultimos7 += Math.max(fin - Math.max(inicio, tsInicioSemana), 0);
          if (!ultimaConexion || fin > ultimaConexion) ultimaConexion = fin;
        });
        return { usuario: u, total, hoy, ultimos7, ultimaConexion };
      })
      .sort((a, b) => b.total - a.total);
  }, [usuarios, sesionesUsuario]);


  const comparacionFabrica = useMemo(() => {
    const filas = [];
    pedidos.forEach((p) => {
      if (p.estado === "Cancelado") return;
      (p.lineas || []).forEach((l) => {
        const oficinaRecibido = p.estado === "Recibido" || l.estado === "Recibido";
        const fabricaConfirmado = !!l.confirmadoFabrica;
        if (!oficinaRecibido && !fabricaConfirmado) return;
        const esLibre = l.modo === "libre";
        const mat = esLibre ? null : materiales.find((m) => m.id === l.materialId);
        const nombre = esLibre ? (l.referencia || "Sin referencia") : (mat ? mat.descripcion : "Material eliminado");
        filas.push({ pedido: p.numero, material: nombre, oficinaRecibido, fabricaConfirmado, coincide: oficinaRecibido === fabricaConfirmado });
      });
    });
    return filas;
  }, [pedidos, materiales]);
  const discrepanciasFabrica = comparacionFabrica.filter((f) => !f.coincide);

  const instalacionesConDatos = useMemo(() => {
    return (instalaciones || []).map((inst) => {
      const proyecto = proyectos.find((p) => p.id === inst.proyectoId);
      const totalHoras = (inst.registroHoras || []).reduce((s, r) => s + (parseFloat(r.horas) || 0), 0);
      const costeManoObra = totalHoras * (parseFloat(inst.costeHora) || 0);
      const totalGastos = (inst.gastos || []).reduce((s, g) => s + (parseFloat(g.importe) || 0), 0);
      const costeTotal = costeManoObra + totalGastos;
      const presupuesto = parseFloat(inst.presupuestoInstalacion) || 0;
      const diferencia = presupuesto - costeTotal;
      return { inst, proyecto, totalHoras, costeTotal, presupuesto, diferencia };
    });
  }, [instalaciones, proyectos]);
  const porEstadoInstalacion = ESTADO_INSTALACION.map((estado) => ({ estado, n: instalacionesConDatos.filter((i) => i.inst.estado === estado).length })).filter((e) => e.n > 0);
  const totalPresupuestadoInstalacion = instalacionesConDatos.reduce((s, i) => s + i.presupuesto, 0);
  const totalCosteInstalacion = instalacionesConDatos.reduce((s, i) => s + i.costeTotal, 0);
  const instalacionesSobrepresupuesto = instalacionesConDatos.filter((i) => i.presupuesto > 0 && i.diferencia < 0);

  // -------- Productividad: horas por instalador, dentro del periodo seleccionado --------
  const productividadInstaladores = useMemo(() => {
    const map = new Map();
    (instalaciones || []).forEach((inst) => {
      (inst.registroHoras || []).forEach((r) => {
        if (!enRango(r.fecha, rango)) return;
        const nombre = (r.instalador || "Sin nombre").trim();
        const actual = map.get(nombre) || { instalador: nombre, horas: 0, dias: new Set() };
        actual.horas += parseFloat(r.horas) || 0;
        actual.dias.add(r.fecha);
        map.set(nombre, actual);
      });
    });
    return Array.from(map.values())
      .map((x) => ({ instalador: x.instalador, horas: x.horas, dias: x.dias.size }))
      .sort((a, b) => b.horas - a.horas);
  }, [instalaciones, rango]);
  const totalHorasPeriodo = productividadInstaladores.reduce((s, x) => s + x.horas, 0);

  // -------- Productividad: horas por vivienda completada, cruzando con el control de montaje --------
  const productividadPorObra = useMemo(() => {
    return instalacionesConDatos
      .map((i) => {
        const datos = controlMontajeData[i.inst.id];
        const viviendas = toArray(datos?.viviendas);
        const viviendasCompletas = viviendas.filter((v) => v.elementos && v.elementos.length > 0 && v.elementos.every((el) => el.instalado)).length;
        const horasPorVivienda = viviendasCompletas > 0 ? i.totalHoras / viviendasCompletas : null;
        return {
          nombre: i.proyecto ? `#${i.proyecto.numero} — ${i.proyecto.nombre}` : (i.inst.nombre || "Instalación sin proyecto"),
          totalHoras: i.totalHoras,
          totalViviendas: viviendas.length,
          viviendasCompletas,
          horasPorVivienda,
        };
      })
      .filter((x) => x.totalViviendas > 0);
  }, [instalacionesConDatos, controlMontajeData]);

  const [filtroClientes, setFiltroClientes] = useState("pendiente");
  const situacionClientes = useMemo(() => {
    return clientes
      .map((c) => {
        const proyectosCliente = proyectos.filter((p) => p.clienteId === c.id);
        if (proyectosCliente.length === 0) return null;
        const idsProyectos = proyectosCliente.map((p) => p.id);
        const presupuestado = proyectosCliente.reduce((s, p) => s + (parseFloat(p.importePresupuesto) || 0), 0);
        const recibido = ingresos.filter((i) => idsProyectos.includes(i.proyectoId)).reduce((s, i) => s + (parseFloat(i.importe) || 0), 0);
        const pendiente = presupuestado - recibido;
        const incidenciasAbiertas = incidencias.filter((i) => idsProyectos.includes(i.proyectoId) && i.estadoIncidencia !== "Solucionado").length;
        const pedidosPendientes = pedidos.filter((pd) => idsProyectos.includes(pd.proyectoId) && pd.estado !== "Recibido" && pd.estado !== "Cancelado").length;
        const limiteCredito = parseFloat(c.limiteCredito) || 0;
        const pctRiesgo = limiteCredito ? (pendiente / limiteCredito) * 100 : null;
        return { cliente: c, nProyectos: proyectosCliente.length, presupuestado, recibido, pendiente, incidenciasAbiertas, pedidosPendientes, pctRiesgo };
      })
      .filter(Boolean)
      .filter((r) => filtroClientes === "todos" || (filtroClientes === "pendiente" ? r.pendiente > 0.01 : r.pendiente <= 0.01))
      .sort((a, b) => b.pendiente - a.pendiente);
  }, [clientes, proyectos, ingresos, incidencias, pedidos, filtroClientes]);
  const totalPendienteClientes = situacionClientes.reduce((s, r) => s + Math.max(r.pendiente, 0), 0);

  const descargarInformeWord = () => {
    const etiquetaPeriodo = PERIODOS_INFORME.find((p) => p.id === periodo)?.label || periodo;
    const filaTabla = (cols) => `<tr>${cols.map((c) => `<td style="border:1px solid #ccc;padding:6px 10px;">${c}</td>`).join("")}</tr>`;
    const cabeceraTabla = (cols) => `<tr>${cols.map((c) => `<th style="border:1px solid #ccc;padding:6px 10px;background:#f1f1f1;text-align:left;">${c}</th>`).join("")}</tr>`;

    const htmlDoc = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
      <head><meta charset="utf-8"><title>Informe ALUMAVEL</title></head>
      <body style="font-family: Calibri, Arial, sans-serif; color:#1a1a1a;">
        <h1 style="color:#2E8B57;">Informe ALUMAVEL — ${etiquetaPeriodo}</h1>
        <p style="color:#666;">Generado el ${fmtDate(new Date().toISOString().slice(0, 10))}</p>

        <h2>Resumen general</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Indicador", "Valor"])}
          ${filaTabla(["Presupuestos aceptados", aceptadosN])}
          ${filaTabla(["Presupuestos sin aceptar / en espera", sinAceptarN])}
          ${filaTabla(["Presupuestos rechazados", rechazadosN])}
          ${filaTabla(["Dinero ingresado", money(totalIngresado)])}
          ${deltaIngresado !== null ? filaTabla(["Variación vs periodo anterior", `${deltaIngresado >= 0 ? "+" : ""}${deltaIngresado.toFixed(0)}%`]) : ""}
          ${filaTabla(["Sin vincular a proyecto", `${money(totalSinVincular)} (${sinVincular.length} entradas)`])}
          ${filaTabla(["Facturado", money(totalFacturado)])}
          ${filaTabla(["Cobrado", money(totalCobrado)])}
          ${filaTabla(["Pendiente de cobro", money(totalPendienteCobro)])}
        </table>

        ${porMotivo.length > 0 ? `
        <h2>Motivos de rechazo</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Motivo", "Nº"])}
          ${porMotivo.map((m) => filaTabla([m.motivo, m.n])).join("")}
        </table>` : ""}

        ${porZona.length > 0 ? `
        <h2>Por zona</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Zona", "Enviados", "Aceptados", "Rechazados"])}
          ${porZona.map((z) => filaTabla([z.zona, z.enviados, z.aceptados, z.rechazados])).join("")}
        </table>` : ""}

        ${porEstadoTrabajo.length > 0 ? `
        <h2>Proyectos por estado de trabajo</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Estado", "Nº"])}
          ${porEstadoTrabajo.map((e) => filaTabla([e.estado, e.n])).join("")}
        </table>` : ""}

        ${porEstadoIncidencia.length > 0 ? `
        <h2>Incidencias</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Estado", "Nº"])}
          ${porEstadoIncidencia.map((e) => filaTabla([e.estado, e.n])).join("")}
        </table>
        ${gastosIncidenciasTotal > 0 ? `<p>Gastos asociados a incidencias: <b>${money(gastosIncidenciasTotal)}</b></p>` : ""}` : ""}

        ${porEstadoPedido.length > 0 ? `
        <h2>Pedidos de materiales</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Estado", "Nº"])}
          ${porEstadoPedido.map((e) => filaTabla([e.estado, e.n])).join("")}
        </table>
        ${importePedidosActual > 0 ? `<p>Importe total de pedidos: <b>${money(importePedidosActual)}</b></p>` : ""}` : ""}

        ${comparacionFabrica.length > 0 ? `
        <h2>Comparación oficina ↔ fábrica</h2>
        <p>${discrepanciasFabrica.length === 0 ? "Todo coincide, sin discrepancias." : `${discrepanciasFabrica.length} discrepancia(s) encontrada(s).`}</p>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Pedido", "Material", "Oficina", "Fábrica", "¿Coincide?"])}
          ${comparacionFabrica.map((f) => filaTabla([`#${f.pedido}`, f.material, f.oficinaRecibido ? "Recibido" : "No recibido", f.fabricaConfirmado ? "Ha llegado" : "No confirmado", f.coincide ? "Sí" : "NO"])).join("")}
        </table>` : ""}

        ${instalacionesConDatos.length > 0 ? `
        <h2>Instalaciones</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Proyecto", "Estado", "Presupuestado", "Coste real", "Diferencia"])}
          ${instalacionesConDatos.map((i) => filaTabla([i.proyecto ? `#${i.proyecto.numero} — ${i.proyecto.nombre}` : (i.inst.nombre || "Instalación sin proyecto"), i.inst.estado, money(i.presupuesto), money(i.costeTotal), i.presupuesto ? money(i.diferencia) : "—"])).join("")}
        </table>` : ""}

        ${situacionClientes.length > 0 ? `
        <h2>Situación de pago por cliente</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Cliente", "Proyectos", "Presupuestado", "Recibido", "Pendiente", "Incidencias abiertas", "Pedidos pendientes"])}
          ${situacionClientes.map((r) => filaTabla([r.cliente.nombre, r.nProyectos, money(r.presupuestado), money(r.recibido), r.pendiente > 0.01 ? money(r.pendiente) : "Al día", r.incidenciasAbiertas, r.pedidosPendientes])).join("")}
        </table>` : ""}

        <h2>Últimos movimientos de dinero</h2>
        <table style="border-collapse:collapse;width:100%;">
          ${cabeceraTabla(["Fecha", "Cliente", "Importe", "Estado"])}
          ${ultimosMovimientos.map((i) => filaTabla([fmtDate(i.fecha), i.clienteNombre, money(i.importe), i.proyectoId ? "Vinculado" : "Sin vincular"])).join("")}
        </table>
      </body>
      </html>
    `;

    const blob = new Blob(["\ufeff", htmlDoc], { type: "application/msword" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `informe_alumavel_${periodo}_${new Date().toISOString().slice(0, 10)}.doc`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="p-8">
      <Header icon={<BarChart3 size={20} className="text-[#2E8B57]" />} title="Informes" manualKey="informes" subtitle="Vista general de proyectos, presupuestos y dinero" />

      <div className="flex flex-wrap items-center gap-2 mb-6">
        {PERIODOS_INFORME.map((p) => (
          <button key={p.id} onClick={() => setPeriodo(p.id)}
            className={`px-4 py-2 rounded-md text-sm font-semibold transition ${periodo === p.id ? "bg-[#2E8B57] text-white" : "bg-white border border-slate-300 text-slate-600 hover:bg-slate-50"}`}>
            {p.label}
          </button>
        ))}
        <button onClick={descargarInformeWord} className="ml-auto flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50">
          <FileText size={14} /> Descargar informe (Word)
        </button>
      </div>

      <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
        <div className="bg-emerald-50 rounded-xl p-4">
          <div className="text-3xl font-extrabold text-emerald-600">{aceptadosN}</div>
          <div className="text-xs uppercase tracking-wide text-emerald-500 mt-1">Presupuestos aceptados</div>
        </div>
        <div className="bg-amber-50 rounded-xl p-4">
          <div className="text-3xl font-extrabold text-amber-600">{sinAceptarN}</div>
          <div className="text-xs uppercase tracking-wide text-amber-500 mt-1">Sin aceptar / en espera</div>
        </div>
        <div className="bg-sky-50 rounded-xl p-4">
          <div className="text-2xl font-extrabold text-sky-700">{money(totalIngresado)}</div>
          <div className="text-xs uppercase tracking-wide text-sky-500 mt-1">
            Ingresado {deltaIngresado !== null && (
              <span className={deltaIngresado >= 0 ? "text-emerald-600" : "text-rose-600"}>
                ({deltaIngresado >= 0 ? "+" : ""}{deltaIngresado.toFixed(0)}% vs anterior)
              </span>
            )}
          </div>
        </div>
        <div className="bg-rose-50 rounded-xl p-4">
          <div className="text-2xl font-extrabold text-rose-600">{money(totalSinVincular)}</div>
          <div className="text-xs uppercase tracking-wide text-rose-500 mt-1">Sin vincular a proyecto ({sinVincular.length})</div>
        </div>
      </div>

      <div className="grid grid-cols-3 gap-4 mb-8">
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Facturado</div>
          <div className="text-xl font-bold text-slate-800">{money(totalFacturado)}</div>
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Cobrado</div>
          <div className="text-xl font-bold text-emerald-600">{money(totalCobrado)}</div>
        </div>
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Pendiente de cobro</div>
          <div className="text-xl font-bold text-rose-600">{money(totalPendienteCobro)}</div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
        {datosPieProyectos.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-4">
            <h2 className="font-display font-bold text-slate-700 text-sm mb-2">Presupuestos por estado</h2>
            <ResponsiveContainer width="100%" height={240}>
              <PieChart>
                <Pie data={datosPieProyectos} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={85} label={(e) => `${e.name}: ${e.value}`}>
                  {datosPieProyectos.map((e) => <Cell key={e.name} fill={COLOR_ESTADO[e.name === "Sin aceptar" ? "Pendiente" : e.name]} />)}
                </Pie>
                <Tooltip />
                <Legend />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
        {evolucionMensual.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-4">
            <h2 className="font-display font-bold text-slate-700 text-sm mb-2">Evolución de ingresos (últimos meses)</h2>
            <ResponsiveContainer width="100%" height={240}>
              <BarChart data={evolucionMensual}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis dataKey="mes" />
                <YAxis />
                <Tooltip formatter={(v) => money(v)} />
                <Bar dataKey="importe" name="Ingresado" fill="#2E8B57" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
        {porMotivo.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-4">
            <h2 className="font-display font-bold text-slate-700 text-sm mb-2">Motivos de rechazo (periodo)</h2>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={porMotivo} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="motivo" width={90} />
                <Tooltip />
                <Bar dataKey="n" name="Nº" fill="#f43f5e" radius={[0, 4, 4, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
        {porZona.length > 0 && (
          <div className="bg-white border border-slate-200 rounded-lg p-4">
            <h2 className="font-display font-bold text-slate-700 text-sm mb-2">Por zona (periodo)</h2>
            <ResponsiveContainer width="100%" height={200}>
              <BarChart data={porZona} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" allowDecimals={false} />
                <YAxis type="category" dataKey="zona" width={90} />
                <Tooltip />
                <Legend />
                <Bar dataKey="aceptados" name="Aceptados" fill={COLOR_ESTADO.Aceptado} />
                <Bar dataKey="rechazados" name="Rechazados" fill={COLOR_ESTADO.Rechazado} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      {porEstadoTrabajo.length > 0 && (
        <div className="bg-white border border-slate-200 rounded-lg p-4 mb-8">
          <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Proyectos por estado de trabajo (dados de alta en el periodo)</h2>
          <div className="flex flex-wrap gap-3">
            {porEstadoTrabajo.map((e) => (
              <div key={e.estado} className="bg-slate-50 rounded-lg px-4 py-2.5">
                <div className="text-lg font-bold text-slate-800">{e.n}</div>
                <div className="text-[11px] text-slate-500">{e.estado}</div>
              </div>
            ))}
          </div>
        </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Incidencias (periodo)</h2>
          {porEstadoIncidencia.length === 0 ? (
            <p className="text-sm text-slate-400 py-3">Sin incidencias en este periodo.</p>
          ) : (
            <div className="flex flex-wrap gap-3 mb-3">
              {porEstadoIncidencia.map((e) => (
                <div key={e.estado} className="bg-slate-50 rounded-lg px-4 py-2.5">
                  <div className="text-lg font-bold text-slate-800">{e.n}</div>
                  <div className="text-[11px] text-slate-500">{e.estado}</div>
                </div>
              ))}
            </div>
          )}
          {gastosIncidenciasTotal > 0 && (
            <div className="text-sm text-slate-600 pt-2 border-t border-slate-100">
              Gastos asociados a incidencias: <span className="font-semibold text-rose-600">{money(gastosIncidenciasTotal)}</span>
            </div>
          )}
        </div>

        <div className="bg-white border border-slate-200 rounded-lg p-4">
          <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Pedidos de materiales (periodo)</h2>
          {porEstadoPedido.length === 0 ? (
            <p className="text-sm text-slate-400 py-3">Sin pedidos en este periodo.</p>
          ) : (
            <>
              <div className="flex flex-wrap gap-3 mb-3">
                {porEstadoPedido.map((e) => (
                  <div key={e.estado} className="bg-slate-50 rounded-lg px-4 py-2.5">
                    <div className="text-lg font-bold text-slate-800">{e.n}</div>
                    <div className="text-[11px] text-slate-500">{e.estado}</div>
                  </div>
                ))}
              </div>
              {importePedidosActual > 0 && (
                <div className="text-sm text-slate-600 pt-2 border-t border-slate-100">
                  Importe total de pedidos: <span className="font-semibold text-slate-800">{money(importePedidosActual)}</span>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      <div className={`rounded-lg p-4 mb-8 border ${discrepanciasFabrica.length > 0 ? "bg-rose-50 border-rose-300" : "bg-emerald-50 border-emerald-200"}`}>
        <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Comparación oficina ↔ fábrica</h2>
        {comparacionFabrica.length === 0 ? (
          <p className="text-sm text-slate-500">Todavía no hay materiales marcados como recibidos ni por oficina ni por fábrica.</p>
        ) : discrepanciasFabrica.length === 0 ? (
          <p className="text-sm text-emerald-700 font-semibold">✓ Todo coincide — {comparacionFabrica.length} material{comparacionFabrica.length === 1 ? "" : "es"} comprobado{comparacionFabrica.length === 1 ? "" : "s"}, sin ningún fallo entre oficina y fábrica.</p>
        ) : (
          <>
            <p className="text-sm text-rose-700 font-semibold mb-3">⚠ {discrepanciasFabrica.length} discrepancia{discrepanciasFabrica.length === 1 ? "" : "s"} — algo dice una cosa en oficina y otra distinta en fábrica.</p>
            <div className="bg-white border border-rose-200 rounded-lg overflow-hidden">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-2.5 font-semibold">Pedido</th>
                    <th className="px-4 py-2.5 font-semibold">Material</th>
                    <th className="px-4 py-2.5 font-semibold">Oficina</th>
                    <th className="px-4 py-2.5 font-semibold">Fábrica</th>
                  </tr>
                </thead>
                <tbody>
                  {discrepanciasFabrica.map((f, i) => (
                    <tr key={i} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-2.5 font-mono-num text-slate-500">#{f.pedido}</td>
                      <td className="px-4 py-2.5 text-slate-700">{f.material}</td>
                      <td className="px-4 py-2.5">{f.oficinaRecibido ? <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Recibido</Badge> : <Badge className="bg-slate-100 text-slate-500 ring-slate-200">No recibido</Badge>}</td>
                      <td className="px-4 py-2.5">{f.fabricaConfirmado ? <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Ha llegado</Badge> : <Badge className="bg-slate-100 text-slate-500 ring-slate-200">No confirmado</Badge>}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      {instalacionesConDatos.length > 0 && (
        <div className="bg-white border border-slate-200 rounded-lg p-4 mb-8">
          <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Instalaciones</h2>
          <div className="flex flex-wrap gap-3 mb-4">
            {porEstadoInstalacion.map((e) => (
              <div key={e.estado} className="bg-slate-50 rounded-lg px-4 py-2.5">
                <div className="text-lg font-bold text-slate-800">{e.n}</div>
                <div className="text-[11px] text-slate-500">{e.estado}</div>
              </div>
            ))}
          </div>
          <div className="grid grid-cols-3 gap-3 mb-4">
            <div className="bg-slate-50 rounded-lg p-3">
              <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Presupuestado</div>
              <div className="font-bold text-slate-800">{money(totalPresupuestadoInstalacion)}</div>
            </div>
            <div className="bg-slate-50 rounded-lg p-3">
              <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Coste real</div>
              <div className="font-bold text-slate-800">{money(totalCosteInstalacion)}</div>
            </div>
            <div className="bg-slate-50 rounded-lg p-3">
              <div className="text-[11px] uppercase tracking-wide text-slate-400 mb-1">Diferencia global</div>
              <div className={`font-bold ${(totalPresupuestadoInstalacion - totalCosteInstalacion) < 0 ? "text-rose-600" : "text-emerald-600"}`}>{money(totalPresupuestadoInstalacion - totalCosteInstalacion)}</div>
            </div>
          </div>
          {instalacionesSobrepresupuesto.length > 0 && (
            <div className="px-3 py-2.5 rounded-md bg-rose-50 border border-rose-200 text-rose-700 text-sm font-semibold">
              ⚠ {instalacionesSobrepresupuesto.length} instalación{instalacionesSobrepresupuesto.length === 1 ? "" : "es"} por encima de su presupuesto: {instalacionesSobrepresupuesto.map((i) => i.proyecto ? `#${i.proyecto.numero}` : (i.inst.nombre || "sin proyecto")).join(", ")}.
            </div>
          )}
        </div>
      )}

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-8">
        <h2 className="font-display font-bold text-slate-700 text-sm mb-3">Productividad ({PERIODOS_INFORME.find((p) => p.id === periodo)?.label || periodo})</h2>

        <h3 className="text-xs uppercase tracking-wide text-slate-400 font-semibold mb-2">Horas por instalador</h3>
        {productividadInstaladores.length === 0 ? (
          <p className="text-sm text-slate-400 mb-4">No hay horas registradas en este periodo.</p>
        ) : (
          <div className="mb-6">
            <ResponsiveContainer width="100%" height={Math.max(160, productividadInstaladores.length * 36)}>
              <BarChart data={productividadInstaladores} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" />
                <XAxis type="number" />
                <YAxis type="category" dataKey="instalador" width={120} />
                <Tooltip formatter={(v) => `${v.toFixed(1)} h`} />
                <Bar dataKey="horas" fill="#2E8B57" />
              </BarChart>
            </ResponsiveContainer>
            <table className="w-full text-sm mt-2">
              <thead>
                <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="py-1.5">Instalador</th>
                  <th className="py-1.5">Horas</th>
                  <th className="py-1.5">Días trabajados</th>
                  <th className="py-1.5">% del total</th>
                </tr>
              </thead>
              <tbody>
                {productividadInstaladores.map((x) => (
                  <tr key={x.instalador} className="border-b border-slate-100 last:border-0">
                    <td className="py-1.5 font-medium text-slate-700">{x.instalador}</td>
                    <td className="py-1.5">{x.horas.toFixed(1)} h</td>
                    <td className="py-1.5">{x.dias}</td>
                    <td className="py-1.5">{totalHorasPeriodo ? ((x.horas / totalHorasPeriodo) * 100).toFixed(0) : 0}%</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}

        <h3 className="text-xs uppercase tracking-wide text-slate-400 font-semibold mb-2">Horas por vivienda completada (control de montaje)</h3>
        {productividadPorObra.length === 0 ? (
          <p className="text-sm text-slate-400">Todavía no hay obras con viviendas importadas en el control de montaje.</p>
        ) : (
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                <th className="py-1.5">Obra</th>
                <th className="py-1.5">Viviendas completas</th>
                <th className="py-1.5">Horas totales</th>
                <th className="py-1.5">Horas / vivienda</th>
              </tr>
            </thead>
            <tbody>
              {productividadPorObra.map((x) => (
                <tr key={x.nombre} className="border-b border-slate-100 last:border-0">
                  <td className="py-1.5 font-medium text-slate-700">{x.nombre}</td>
                  <td className="py-1.5">{x.viviendasCompletas} / {x.totalViviendas}</td>
                  <td className="py-1.5">{x.totalHoras.toFixed(1)} h</td>
                  <td className="py-1.5">{x.horasPorVivienda !== null ? `${x.horasPorVivienda.toFixed(1)} h` : "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        <p className="text-xs text-slate-400 mt-2">Horas totales de todo el histórico de la instalación (no solo del periodo filtrado arriba), para comparar contra viviendas ya terminadas.</p>
      </div>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden mb-8">
        <div className="p-4 pb-3 flex flex-wrap items-center justify-between gap-3">
          <h2 className="font-display font-bold text-slate-700 text-sm">Situación de pago por cliente</h2>
          <div className="flex gap-1">
            {[{ id: "pendiente", label: "Con pendiente" }, { id: "aldia", label: "Al día" }, { id: "todos", label: "Todos" }].map((f) => (
              <button key={f.id} onClick={() => setFiltroClientes(f.id)}
                className={`px-3 py-1.5 rounded-md text-xs font-semibold transition ${filtroClientes === f.id ? "bg-[#2E8B57] text-white" : "bg-slate-100 text-slate-500 hover:bg-slate-200"}`}>
                {f.label}
              </button>
            ))}
          </div>
        </div>
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Cliente</th>
              <th className="px-4 py-2.5 font-semibold text-right">Proyectos</th>
              <th className="px-4 py-2.5 font-semibold text-right">Presupuestado</th>
              <th className="px-4 py-2.5 font-semibold text-right">Recibido</th>
              <th className="px-4 py-2.5 font-semibold text-right">Pendiente</th>
              <th className="px-4 py-2.5 font-semibold text-right">Incidencias abiertas</th>
              <th className="px-4 py-2.5 font-semibold text-right">Pedidos pendientes</th>
              <th className="px-4 py-2.5 font-semibold text-right">Riesgo</th>
            </tr>
          </thead>
          <tbody>
            {situacionClientes.map((r) => (
              <tr key={r.cliente.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50">
                <td className="px-4 py-2.5 font-medium text-slate-800">{r.cliente.nombre}</td>
                <td className="px-4 py-2.5 text-right font-mono-num text-slate-500">{r.nProyectos}</td>
                <td className="px-4 py-2.5 text-right font-mono-num text-slate-500">{money(r.presupuestado)}</td>
                <td className="px-4 py-2.5 text-right font-mono-num text-emerald-600">{money(r.recibido)}</td>
                <td className={`px-4 py-2.5 text-right font-mono-num font-bold ${r.pendiente > 0.01 ? "text-amber-700" : "text-emerald-600"}`}>{r.pendiente > 0.01 ? money(r.pendiente) : "Al día"}</td>
                <td className="px-4 py-2.5 text-right font-mono-num">{r.incidenciasAbiertas > 0 ? <Badge className="bg-rose-50 text-rose-700 ring-rose-200">{r.incidenciasAbiertas}</Badge> : "—"}</td>
                <td className="px-4 py-2.5 text-right font-mono-num">{r.pedidosPendientes > 0 ? <Badge className="bg-sky-50 text-sky-700 ring-sky-200">{r.pedidosPendientes}</Badge> : "—"}</td>
                <td className="px-4 py-2.5 text-right">
                  {r.pctRiesgo !== null ? (
                    <Badge className={r.pctRiesgo >= 100 ? "bg-rose-50 text-rose-700 ring-rose-200" : r.pctRiesgo >= 80 ? "bg-amber-50 text-amber-700 ring-amber-200" : "bg-emerald-50 text-emerald-700 ring-emerald-200"}>
                      {r.pctRiesgo.toFixed(0)}%
                    </Badge>
                  ) : "—"}
                </td>
              </tr>
            ))}
            {situacionClientes.length === 0 && (
              <tr><td colSpan={8} className="px-4 py-8 text-center text-slate-400 text-sm">No hay clientes que coincidan con este filtro.</td></tr>
            )}
          </tbody>
        </table>
        {totalPendienteClientes > 0 && (
          <p className="text-xs text-slate-400 px-4 py-3 border-t border-slate-100">Total pendiente de cobro en estos clientes: <span className="font-semibold text-slate-600">{money(totalPendienteClientes)}</span></p>
        )}
      </div>

      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <h2 className="font-display font-bold text-slate-700 text-sm p-4 pb-0">Últimos movimientos de dinero</h2>
        <table className="w-full text-sm mt-3">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-2.5 font-semibold">Fecha</th>
              <th className="px-4 py-2.5 font-semibold">Cliente</th>
              <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
              <th className="px-4 py-2.5 font-semibold">Estado</th>
            </tr>
          </thead>
          <tbody>
            {ultimosMovimientos.map((i) => (
              <tr key={i.id} className="border-b border-slate-100 last:border-0">
                <td className="px-4 py-2.5 text-slate-600">{fmtDate(i.fecha)}</td>
                <td className="px-4 py-2.5 font-medium text-slate-800">{i.clienteNombre}</td>
                <td className="px-4 py-2.5 text-right font-mono-num">{money(i.importe)}</td>
                <td className="px-4 py-2.5">
                  {i.proyectoId ? <Badge className="bg-emerald-50 text-emerald-700 ring-emerald-200">Vinculado</Badge> : <Badge className="bg-amber-50 text-amber-700 ring-amber-200">Sin vincular</Badge>}
                </td>
              </tr>
            ))}
            {ultimosMovimientos.length === 0 && (
              <tr><td colSpan={4} className="px-4 py-10 text-center text-slate-400 text-sm">Todavía no hay entradas de dinero registradas.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {isAdmin && (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <h2 className="font-display font-bold text-slate-700 text-sm p-4 pb-0">Tiempo conectado por usuario</h2>
          <p className="text-xs text-slate-400 px-4 pt-1">Tiempo que cada usuario ha tenido el CRM abierto en el navegador (aproximado — no es lo mismo que las horas fichadas en Fichajes).</p>
          <table className="w-full text-sm mt-3">
            <thead>
              <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                <th className="px-4 py-2.5 font-semibold">Usuario</th>
                <th className="px-4 py-2.5 font-semibold text-right">Hoy</th>
                <th className="px-4 py-2.5 font-semibold text-right">Últimos 7 días</th>
                <th className="px-4 py-2.5 font-semibold text-right">Total acumulado</th>
                <th className="px-4 py-2.5 font-semibold">Última conexión</th>
              </tr>
            </thead>
            <tbody>
              {resumenConexion.map((r) => (
                <tr key={r.usuario.id} className="border-b border-slate-100 last:border-0">
                  <td className="px-4 py-2.5 font-medium text-slate-800">{r.usuario.nombre} {r.usuario.apellidos || ""}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num text-slate-600">{formatoHorasMin(r.hoy)}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num text-slate-600">{formatoHorasMin(r.ultimos7)}</td>
                  <td className="px-4 py-2.5 text-right font-mono-num font-semibold text-slate-800">{formatoHorasMin(r.total)}</td>
                  <td className="px-4 py-2.5 text-slate-500">{r.ultimaConexion ? new Date(r.ultimaConexion).toLocaleString("es-ES") : "—"}</td>
                </tr>
              ))}
              {resumenConexion.length === 0 && (
                <tr><td colSpan={5} className="px-4 py-8 text-center text-slate-400 text-sm">Todavía no hay datos de conexión registrados.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function EnviarAvisoEmailPanel({ llamarHoy, contactarVencidos }) {
  const [email, setEmail] = useState("");
  const [enviando, setEnviando] = useState(false);
  const [resultado, setResultado] = useState(null);

  useEffect(() => {
    const v = localStorage.getItem("alumavel_email_avisos_presupuestos");
    if (v) setEmail(v);
  }, []);

  const guardarEmail = (valor) => {
    setEmail(valor);
    localStorage.setItem("alumavel_email_avisos_presupuestos", valor);
  };

  const enviarAviso = async () => {
    if (!email.trim()) { setResultado({ ok: false, msg: "Escribe primero tu email." }); return; }
    setEnviando(true);
    setResultado(null);
    try {
      const lineasLlamar = llamarHoy.map((p) => `- ${p.numero} — ${p.clienteNombre} (${diasSinRespuestaDe(p)} días sin respuesta)${p.telefono ? ` — Tel: ${p.telefono}` : ""}`).join("\n") || "Ninguno";
      const lineasContactar = contactarVencidos.map((p) => `- ${p.numero} — ${p.clienteNombre} (fecha prevista de confirmación ya pasada)`).join("\n") || "Ninguno";
      const cuerpo = `Aviso diario de presupuestos ALUMAVEL\n\nPresupuestos para llamar hoy (mas de 7 dias sin respuesta):\n${lineasLlamar}\n\nPresupuestos en espera con seguimiento vencido:\n${lineasContactar}`;
      const prompt = `Envía un correo usando la herramienta Gmail:send_message con to="${email.trim()}", subject="Aviso diario de presupuestos ALUMAVEL", y body="${cuerpo.replace(/"/g, "'")}". No hagas nada más ni preguntes nada. Al terminar responde únicamente "OK".`;

      const response = await fetch("/.netlify/functions/anthropic-proxy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-5",
          max_tokens: 500,
          messages: [{ role: "user", content: prompt }],
          mcp_servers: [{ type: "url", url: "https://gmailmcp.googleapis.com/mcp/v1", name: "gmail-mcp" }],
        }),
      });
      if (!response.ok) throw new Error("Respuesta no válida de la API");
      setResultado({ ok: true, msg: `Aviso enviado a ${email.trim()}.` });
    } catch (err) {
      setResultado({ ok: false, msg: "No se pudo enviar el correo. Comprueba que Gmail esté conectado y vuelve a intentarlo." });
    } finally {
      setEnviando(false);
    }
  };

  return (
    <div className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
      <label className="text-sm font-semibold text-slate-700 block">Tu email para recibir el aviso</label>
      <div className="flex flex-wrap gap-2">
        <TextInput type="email" value={email} onChange={(e) => guardarEmail(e.target.value)} placeholder="tucorreo@ejemplo.com" className="flex-1 min-w-[200px]" />
        <button onClick={enviarAviso} disabled={enviando} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1.5 text-sm font-semibold disabled:bg-slate-300 px-4 py-2 rounded-md hover:opacity-90">
          <Mail size={14} /> {enviando ? "Enviando..." : "Enviar ahora"}
        </button>
      </div>
      {resultado && (
        <div className={`text-sm font-semibold px-3 py-2 rounded-md ${resultado.ok ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"}`}>
          {resultado.msg}
        </div>
      )}
    </div>
  );
}

function PresupuestoForm({ initial, clientes, presupuestosExistentes, onCrearClienteRapido, onLeerDatos, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || {
      id: null, numero: "", fechaEnvio: new Date().toISOString().slice(0, 10), clienteNombre: "", telefono: "",
      descripcion: "", importe: "", estado: "Pendiente", motivoRechazo: "", fechaRespuesta: "",
      comentarios: "", fechaPrevistaConfirmacion: "", envio: false, direccionEnvio: "", montaje: false, zona: "",
    }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const [leyendoFotoForm, setLeyendoFotoForm] = useState(false);
  const [errorFotoForm, setErrorFotoForm] = useState("");
  const inputFotoFormRef = useRef(null);

  const rellenarDesdeArchivo = async (file) => {
    if (!file || !onLeerDatos) return;
    setLeyendoFotoForm(true);
    setErrorFotoForm("");
    try {
      const datos = await onLeerDatos(file);
      setF((prev) => ({
        ...prev,
        clienteNombre: datos.clienteNombre || prev.clienteNombre,
        telefono: datos.telefono || prev.telefono,
        descripcion: datos.descripcion ? (prev.descripcion ? `${prev.descripcion}\n\n${datos.descripcion}` : datos.descripcion) : prev.descripcion,
        importe: datos.importe || prev.importe,
        zona: datos.zona || prev.zona,
      }));
    } catch (err) {
      setErrorFotoForm("No se pudo leer el archivo. Prueba con una foto más clara, con más luz, o inténtalo de nuevo. (" + err.message + ")");
    } finally {
      setLeyendoFotoForm(false);
    }
  };

  const clientesDisponibles = useMemo(() => clientes.map((c) => c.nombre).sort(), [clientes]);
  const [confirmarDuplicado, setConfirmarDuplicado] = useState(false);
  const [limiteCreditoRapido, setLimiteCreditoRapido] = useState("");
  const [errorClienteRapido, setErrorClienteRapido] = useState("");

  const duplicado = !f.id && (presupuestosExistentes || []).find((p) => p.numero.trim().toLowerCase() === f.numero.trim().toLowerCase());
  const clienteExiste = (nombre) => clientes.some((c) => c.nombre.trim().toLowerCase() === (nombre || "").trim().toLowerCase());
  const clienteFaltante = f.clienteNombre.trim() && !clienteExiste(f.clienteNombre);

  const crearClienteAhora = () => {
    if (!limiteCreditoRapido || parseFloat(limiteCreditoRapido) <= 0) {
      setErrorClienteRapido("Pon un límite de crédito mayor que 0 para poder darlo de alta.");
      return;
    }
    setErrorClienteRapido("");
    onCrearClienteRapido(f.clienteNombre.trim(), limiteCreditoRapido);
    setLimiteCreditoRapido("");
  };

  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.numero.trim()) { setErrorMsg("Falta el Nº de presupuesto (es obligatorio)."); return; }
    if (!f.clienteNombre.trim()) { setErrorMsg("Falta el nombre del cliente (es obligatorio)."); return; }
    if (!clienteExiste(f.clienteNombre)) {
      setErrorMsg(`El cliente "${f.clienteNombre}" no existe todavía en el CRM. Ve a Clientes y date de alta primero, y luego vuelve aquí a elegirlo de la lista.`);
      return;
    }
    if (!f.fechaEnvio) { setErrorMsg("Falta la fecha de envío (es obligatoria)."); return; }
    if (duplicado && !confirmarDuplicado) {
      setErrorMsg(`Ya existe un presupuesto con el número "${f.numero}" (cliente: ${duplicado.clienteNombre}). Si de verdad quieres crear otro con el mismo número, cambia el número o marca la casilla de abajo para confirmarlo.`);
      return;
    }
    setErrorMsg("");
    onSave({ ...f, importe: parseFloat(f.importe) || 0 });
  };

  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-4">{initial?.id ? `Editar presupuesto ${initial.numero}` : "Nuevo presupuesto"}</h1>

      {onLeerDatos && (
        <div className="mb-5">
          <input
            ref={inputFotoFormRef}
            type="file"
            accept="image/*,application/pdf"
            className="hidden"
            onChange={(e) => { rellenarDesdeArchivo(e.target.files[0]); e.target.value = ""; }}
          />
          <button
            type="button"
            onClick={() => inputFotoFormRef.current?.click()}
            disabled={leyendoFotoForm}
            className="flex items-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50 disabled:opacity-60"
          >
            {leyendoFotoForm ? "Leyendo el archivo..." : "📷 Rellenar (o completar) desde foto/PDF"}
          </button>
          <p className="text-xs text-slate-400 mt-1">Rellena los campos vacíos con lo que encuentre en la foto/PDF, sin borrar lo que ya tengas escrito. Útil también al duplicar un presupuesto.</p>
          {errorFotoForm && <p className="text-xs text-rose-600 font-semibold mt-1">⚠ {errorFotoForm}</p>}
        </div>
      )}

      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}

      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-3 gap-4">
          <Field label="Nº Presupuesto" required>
            <TextInput value={f.numero} onChange={set("numero")} placeholder="Ej: PPT-2026-1234.1" />
          </Field>
          <Field label="Fecha envío" required>
            <TextInput type="date" value={f.fechaEnvio} onChange={set("fechaEnvio")} />
          </Field>
          <Field label="Importe (€)">
            <TextInput type="number" step="0.01" value={f.importe} onChange={set("importe")} />
          </Field>
        </div>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Cliente" required>
            <TextInput value={f.clienteNombre} onChange={set("clienteNombre")} list="presupuesto-clientes-datalist" placeholder="Nombre del cliente" />
            <datalist id="presupuesto-clientes-datalist">
              {clientesDisponibles.map((n) => <option key={n} value={n} />)}
            </datalist>
          </Field>
          <Field label="Teléfono (WhatsApp)">
            <TextInput value={f.telefono} onChange={set("telefono")} placeholder="Ej: 611223344" />
          </Field>
        </div>

        {clienteFaltante && (
          <div className="px-4 py-3 rounded-md bg-amber-50 border border-amber-300 space-y-2">
            <p className="text-sm font-semibold text-amber-800">⚠ "{f.clienteNombre}" no existe todavía en el CRM.</p>
            <div className="flex flex-wrap items-end gap-2">
              <div className="flex-1 min-w-[180px]">
                <label className="text-xs font-semibold text-amber-800 block mb-1">Límite de crédito asegurado (€)</label>
                <TextInput type="number" step="0.01" value={limiteCreditoRapido} onChange={(e) => setLimiteCreditoRapido(e.target.value)} placeholder="Ej: 40000" />
              </div>
              <button type="button" onClick={crearClienteAhora} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="text-sm font-semibold hover:opacity-90 px-4 py-2.5 rounded-md">
                Crear cliente ahora
              </button>
            </div>
            {errorClienteRapido && <p className="text-xs text-rose-600 font-semibold">⚠ {errorClienteRapido}</p>}
          </div>
        )}

        <div className="grid grid-cols-2 gap-4">
          <Field label="Zona">
            <TextInput value={f.zona} onChange={set("zona")} placeholder="Ej: Almería, Granada..." />
          </Field>
        </div>

        <Field label="Descripción / Obra">
          <TextArea rows={2} value={f.descripcion} onChange={set("descripcion")} placeholder="Ej: Solicitud de 4 ventanas de PVC" />
        </Field>

        <div className="grid grid-cols-2 gap-4">
          <Field label="Estado">
            <Select value={f.estado} onChange={set("estado")}>
              {ESTADO_PRESUPUESTO_TRACKER.map((t) => <option key={t}>{t}</option>)}
            </Select>
          </Field>
          {f.estado === "Rechazado" && (
            <Field label="Motivo rechazo">
              <Select value={f.motivoRechazo} onChange={set("motivoRechazo")}>
                <option value="">Selecciona...</option>
                {MOTIVO_RECHAZO.map((t) => <option key={t}>{t}</option>)}
              </Select>
            </Field>
          )}
          {(f.estado === "Aceptado" || f.estado === "Rechazado") && (
            <Field label="Fecha respuesta">
              <TextInput type="date" value={f.fechaRespuesta} onChange={set("fechaRespuesta")} />
            </Field>
          )}
          {f.estado === "En espera" && (
            <Field label="Fecha prevista confirmación">
              <TextInput type="date" value={f.fechaPrevistaConfirmacion} onChange={set("fechaPrevistaConfirmacion")} />
            </Field>
          )}
        </div>

        <div className="grid grid-cols-2 gap-4 items-end">
          <label className="flex items-center gap-2 text-sm text-slate-700 font-medium">
            <input type="checkbox" checked={f.envio} onChange={set("envio")} className="w-4 h-4" /> ¿Envío?
          </label>
          <label className="flex items-center gap-2 text-sm text-slate-700 font-medium">
            <input type="checkbox" checked={f.montaje} onChange={set("montaje")} className="w-4 h-4" /> ¿Montaje?
          </label>
        </div>
        {f.envio && (
          <Field label="Dirección de envío">
            <TextInput value={f.direccionEnvio} onChange={set("direccionEnvio")} />
          </Field>
        )}

        <Field label="Comentarios">
          <TextArea rows={2} value={f.comentarios} onChange={set("comentarios")} />
        </Field>

        {duplicado && (
          <label className="flex items-start gap-2 text-sm text-amber-800 bg-amber-50 border border-amber-300 rounded-md px-3 py-2.5">
            <input type="checkbox" checked={confirmarDuplicado} onChange={(e) => setConfirmarDuplicado(e.target.checked)} className="mt-0.5" />
            Sé que ya existe un presupuesto con este número y quiero crear otro igualmente.
          </label>
        )}

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar presupuesto</button>
        </div>
      </form>
    </div>
  );
}

function PresupuestoDetail({ presupuesto, onBack, onEdit, onDelete, onAddLlamada, onDeleteLlamada, onCrearProyecto, onDuplicar, replicas, onAbrirReplica, isAdmin }) {
  const estadoActual = presupuesto.estado || "Pendiente";
  const dias = diasSinRespuestaDe(presupuesto);
  const diasResp = diasEntre(presupuesto.fechaEnvio, presupuesto.fechaRespuesta);
  const llamadas = presupuesto.llamadas || [];
  const versiones = presupuesto.versiones || [];
  const [versionAbiertaId, setVersionAbiertaId] = useState(null);
  // Cada entrada del historial guarda cómo estaba ANTES de ese cambio. El valor de
  // "después" es el de la revisión más reciente (índice anterior), o el valor actual
  // del presupuesto si es la revisión más reciente de todas.
  const despuesDe = (index) => {
    if (index === 0) return { importe: presupuesto.importe, descripcion: presupuesto.descripcion };
    return { importe: versiones[index - 1].importe, descripcion: versiones[index - 1].descripcion };
  };
  const [lForm, setLForm] = useState({ fecha: new Date().toISOString().slice(0, 10), notas: "", enlaceGrabacion: "" });
  const [errorLlamada, setErrorLlamada] = useState("");
  const registrarLlamada = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!lForm.notas.trim() && !lForm.enlaceGrabacion.trim()) {
      setErrorLlamada("Escribe algo en las notas o pega un enlace a la grabación antes de registrar la llamada.");
      return;
    }
    setErrorLlamada("");
    onAddLlamada({ id: uid(), ...lForm });
    setLForm({ fecha: new Date().toISOString().slice(0, 10), notas: "", enlaceGrabacion: "" });
  };
  return (
    <div className="p-8 max-w-3xl">
      <button onClick={onBack} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>

      <div className="flex items-start justify-between mb-6 gap-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <h1 className="font-display text-2xl font-extrabold text-slate-900">{presupuesto.numero}</h1>
            <Badge className={ESTADO_PRESUPUESTO_TRACKER_STYLE[estadoActual]}>{estadoActual}</Badge>
          </div>
          <p className="text-slate-500 text-sm">{presupuesto.clienteNombre} · {fmtDate(presupuesto.fechaEnvio)}</p>
        </div>
        <div className="flex gap-2 shrink-0">
          {enlaceLlamar(presupuesto) && (
            <a href={enlaceLlamar(presupuesto)} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="flex items-center gap-1.5 text-sm font-semibold hover:opacity-90 px-3.5 py-2 rounded-md">
              <Phone size={14} /> Llamar
            </a>
          )}
          {enlaceWhatsapp(presupuesto) && (
            <a href={enlaceWhatsapp(presupuesto)} target="_blank" rel="noopener noreferrer" className="flex items-center gap-1.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 px-3.5 py-2 rounded-md">
              <MessageCircle size={14} /> WhatsApp
            </a>
          )}
          <button onClick={onDuplicar} title="Crea una réplica de este presupuesto con su propio número (ej. 4192 → 4192-1), lista para modificar" className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Copy size={14} /> Duplicar (nueva réplica)</button>
          <button onClick={onEdit} className="flex items-center gap-1.5 text-sm font-semibold text-slate-600 border border-slate-300 px-3.5 py-2 rounded-md hover:bg-slate-50"><Pencil size={14} /> Editar</button>
          {isAdmin && (
            <button onClick={onDelete} className="flex items-center gap-1.5 text-sm font-semibold text-rose-600 border border-rose-200 px-3.5 py-2 rounded-md hover:bg-rose-50"><Trash2 size={14} /> Eliminar</button>
          )}
        </div>
      </div>

      {dias !== null && dias > 7 && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ Llevan {dias} días sin respuesta — llamar hoy.
        </div>
      )}

      {estadoActual === "Aceptado" && !presupuesto.proyectoCreadoId && (
        <button
          onClick={onCrearProyecto}
          style={{ backgroundColor: "#2E8B57", color: "#ffffff" }}
          className="w-full flex items-center justify-center gap-2 hover:opacity-90 text-base font-bold py-4 rounded-lg mb-6 shadow-md cursor-pointer select-none"
        >
          <Briefcase size={20} /> CREAR PROYECTO DESDE ESTE PRESUPUESTO
        </button>
      )}
      {presupuesto.proyectoCreadoId && (
        <div className="mb-6 px-4 py-3 rounded-md bg-emerald-50 border border-emerald-200 text-emerald-700 text-sm font-semibold">
          ✓ Ya se creó un proyecto a partir de este presupuesto (misma tarjeta de Trello).
        </div>
      )}

      <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6 mb-6">
        <div className="grid grid-cols-2 gap-x-8 gap-y-4 text-sm">
          <InfoRow icon={<Euro size={14} />} label="Importe" value={money(presupuesto.importe)} />
          <InfoRow icon={<Phone size={14} />} label="Teléfono" value={presupuesto.telefono || "—"} />
          <InfoRow icon={<MapPin size={14} />} label="Zona" value={presupuesto.zona || "—"} />
          <InfoRow icon={<FileText size={14} />} label="Descripción / Obra" value={presupuesto.descripcion || "—"} />
          {presupuesto.fechaRespuesta && <InfoRow icon={<CalendarDays size={14} />} label="Fecha respuesta" value={`${fmtDate(presupuesto.fechaRespuesta)} (${diasResp}d)`} />}
          {estadoActual === "Rechazado" && <InfoRow icon={<AlertOctagon size={14} />} label="Motivo rechazo" value={presupuesto.motivoRechazo || "—"} />}
          {estadoActual === "En espera" && <InfoRow icon={<CalendarDays size={14} />} label="Fecha prevista confirmación" value={fmtDate(presupuesto.fechaPrevistaConfirmacion)} />}
          <InfoRow icon={<CheckCircle2 size={14} />} label="¿Envío?" value={presupuesto.envio ? "Sí" : "No"} />
          {presupuesto.envio && <InfoRow icon={<MapPin size={14} />} label="Dirección de envío" value={presupuesto.direccionEnvio || "—"} />}
          <InfoRow icon={<CheckCircle2 size={14} />} label="¿Montaje?" value={presupuesto.montaje ? "Sí" : "No"} />
        </div>
        {presupuesto.comentarios && (
          <div className="mt-4 pt-4 border-t border-slate-100 text-sm text-slate-600">
            <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Comentarios</span>
            {presupuesto.comentarios}
          </div>
        )}
      </CornerFrame>

      {(replicas || []).length > 0 && (
        <div className="mb-8">
          <h2 className="font-display font-bold text-slate-800 mb-3">Réplicas de este presupuesto ({replicas.length})</h2>
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            <table className="w-full text-sm">
              <thead>
                <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                  <th className="px-4 py-2.5 font-semibold">Nº</th>
                  <th className="px-4 py-2.5 font-semibold">Fecha</th>
                  <th className="px-4 py-2.5 font-semibold text-right">Importe</th>
                  <th className="px-4 py-2.5 font-semibold">Estado</th>
                </tr>
              </thead>
              <tbody>
                {replicas.map((r) => (
                  <tr key={r.id} onClick={() => onAbrirReplica(r.id)} className="border-b border-slate-100 last:border-0 hover:bg-slate-50 cursor-pointer transition">
                    <td className="px-4 py-2.5 font-mono-num font-medium text-slate-800">{r.numero}</td>
                    <td className="px-4 py-2.5 text-slate-600">{fmtDate(r.fechaEnvio)}</td>
                    <td className="px-4 py-2.5 text-right font-mono-num">{money(r.importe)}</td>
                    <td className="px-4 py-2.5"><Badge className={ESTADO_PRESUPUESTO_TRACKER_STYLE[r.estado || "Pendiente"]}>{r.estado || "Pendiente"}</Badge></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-slate-400 mt-1.5">Estas réplicas no aparecen sueltas en el listado principal — solo aquí, y si las buscas por su número.</p>
        </div>
      )}

      {versiones.length > 0 && (
        <div className="mb-8">
          <h2 className="font-display font-bold text-slate-800 mb-3">Historial de versiones ({versiones.length})</h2>
          <div className="space-y-2">
            {versiones.map((v, i) => {
              const despues = despuesDe(i);
              const abierta = versionAbiertaId === v.id;
              const cambioImporte = parseFloat(v.importe) !== parseFloat(despues.importe);
              const cambioDescripcion = (v.descripcion || "") !== (despues.descripcion || "");
              return (
                <div key={v.id} className="bg-white border border-slate-200 rounded-lg overflow-hidden">
                  <button
                    onClick={() => setVersionAbiertaId(abierta ? null : v.id)}
                    className="w-full flex items-center justify-between gap-3 px-4 py-3 text-left hover:bg-slate-50"
                  >
                    <span className="text-sm font-medium text-slate-700">Revisión del {fmtDate(v.fecha)}</span>
                    <span className="text-xs text-slate-400">{abierta ? "Ocultar ▲" : "Ver cambios ▼"}</span>
                  </button>
                  {abierta && (
                    <div className="px-4 pb-4 pt-1 border-t border-slate-100 grid grid-cols-2 gap-4 text-sm">
                      <div>
                        <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Antes</span>
                        <div className={cambioImporte ? "text-rose-600 font-semibold" : "text-slate-600"}>{money(v.importe)}</div>
                        <div className="text-slate-600 mt-1">{v.descripcion || "—"}</div>
                      </div>
                      <div>
                        <span className="text-[11px] font-semibold uppercase tracking-wide text-slate-400 block mb-1">Después{i === 0 ? " (actual)" : ""}</span>
                        <div className={cambioImporte ? "text-emerald-600 font-semibold" : "text-slate-600"}>{money(despues.importe)}</div>
                        <div className="text-slate-600 mt-1">{despues.descripcion || "—"}</div>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      <h2 className="font-display font-bold text-slate-800 mb-3">Registro de llamadas ({llamadas.length})</h2>
      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-4 space-y-3">
        {llamadas.length === 0 ? (
          <p className="text-sm text-slate-400 text-center py-3">Todavía no hay llamadas registradas.</p>
        ) : (
          llamadas.map((l) => (
            <div key={l.id} className="flex items-start justify-between gap-3 border-b border-slate-100 last:border-0 pb-3 last:pb-0">
              <div className="text-sm">
                <div className="font-semibold text-slate-700">{fmtDate(l.fecha)}</div>
                {l.notas && <div className="text-slate-600">{l.notas}</div>}
                {l.enlaceGrabacion && (
                  <a href={l.enlaceGrabacion} target="_blank" rel="noopener noreferrer" className="text-sky-600 hover:underline text-xs">🎧 Escuchar grabación</a>
                )}
              </div>
              <button onClick={() => onDeleteLlamada(l.id)} className="text-slate-400 hover:text-rose-600 shrink-0"><X size={15} /></button>
            </div>
          ))
        )}
      </div>

      <form onSubmit={registrarLlamada} className="bg-white border border-slate-200 rounded-lg p-4 space-y-3">
        {errorLlamada && (
          <div className="px-3 py-2.5 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">⚠ {errorLlamada}</div>
        )}
        <div className="grid grid-cols-2 gap-3">
          <Field label="Fecha">
            <TextInput type="date" value={lForm.fecha} onChange={(e) => setLForm({ ...lForm, fecha: e.target.value })} />
          </Field>
          <Field label="Enlace a la grabación (si tu centralita lo da)">
            <TextInput value={lForm.enlaceGrabacion} onChange={(e) => setLForm({ ...lForm, enlaceGrabacion: e.target.value })} placeholder="https://..." />
          </Field>
        </div>
        <Field label="Notas de la llamada">
          <TextArea rows={2} value={lForm.notas} onChange={(e) => setLForm({ ...lForm, notas: e.target.value })} placeholder="Qué se habló, próximos pasos..." />
        </Field>
        <div className="flex justify-end">
          <button type="submit" onClick={registrarLlamada} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-4 py-2 rounded-md"><Plus size={15} /> Registrar llamada</button>
        </div>
      </form>
    </div>
  );
}

/* ================= FICHAJES ================= */

function fmtHora(iso) {
  if (!iso) return "—";
  return new Date(iso).toLocaleTimeString("es-ES", { hour: "2-digit", minute: "2-digit" });
}

function horasTrabajadas(registro) {
  if (!registro.entrada) return 0;
  const fin = registro.salida ? new Date(registro.salida) : new Date();
  let ms = fin - new Date(registro.entrada);
  (registro.descansos || []).forEach((d) => {
    const finD = d.fin ? new Date(d.fin) : new Date();
    ms -= (finD - new Date(d.inicio));
  });
  return Math.max(ms / 3600000, 0);
}

function FichajesModulo({ fichajes, empleadoActual, setEmpleadoActual, onFichar, onDeleteFichaje, isAdmin }) {
  const [tab, setTab] = useState("mi_fichaje");

  return (
    <div className="p-8 max-w-6xl overflow-x-hidden">
      <div className="flex items-center gap-2 mb-1">
        <Badge className="bg-slate-100 text-slate-500 ring-slate-200">Fase posterior</Badge>
      </div>
      <Header icon={<Timer size={20} className="text-[#2E8B57]" />} title="Fichajes" manualKey="fichajes" subtitle="Control horario de empleados" />

      <div className="flex gap-1 mb-6 border-b border-slate-200">
        {[
          { id: "mi_fichaje", label: "Fichajes usuario", icon: LogIn },
          ...(isAdmin ? [{ id: "control", label: "Control de fichajes", icon: FileSpreadsheet }] : []),
        ].map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
            <t.icon size={14} /> {t.label}
          </button>
        ))}
      </div>

      {tab === "mi_fichaje" && <MiFichaje fichajes={fichajes} empleadoActual={empleadoActual} setEmpleadoActual={setEmpleadoActual} onFichar={onFichar} />}
      {tab === "control" && isAdmin && <ControlFichajes fichajes={fichajes} onDeleteFichaje={onDeleteFichaje} />}
    </div>
  );
}

function MiFichaje({ fichajes, empleadoActual, setEmpleadoActual, onFichar }) {
  const hoy = new Date().toISOString().slice(0, 10);
  const registro = fichajes.find((f) => f.empleado === empleadoActual && f.fecha === hoy);
  const estado = registro?.estado || "sin_entrada";

  return (
    <div className="max-w-lg">
      <Field label="Empleado">
        <TextInput value={empleadoActual} onChange={(e) => setEmpleadoActual(e.target.value)} placeholder="Escribe tu nombre para fichar" />
      </Field>

      {!empleadoActual.trim() ? (
        <p className="text-sm text-slate-400 mt-4">Indica tu nombre para poder fichar.</p>
      ) : (
        <>
          <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6 mt-5 text-center">
            <p className="text-sm text-slate-500 mb-1">{new Date().toLocaleDateString("es-ES", { weekday: "long", day: "numeric", month: "long" })}</p>
            <p className="font-display text-3xl font-extrabold text-slate-900 mb-4">
              {estado === "sin_entrada" && "Sin fichar"}
              {estado === "trabajando" && "Trabajando"}
              {estado === "descanso" && "En descanso"}
              {estado === "finalizado" && "Jornada finalizada"}
            </p>

            <div className="grid grid-cols-2 gap-3">
              <button
                onClick={() => onFichar(empleadoActual, "entrada")}
                disabled={estado !== "sin_entrada"}
                className="flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-semibold py-3 rounded-md transition"
              >
                <LogIn size={16} /> Entrada
              </button>
              <button
                onClick={() => onFichar(empleadoActual, "salida")}
                disabled={estado !== "trabajando"}
                className="flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-semibold py-3 rounded-md transition"
              >
                <LogOut size={16} /> Salida
              </button>
              <button
                onClick={() => onFichar(empleadoActual, "inicio_descanso")}
                disabled={estado !== "trabajando"}
                className="flex items-center justify-center gap-2 bg-amber-500 hover:bg-amber-600 disabled:bg-slate-200 disabled:text-slate-400 text-white font-semibold py-3 rounded-md transition"
              >
                <Coffee size={16} /> Inicio descanso
              </button>
              <button
                onClick={() => onFichar(empleadoActual, "fin_descanso")}
                disabled={estado !== "descanso"}
                className="flex items-center justify-center gap-2 bg-sky-600 hover:bg-sky-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-semibold py-3 rounded-md transition"
              >
                <Coffee size={16} /> Fin descanso
              </button>
            </div>
          </CornerFrame>

          {registro && (
            <div className="mt-5 bg-white border border-slate-200 rounded-lg p-4 text-sm space-y-2">
              <div className="flex justify-between"><span className="text-slate-500">Entrada</span><span className="font-mono-num">{fmtHora(registro.entrada)}</span></div>
              {(registro.descansos || []).map((d, i) => (
                <div key={d.id} className="flex justify-between text-slate-500">
                  <span>Descanso {i + 1}</span>
                  <span className="font-mono-num">{fmtHora(d.inicio)} – {fmtHora(d.fin)}</span>
                </div>
              ))}
              <div className="flex justify-between"><span className="text-slate-500">Salida</span><span className="font-mono-num">{fmtHora(registro.salida)}</span></div>
              <div className="flex justify-between pt-2 border-t border-slate-100 font-semibold"><span className="text-slate-700">Horas trabajadas</span><span className="font-mono-num">{horasTrabajadas(registro).toFixed(2)} h</span></div>
            </div>
          )}
        </>
      )}
      <p className="text-xs text-slate-400 mt-4">Si un fichaje sigue en curso a las 23:59h, el sistema completo lo cerraría automáticamente como salida.</p>
    </div>
  );
}

function ControlFichajes({ fichajes, onDeleteFichaje }) {
  const [empleado, setEmpleado] = useState("");
  const [desde, setDesde] = useState("");
  const [hasta, setHasta] = useState("");

  const filtered = useMemo(() => {
    return fichajes.filter((f) => {
      if (empleado && !f.empleado.toLowerCase().includes(empleado.toLowerCase())) return false;
      if (desde && f.fecha < desde) return false;
      if (hasta && f.fecha > hasta) return false;
      return true;
    }).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  }, [fichajes, empleado, desde, hasta]);

  const exportarExcel = () => {
    const rows = filtered.map((f) => ({
      Empleado: f.empleado,
      Fecha: f.fecha,
      Entrada: fmtHora(f.entrada),
      "Fin descanso / Inicio descanso": (f.descansos || []).map((d) => `${fmtHora(d.inicio)}–${fmtHora(d.fin)}`).join(" / "),
      Salida: fmtHora(f.salida),
      "Horas trabajadas": horasTrabajadas(f).toFixed(2),
      Estado: f.estado,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Fichajes");
    XLSX.writeFile(wb, `fichajes_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  return (
    <div>
      <p className="text-xs text-slate-400 mb-4">Esta subsección solo debería ser visible para el rol Administrador.</p>
      <div className="flex flex-wrap gap-3 mb-4">
        <TextInput placeholder="Filtrar por empleado" value={empleado} onChange={(e) => setEmpleado(e.target.value)} className="max-w-[220px]" />
        <TextInput type="date" value={desde} onChange={(e) => setDesde(e.target.value)} className="max-w-[170px]" />
        <TextInput type="date" value={hasta} onChange={(e) => setHasta(e.target.value)} className="max-w-[170px]" />
        <button onClick={exportarExcel} disabled={filtered.length === 0} className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 disabled:bg-slate-300 text-white text-sm font-semibold px-4 py-2 rounded-md ml-auto">
          <Download size={15} /> Exportar a Excel
        </button>
      </div>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Empleado</th>
              <th className="px-4 py-3 font-semibold">Fecha</th>
              <th className="px-4 py-3 font-semibold">Entrada</th>
              <th className="px-4 py-3 font-semibold">Descansos</th>
              <th className="px-4 py-3 font-semibold">Salida</th>
              <th className="px-4 py-3 font-semibold text-right">Horas</th>
              <th className="px-4 py-3 font-semibold">Estado</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr><td colSpan={8} className="px-4 py-10 text-center text-slate-400 text-sm">No hay fichajes que coincidan con la búsqueda.</td></tr>
            )}
            {filtered.map((f) => (
              <tr key={f.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50">
                <td className="px-4 py-3 font-medium text-slate-800">{f.empleado}</td>
                <td className="px-4 py-3 text-slate-500">{fmtDate(f.fecha)}</td>
                <td className="px-4 py-3 font-mono-num">{fmtHora(f.entrada)}</td>
                <td className="px-4 py-3 text-slate-500">{(f.descansos || []).length ? `${f.descansos.length} descanso${f.descansos.length === 1 ? "" : "s"}` : "—"}</td>
                <td className="px-4 py-3 font-mono-num">{fmtHora(f.salida)}</td>
                <td className="px-4 py-3 text-right font-mono-num">{horasTrabajadas(f).toFixed(2)}</td>
                <td className="px-4 py-3">
                  <Badge className={f.estado === "finalizado" ? "bg-emerald-50 text-emerald-700 ring-emerald-200" : f.estado === "descanso" ? "bg-amber-50 text-amber-700 ring-amber-200" : "bg-sky-50 text-sky-700 ring-sky-200"}>
                    {f.estado === "finalizado" ? "Finalizado" : f.estado === "descanso" ? "En descanso" : "Trabajando"}
                  </Badge>
                </td>
                <td className="px-4 py-3 text-right">
                  <button onClick={() => { if (confirm(`¿Eliminar el fichaje de ${f.empleado} del ${fmtDate(f.fecha)}?`)) onDeleteFichaje(f.id); }} className="text-slate-300 hover:text-rose-500">
                    <Trash2 size={15} />
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ================= LOGIN / ADMINISTRACIÓN ================= */

function LoginGate({ usuarios, clientes, onCreateFirstAdmin, onLogin, onLoginCliente }) {
  const primeraVez = usuarios.length === 0;
  const [modo, setModo] = useState("empleado"); // empleado | cliente
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");

  const [f, setF] = useState({ nombre: "", apellidos: "", email: "", password: "", telefono: "" });
  const setField = (k) => (e) => setF({ ...f, [k]: e.target.value });

  const submitLogin = (e) => {
    e.preventDefault();
    const user = usuarios.find((u) => u.email.toLowerCase() === email.trim().toLowerCase());
    if (!user || user.password !== password) {
      setError("Email o contraseña incorrectos.");
      return;
    }
    setError("");
    onLogin(user.id);
  };

  const submitLoginCliente = (e) => {
    e.preventDefault();
    const cliente = clientes.find((c) => c.portalActivo && c.email && c.email.toLowerCase() === email.trim().toLowerCase());
    if (!cliente || cliente.portalPassword !== password) {
      setError("Email o contraseña incorrectos, o el portal no está activado para este cliente.");
      return;
    }
    setError("");
    onLoginCliente(cliente.id);
  };

  const submitFirstAdmin = (e) => {
    e.preventDefault();
    if (!f.nombre.trim() || !f.email.trim() || !f.password.trim()) return;
    onCreateFirstAdmin(f);
  };

  return (
    <div className="min-h-screen bg-[#F4F5F3] flex items-start justify-center p-6 overflow-y-auto" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .font-mono-num{ font-family:'IBM Plex Mono', monospace; }
        .font-display{ font-family:'Inter', system-ui, sans-serif; letter-spacing:-0.02em; }
      `}</style>
      <div className="w-full max-w-sm pt-6 pb-24">
        <div className="flex items-center gap-2 justify-center mb-6">
          <div className="w-9 h-9 rounded bg-[#2E8B57] flex items-center justify-center font-mono-num font-bold text-white">A</div>
          <div>
            <div className="font-display font-extrabold text-lg text-slate-900 leading-none">ALUMAVEL</div>
            <div className="text-[10px] tracking-[0.15em] uppercase text-slate-400 mt-1">{modo === "cliente" ? "Portal de cliente" : "Panel de gestión"}</div>
          </div>
        </div>

        {!primeraVez && (
          <div className="flex gap-1 mb-4 bg-slate-200/60 rounded-md p-1">
            <button onClick={() => { setModo("empleado"); setError(""); }} className={`flex-1 flex items-center justify-center gap-1.5 text-xs font-semibold py-2 rounded transition ${modo === "empleado" ? "bg-white text-slate-800 shadow-sm" : "text-slate-500"}`}>
              <UserCog size={13} /> Empleado
            </button>
            <button onClick={() => { setModo("cliente"); setError(""); }} className={`flex-1 flex items-center justify-center gap-1.5 text-xs font-semibold py-2 rounded transition ${modo === "cliente" ? "bg-white text-slate-800 shadow-sm" : "text-slate-500"}`}>
              <Globe size={13} /> Cliente
            </button>
          </div>
        )}

        <CornerFrame className="bg-white border border-slate-200 rounded-lg p-6">
          {primeraVez ? (
            <>
              <div className="flex items-center gap-2 mb-1">
                <Lock size={16} className="text-[#2E8B57]" />
                <h1 className="font-display font-bold text-lg text-slate-900">Crear administrador</h1>
              </div>
              <p className="text-xs text-slate-400 mb-5">Es la primera vez que se abre este CRM. Crea el usuario administrador para empezar.</p>
              <form onSubmit={submitFirstAdmin} className="space-y-3">
                <div className="grid grid-cols-2 gap-3">
                  <Field label="Nombre" required><TextInput value={f.nombre} onChange={setField("nombre")} required /></Field>
                  <Field label="Apellidos"><TextInput value={f.apellidos} onChange={setField("apellidos")} /></Field>
                </div>
                <Field label="Email" required><TextInput type="email" value={f.email} onChange={setField("email")} required /></Field>
                <Field label="Contraseña" required><TextInput type="password" value={f.password} onChange={setField("password")} required /></Field>
                <Field label="Teléfono"><TextInput value={f.telefono} onChange={setField("telefono")} onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); submitFirstAdmin(e); } }} /></Field>
                <button type="submit" onClick={submitFirstAdmin} style={{ backgroundColor: "#2E8B57", color: "#ffffff" }} className="w-full flex items-center justify-center gap-2 hover:opacity-90 text-base font-bold py-4 rounded-md mt-2 cursor-pointer select-none">
                  <ShieldCheck size={18} /> Crear administrador y entrar
                </button>
              </form>
            </>
          ) : modo === "empleado" ? (
            <>
              <div className="flex items-center gap-2 mb-5">
                <Lock size={16} className="text-[#2E8B57]" />
                <h1 className="font-display font-bold text-lg text-slate-900">Acceso empleados</h1>
              </div>
              <form onSubmit={submitLogin} className="space-y-3">
                <Field label="Email" required><TextInput type="email" value={email} onChange={(e) => setEmail(e.target.value)} required autoFocus /></Field>
                <Field label="Contraseña" required><TextInput type="password" value={password} onChange={(e) => setPassword(e.target.value)} required /></Field>
                {error && <p className="text-xs text-rose-600">{error}</p>}
                <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="w-full flex items-center justify-center gap-1.5 text-sm font-semibold py-2.5 rounded-md mt-2">
                  <LogIn size={15} /> Entrar
                </button>
              </form>
            </>
          ) : (
            <>
              <div className="flex items-center gap-2 mb-5">
                <Globe size={16} className="text-[#2E8B57]" />
                <h1 className="font-display font-bold text-lg text-slate-900">Portal de cliente</h1>
              </div>
              <form onSubmit={submitLoginCliente} className="space-y-3">
                <Field label="Email" required><TextInput type="email" value={email} onChange={(e) => setEmail(e.target.value)} required autoFocus /></Field>
                <Field label="Contraseña" required><TextInput type="password" value={password} onChange={(e) => setPassword(e.target.value)} required /></Field>
                {error && <p className="text-xs text-rose-600">{error}</p>}
                <button type="submit" style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="w-full flex items-center justify-center gap-1.5 text-sm font-semibold py-2.5 rounded-md mt-2">
                  <LogIn size={15} /> Entrar
                </button>
              </form>
            </>
          )}
        </CornerFrame>
        <p className="text-center text-[11px] text-slate-400 mt-4">
          {primeraVez
            ? "Podrás dar de alta a más usuarios luego, desde Administración."
            : modo === "empleado"
            ? "¿No tienes cuenta? Pídele a un administrador que te dé de alta."
            : "El acceso al portal lo activa ALUMAVEL desde la ficha de cliente."}
        </p>
      </div>
    </div>
  );
}

/* ---- Portal de cliente (solo lectura) ---- */

function ClientePortal({ cliente, proyectos, facturas, incidencias, onLogout }) {
  const [tab, setTab] = useState("proyectos");

  return (
    <div className="min-h-screen bg-[#F4F5F3]" style={{ fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@500;600&display=swap');
        .font-mono-num{ font-family:'IBM Plex Mono', monospace; }
        .font-display{ font-family:'Inter', system-ui, sans-serif; letter-spacing:-0.02em; }
      `}</style>

      <header className="bg-[#2A1F3D] text-slate-200">
        <div className="max-w-4xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded bg-[#2E8B57] flex items-center justify-center font-mono-num font-bold text-sm">A</div>
            <div>
              <div className="font-display font-bold text-[15px] leading-none">ALUMAVEL</div>
              <div className="text-[10px] tracking-[0.15em] uppercase text-slate-400 mt-1 flex items-center gap-1"><Globe size={10} /> Portal de cliente</div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <span className="text-sm text-slate-300 hidden sm:inline">{cliente.nombre}</span>
            <button onClick={onLogout} className="flex items-center gap-1.5 text-xs font-semibold text-slate-400 hover:text-white border border-white/10 hover:border-white/20 rounded-md px-3 py-1.5 transition">
              <LogOut size={13} /> Salir
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-6 py-8">
        <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-1">Hola, {cliente.nombre}</h1>
        <p className="text-sm text-slate-500 mb-6">Aquí puedes consultar el estado de tus trabajos y facturas.</p>

        <div className="flex gap-1 mb-6 border-b border-slate-200">
          {[
            { id: "proyectos", label: `Mis trabajos (${proyectos.length})`, icon: Briefcase },
            { id: "facturas", label: `Mis facturas (${facturas.length})`, icon: Receipt },
            { id: "incidencias", label: `Incidencias (${incidencias.length})`, icon: AlertOctagon },
          ].map((t) => (
            <button key={t.id} onClick={() => setTab(t.id)} className={`flex items-center gap-1.5 px-4 py-2.5 text-sm font-semibold border-b-2 -mb-px transition ${tab === t.id ? "border-[#2E8B57] text-[#2E8B57]" : "border-transparent text-slate-500 hover:text-slate-700"}`}>
              <t.icon size={14} /> {t.label}
            </button>
          ))}
        </div>

        {tab === "proyectos" && (
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {proyectos.length === 0 ? (
              <p className="px-4 py-10 text-center text-sm text-slate-400">Todavía no tienes trabajos registrados.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-3 font-semibold">Nº</th>
                    <th className="px-4 py-3 font-semibold">Trabajo</th>
                    <th className="px-4 py-3 font-semibold">Estado presupuesto</th>
                    <th className="px-4 py-3 font-semibold">Estado trabajo</th>
                    <th className="px-4 py-3 font-semibold text-right">Importe</th>
                    <th className="px-4 py-3 font-semibold">Entrega prevista</th>
                  </tr>
                </thead>
                <tbody>
                  {proyectos.map((p) => (
                    <tr key={p.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-3 font-mono-num text-slate-500">#{p.numero}</td>
                      <td className="px-4 py-3 font-medium text-slate-800">{p.nombre}</td>
                      <td className="px-4 py-3"><Badge className={ESTADO_PRESUPUESTO_STYLE[p.estadoPresupuesto]}>{p.estadoPresupuesto}</Badge></td>
                      <td className="px-4 py-3"><Badge className={ESTADO_TRABAJO_STYLE[p.estadoTrabajo]}>{p.estadoTrabajo}</Badge></td>
                      <td className="px-4 py-3 text-right font-mono-num">{money(p.importePresupuesto)}</td>
                      <td className="px-4 py-3 text-slate-500">{fmtDate(p.fechaEntregaPrevista)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}

        {tab === "facturas" && (
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {facturas.length === 0 ? (
              <p className="px-4 py-10 text-center text-sm text-slate-400">Todavía no tienes facturas emitidas.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-3 font-semibold">Nº Factura</th>
                    <th className="px-4 py-3 font-semibold">Tipo</th>
                    <th className="px-4 py-3 font-semibold">Fecha</th>
                    <th className="px-4 py-3 font-semibold text-right">Total</th>
                    <th className="px-4 py-3 font-semibold text-right">Pagado</th>
                    <th className="px-4 py-3 font-semibold text-right">Saldo</th>
                    <th className="px-4 py-3 font-semibold">Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {facturas.map((f) => {
                    const pagado = (f.pagos || []).reduce((s, p) => s + (parseFloat(p.importe) || 0), 0);
                    const saldo = (parseFloat(f.total) || 0) - pagado;
                    return (
                      <tr key={f.id} className="border-b border-slate-100 last:border-0">
                        <td className="px-4 py-3 font-mono-num text-slate-500">{f.numero}</td>
                        <td className="px-4 py-3 text-slate-600">{f.tipo}</td>
                        <td className="px-4 py-3 text-slate-500">{fmtDate(f.fecha)}</td>
                        <td className="px-4 py-3 text-right font-mono-num">{money(f.total)}</td>
                        <td className="px-4 py-3 text-right font-mono-num">{money(pagado)}</td>
                        <td className="px-4 py-3 text-right font-mono-num">{money(Math.max(saldo, 0))}</td>
                        <td className="px-4 py-3"><Badge className={ESTADO_FACTURA_STYLE[estadoFacturaCalc(f)]}>{estadoFacturaCalc(f)}</Badge></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

        {tab === "incidencias" && (
          <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
            {incidencias.length === 0 ? (
              <p className="px-4 py-10 text-center text-sm text-slate-400">No tienes incidencias registradas.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
                    <th className="px-4 py-3 font-semibold">Nº</th>
                    <th className="px-4 py-3 font-semibold">Fecha</th>
                    <th className="px-4 py-3 font-semibold">Estado</th>
                  </tr>
                </thead>
                <tbody>
                  {incidencias.map((i) => (
                    <tr key={i.id} className="border-b border-slate-100 last:border-0">
                      <td className="px-4 py-3 font-mono-num text-slate-500">#{i.numero}</td>
                      <td className="px-4 py-3 text-slate-500">{fmtDate(i.fecha)}</td>
                      <td className="px-4 py-3"><Badge className={ESTADO_INCIDENCIA_STYLE[i.estadoIncidencia]}>{i.estadoIncidencia}</Badge></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

function AdministracionModulo({ usuarios, currentUser, onUpsert, onDelete }) {
  const [view, setView] = useState("list");
  const [editId, setEditId] = useState(null);
  const [descargandoBackup, setDescargandoBackup] = useState(false);

  const descargarCopiaSeguridad = async () => {
    setDescargandoBackup(true);
    try {
      const snap = await fbGet(ref(fbDb, "/"));
      const datos = JSON.stringify(snap.val(), null, 2);
      const blob = new Blob([datos], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `backup-alumavel-crm-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      alert("No se pudo descargar la copia de seguridad: " + err.message);
    } finally {
      setDescargandoBackup(false);
    }
  };

  if (view === "form") {
    const editing = usuarios.find((u) => u.id === editId) || null;
    return (
      <UsuarioForm
        initial={editing}
        onCancel={() => setView("list")}
        onSave={(data) => { onUpsert(data); setView("list"); }}
      />
    );
  }

  return (
    <div className="p-8 max-w-4xl">
      <Header
        icon={<UserCog size={20} className="text-[#2E8B57]" />}
        title="Administración de usuarios"
        manualKey="administracion"
        subtitle={`${usuarios.length} usuario${usuarios.length === 1 ? "" : "s"} con acceso al CRM`}
      />

      <div className="bg-white border border-slate-200 rounded-lg p-4 mb-6 flex items-center justify-between flex-wrap gap-3">
        <div>
          <h2 className="font-display font-bold text-slate-800 text-sm mb-1">Copia de seguridad</h2>
          <p className="text-xs text-slate-500">Descarga ahora mismo todos los datos del CRM en un archivo. Además, cada día se envía automáticamente una copia por email.</p>
        </div>
        <button
          onClick={descargarCopiaSeguridad}
          disabled={descargandoBackup}
          className="flex items-center gap-2 text-sm font-semibold text-slate-600 border border-slate-300 px-4 py-2.5 rounded-md hover:bg-slate-50 disabled:opacity-60 shrink-0"
        >
          <Download size={15} /> {descargandoBackup ? "Descargando..." : "Descargar copia de seguridad ahora"}
        </button>
      </div>

      <button
        onClick={() => { setEditId(null); setView("form"); }}
        className="w-full flex items-center justify-center gap-2 bg-rose-600 hover:bg-rose-700 active:bg-rose-800 text-white text-lg font-bold py-5 rounded-lg mb-6 shadow-md cursor-pointer select-none"
      >
        <Plus size={22} /> NUEVO USUARIO
      </button>

      <div className="bg-white rounded-lg border border-slate-200 overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="bg-slate-50 text-left text-[11px] uppercase tracking-wide text-slate-500 border-b border-slate-200">
              <th className="px-4 py-3 font-semibold">Nombre</th>
              <th className="px-4 py-3 font-semibold">Email</th>
              <th className="px-4 py-3 font-semibold">Teléfono</th>
              <th className="px-4 py-3 font-semibold">Rol</th>
              <th className="px-4 py-3"></th>
            </tr>
          </thead>
          <tbody>
            {usuarios.map((u) => (
              <tr key={u.id} className="border-b border-slate-100 last:border-0 hover:bg-slate-50">
                <td className="px-4 py-3 font-medium text-slate-800">
                  {u.nombre} {u.apellidos} {u.id === currentUser.id && <span className="text-xs text-slate-400">(tú)</span>}
                </td>
                <td className="px-4 py-3 text-slate-600">{u.email}</td>
                <td className="px-4 py-3 text-slate-600">{u.telefono || "—"}</td>
                <td className="px-4 py-3">
                  <Badge className={u.rol === "Administrador" ? "bg-sky-50 text-sky-700 ring-sky-200" : "bg-slate-100 text-slate-600 ring-slate-200"}>
                    {u.rol === "Administrador" && <ShieldCheck size={11} />} {u.rol}
                  </Badge>
                </td>
                <td className="px-4 py-3 text-right">
                  <div className="flex items-center justify-end gap-3">
                    <button onClick={() => { setEditId(u.id); setView("form"); }} className="text-slate-300 hover:text-[#2E8B57]"><Pencil size={15} /></button>
                    {u.id !== currentUser.id && (
                      <button onClick={() => { if (confirm(`¿Eliminar a ${u.nombre}?`)) onDelete(u.id); }} className="text-slate-300 hover:text-rose-500"><Trash2 size={15} /></button>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function UsuarioForm({ initial, onCancel, onSave }) {
  const [f, setF] = useState(
    initial || { id: null, nombre: "", apellidos: "", email: "", password: "", telefono: "", rol: "Usuario", modulos: MODULOS_DISPONIBLES.map((m) => m.id) }
  );
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const [errorMsg, setErrorMsg] = useState("");
  const modulosActuales = f.modulos || MODULOS_DISPONIBLES.map((m) => m.id);
  const toggleModulo = (id) => {
    setF({ ...f, modulos: modulosActuales.includes(id) ? modulosActuales.filter((m) => m !== id) : [...modulosActuales, id] });
  };
  const submit = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!f.nombre.trim() || !f.email.trim() || !f.password.trim()) {
      setErrorMsg("Faltan campos obligatorios: Nombre, Email y Contraseña.");
      return;
    }
    setErrorMsg("");
    onSave({ ...f, modulos: modulosActuales });
  };

  return (
    <div className="p-8 max-w-2xl">
      <button onClick={onCancel} className="flex items-center gap-1 text-sm text-slate-500 hover:text-slate-800 mb-5"><ChevronLeft size={16} /> Volver</button>
      <h1 className="font-display text-2xl font-extrabold text-slate-900 mb-6">{initial ? "Editar usuario" : "Alta de usuario"}</h1>
      {errorMsg && (
        <div className="mb-4 px-4 py-3 rounded-md bg-rose-50 border border-rose-300 text-rose-700 text-sm font-semibold">
          ⚠ {errorMsg}
        </div>
      )}
      <form onSubmit={submit} className="bg-white border border-slate-200 rounded-lg p-6 space-y-5">
        <div className="grid grid-cols-2 gap-4">
          <Field label="Nombre" required><TextInput value={f.nombre} onChange={set("nombre")} required /></Field>
          <Field label="Apellidos"><TextInput value={f.apellidos} onChange={set("apellidos")} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-4">
          <Field label="Email" required><TextInput type="email" value={f.email} onChange={set("email")} required /></Field>
          <Field label="Teléfono"><TextInput value={f.telefono} onChange={set("telefono")} /></Field>
        </div>
        <div className="grid grid-cols-2 gap-4">
          <Field label="Contraseña" required><TextInput type="password" value={f.password} onChange={set("password")} required /></Field>
          <Field label="Rol">
            <Select value={f.rol} onChange={set("rol")}>
              {ROLES.map((r) => <option key={r}>{r}</option>)}
            </Select>
          </Field>
        </div>

        {f.rol === "Administrador" ? (
          <p className="text-xs text-slate-400 px-1">Los administradores siempre tienen acceso a todos los módulos.</p>
        ) : (
          <Field label="Módulos a los que tiene acceso este usuario">
            <div className="grid grid-cols-2 gap-2 bg-slate-50 border border-slate-200 rounded-md p-3">
              {MODULOS_DISPONIBLES.map((m) => (
                <label key={m.id} className="flex items-center gap-2 text-sm text-slate-700">
                  <input type="checkbox" checked={modulosActuales.includes(m.id)} onChange={() => toggleModulo(m.id)} className="rounded border-slate-300 text-[#2E8B57] focus:ring-[#2E8B57]" />
                  {m.label}
                </label>
              ))}
            </div>
            <div className="flex gap-2 mt-2">
              <button type="button" onClick={() => setF({ ...f, modulos: MODULOS_DISPONIBLES.map((m) => m.id) })} className="text-xs font-semibold text-[#2E8B57] hover:underline">Marcar todos</button>
              <button type="button" onClick={() => setF({ ...f, modulos: [] })} className="text-xs font-semibold text-slate-400 hover:underline">Desmarcar todos</button>
            </div>
          </Field>
        )}

        <div className="flex flex-wrap justify-end gap-2 pt-2">
          <button type="button" onClick={onCancel} className="px-4 py-2.5 rounded-md text-sm font-semibold text-slate-600 hover:bg-slate-100">Cancelar</button>
          <button type="submit" onClick={submit} style={{ backgroundColor: "#2E8B57", color: "#ffffff", border: "2px solid #256E46" }} className="flex items-center gap-1.5 text-sm font-semibold px-5 py-2.5 rounded-md"><Save size={15} /> Guardar usuario</button>
        </div>
      </form>
    </div>
  );
}
