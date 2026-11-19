/* ============================================================
   ACI 211 · Tecnología del Concreto — interfaz (Equipo 6)
   Aquí armamos toda la vista: portada, formulario y ficha técnica.
   Nosotros NO calculamos nada: cada valor que se pinta sale del
   script original diseno_mezclas_aci211.py, que la API ejecuta y
   nos devuelve como JSON. En producción Vercel compila este
   archivo con Vite y lo publica en el dominio de la web.
   ============================================================ */
import React, { useEffect, useState } from "react";
import { disenar } from "./lib/api";
import {
  supabase,
  supabaseConfigurado,
  guardarDiseno,
  listarDisenos,
  eliminarDiseno,
} from "./lib/supabaseClient";

/* ============================================================
   Cargamos el formulario con el ejemplo "CORREGIDO" del script:
   dejar un campo vacío equivale a ENTER = valor por defecto.
   ============================================================ */
const DEFAULTS = {
  descripcion: "losa de concreto",
  fc: 210,
  responsable: "ING. FLAVIO CESAR TRUJILLO BAILON",
  consistencia: "2",
  aire: false,
  exposicion: "1",
  tipo_cemento: "Andino Tipo V",
  pe_cemento: 3.12,
  arena: { puss: 1580, pucs: 1710, pe: 2.65, humedad: 3.5, absorcion: 1.2, mf: 2.7 },
  piedra: { puss: 1550, pucs: 1660, pe: 2.62, humedad: 0.4, absorcion: 1.0 },
  tmn: "3",
  probeta: { diametro: 15, altura: 30, numero: 3, desperdicio: 20 },
  confirmar: true,
  informe: true,
};

/* Las 8 opciones de TMN que el script acepta (PASO 3) */
const TMN_OPCIONES = [
  ["1", '3/8"'], ["2", '1/2"'], ["3", '3/4"'], ["4", '1"'],
  ["5", '1 1/2"'], ["6", '2"'], ["7", '3"'], ["8", '6"'],
];

/* Los 15 pasos del método, listos para pintarse como chips en la portada */
const METODOLOGIA = [
  ["1", "f'cr requerida"], ["2", "Asentamiento"], ["3", "TMN"],
  ["4", "Agua unitaria"], ["5", "Relación a/c"], ["6", "Aire"],
  ["7", "b/bo grueso"], ["8", "Volumen absoluto"], ["9-12", "Humedad → agua efectiva"],
  ["13", "Proporción en peso"], ["14", "Por saco de cemento"], ["15", "Probetas"],
];

/* Formateamos los valores que llegan del script: "—" si falta algo,
   y num() nos da el número limpio (acepta coma decimal) para el gráfico */
const fmt = (v) => (v === null || v === undefined || v === "" ? "—" : v);
const num = (v) => (fmt(v) === "—" ? 0 : parseFloat(String(v).replace(",", ".")) || 0);

/* ============================ piezas ============================
   Componentes chiquitos que reutilizamos en toda la vista */

function Chip({ children, viva }) {
  return <span className={"chip" + (viva ? " viva" : "")}>{children}</span>;
}

/* Bloque numerado del formulario (01–05): cabecera con el PASO del
   método al que corresponde y cuerpo con sus campos */
function Grupo({ num, titulo, referencia, children }) {
  return (
    <section className="grupo">
      <div className="grupo-cab">
        <span className="num">{num}</span>
        <h2>{titulo}</h2>
        {referencia && <span className="ref">{referencia}</span>}
      </div>
      <div className="grupo-cuerpo">{children}</div>
    </section>
  );
}

/* Los 4 tipos de campo (texto, número, lista, interruptor) comparten
   el mismo .campo: etiqueta + unidad + control + ayuda opcional */
function Texto({ etiqueta, unidad, valor, onChange, ayuda, ancho }) {
  return (
    <label className={"campo" + (ancho ? " ancho" : "")}>
      <span>{etiqueta}{unidad && <em>{unidad}</em>}</span>
      <input type="text" value={valor} onChange={(e) => onChange(e.target.value)} />
      {ayuda && <span className="ayuda">{ayuda}</span>}
    </label>
  );
}

function Numero({ etiqueta, unidad, valor, onChange, ayuda, paso = "any", ancho }) {
  return (
    <label className={"campo num" + (ancho ? " ancho" : "")}>
      <span>{etiqueta}{unidad && <em>{unidad}</em>}</span>
      <input type="number" step={paso} value={valor}
        onChange={(e) => onChange(e.target.value)} />
      {ayuda && <span className="ayuda">{ayuda}</span>}
    </label>
  );
}

function Lista({ etiqueta, valor, onChange, opciones }) {
  return (
    <label className="campo">
      <span>{etiqueta}</span>
      <select value={valor} onChange={(e) => onChange(e.target.value)}>
        {opciones.map(([v, t]) => <option key={v} value={v}>{t}</option>)}
      </select>
    </label>
  );
}

function Interruptor({ etiqueta, valor, onChange }) {
  return (
    <label className="campo ancho">
      <span className="switch">
        <input type="checkbox" checked={valor} onChange={(e) => onChange(e.target.checked)} />
        <span className="pista" />
        <span>{etiqueta}</span>
      </span>
    </label>
  );
}

/* Tarjeta de indicador grande para la cabecera de la ficha */
function Kpi({ etiqueta, valor, unidad, nota, tono }) {
  return (
    <div className={"kpi" + (tono ? " tono-" + tono : "")}>
      <span className="valor">{valor}<small>{unidad}</small></span>
      <span className="etiqueta">{etiqueta}</span>
      {nota && <span className="nota">{nota}</span>}
    </div>
  );
}

/* Sección con título de la ficha técnica (tabla, gráfico, consola…) */
function Seccion({ titulo, chip, tono, children }) {
  return (
    <section className={"seccion" + (tono ? " tono-" + tono : "")}>
      <h3>{titulo}{chip}</h3>
      <div className="seccion-cuerpo">{children}</div>
    </section>
  );
}

/* Barras de distribución por m³: dibujamos las proporciones que
   devolvió el script (cemento, fino, grueso, agua) */
function GraficoDosis({ D }) {
  const filas = [
    ["Cemento", num(D.cemento), "var(--bar-1)", "kg/m³"],
    ["Ag. fino", num(D.fino), "var(--bar-2)", "kg/m³"],
    ["Ag. grueso", num(D.grueso), "var(--bar-3)", "kg/m³"],
    ["Agua efectiva", num(D.agua_efectiva), "var(--bar-4)", "lt/m³"],
  ];
  const max = Math.max(...filas.map((f) => f[1]), 1);
  return (
    <Seccion tono="cemento" titulo="Distribución de la mezcla" chip={<Chip>por m³</Chip>}>
      <div className="grafico">
        {filas.map(([nombre, v, color, unidad]) => (
          <div className="fila-g" key={nombre}>
            <span className="nombre">{nombre}</span>
            <span className="rail">
              <span className="relleno" style={{ width: (v / max) * 100 + "%", background: color }} />
            </span>
            <span className="val">{v.toFixed(2)} {unidad}</span>
          </div>
        ))}
      </div>
    </Seccion>
  );
}

/* ============================= app =============================
   Aquí vive todo el estado de la página y el flujo de una consulta */
export default function App() {
  const [datos, setDatos] = useState(DEFAULTS);
  const [resultado, setResultado] = useState(null);
  const [cargando, setCargando] = useState(false);
  const [error, setError] = useState(null);
  const [verConsola, setVerConsola] = useState(false);

  /* ---------- Supabase: cuenta e historial "Mis diseños" ---------- */
  const [sesion, setSesion] = useState(null);
  const [cuentaAbierta, setCuentaAbierta] = useState(false);
  const [modoAuth, setModoAuth] = useState("login"); // "login" | "registro"
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [authError, setAuthError] = useState(null);
  const [authMsg, setAuthMsg] = useState(null);
  const [guardando, setGuardando] = useState(false);
  const [guardado, setGuardado] = useState(null); // "ok" o mensaje de error
  const [misDisenos, setMisDisenos] = useState([]);
  const [verMisDisenos, setVerMisDisenos] = useState(false);
  const [cargandoLista, setCargandoLista] = useState(false);

  /* Al abrir la página recuperamos la sesión guardada en el navegador
     y nos suscribimos a los cambios de Supabase Auth */
  useEffect(() => {
    if (!supabase) return;
    supabase.auth.getSession().then(({ data }) => setSesion(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_evento, s) => setSesion(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  function abrirCuenta() {
    /* Cada vez que abrimos el panel empezamos en "Iniciar sesión":
       si el usuario había cambiado a "Crear cuenta" y cerró la ventana,
       no queremos que el modo de registro se quede pegado. */
    setModoAuth("login");
    setCuentaAbierta(true);
    setAuthError(null);
    setAuthMsg(null);
  }

  /* Registro o inicio de sesión con Supabase Auth */
  async function entrar() {
    setAuthError(null);
    setAuthMsg(null);
    const credenciales = { email: email.trim(), password: pass };
    const { error } =
      modoAuth === "login"
        ? await supabase.auth.signInWithPassword(credenciales)
        : await supabase.auth.signUp(credenciales);
    if (error) {
      setAuthError(error.message);
    } else if (modoAuth === "registro") {
      setAuthMsg("Cuenta creada. Si te pide confirmar el correo, revisa tu bandeja; luego inicia sesión.");
    } else {
      setCuentaAbierta(false);
      setPass("");
    }
  }

  async function salir() {
    await supabase.auth.signOut();
    setSesion(null);
    setMisDisenos([]);
    setVerMisDisenos(false);
  }

  /* Guardamos la ficha actual en Supabase (botón "Guardar diseño") */
  async function guardar() {
    setGuardando(true);
    setGuardado(null);
    try {
      await guardarDiseno(datos, resultado);
      setGuardado("ok");
      cargarMisDisenos();
    } catch (e) {
      setGuardado(e.message || "error desconocido");
    }
    setGuardando(false);
  }

  async function cargarMisDisenos() {
    setCargandoLista(true);
    try {
      setMisDisenos(await listarDisenos());
    } catch (e) {
      setMisDisenos([]);
    }
    setCargandoLista(false);
  }

  function abrirMisDisenos() {
    setVerMisDisenos(true);
    cargarMisDisenos();
  }

  /* Cargamos un diseño guardado de vuelta en el formulario */
  function cargarDiseno(d) {
    if (d.datos_entrada) setDatos(d.datos_entrada);
    setVerMisDisenos(false);
  }

  async function borrarDiseno(id) {
    try {
      await eliminarDiseno(id);
      setMisDisenos((lista) => lista.filter((d) => d.id !== id));
    } catch (e) {
      setGuardado(e.message || "no se pudo eliminar");
    }
  }

  /* Dos fábricas de manejadores: f() edita una sub-rama (arena,
     piedra, probeta) y directo() edita una clave suelta del formulario */
  const f = (rama, clave) => (v) =>
    setDatos((d) => ({ ...d, [rama]: { ...d[rama], [clave]: v } }));
  const directo = (clave) => (v) => setDatos((d) => ({ ...d, [clave]: v }));

  /* Al pulsar "Calcular dosificación": enviamos el formulario completo
     a la API (en producción: la función serverless de Vercel) y con la
     respuesta pintamos la ficha técnica o el error */
  async function calcular() {
    setCargando(true);
    setError(null);
    setResultado(null);
    try {
      const j = await disenar(datos);
      if (!j.ok) setError(j.detalle || "El script no devolvió resultados.");
      else setResultado(j);
    } catch (e) {
      setError("No se pudo contactar la API" +
        (e && e.message ? " · " + e.message : "") +
        ". En local corre backend/api.py (iniciar_web.bat); en producción revisa VITE_API_URL.");
    }
    setCargando(false);
  }

  const D = resultado?.dosificacion;

  return (
    <div className="app">
      {/* -------- nav sticky -------- */}
      <nav className="nav">
        <div className="nav-in">
          <div className="marca">
            <span className="sello">211</span>
            <span>
              <b>Tecnología del Concreto</b>
              <small>Diseño de mezclas · ACI 211.1</small>
            </span>
          </div>
          <span className="nav-badge">
            <span className="punto" />
            Equipo 6
          </span>
          {supabaseConfigurado && (
            <div className="nav-cuenta">
              {sesion ? (
                <>
                  <button className="nav-boton" onClick={abrirMisDisenos}>
                    Mis diseños
                  </button>
                  <span className="nav-email" title={sesion.user.email}>
                    {sesion.user.email}
                  </span>
                  <button className="nav-boton" onClick={salir}>Salir</button>
                </>
              ) : (
                <button className="nav-boton" onClick={abrirCuenta}>
                  Iniciar sesión
                </button>
              )}
            </div>
          )}
        </div>
      </nav>

      {/* -------- portada -------- */}
      <div className="portada">
        {/* orbes decorativos animados */}
        <div className="hero-orbs" aria-hidden="true">
          <span className="orb orb-1" />
          <span className="orb orb-2" />
          <span className="orb orb-3" />
        </div>

        <header className="hero">
          <div className="hero-content">
            <div className="hero-main">
              <h1>
                <span className="hero-pre">Diseño de mezclas de concreto</span>
                <span className="hero-accent">ACI 211</span>
              </h1>
              <p className="hero-desc">
                Calcula la dosificación por m³ en segundos.
                Ingresa tus materiales, ejecuta los <strong>15 pasos</strong> del
                método y descarga tu informe Excel al instante.
              </p>
              <div className="hero-actions">
                <a href="#formulario" className="hero-cta">
                  Comenzar diseño
                  <span className="hero-cta-arrow">↓</span>
                </a>
              </div>
            </div>

            {/* tarjetas flotantes de vidrio */}
            <div className="hero-cards">
              <div className="glass-card">
                <span className="glass-card-num">15</span>
                <span className="glass-card-label">Pasos del<br/>método</span>
              </div>
              <div className="glass-card glass-card--alt">
                <span className="glass-card-num">4</span>
                <span className="glass-card-label">Normas<br/>técnicas</span>
              </div>
              <div className="glass-card glass-card--accent">
                <span className="glass-card-num">5</span>
                <span className="glass-card-label">Hojas +<br/>3 gráficos</span>
              </div>
            </div>
          </div>
        </header>
      </div>

      {/* -------- principal -------- */}
      <main id="formulario" className="layout">
        {/* formulario */}
        <div>
          <Grupo num="01" titulo="Datos de diseño" referencia="PASOS 1–2">
            <Texto etiqueta="Descripción del elemento" valor={datos.descripcion}
              onChange={directo("descripcion")} ancho
              ayuda="losa, viga, columna, zapata…" />
            <Numero etiqueta="Resistencia especificada f'c" unidad="kg/cm²"
              valor={datos.fc} onChange={directo("fc")}
              ayuda="El script eleva a f'cr según ACI 318-19 (PASO 1)" />
            <Texto etiqueta="Responsable / docente" valor={datos.responsable}
              onChange={directo("responsable")} />
            <Lista etiqueta="Consistencia de la mezcla" valor={datos.consistencia}
              onChange={directo("consistencia")}
              opciones={[["1", 'Seca · 0" a 2"'], ["2", 'Plástica · 3" a 4"'], ["3", 'Fluida · ≥ 5"']]} />
            <Interruptor etiqueta="Concreto con aire incorporado"
              valor={datos.aire} onChange={directo("aire")} />
            {datos.aire && (
              <Lista etiqueta="Condición de exposición" valor={datos.exposicion}
                onChange={directo("exposicion")}
                opciones={[["1", "1 · Suave"], ["2", "2 · Moderada"], ["3", "3 · Severa"]]} />
            )}
          </Grupo>

          <Grupo num="02" titulo="Cemento" referencia="PASO 12">
            <Texto etiqueta="Tipo de cemento" valor={datos.tipo_cemento}
              onChange={directo("tipo_cemento")} />
            <Numero etiqueta="Peso específico" unidad="gr/cm³"
              valor={datos.pe_cemento} onChange={directo("pe_cemento")}
              ayuda="Volumen absoluto del cemento" />
          </Grupo>

          <Grupo num="03" titulo="Agregado fino · arena" referencia="PASOS 8–11">
            <Numero etiqueta="Peso unitario suelto seco (PUSS)" unidad="kg/m³"
              valor={datos.arena.puss} onChange={f("arena", "puss")} />
            <Numero etiqueta="Peso unitario compactado (PUCS)" unidad="kg/m³"
              valor={datos.arena.pucs} onChange={f("arena", "pucs")} />
            <Numero etiqueta="Peso específico de masa" valor={datos.arena.pe}
              onChange={f("arena", "pe")} />
            <Numero etiqueta="Módulo de fineza (MF)" valor={datos.arena.mf}
              onChange={f("arena", "mf")} ayuda="Tabla 04 · b/bo (PASO 7)" />
            <Numero etiqueta="% de humedad" unidad="%"
              valor={datos.arena.humedad} onChange={f("arena", "humedad")} />
            <Numero etiqueta="% de absorción" unidad="%"
              valor={datos.arena.absorcion} onChange={f("arena", "absorcion")} />
          </Grupo>

          <Grupo num="04" titulo="Agregado grueso · piedra" referencia="PASOS 3, 7 y 10">
            <Lista etiqueta="Tamaño máximo nominal (TMN)" valor={datos.tmn}
              onChange={directo("tmn")} opciones={TMN_OPCIONES} />
            <Numero etiqueta="Peso unitario suelto seco (PUSS)" unidad="kg/m³"
              valor={datos.piedra.puss} onChange={f("piedra", "puss")} />
            <Numero etiqueta="Peso unitario compactado (PUCS)" unidad="kg/m³"
              valor={datos.piedra.pucs} onChange={f("piedra", "pucs")} />
            <Numero etiqueta="Peso específico de masa" valor={datos.piedra.pe}
              onChange={f("piedra", "pe")} />
            <Numero etiqueta="% de humedad" unidad="%"
              valor={datos.piedra.humedad} onChange={f("piedra", "humedad")} />
            <Numero etiqueta="% de absorción" unidad="%"
              valor={datos.piedra.absorcion} onChange={f("piedra", "absorcion")} />
          </Grupo>

          <Grupo num="05" titulo="Probetas" referencia="PASO 15 · ASTM C192">
            <Numero etiqueta="Diámetro" unidad="cm" paso="1"
              valor={datos.probeta.diametro} onChange={f("probeta", "diametro")} />
            <Numero etiqueta="Altura" unidad="cm" paso="1"
              valor={datos.probeta.altura} onChange={f("probeta", "altura")} />
            <Numero etiqueta="Número de probetas" paso="1"
              valor={datos.probeta.numero} onChange={f("probeta", "numero")} />
            <Numero etiqueta="% de desperdicio" unidad="%"
              valor={datos.probeta.desperdicio} onChange={f("probeta", "desperdicio")} />
            <Interruptor etiqueta="Generar informe Excel · 5 hojas + 3 gráficos"
              valor={datos.informe} onChange={directo("informe")} />
          </Grupo>

          <div className="acciones">
            <button className="boton" onClick={calcular} disabled={cargando}>
              {cargando ? "Ejecutando el script…" : "Calcular dosificación"}
              <span className="flecha">→</span>
            </button>
            {resultado && (
              <button className="boton-fantasma" onClick={() => window.print()}>
                🖨️ Imprimir ficha
              </button>
            )}
          </div>
          {error && <p className="error">⚠️ {error}</p>}
        </div>

        {/* resultados */}
        <div>
          {!resultado && !cargando && (
            <div className="vacio">
              <div className="vacio-motas">
                <i className="mota m-cemento" /><i className="mota m-arena" />
                <i className="mota m-piedra" /><i className="mota m-agua" />
              </div>
              <h2>Ficha técnica de dosificación</h2>
              <p>Aún no hay cálculos. Pulsa <b>Calcular dosificación</b>;
              los campos vacíos usan los valores por defecto.</p>
            </div>
          )}

          {cargando && (
            <div className="vacio vacio-loading">
              <div className="spinner" />
              <h2>Ejecutando el diseño…</h2>
              <p>Resolviendo los 15 pasos del método y generando el informe Excel.</p>
            </div>
          )}

          {D && (
            <div className="resultados-wrapper">
              <div className="ficha-cab">
                <h2>Ficha técnica de dosificación</h2>
                <span className="meta">
                  {resultado.informe ? resultado.informe.nombre : "sin Excel"}
                </span>
              </div>

              <div className="kpis">
                <Kpi tono="ambar" etiqueta="f'cr" valor={fmt(D.fcr)} unidad="kg/cm²"
                  nota={"desde f'c = " + fmt(D.fc)} />
                <Kpi tono="cielo" etiqueta="a/c" valor={fmt(D.ac)} unidad=""
                  nota="Tabla 05 · PASO 5" />
                <Kpi tono="cemento" etiqueta="Cemento" valor={fmt(D.cemento)} unidad="kg/m³"
                  nota={fmt(D.sacos) + " sacos · 42.5 kg"} />
                <Kpi tono="agua" etiqueta="Agua efectiva" valor={fmt(D.agua_efectiva)} unidad="lt/m³"
                  nota="peso húmedo corregido" />
              </div>

              <Seccion tono="ambar" titulo="Dosificación por m³ de concreto"
                chip={<Chip viva>peso húmedo corregido</Chip>}>
                <table className="tabla">
                  <thead>
                    <tr><th>Componente</th><th className="der">Cantidad</th></tr>
                  </thead>
                  <tbody>
                    <tr><td className="nombre"><i className="mota m-cemento" />Cemento</td>
                      <td className="dato">{fmt(D.cemento)} kg/m³ · {fmt(D.sacos)} sacos</td></tr>
                    <tr><td className="nombre"><i className="mota m-arena" />Agregado fino</td>
                      <td className="dato">{fmt(D.fino)} kg/m³</td></tr>
                    <tr><td className="nombre"><i className="mota m-piedra" />Agregado grueso</td>
                      <td className="dato">{fmt(D.grueso)} kg/m³</td></tr>
                    <tr className="total"><td className="nombre"><i className="mota m-agua" />Agua efectiva</td>
                      <td className="dato">{fmt(D.agua_efectiva)} lt/m³</td></tr>
                  </tbody>
                </table>

                <div className="proporcion">
                  <span className="dosis">
                    1 <i>:</i> {fmt(D.prop_fino)} <i>:</i> {fmt(D.prop_grueso)}
                  </span>
                  <span className="nota">
                    <b>Proporción en peso</b> cemento : fino : grueso
                    · a/c = <b>{fmt(D.ac)}</b> · consistencia <b>{fmt(D.consistencia)}</b> ({fmt(D.slump_rango)})
                  </span>
                </div>
              </Seccion>

              <GraficoDosis D={D} />

              <Seccion tono="agua" titulo="Volumen por saco de cemento" chip={<Chip>PASO 14 · 42.5 kg</Chip>}>
                <table className="tabla">
                  <thead>
                    <tr><th>Componente</th><th className="der">Volumen</th></tr>
                  </thead>
                  <tbody>
                    <tr><td className="nombre"><i className="mota m-arena" />Agregado fino</td>
                      <td className="dato">{fmt(D.vol_fino_saco)} m³</td></tr>
                    <tr><td className="nombre"><i className="mota m-piedra" />Agregado grueso</td>
                      <td className="dato">{fmt(D.vol_grueso_saco)} m³</td></tr>
                    <tr className="total"><td className="nombre"><i className="mota m-agua" />Agua</td>
                      <td className="dato">{fmt(D.agua_saco)} lt</td></tr>
                  </tbody>
                </table>
              </Seccion>

              <Seccion tono="piedra" titulo="Dosificación para probetas"
                chip={<Chip>PASO 15 · ⌀{fmt(resultado.datos?.["PROBETA - Diametro (cm)"])}×{fmt(resultado.datos?.["PROBETA - Altura (cm)"])} cm</Chip>}>
                <table className="tabla">
                  <thead>
                    <tr><th>Material</th><th className="der">Cantidad</th></tr>
                  </thead>
                  <tbody>
                    <tr><td className="nombre"><i className="mota m-cemento" />Cemento</td>
                      <td className="dato">{fmt(D.cant_cemento)} kg · {fmt(D.num_sacos)} sacos</td></tr>
                    <tr><td className="nombre"><i className="mota m-arena" />Agregado fino</td>
                      <td className="dato">{fmt(D.cant_fino)} kg</td></tr>
                    <tr><td className="nombre"><i className="mota m-piedra" />Agregado grueso</td>
                      <td className="dato">{fmt(D.cant_grueso)} kg</td></tr>
                    <tr className="total"><td className="nombre"><i className="mota m-agua" />Agua efectiva</td>
                      <td className="dato">{fmt(D.cant_agua)} lt</td></tr>
                  </tbody>
                </table>
                <p style={{ margin: "12px 0 0", fontSize: "var(--fs-xs)", color: "var(--ink-3)" }}>
                  Volumen total V = {fmt(D.v_total)} m³ · desperdicio incluido:
                  {" "}{fmt(resultado.datos?.["PROBETA - Desperdicio (%)"])}% ·
                  {" "}{fmt(resultado.datos?.["PROBETA - Numero"])} probetas
                </p>
              </Seccion>

              {resultado.informe && (
                <a className="boton-oscuro" href={resultado.informe.url}
                  download={resultado.informe.nombre || "informe-aci211.xlsx"}>
                  ⬇️ Descargar informe Excel · {resultado.informe.nombre}
                </a>
              )}

              {/* Guardado en Supabase ("Mis diseños") */}
              {supabaseConfigurado && (
                sesion ? (
                  <button className="boton-guardar" onClick={guardar} disabled={guardando}>
                    💾 {guardando ? "Guardando…" : "Guardar diseño"}
                  </button>
                ) : (
                  <p className="aviso-cuenta">
                    ¿Quieres conservar este diseño?{" "}
                    <button className="enlace" onClick={abrirCuenta}>Inicia sesión</button>{" "}
                    y quedará en tu historial.
                  </p>
                )
              )}
              {guardado === "ok" && (
                <p className="guardado-ok">✓ Guardado en "Mis diseños" (Supabase)</p>
              )}
              {guardado && guardado !== "ok" && (
                <p className="guardado-error">No se pudo guardar: {guardado}</p>
              )}

              <Seccion titulo="Datos ingresados"
                chip={<Chip>{Object.keys(resultado.datos).length} campos</Chip>}>
                <table className="tabla">
                  <tbody>
                    {Object.entries(resultado.datos).map(([k, v]) => (
                      <tr key={k}><td className="nombre">{k}</td>
                        <td className="dato">{v}</td></tr>
                    ))}
                  </tbody>
                </table>
              </Seccion>

              <Seccion titulo="Salida de consola" chip={<Chip>sin alterar</Chip>}>
                <button className="mini-boton" onClick={() => setVerConsola(!verConsola)}>
                  {verConsola ? "▾ Ocultar consola" : "▸ Mostrar consola"}
                </button>
                {verConsola && (
                  <div className="terminal">
                    <div className="terminal-barra">
                      Consola del programa
                    </div>
                    <pre>{resultado.salida}</pre>
                  </div>
                )}
              </Seccion>
            </div>
          )}
        </div>
      </main>

      {/* -------- modal de cuenta (Supabase Auth) -------- */}
      {cuentaAbierta && (
        <div className="modal-fondo" onClick={() => setCuentaAbierta(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-cab">
              <h3>{modoAuth === "login" ? "Iniciar sesión" : "Crear cuenta"}</h3>
              <button className="modal-x" onClick={() => setCuentaAbierta(false)}>×</button>
            </div>
            <p className="modal-sub">
              Tu historial se guarda en Supabase; cada cuenta ve solo sus diseños.
            </p>
            <label className="campo">
              <span>Correo electrónico</span>
              <input type="email" value={email} autoComplete="email"
                onChange={(e) => setEmail(e.target.value)} />
            </label>
            <label className="campo">
              <span>Contraseña</span>
              <input type="password" value={pass}
                autoComplete={modoAuth === "login" ? "current-password" : "new-password"}
                onChange={(e) => setPass(e.target.value)} />
            </label>
            {authError && <p className="error" style={{ margin: 0 }}>⚠️ {authError}</p>}
            {authMsg && <p className="guardado-ok" style={{ margin: 0 }}>{authMsg}</p>}
            <button className="boton modal-boton" onClick={entrar}>
              {modoAuth === "login" ? "Entrar" : "Registrarme"}
            </button>
            <button className="enlace modal-cambio"
              onClick={() => {
                setModoAuth(modoAuth === "login" ? "registro" : "login");
                setAuthError(null);
                setAuthMsg(null);
              }}>
              {modoAuth === "login" ? "¿No tienes cuenta? Crear una" : "¿Ya tienes cuenta? Iniciar sesión"}
            </button>
          </div>
        </div>
      )}

      {/* -------- modal "Mis diseños" (historial en Supabase) -------- */}
      {verMisDisenos && (
        <div className="modal-fondo" onClick={() => setVerMisDisenos(false)}>
          <div className="modal modal-ancho" onClick={(e) => e.stopPropagation()}>
            <div className="modal-cab">
              <h3>Mis diseños</h3>
              <button className="modal-x" onClick={() => setVerMisDisenos(false)}>×</button>
            </div>
            {cargandoLista ? (
              <p className="modal-sub">Cargando…</p>
            ) : misDisenos.length === 0 ? (
              <p className="modal-sub">
                Aún no has guardado diseños. Calcula uno y pulsa "Guardar diseño".
              </p>
            ) : (
              <div className="lista-disenos">
                {misDisenos.map((d) => (
                  <div className="diseno-item" key={d.id}>
                    <div className="diseno-info">
                      <b>{d.descripcion || "Sin descripción"}</b>
                      <span>
                        {new Date(d.creado_en).toLocaleString("es-PE")} · f'c {d.fc ?? "—"} kg/cm²
                        {d.dosificacion?.fcr ? " · f'cr " + d.dosificacion.fcr : ""}
                      </span>
                    </div>
                    <div className="diseno-acciones">
                      <button className="mini-boton" onClick={() => cargarDiseno(d)}>Cargar</button>
                      <button className="mini-boton peligro" onClick={() => borrarDiseno(d.id)}>Eliminar</button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* -------- pie -------- */}
      <footer className="pie">
        <div className="pie-in">
          <div>
            <div className="normas">
              <span className="norma">ACI 211.1-91 · proporciones</span>
              <span className="norma">ACI 318-19 · f'cr</span>
              <span className="norma">ASTM C33 · agregados</span>
              <span className="norma">ASTM C192/C39 · probetas</span>
            </div>
          </div>
          <div>
            <p className="equipo">
              <b>Equipo 6</b> — Xiomara Nuñez · Aymi Juzarit · José Cunya ·
              Anderson Flores · Luis Espinoza · Christian Llerena<br />
              Docente: ING. FLAVIO CESAR TRUJILLO BAILON
            </p>
            <p className="pie-copy">© 2024 Tecnología del Concreto · Universidad</p>
          </div>
        </div>
      </footer>
    </div>
  );
}
