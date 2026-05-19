let filasDatos = [];
let columnas = {};
let zonas = [];

let filasEstadisticaReal = [];
let columnasEstadisticaReal = {};

let filasEstadisticaMeta = [];
let columnasEstadisticaMeta = {};

const indicadores = [
  {
    id: 1,
    titulo: "Pérdidas AT",
    icono: "torre",
    origen: "perdidasAT",
    formato: "porcentaje",
    comparacion: "menorMejor"
  },
  {
    id: 2,
    titulo: "Pérdidas MT",
    icono: "poste",
    origen: "estadistica",
    actividadReal: "INDICE MT",
    tipoReal: "M",
    actividadMeta: "% Energía Perdida MT",
    tipoMeta: "M",
    formato: "porcentaje",
    comparacion: "menorMejor"
  },
  {
    id: 3,
    titulo: "Energía Perdida",
    icono: "energia",
    origen: "energiaPerdida",
    formato: "gwh",
    mostrarMeta: false
  },
  {
    id: 4,
    titulo: "Detección",
    icono: "lupa",
    origen: "datos",
    actividad: "1.4.1 Detección de Energía",
    comparacion: "mayorMejor"
  },
  {
    id: 5,
    titulo: "Recuperación",
    icono: "dinero",
    origen: "datos",
    actividad: "1.4.2 Cobro de Energía",
    comparacion: "mayorMejor"
  },
  {
    id: 6,
    titulo: "Modernización BT",
    icono: "medidor",
    origen: "datos",
    actividad: "2.2.1 Reemplazo de Medidores en Baja Tensión",
    comparacion: "mayorMejor"
  },
  {
    id: 7,
    titulo: "Modernización MT",
    icono: "medidor",
    origen: "datos",
    actividad: "2.1.1 Reemplazo de Medidores en Media Tensión",
    comparacion: "mayorMejor"
  },
  {
    id: 8,
    titulo: "Puntos de Intercambio",
    icono: "intercambio",
    origen: "datos",
    actividad: "PUNTOS DE INTERCAMBIO",
    comparacion: "mayorMejor"
  },
  {
    id: 9,
    titulo: "Servicios Fuerza",
    icono: "rayo",
    origen: "datos",
    actividad: "SERVICIOS DE FUERZA",
    comparacion: "mayorMejor"
  },
  {
    id: 10,
    titulo: "Verificador Calibrador MT",
    icono: "verificador",
    origen: "datos",
    actividad: "1.5.1 Productividad de Verificador Calibrador Media Tensión",
    comparacion: "mayorMejor"
  },
  {
    id: 11,
    titulo: "Verificador Calibrador BT",
    icono: "verificador",
    origen: "datos",
    actividad: "1.5.2 Productividad de Verificador Calibrador Baja Tensión",
    comparacion: "mayorMejor"
  },
  {
    id: 12,
    titulo: "Adicionales MED",
    icono: "clipboard",
    origen: "datos",
    actividad: "Adicionales MED",
    mostrarMeta: false
  },
  {
    id: 13,
    titulo: "Adicionales CCC",
    icono: "documento",
    origen: "datos",
    actividad: "Adicionales CCC",
    mostrarMeta: false
  },
  {
    id: 14,
    titulo: "Verificadores MT",
    icono: "checklist",
    origen: "datos",
    actividad: "Verificadores MT",
    mostrarMeta: false
  },
  {
    id: 15,
    titulo: "Verificadores BT",
    icono: "checklist",
    origen: "datos",
    actividad: "Verificadores BT",
    mostrarMeta: false
  },
  {
    id: 16,
    titulo: "Linieros",
    icono: "casco",
    origen: "datos",
    actividad: "Linieros",
    mostrarMeta: false
  },
  {
    id: 17,
    titulo: "Clientes",
    icono: "clientes",
    origen: "datos",
    actividad: "Clientes",
    mostrarMeta: false
  },
  {
    id: 18,
    titulo: "Ventas Totales",
    icono: "ventas",
    origen: "datos",
    actividad: "Ventas Totales",
    mostrarMeta: false
  }
];

const mapaZonasEstadistica = {
  "COR": ["DV000", "DIVISION"],
  "TLAXCA": ["DV02", "TLAXCALA"],
  "TLAXCALA": ["DV02", "TLAXCALA"],
  "TEHUAC": ["DV03", "TEHUACAN"],
  "TEHUACAN": ["DV03", "TEHUACAN"],
  "MATAMO": ["DV04", "MATAMOROS", "IZUCAR DE MATAMOROS"],
  "MATAMOROS": ["DV04", "MATAMOROS", "IZUCAR DE MATAMOROS"],
  "SN MARTIN": ["DV05", "SAN MARTIN", "SAN MARTIN TEXMELUCAN"],
  "SAN MARTIN": ["DV05", "SAN MARTIN", "SAN MARTIN TEXMELUCAN"],
  "SN MARTI": ["DV05", "SAN MARTIN", "SAN MARTIN TEXMELUCAN"],
  "TECAMA": ["DV06", "TECAMACHALCO"],
  "TECAMAC": ["DV06", "TECAMACHALCO"],
  "TECAMACHALCO": ["DV06", "TECAMACHALCO"],
  "PUEBLA PTE": ["DV07", "PUEBLA PONIENTE"],
  "PUEBLA P": ["DV07", "PUEBLA PONIENTE"],
  "PUEBLA PONIENTE": ["DV07", "PUEBLA PONIENTE"],
  "PUEBLA OTE": ["DV08", "PUEBLA ORIENTE"],
  "PUEBLA O": ["DV08", "PUEBLA ORIENTE"],
  "PUEBLA ORIENTE": ["DV08", "PUEBLA ORIENTE"],
  "PACHUC": ["DV11", "PACHUCA"],
  "PACHUCA": ["DV11", "PACHUCA"],
  "TULAN": ["DV12", "TULANCINGO"],
  "TULANCINGO": ["DV12", "TULANCINGO"],
  "TULA": ["DV13", "DV1", "TULA"]
};

const archivoInput = document.getElementById("archivoExcel");
const selectMes = document.getElementById("selectMes");
const selectZona = document.getElementById("selectZona");
const cards = document.getElementById("cards");
const controls = document.getElementById("controls");
const errorBox = document.getElementById("error");

archivoInput.addEventListener("change", leerArchivo);
selectMes.addEventListener("change", mostrarIndicadores);
selectZona.addEventListener("change", mostrarIndicadores);

function normalizar(texto) {
  return String(texto ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toUpperCase();
}

function mostrarError(mensaje) {
  errorBox.style.display = "block";
  errorBox.textContent = mensaje;
}

function limpiarError() {
  errorBox.style.display = "none";
  errorBox.textContent = "";
}

function leerArchivo(event) {
  limpiarError();
  mostrarLoading();

  const archivo = event.target.files[0];

  if (!archivo) {
    ocultarLoading();
    return;
  }

  const lector = new FileReader();

  lector.onload = function(e) {
    setTimeout(() => {
      try {
        const datos = new Uint8Array(e.target.result);
        const libro = XLSX.read(datos, { type: "array" });

        const nombreHojaDatos = libro.SheetNames.find(
          nombre => normalizar(nombre) === "DATOS"
        );

        if (!nombreHojaDatos) {
          mostrarError("No se encontró una hoja llamada DATOS.");
          return;
        }

        const hoja = libro.Sheets[nombreHojaDatos];

        const filas = XLSX.utils.sheet_to_json(hoja, {
          header: 1,
          defval: "",
          raw: false
        });

        const hojaReal = libro.Sheets["Estadistica Real"];
        const hojaMeta = libro.Sheets["Estadistica Meta"];

        if (!hojaReal) {
          mostrarError("No se encontró la hoja Estadistica Real.");
          return;
        }

        if (!hojaMeta) {
          mostrarError("No se encontró la hoja Estadistica Meta.");
          return;
        }

        const filasReal = XLSX.utils.sheet_to_json(hojaReal, {
          header: 1,
          defval: "",
          raw: false
        });

        const filasMeta = XLSX.utils.sheet_to_json(hojaMeta, {
          header: 1,
          defval: "",
          raw: false
        });

        procesarEstadisticaReal(filasReal);
        procesarEstadisticaMeta(filasMeta);
        procesarFilas(filas);

      } catch (error) {
        console.error(error);
        mostrarError("Hubo un problema al leer el archivo Excel.");
      } finally {
        ocultarLoading();
      }
    }, 300);
  };

  lector.readAsArrayBuffer(archivo);
}

function procesarFilas(filas) {
  let filaEncabezados = -1;

  for (let i = 0; i < filas.length; i++) {
    const filaNormalizada = filas[i].map(celda => normalizar(celda));

    const tieneMes = filaNormalizada.includes("MES");
    const tieneActividad = filaNormalizada.some(celda =>
      celda.includes("INDICE") && celda.includes("ACTIVIDAD")
    );
    const tieneTipo = filaNormalizada.includes("TIPO");

    if (tieneMes && tieneActividad && tieneTipo) {
      filaEncabezados = i;
      break;
    }
  }

  if (filaEncabezados === -1) {
    mostrarError("No se encontró la fila de encabezados: Mes, Índice-Actividad y Tipo.");
    return;
  }

  const encabezados = filas[filaEncabezados];
  const encabezadosNormalizados = encabezados.map(celda => normalizar(celda));

  columnas.mes = encabezadosNormalizados.indexOf("MES");
  columnas.tipo = encabezadosNormalizados.indexOf("TIPO");
  columnas.actividad = encabezadosNormalizados.findIndex(celda =>
    celda.includes("INDICE") && celda.includes("ACTIVIDAD")
  );

  zonas = [];

  for (let i = columnas.tipo + 1; i < encabezados.length; i++) {
    const nombreZona = String(encabezados[i] ?? "").trim();

    if (nombreZona !== "") {
      zonas.push({
        nombre: nombreZona,
        indice: i
      });
    }
  }

  filasDatos = filas
    .slice(filaEncabezados + 1)
    .filter(fila => fila.some(celda => String(celda).trim() !== ""));

  cargarMeses();
  cargarZonas();

  controls.style.display = "grid";

  mostrarIndicadores();
}

function procesarEstadisticaReal(filas) {
  filasEstadisticaReal = limpiarFilas(filas);

  columnasEstadisticaReal = {
    nombre: 2,
    actividad: 3,
    tipo: 4,
    primeraFecha: 5
  };
}

function procesarEstadisticaMeta(filas) {
  filasEstadisticaMeta = limpiarFilas(filas);

  columnasEstadisticaMeta = {
    nombre: 2,
    actividad: 3,
    tipo: 4,
    primeraFecha: 5
  };
}

function limpiarFilas(filas) {
  return filas.filter(fila =>
    fila.some(celda => String(celda).trim() !== "")
  );
}

function mesTextoAFecha(mesTexto) {
  const meses = {
    "ENE": 0,
    "FEB": 1,
    "MAR": 2,
    "ABR": 3,
    "MAY": 4,
    "JUN": 5,
    "JUL": 6,
    "AGO": 7,
    "SEP": 8,
    "OCT": 9,
    "NOV": 10,
    "DIC": 11,
    "JAN": 0,
    "APR": 3,
    "AUG": 7,
    "DEC": 11
  };

  const fechaDirecta = celdaAFecha(mesTexto);
  if (fechaDirecta) return fechaDirecta;

  const limpio = normalizar(mesTexto)
    .replace(".", "")
    .replace("/", "-");

  const partes = limpio.split("-");

  if (partes.length < 2) return null;

  const mesTextoCorto = partes[0].substring(0, 3);
  const mes = meses[mesTextoCorto];

  let anio = Number(partes[1]);

  if (anio < 100) {
    anio = 2000 + anio;
  }

  if (mes === undefined || isNaN(anio)) return null;

  return new Date(anio, mes, 1);
}

function celdaAFecha(valor) {
  if (valor === undefined || valor === null || valor === "") return null;

  if (valor instanceof Date) {
    return new Date(valor.getFullYear(), valor.getMonth(), 1);
  }

  const numero = Number(valor);

  if (!isNaN(numero) && numero > 1000) {
    const fechaBase = new Date(Date.UTC(1899, 11, 30));
    const fecha = new Date(fechaBase.getTime() + numero * 24 * 60 * 60 * 1000);
    return new Date(fecha.getUTCFullYear(), fecha.getUTCMonth(), 1);
  }

  const texto = normalizar(valor)
    .replace(".", "")
    .replace(/\//g, "-");

  const meses = {
    "ENE": 0,
    "FEB": 1,
    "MAR": 2,
    "ABR": 3,
    "MAY": 4,
    "JUN": 5,
    "JUL": 6,
    "AGO": 7,
    "SEP": 8,
    "OCT": 9,
    "NOV": 10,
    "DIC": 11,
    "JAN": 0,
    "APR": 3,
    "AUG": 7,
    "DEC": 11
  };

  const partes = texto.split("-");

  if (partes.length >= 2) {
    const mesTextoCorto = partes[0].substring(0, 3);

    if (meses[mesTextoCorto] !== undefined) {
      const mes = meses[mesTextoCorto];
      const anio = partes[1].length === 2 ? 2000 + Number(partes[1]) : Number(partes[1]);

      if (!isNaN(anio)) {
        return new Date(anio, mes, 1);
      }
    }

    const p1 = Number(partes[0]);
    const p2 = Number(partes[1]);
    const p3 = Number(partes[2]);

    if (!isNaN(p1) && !isNaN(p2)) {
      let anio;
      let mes;

      if (!isNaN(p3)) {
        anio = p3 < 100 ? 2000 + p3 : p3;
        mes = p2 - 1;
      } else {
        anio = p2 < 100 ? 2000 + p2 : p2;
        mes = p1 - 1;
      }

      if (!isNaN(anio) && !isNaN(mes)) {
        return new Date(anio, mes, 1);
      }
    }
  }

  return null;
}

function esFilaEncabezadoEstadistica(fila) {
  return fila.some(celda => {
    const texto = normalizar(celda);
    return (
      texto.includes("NOMBREACTIVIDAD") ||
      texto.includes("CONCATENADO") ||
      texto === "NO." ||
      texto === "ZONA"
    );
  });
}

function buscarColumnaMesEstadistica(filas, columnasEstadistica, mesSeleccionado) {
  const fechaMes = mesTextoAFecha(mesSeleccionado);

  if (!fechaMes) return -1;

  for (const fila of filas) {
    if (!esFilaEncabezadoEstadistica(fila)) continue;

    for (let i = columnasEstadistica.primeraFecha; i < fila.length; i++) {
      const fechaColumna = celdaAFecha(fila[i]);

      if (
        fechaColumna &&
        fechaColumna.getFullYear() === fechaMes.getFullYear() &&
        fechaColumna.getMonth() === fechaMes.getMonth()
      ) {
        return i;
      }
    }
  }

  return -1;
}

function obtenerCodigosZona(zonaSeleccionada) {
  const zonaNormalizada = normalizar(zonaSeleccionada);
  return mapaZonasEstadistica[zonaNormalizada] || [zonaNormalizada];
}

function buscarValorEstadisticaEn(filas, columnasEstadistica, actividad, tipo, mesSeleccionado, zonaSeleccionada) {
  const columnaMes = buscarColumnaMesEstadistica(filas, columnasEstadistica, mesSeleccionado);

  if (columnaMes === -1) return "—";

  const actividadBuscada = normalizar(actividad);
  const tipoBuscado = normalizar(tipo);
  const codigosPermitidos = obtenerCodigosZona(zonaSeleccionada);

  const coincidencias = filas.filter(fila => {
    if (esFilaEncabezadoEstadistica(fila)) return false;

    const codigoFila = normalizar(fila[1]);
    const nombreFila = normalizar(fila[columnasEstadistica.nombre]);
    const actividadFila = normalizar(fila[columnasEstadistica.actividad]);
    const tipoFila = normalizar(fila[columnasEstadistica.tipo]);

    const coincideZona = codigosPermitidos.some(codigo => {
      const codigoNormalizado = normalizar(codigo);
      return (
        codigoNormalizado === codigoFila ||
        codigoNormalizado === nombreFila ||
        nombreFila.includes(codigoNormalizado) ||
        codigoNormalizado.includes(nombreFila)
      );
    });

    const coincideActividadExacta = actividadFila === actividadBuscada;
    const coincideActividadParcial =
      actividadFila.includes(actividadBuscada) ||
      actividadBuscada.includes(actividadFila);

    const coincideTipo = tipoBuscado === "" || tipoFila === tipoBuscado;

    return coincideZona && (coincideActividadExacta || coincideActividadParcial) && coincideTipo;
  });

  if (coincidencias.length === 0) return "—";

  const coincidenciaExacta = coincidencias.find(fila =>
    normalizar(fila[columnasEstadistica.actividad]) === actividadBuscada
  );

  const filaEncontrada = coincidenciaExacta || coincidencias[0];

  return filaEncontrada[columnaMes];
}

function buscarValorEstadisticaReal(actividad, tipo, mesSeleccionado, zonaSeleccionada) {
  return buscarValorEstadisticaEn(
    filasEstadisticaReal,
    columnasEstadisticaReal,
    actividad,
    tipo,
    mesSeleccionado,
    zonaSeleccionada
  );
}

function buscarValorEstadisticaMeta(actividad, tipo, mesSeleccionado, zonaSeleccionada) {
  return buscarValorEstadisticaEn(
    filasEstadisticaMeta,
    columnasEstadisticaMeta,
    actividad,
    tipo,
    mesSeleccionado,
    zonaSeleccionada
  );
}

function calcularPerdidasATReal(mesSeleccionado, zonaSeleccionada) {
  const energiaPerdida = buscarValorEstadisticaReal(
    "Energía Perdida",
    "",
    mesSeleccionado,
    zonaSeleccionada
  );

  const energiaRecibida = buscarValorEstadisticaReal(
    "Energía Recibida",
    "",
    mesSeleccionado,
    zonaSeleccionada
  );

  const perdidaNum = convertirNumero(energiaPerdida);
  const recibidaNum = convertirNumero(energiaRecibida);

  if (isNaN(perdidaNum) || isNaN(recibidaNum) || recibidaNum === 0) {
    return "—";
  }

  return (perdidaNum / recibidaNum) * 100;
}

function calcularPerdidasATMeta(mesSeleccionado, zonaSeleccionada) {
  return buscarValorEstadisticaMeta(
    "% Energía Perdida",
    "M",
    mesSeleccionado,
    zonaSeleccionada
  );
}

function obtenerEnergiaPerdidaReal(mesSeleccionado, zonaSeleccionada) {
  return buscarValorEstadisticaReal(
    "Energía Perdida",
    "",
    mesSeleccionado,
    zonaSeleccionada
  );
}

function obtenerEnergiaPerdidaMeta(mesSeleccionado, zonaSeleccionada) {
  return buscarValorEstadisticaMeta(
    "Energía Perdida",
    "M",
    mesSeleccionado,
    zonaSeleccionada
  );
}

function cargarMeses() {
  const meses = [];

  filasDatos.forEach(fila => {
    const mes = String(fila[columnas.mes] ?? "").trim();

    if (mes !== "" && !meses.includes(mes)) {
      meses.push(mes);
    }
  });

  selectMes.innerHTML = "";

  meses.forEach(mes => {
    const option = document.createElement("option");
    option.value = mes;
    option.textContent = mes;
    selectMes.appendChild(option);
  });
}

function cargarZonas() {
  selectZona.innerHTML = "";

  zonas.forEach(zona => {
    const option = document.createElement("option");
    option.value = zona.nombre;
    option.textContent = zona.nombre;
    selectZona.appendChild(option);
  });
}

function convertirNumero(valor) {
  if (valor === undefined || valor === null || valor === "" || valor === "—") {
    return NaN;
  }

  const texto = String(valor)
    .replace(/,/g, "")
    .replace("%", "")
    .trim();

  return Number(texto);
}

function buscarValorDatos(indicador, tipo, mesSeleccionado, zonaSeleccionada) {
  const zona = zonas.find(z => z.nombre === zonaSeleccionada);

  if (!zona) return "—";

  const actividadBuscada = normalizar(indicador.actividad);
  const tipoBuscado = normalizar(tipo);
  const mesBuscado = normalizar(mesSeleccionado);

  const filaEncontrada = filasDatos.find(fila => {
    const mesFila = normalizar(fila[columnas.mes]);
    const actividadFila = normalizar(fila[columnas.actividad]);
    const tipoFila = normalizar(fila[columnas.tipo]);

    return (
      mesFila === mesBuscado &&
      actividadFila.includes(actividadBuscada) &&
      tipoFila === tipoBuscado
    );
  });

  if (!filaEncontrada) return "—";

  return filaEncontrada[zona.indice];
}

function buscarValores(indicador, mesSeleccionado, zonaSeleccionada) {
  if (indicador.origen === "perdidasAT") {
    return {
      meta: calcularPerdidasATMeta(mesSeleccionado, zonaSeleccionada),
      real: calcularPerdidasATReal(mesSeleccionado, zonaSeleccionada)
    };
  }

  if (indicador.origen === "energiaPerdida") {
    return {
      meta: obtenerEnergiaPerdidaMeta(mesSeleccionado, zonaSeleccionada),
      real: obtenerEnergiaPerdidaReal(mesSeleccionado, zonaSeleccionada)
    };
  }

  if (indicador.origen === "estadistica") {
    return {
      meta: buscarValorEstadisticaMeta(
        indicador.actividadMeta,
        indicador.tipoMeta,
        mesSeleccionado,
        zonaSeleccionada
      ),
      real: buscarValorEstadisticaReal(
        indicador.actividadReal,
        indicador.tipoReal,
        mesSeleccionado,
        zonaSeleccionada
      )
    };
  }

  return {
    meta: buscarValorDatos(indicador, "M", mesSeleccionado, zonaSeleccionada),
    real: buscarValorDatos(indicador, "R", mesSeleccionado, zonaSeleccionada)
  };
}

function formatearValor(valor, formato) {
  if (valor === undefined || valor === null || valor === "" || valor === "—") {
    return "—";
  }

  const numero = Number(String(valor).replace(/,/g, ""));

  if (isNaN(numero)) return valor;

  if (formato === "porcentaje") {
    return numero.toLocaleString("es-MX", {
      maximumFractionDigits: 2
    }) + "%";
  }

  if (formato === "gwh") {
    return numero.toLocaleString("es-MX", {
      maximumFractionDigits: 2
    }) + " GWh";
  }

  return numero.toLocaleString("es-MX", {
    maximumFractionDigits: 2
  });
}

function claseComparacion(indicador, valores) {
  if (!indicador.comparacion) return "";

  const real = convertirNumero(valores.real);
  const meta = convertirNumero(valores.meta);

  if (isNaN(real) || isNaN(meta)) return "";

  if (indicador.comparacion === "menorMejor") {
    return real <= meta ? "valor-ok" : "valor-mal";
  }

  if (indicador.comparacion === "mayorMejor") {
    return real >= meta ? "valor-ok" : "valor-mal";
  }

  return "";
}

function crearBloqueValor(etiqueta, valor, formato, claseExtra = "") {
  return `
    <div class="valor-bloque">
      <span class="valor-label">${etiqueta}</span>
      <span class="valor ${claseExtra}">${formatearValor(valor, formato)}</span>
    </div>
  `;
}

function obtenerSvgIcono(tipo) {
  const iconos = {
    torre: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M32 6 20 22h8l-6 10h8l-6 10h8l-6 16h16l-6-16h8l-6-10h8l-6-10h8L32 6zm-8 18h16m-14 10h12m-10 10h8" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    poste: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M30 8v48M18 18h24M14 26h32M22 18l-5 8M42 18l5 8M20 56h20" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    energia: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M22 18h22a6 6 0 0 1 6 6v16a6 6 0 0 1-6 6H22a6 6 0 0 1-6-6V24a6 6 0 0 1 6-6zm28 10h4v8h-4" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M33 22 26 34h7l-2 8 9-14h-7l0-6z" fill="currentColor"/></svg>`,
    lupa: `<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="28" cy="28" r="14" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M38 38l12 12" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/><path d="M28 21v14M21 28h14" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round"/></svg>`,
    dinero: `<svg viewBox="0 0 64 64" aria-hidden="true"><rect x="10" y="18" width="44" height="28" rx="6" fill="none" stroke="currentColor" stroke-width="3.5"/><circle cx="32" cy="32" r="6" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M18 24h.01M46 40h.01M32 10v6M32 48v6" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`,
    medidor: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M16 40a16 16 0 1 1 32 0v10H16V40z" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linejoin="round"/><path d="M32 40l8-8" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/><path d="M24 50v4M40 50v4" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`,
    intercambio: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M14 22h28l-6-6m6 6-6 6M50 42H22l6-6m-6 6 6 6" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    rayo: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M34 6 18 34h12l-4 24 20-30H34l0-22z" fill="currentColor"/></svg>`,
    verificador: `<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="24" cy="20" r="8" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M12 46c2-8 8-12 12-12s10 4 12 12" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/><path d="M40 34l5 5 9-11" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    clipboard: `<svg viewBox="0 0 64 64" aria-hidden="true"><rect x="18" y="14" width="28" height="40" rx="4" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M26 14v-2a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v2M24 28h16M24 36h16M24 44h10" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`,
    documento: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M20 10h18l10 10v34H20V10z" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linejoin="round"/><path d="M38 10v10h10M26 30h16M26 38h16M26 46h10" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`,
    checklist: `<svg viewBox="0 0 64 64" aria-hidden="true"><rect x="18" y="14" width="28" height="40" rx="4" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M24 28l3 3 5-6M24 40l3 3 5-6M34 28h6M34 40h6" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    casco: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M16 36a16 16 0 0 1 32 0v4H16v-4zm6 4v6h20v-6M32 20v8" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
    clientes: `<svg viewBox="0 0 64 64" aria-hidden="true"><circle cx="24" cy="24" r="7" fill="none" stroke="currentColor" stroke-width="3.5"/><circle cx="42" cy="26" r="6" fill="none" stroke="currentColor" stroke-width="3.5"/><path d="M12 48c2-8 8-12 12-12s10 4 12 12M36 48c1.5-6 6-9 10-9 3 0 7 2 9 9" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`,
    ventas: `<svg viewBox="0 0 64 64" aria-hidden="true"><path d="M14 50h36M18 42l10-10 8 8 14-16" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round" stroke-linejoin="round"/><path d="M46 24h4v4" fill="none" stroke="currentColor" stroke-width="3.5" stroke-linecap="round"/></svg>`
  };

  return iconos[tipo] || iconos.documento;
}

function mostrarIndicadores() {
  const mesSeleccionado = selectMes.value;
  const zonaSeleccionada = selectZona.value;

  cards.innerHTML = "";

  indicadores.forEach(indicador => {
    let valores = {
      meta: "—",
      real: "—"
    };

    try {
      valores = buscarValores(indicador, mesSeleccionado, zonaSeleccionada);
    } catch (error) {
      console.error("Error en indicador", indicador.id, error);
      valores = {
        meta: "ERROR",
        real: "ERROR"
      };
    }

    const card = document.createElement("div");
    card.className = "card";

    const mostrarMeta = indicador.mostrarMeta !== false;
    const claseReal = claseComparacion(indicador, valores);
    const claseContenedor = mostrarMeta ? "valores-dobles" : "valores-solo";
    const bloquesValores = mostrarMeta
      ? crearBloqueValor("Meta", valores.meta, indicador.formato) +
        crearBloqueValor("Real", valores.real, indicador.formato, claseReal)
      : crearBloqueValor("Real", valores.real, indicador.formato, claseReal);

    card.innerHTML = `
      <div class="card-num">${indicador.id}</div>
      <div class="card-icon-wrap">
        <div class="card-icon">${obtenerSvgIcono(indicador.icono)}</div>
      </div>
      <h3>${indicador.titulo}</h3>
      <div class="${claseContenedor}">
        ${bloquesValores}
      </div>
    `;

    cards.appendChild(card);
  });
}
