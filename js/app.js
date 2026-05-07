    let filasDatos = [];
    let columnas = {};
    let zonas = [];

    let filasEstadistica = [];
    let columnasEstadistica = {};

    const indicadores = [
    {
      id: 1,
      titulo: "Pérdidas AT",
      origen: "perdidasAT",
      actividad: "Pérdidas AT",
      tipo: "",
      formato: "porcentaje"
    },
    {
      id: 2,
      titulo: "Pérdidas MT",
      origen: "estadistica",
      actividad: "INDICE MT",
      tipo: "M",
      formato: "porcentaje"
    },
    {
      id: 3,
      titulo: "Energía Perdida",
      origen: "energiaPerdida",
      actividad: "Energía Perdida",
      tipo: "",
      formato: "gwh"
    },
    {
        id: 4,
        titulo: "Detección",
        origen: "datos",
        actividad: "1.4.1 Detección de Energía",
        tipo: "M"
    },
    {
        id: 5,
        titulo: "Recuperación",
        origen: "datos",
        actividad: "1.4.2 Cobro de Energía",
        tipo: "R"
    },
    {
        id: 6,
        titulo: "Modernización BT",
        origen: "datos",
        actividad: "2.2.1 Reemplazo de Medidores en Baja Tensión",
        tipo: "R"
    },
    {
        id: 7,
        titulo: "Modernización MT",
        origen: "datos",
        actividad: "2.1.1 Reemplazo de Medidores en Media Tensión",
        tipo: "R"
    },
    {
        id: 8,
        titulo: "Puntos de Intercambio",
        origen: "datos",
        actividad: "PUNTOS DE INTERCAMBIO",
        tipo: "M"
    },
    {
        id: 9,
        titulo: "Servicios Fuerza",
        origen: "datos",
        actividad: "SERVICIOS DE FUERZA",
        tipo: "M"
    },
    {
        id: 10,
        titulo: "Verificador Calibrador MT",
        origen: "datos",
        actividad: "1.5.1 Productividad de Verificador Calibrador Media Tensión",
        tipo: "R"
    },
    {
        id: 11,
        titulo: "Verificador Calibrador BT",
        origen: "datos",
        actividad: "1.5.2 Productividad de Verificador Calibrador Baja Tensión",
        tipo: "R"
    }
    ];

    const mapaZonasEstadistica = {
      "COR": ["DV000"],
      "TLAXCA": ["DV02"],
      "TEHUAC": ["DV03"],
      "MATAMO": ["DV04"],
      "SN MARTI": ["DV05"],
      "TECAMAC": ["DV06"],
      "PUEBLA P": ["DV07"],
      "PUEBLA O": ["DV08"],
      "PACHUC": ["DV11"],
      "TULAN": ["DV12"],
      "TULA": ["DV13", "DV1"]
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

            const hojaEstadistica = libro.Sheets["Estadistica Real"];

            if (!hojaEstadistica) {
            mostrarError("No se encontró la hoja Estadistica Real.");
            return;
            }

            const filasEst = XLSX.utils.sheet_to_json(hojaEstadistica, {
            header: 1,
            defval: "",
            raw: false
            });

            procesarEstadisticaReal(filasEst);
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
      filasEstadistica = filas.filter(fila =>
        fila.some(celda => String(celda).trim() !== "")
      );

      columnasEstadistica = {
        nombre: 2,
        actividad: 3,
        tipo: 4,
        primeraFecha: 5
      };
    }

    function mesTextoAFecha(mesTexto) {
      const meses = {
        // Español
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

        // Inglés
        "JAN": 0,
        "FEB": 1,
        "MAR": 2,
        "APR": 3,
        "MAY": 4,
        "JUN": 5,
        "JUL": 6,
        "AUG": 7,
        "SEP": 8,
        "OCT": 9,
        "NOV": 10,
        "DEC": 11
      };

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
        // Español
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

        // Inglés
        "JAN": 0,
        "FEB": 1,
        "MAR": 2,
        "APR": 3,
        "MAY": 4,
        "JUN": 5,
        "JUL": 6,
        "AUG": 7,
        "SEP": 8,
        "OCT": 9,
        "NOV": 10,
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

    function buscarColumnaMesEstadistica(mesSeleccionado) {
      const fechaMes = mesTextoAFecha(mesSeleccionado);

      if (!fechaMes) return -1;

      for (const fila of filasEstadistica) {
        const esEncabezado =
          normalizar(fila[0]).includes("NOMBREACTIVIDAD") ||
          normalizar(fila[0]).includes("CONCATENADO");

        if (!esEncabezado) continue;

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

    function buscarValorEstadistica(actividad, tipo, mesSeleccionado, zonaSeleccionada) {
      const columnaMes = buscarColumnaMesEstadistica(mesSeleccionado);

      if (columnaMes === -1) return "—";

      const actividadBuscada = normalizar(actividad);
      const tipoBuscado = normalizar(tipo);

      const zonaNormalizada = normalizar(zonaSeleccionada);
      const codigosPermitidos = mapaZonasEstadistica[zonaNormalizada] || [zonaSeleccionada];

      const filaEncontrada = filasEstadistica.find(fila => {
        const codigoFila = normalizar(fila[1]);
        const nombreFila = normalizar(fila[columnasEstadistica.nombre]);
        const actividadFila = normalizar(fila[columnasEstadistica.actividad]);
        const tipoFila = normalizar(fila[columnasEstadistica.tipo]);

        const coincideZona =
          codigosPermitidos.some(codigo => normalizar(codigo) === codigoFila) ||
          codigosPermitidos.some(codigo => normalizar(codigo) === nombreFila);

        const coincideActividad = actividadFila === actividadBuscada;

        const coincideTipo =
          tipoBuscado === "" ||
          tipoFila === tipoBuscado;

        return coincideZona && coincideActividad && coincideTipo;
      });

      if (!filaEncontrada) return "—";

      return filaEncontrada[columnaMes];
    }

    function calcularPerdidasAT(mesSeleccionado, zonaSeleccionada) {
      const energiaPerdida = buscarValorEstadistica(
        "Energía Perdida",
        "",
        mesSeleccionado,
        zonaSeleccionada
      );

      const energiaRecibida = buscarValorEstadistica(
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

    function obtenerEnergiaPerdida(mesSeleccionado, zonaSeleccionada) {
      return buscarValorEstadistica(
        "Energía Perdida",
        "",
        mesSeleccionado,
        zonaSeleccionada
      );
    }

    function calcularEnergiaPerdida(indicador, mesSeleccionado, zonaSeleccionada) {
      const actual = buscarValorEstadistica(
        indicador.actividad,
        indicador.tipo,
        mesSeleccionado,
        zonaSeleccionada
      );

      const anterior = buscarValorEstadistica(
        indicador.actividadAnterior,
        indicador.tipo,
        mesSeleccionado,
        zonaSeleccionada
      );

      const actualNum = Number(actual);
      const anteriorNum = Number(anterior);

      if (isNaN(actualNum) || isNaN(anteriorNum)) return "—";

      return actualNum - anteriorNum;
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

    function buscarValor(indicador, mesSeleccionado, zonaSeleccionada) {
      if (indicador.origen === "perdidasAT") {
        return calcularPerdidasAT(mesSeleccionado, zonaSeleccionada);
      }

      if (indicador.origen === "energiaPerdida") {
        return obtenerEnergiaPerdida(mesSeleccionado, zonaSeleccionada);
      }

      if (indicador.origen === "estadistica") {
        return buscarValorEstadistica(
          indicador.actividad,
          indicador.tipo,
          mesSeleccionado,
          zonaSeleccionada
        );
      }

      const zona = zonas.find(z => z.nombre === zonaSeleccionada);

      if (!zona) return "—";

      const actividadBuscada = normalizar(indicador.actividad);
      const tipoBuscado = normalizar(indicador.tipo);
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

      if (!filaEncontrada) {
        return "—";
      }

      return filaEncontrada[zona.indice];
    }

    function formatearValor(valor, formato) {
      if (valor === undefined || valor === null || valor === "") {
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

    function mostrarIndicadores() {
      const mesSeleccionado = selectMes.value;
      const zonaSeleccionada = selectZona.value;

      cards.innerHTML = "";

      indicadores.forEach(indicador => {
        let valor = "—";

        try {
          valor = buscarValor(indicador, mesSeleccionado, zonaSeleccionada);
        } catch (error) {
          console.error("Error en indicador", indicador.id, error);
          valor = "ERROR";
        }

        const card = document.createElement("div");
        card.className = "card";

        card.innerHTML = `
          <div class="card-num">${indicador.id}</div>
          <h3>${indicador.titulo}</h3>
          <div class="valor">${formatearValor(valor, indicador.formato)}</div>
          
        `;

        cards.appendChild(card);
      });
    }