const SHEET_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vRjYETNmIABbU9kSn86fRD9p1v_TCcCXeTzC5qBRoMiAWgvO-HJeljEGgwy5H-dVCCkH5dDDlGEwhur/pub?gid=66268568&single=true&output=csv";

let polizas = [];
const dashboardCharts = {};
const norm = s => (s || "").toString().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

// -------------------- Cargar datos (actualización casi en tiempo real) --------------------
async function cargarDatos() {
  polizas = []; // limpia cache
  return new Promise((resolve, reject) => {
    const urlConCacheBust = `${SHEET_URL}&cacheBust=${Date.now()}`;
    Papa.parse(urlConCacheBust, {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: results => {
        const data = results.data.map(r => {
          for (let k in r) if (typeof r[k] === "string") r[k] = r[k].trim();
          return r;
        });
        polizas = data;
        console.log("✅ Datos actualizados:", polizas.length, "registros");
        resolve(polizas);
      },
      error: err => {
        console.error("❌ Error al cargar CSV:", err);
        reject(err);
      }
    });
  });
}

// -------------------- Refresco automático cada 2 minutos --------------------
setInterval(async () => {
  console.log("♻️ Refrescando datos desde Google Sheets...");
  await cargarDatos();
  renderResumen(polizas);
  renderTabla(polizas);
}, 120000);

// -------------------- Resumen general --------------------
function renderResumen(lista) {
  const grupos = {};
  lista.forEach(p => {
    const contrato = (p["No. De OM / Contrato"] || "").trim() || "Sin OM asignado";
    const cliente = (p["Cliente"] || "").trim();
    const estado = norm(p["Estado"]);
    const key = `${contrato}__${cliente}__${estado}`;
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(p);
  });

  let total = 0, expedidas = 0, enExp = 0, sinExp = 0, pag = 0, noPag = 0;

  for (const key in grupos) {
    total++;
    const filas = grupos[key];
    const estado = norm(filas[0]["Estado"]);
    const pagos = filas.map(f => norm(f["Pago"]));

    if (estado === "expedida") expedidas++;
    else if (estado === "en expedicion") enExp++;
    else if (estado === "sin expedir") sinExp++;

    if (pagos.every(p => p === "si")) pag++;
    else noPag++;
  }

  const cont = document.getElementById("resumen");
  cont.innerHTML = `
    <div class="resumen" style="display:flex; flex-wrap:wrap; justify-content:center;">
      <div class="item total" data-filtro="todos"><div class="numero">${total}</div><div class="etiqueta">📊 Total</div></div>
      <div class="item expedida" data-filtro="expedida"><div class="numero">${expedidas}</div><div class="etiqueta">✅ Expedidas</div></div>
      <div class="item en-expedicion" data-filtro="en expedicion"><div class="numero">${enExp}</div><div class="etiqueta">⏳ En expedición</div></div>
      <div class="item sin-expedir" data-filtro="sin expedir"><div class="numero">${sinExp}</div><div class="etiqueta">❌ Sin expedir</div></div>
      <div class="item pagados" data-filtro="pagados"><div class="numero">${pag}</div><div class="etiqueta">💰 Pagados</div></div>
      <div class="item no-pagados" data-filtro="no pagados"><div class="numero">${noPag}</div><div class="etiqueta">🚫 No pagados</div></div>
    </div>
  `;

  document.querySelectorAll("#resumen .item").forEach(item => {
    item.onclick = () => {
      const filtro = item.dataset.filtro;
      const resultado = [];

      for (const key in grupos) {
        const filas = grupos[key];
        const estado = norm(filas[0]["Estado"]);
        const pagos = filas.map(f => norm(f["Pago"]));
        let inc = false;

        if (filtro === "todos") inc = true;
        if (filtro === "expedida" && estado === "expedida") inc = true;
        if (filtro === "en expedicion" && estado === "en expedicion") inc = true;
        if (filtro === "sin expedir" && estado === "sin expedir") inc = true;
        if (filtro === "pagados" && pagos.every(p => p === "si")) inc = true;
        if (filtro === "no pagados" && pagos.some(p => p === "no")) inc = true;

        if (inc) resultado.push(...filas);
      }

      renderTabla(resultado);
    };
  });
}

// -------------------- Tabla principal --------------------
function renderTabla(lista) {
  const tbody = document.querySelector("#tablaResultados tbody");
  tbody.innerHTML = "";

  if (!lista || lista.length === 0) {
    tbody.innerHTML = `<tr><td colspan="8" style="text-align:center;">No se encontraron registros.</td></tr>`;
    return;
  }

  const grupos = {};
  lista.forEach(p => {
    const contrato = (p["No. De OM / Contrato"] || "").trim() || "Sin OM asignado";
    const cliente = (p["Cliente"] || "").trim();
    const key = `${contrato}__${cliente}`;
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(p);
  });

  for (const key in grupos) {
    const filas = grupos[key];
    const primera = filas[0];
    const tr = document.createElement("tr");

    const pagos = filas.map(f => norm(f["Pago"]));
    let pagoIcon = "❌";
    if (pagos.every(p => p === "si")) pagoIcon = "✅";
    else if (pagos.some(p => p === "si") && pagos.some(p => p === "no")) pagoIcon = "⚠️";

    const campos = [
      primera["No. De OM / Contrato"] || "Sin OM asignado",
      primera["Identificación"] || "",
      primera["Cliente"] || "",
      primera["Tipo Compra"] || "",
      primera["Comprador responsable"] || "",
      pagoIcon,
      primera["Estado"] || ""
    ];

    const labels = [
      "No. De OM / Contrato",
      "Identificación",
      "Cliente",
      "Tipo Compra",
      "Comprador responsable",
      "Pago",
      "Estado"
    ];

    campos.forEach((v, idx) => {
      const td = document.createElement("td");
      td.textContent = v;
      if (idx === 5) {
        td.classList.add("icono-pago");
        if (v === "✅") td.style.color = "green";
        else if (v === "⚠️") td.style.color = "#d88f00";
        else td.style.color = "red";
      }
      if (idx === 6) {
        td.style.cursor = "pointer";
        td.addEventListener("click", () => abrirModalDetalle(filas));
      }
      td.setAttribute("data-label", labels[idx]);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }
}

// -------------------- Modal detalle --------------------
function abrirModalDetalle(filas) {
  const body = document.getElementById("modalBody");
  body.innerHTML = "";

  filas.forEach(p => {
    const div = document.createElement("div");
    div.className = "detalle-certificado";
    div.style.marginBottom = "1rem";

    const info = {
      "Mes": p["Mes"],
      "Descripción": p["Descripcion"],
      "Moneda": p["Moneda"],
      "Valor Unitario": p["Valor Unitario"],
      "Valor Total": p["Valor Total"],
      "Tipo de trámite": p["Tipo de tramite"],
      "Anual o Puntual": p["Anual o Puntual"],
      "Vigencia inicio": p["Vigencia inicio"],
      "Vigencia fin": p["Vigencia fin"],
      "Ramo": p["Ramo"],
      "Compañía": p["Compañía"],
      "Número de póliza": p["Número de póliza"],
      "Certificado": p["Certificado"],
      "Tipo Movimiento": p["Tipo Movimiento"],
      "Valor póliza": p["Valor póliza"],
      "Pago": p["Pago"],
      "Estado": p["Estado"]
    };

    div.innerHTML = Object.entries(info)
      .filter(([k, v]) => v)
      .map(([k, v]) => `<div><strong>${k}:</strong> ${v}</div>`)
      .join("");

    body.appendChild(div);
  });

  document.getElementById("btnDescargarCSV").onclick = () => descargarCSV(filas);
  document.getElementById("btnDescargarPDF").onclick = () => descargarPDF(filas);

  document.getElementById("modalDetalle").classList.add("show");
}

function cerrarModal() {
  document.getElementById("modalDetalle").classList.remove("show");
}

// -------------------- Exportar --------------------
function descargarCSV(filas) {
  const csv = Papa.unparse(filas);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Detalle_${filas[0]["No. De OM / Contrato"] || "Sin_OM"}.csv`;
  a.click();
}

function descargarPDF(filas) {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit: "pt", format: "A4" });
  const width = doc.internal.pageSize.getWidth();

  doc.setFillColor(0, 88, 163);
  doc.rect(0, 0, width, 50, "F");
  doc.setFontSize(16);
  doc.setTextColor(255, 255, 255);
  doc.text(`Detalle del contrato ${filas[0]["No. De OM / Contrato"] || "Sin OM asignado"}`, width / 2, 30, { align: "center" });

  doc.setFontSize(12);
  doc.setTextColor(0, 0, 0);
  let y = 70;

  filas.forEach(p => {
    Object.entries(p).forEach(([k, v]) => {
      if (v && y < 770) {
        doc.text(`${k}: ${v}`, 40, y);
        y += 15;
      } else if (y >= 770) {
        doc.addPage();
        y = 40;
      }
    });
    y += 10;
  });

  doc.save(`Detalle_${filas[0]["No. De OM / Contrato"] || "Sin_OM"}.pdf`);
}

// -------------------- Dashboard (5 indicadores) --------------------
function mostrarDashboard(lista = polizas) {
  if (!lista || lista.length === 0) return;

  const grupos = {};
  lista.forEach(p => {
    const contrato = (p["No. De OM / Contrato"] || "").trim() || "Sin OM asignado";
    const cliente = (p["Cliente"] || "").trim();
    const estado = norm(p["Estado"]);
    const key = `${contrato}__${cliente}__${estado}`;
    if (!grupos[key]) grupos[key] = [];
    grupos[key].push(p);
  });

  let total = 0, expedidas = 0, enExp = 0, sinExp = 0, pagados = 0, noPagados = 0;
  for (const key in grupos) {
    total++;
    const filas = grupos[key];
    const estado = norm(filas[0]["Estado"]);
    const pagos = filas.map(f => norm(f["Pago"]));

    if (estado === "expedida") expedidas++;
    else if (estado === "en expedicion") enExp++;
    else if (estado === "sin expedir") sinExp++;

    if (pagos.every(p => p === "si")) pagados++;
    else noPagados++;
  }

  const charts = [
    { id: "graficoExpedidos", valor: expedidas, color: "#28a745", labelId: "labelExpedidos" },
    { id: "graficoEnExpedicion", valor: enExp, color: "#17a2b8", labelId: "labelEnExpedicion" },
    { id: "graficoSinExpedir", valor: sinExp, color: "#dc3545", labelId: "labelSinExpedir" },
    { id: "graficoPagados", valor: pagados, color: "#007bff", labelId: "labelPagados" },
    { id: "graficoNoPagados", valor: noPagados, color: "#ffc107", labelId: "labelNoPagados" },
  ];

  charts.forEach(c => {
    const porcentaje = total ? ((c.valor / total) * 100).toFixed(1) : 0;
    document.getElementById(c.labelId).textContent = `${porcentaje}% (${c.valor})`;
    const ctx = document.getElementById(c.id).getContext("2d");
    if (dashboardCharts[c.id]) dashboardCharts[c.id].destroy();

    dashboardCharts[c.id] = new Chart(ctx, {
      type: "doughnut",
      data: { labels: ["Completado", "Restante"], datasets: [{ data: [c.valor, total - c.valor], backgroundColor: [c.color, "#e0e0e0"], borderWidth: 0 }] },
      options: { rotation: -90, circumference: 180, cutout: "70%", plugins: { legend: { display: false }, tooltip: { enabled: false } }, responsive: false, maintainAspectRatio: false }
    });
  });

  document.getElementById("modalIndicadores").classList.add("show");
}

// -------------------- Eventos --------------------
function ejecutarBusqueda() {
  const q = document.getElementById("identificacion").value.trim().toLowerCase();
  let filtrada = polizas;
  if (q) filtrada = polizas.filter(p =>
    (p["Identificación"] || "").toLowerCase().includes(q) ||
    (p["No. De OM / Contrato"] || "").toLowerCase().includes(q)
  );
  renderResumen(filtrada);
  renderTabla(filtrada);
}

document.getElementById("btnConsultar").onclick = ejecutarBusqueda;
document.getElementById("identificacion").addEventListener("keydown", e => { if (e.key === "Enter") ejecutarBusqueda(); });
document.getElementById("btnIndicadores").onclick = () => mostrarDashboard(polizas);
document.getElementById("modalClose").onclick = cerrarModal;
document.getElementById("modalDetalle").addEventListener("click", e => { if (e.target.classList.contains("modal-backdrop")) cerrarModal(); });
document.getElementById("modalIndicadoresClose").onclick = () => document.getElementById("modalIndicadores").classList.remove("show");
document.getElementById("modalIndicadores").addEventListener("click", e => { if (e.target.classList.contains("modal-backdrop")) document.getElementById("modalIndicadores").classList.remove("show"); });

// -------------------- Inicialización --------------------
window.addEventListener("load", async () => {
  await cargarDatos();
  renderResumen(polizas);
  renderTabla(polizas);
});
