let nomina = JSON.parse(localStorage.getItem('nomina')) || [];
let registrosHoy = JSON.parse(localStorage.getItem('registrosHoy')) || [];
let historial = JSON.parse(localStorage.getItem('historialAsistencia')) || {};
let html5QrcodeScanner = null;

document.addEventListener('DOMContentLoaded', () => {
  actualizarContadores();
  renderTablaHoy();
  renderHistorial();

  document.getElementById('excel-input').addEventListener('change', cargarExcel);

  document.getElementById('btn-borrar-nomina').addEventListener('click', () => {
    if (confirm('¿Desea borrar la nómina guardada?')) {
      localStorage.removeItem('nomina');
      nomina = [];
      actualizarContadores();
      alert('Nómina borrada.');
    }
  });

  document.getElementById('btn-generar-qr').addEventListener('click', generarCarnetsQR);
  document.getElementById('btn-abrir-camara').addEventListener('click', iniciarCamara);
  document.getElementById('btn-cerrar-camara').addEventListener('click', detenerCamara);

  document.getElementById('dni-manual').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      const dniInput = e.target.value.trim();
      if (dniInput) {
        registrarAsistencia(dniInput);
        e.target.value = '';
      }
    }
  });

  document.getElementById('btn-descargar-excel').addEventListener('click', () => descargarExcelHoy());
  document.getElementById('btn-cerrar-dia').addEventListener('click', cerrarDia);
  document.getElementById('btn-reiniciar-dia').addEventListener('click', () => {
    if (confirm('¿Está seguro de reiniciar los registros de hoy?')) {
      registrosHoy = [];
      localStorage.removeItem('registrosHoy');
      actualizarContadores();
      renderTablaHoy();
    }
  });
});

function cargarExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (evt) => {
    const data = new Uint8Array(evt.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    nomina = XLSX.utils.sheet_to_json(sheet);
    localStorage.setItem('nomina', JSON.stringify(nomina));
    actualizarContadores();
    alert(`Nómina cargada con éxito: ${nomina.length} estudiantes.`);
  };
  reader.readAsArrayBuffer(file);
}

function registrarAsistencia(dni) {
  const estudiante = nomina.find(e => String(e.DNI || e.dni).trim() === String(dni).trim());

  if (!estudiante) {
    alert(`El DNI ${dni} no se encuentra en la nómina.`);
    return;
  }

  const yaRegistrado = registrosHoy.some(r => String(r.dni) === String(dni).trim());
  if (yaRegistrado) {
    alert(`El estudiante ${estudiante.ESTUDIANTE || estudiante.estudiante} ya fue registrado hoy.`);
    return;
  }

  const ahora = new Date();
  const fechaStr = ahora.toLocaleDateString('es-PE');
  const horaStr = ahora.toLocaleTimeString('es-PE', { hour: '2-digit', minute: '2-digit', second: '2-digit' });

  const hora = ahora.getHours();
  const minuto = ahora.getMinutes();
  let estado = 'ASISTIÓ';
  if (hora > 7 || (hora === 7 && minuto > 50)) {
    estado = 'TARDE';
  }

  const registro = {
    dni: String(dni).trim(),
    estudiante: estudiante.ESTUDIANTE || estudiante.estudiante || 'Sin nombre',
    grado: estudiante.GRADO || estudiante.grado || '',
    seccion: estudiante.SECCIÓN || estudiante.SECCION || estudiante.seccion || '',
    estado: estado,
    fecha: fechaStr,
    hora: horaStr
  };

  registrosHoy.unshift(registro);
  localStorage.setItem('registrosHoy', JSON.stringify(registrosHoy));

  actualizarContadores();
  renderTablaHoy();
}

function renderTablaHoy() {
  const tbody = document.getElementById('tabla-registros');
  tbody.innerHTML = '';

  registrosHoy.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.dni}</td>
      <td>${r.estudiante}</td>
      <td>${r.grado}</td>
      <td>${r.seccion}</td>
      <td class="${r.estado === 'ASISTIÓ' ? 'estado-asistio' : 'estado-tarde'}">${r.estado}</td>
      <td>${r.fecha}</td>
      <td>${r.hora}</td>
    `;
    tbody.appendChild(tr);
  });
}

function actualizarContadores() {
  const total = nomina.length;
  const regCount = registrosHoy.length;
  const sinRegCount = total - regCount;

  document.getElementById('total-estudiantes').textContent = total;
  document.getElementById('total-registrados').textContent = regCount;
  document.getElementById('total-sin-registrar').textContent = sinRegCount < 0 ? 0 : sinRegCount;
}

function descargarExcelHoy(data = registrosHoy, nombreArchivo = `Asistencia_${new Date().toLocaleDateString('es-PE').replace(/\//g, '-')}.xlsx`) {
  if (data.length === 0) {
    alert('No hay registros para exportar.');
    return;
  }

  const dataExport = data.map(r => ({
    'DNI': r.dni,
    'ESTUDIANTE': r.estudiante,
    'GRADO': r.grado,
    'SECCIÓN': r.seccion,
    'ESTADO': r.estado,
    'FECHA': r.fecha,
    'HORA DE INGRESO': r.hora
  }));

  const ws = XLSX.utils.json_to_sheet(dataExport);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Asistencia");
  XLSX.writeFile(wb, nombreArchivo);
}

function cerrarDia() {
  if (registrosHoy.length === 0) {
    alert('No hay registros el día de hoy.');
    return;
  }

  const hoyStr = new Date().toLocaleDateString('es-PE');
  descargarExcelHoy();

  historial[hoyStr] = [...registrosHoy];
  localStorage.setItem('historialAsistencia', JSON.stringify(historial));

  registrosHoy = [];
  localStorage.removeItem('registrosHoy');

  actualizarContadores();
  renderTablaHoy();
  renderHistorial();
  alert('Día cerrado e historial guardado.');
}

function renderHistorial() {
  const tbody = document.getElementById('tabla-historial');
  tbody.innerHTML = '';

  Object.keys(historial).forEach(fecha => {
    const list = historial[fecha];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${fecha}</td>
      <td>${list.length} registros</td>
      <td><button class="btn btn-slate" style="padding: 4px 10px; font-size: 12px;" onclick="descargarHistorialFecha('${fecha}')">Descargar Excel</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function descargarHistorialFecha(fecha) {
  const data = historial[fecha];
  if (data) {
    descargarExcelHoy(data, `Asistencia_${fecha.replace(/\//g, '-')}.xlsx`);
  }
}

function generarCarnetsQR() {
  if (nomina.length === 0) {
    alert('Primero debe cargar un archivo Excel con la nómina.');
    return;
  }

  const contenedor = document.getElementById('contenedor-carnets');
  contenedor.innerHTML = '';

  nomina.forEach(alumno => {
    const dni = String(alumno.DNI || alumno.dni || '').trim();
    const nombre = alumno.ESTUDIANTE || alumno.estudiante || 'Sin nombre';
    const grado = alumno.GRADO || alumno.grado || '';
    const seccion = alumno.SECCIÓN || alumno.SECCION || alumno.seccion || '';

    if (dni) {
      const card = document.createElement('div');
      card.className = 'carnet-item';

      const title = document.createElement('strong');
      title.textContent = nombre;
      card.appendChild(title);

      const details = document.createElement('p');
      details.style.fontSize = '12px';
      details.style.margin = '4px 0';
      details.textContent = `DNI: ${dni} | ${grado} ${seccion}`;
      card.appendChild(details);

      const qrDiv = document.createElement('div');
      qrDiv.className = 'carnet-qr';
      card.appendChild(qrDiv);

      new QRCode(qrDiv, { text: dni, width: 120, height: 120 });
      contenedor.appendChild(card);
    }
  });

  document.getElementById('seccion-carnets').style.display = 'block';
  document.getElementById('seccion-carnets').scrollIntoView({ behavior: 'smooth' });
}

function iniciarCamara() {
  if (html5QrcodeScanner) return;

  html5QrcodeScanner = new Html5Qrcode("reader");
  html5QrcodeScanner.start(
    { facingMode: "environment" },
    { fps: 30, qrbox: { width: 300, height: 300 }, aspectRatio: 1.0 };
    (decodedText) => {
      registrarAsistencia(decodedText);
    },
    (errorMessage) => {}
  ).catch(err => {
    console.error(err);
    alert("No se pudo iniciar la cámara.");
  });
}

function detenerCamara() {
  if (html5QrcodeScanner) {
    html5QrcodeScanner.stop().then(() => {
      html5QrcodeScanner = null;
    }).catch(err => console.error(err));
  }
}
