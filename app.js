
const $ = id => document.getElementById(id);
let validationIssues = [], excelData = [];

document.querySelectorAll('.tab').forEach(b => b.addEventListener('click', () => {
  document.querySelectorAll('.tab').forEach(x => x.classList.toggle('active', x === b));
  document.querySelectorAll('.page').forEach(p => p.classList.toggle('active', p.id === b.dataset.page));
}));

const norm = s => String(s ?? '').trim();

function parseDate(v) {
  if (v instanceof Date && !isNaN(v))
    return `${String(v.getDate()).padStart(2,'0')}/${String(v.getMonth()+1).padStart(2,'0')}/${v.getFullYear()}`;

  if (typeof v === 'number' && window.XLSX) {
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return `${String(d.d).padStart(2,'0')}/${String(d.m).padStart(2,'0')}/${d.y}`;
  }

  let s = norm(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    const [y,m,d] = s.slice(0,10).split('-');
    return `${d}/${m}/${y}`;
  }
  return s;
}

function dateValid(s) {
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!m) return false;
  const d = new Date(+m[3], +m[2]-1, +m[1]);
  return d.getFullYear() === +m[3] &&
    d.getMonth() === +m[2]-1 &&
    d.getDate() === +m[1];
}

function validRuc(s) {
  if (!/^\d{11}$/.test(s)) return false;
  const w = [5,4,3,2,7,6,5,4,3,2];
  let sum = 0;
  for (let i=0; i<10; i++) sum += Number(s[i]) * w[i];
  const r = 11 - (sum % 11);
  const d = r === 10 ? 0 : r === 11 ? 1 : r;
  return d === Number(s[10]);
}

function addIssue(row, field, message, value = '') {
  validationIssues.push({row, field, message, value});
}

async function readFile(file) {
  if (file.name.toLowerCase().endsWith('.zip')) {
    const z = await JSZip.loadAsync(file);
    const entries = Object.values(z.files).filter(x =>
      !x.dir && x.name.toLowerCase().endsWith('.txt')
    );
    if (!entries.length) throw Error('El ZIP no contiene archivos TXT.');
    let out = [];
    for (const e of entries)
      out.push({name:e.name, text:await e.async('string')});
    return out;
  }
  return [{name:file.name, text:await file.text()}];
}

$('txtFile').addEventListener('change', e => {
  $('fileInfo').textContent = e.target.files[0]?.name || 'Aún no has seleccionado un archivo.';
});

$('validateBtn').addEventListener('click', async () => {
  const file = $('txtFile').files[0];
  const ruc = norm($('ruc').value);
  const period = $('period').value.replace('-','');

  if (!file) {
    alert('Selecciona un archivo TXT o ZIP.');
    return;
  }
  if (!validRuc(ruc)) {
    alert('Ingresa un RUC válido de 11 dígitos.');
    return;
  }
  if (!/^\d{6}$/.test(period)) {
    alert('Selecciona el periodo tributario.');
    return;
  }

  try {
    validationIssues = [];
    const files = await readFile(file);
    let total = 0;

    for (const f of files) {
      const lines = f.text.replace(/^\uFEFF/,'').split(/\r?\n/);

      lines.forEach((raw,i) => {
        if (!raw.trim()) return;
        total++;
        const cols = raw.split('|');
        const row = i+1;

        if (cols.length < 2) {
          addIssue(row,'Estructura','No se detectaron campos separados por |',raw);
          return;
        }

        if (norm(cols[0]) !== ruc)
          addIssue(row,'Campo 1 - RUC','El RUC de la fila no coincide con el RUC ingresado',cols[0]);

        if (cols.some(c => c.includes('\t')))
          addIssue(row,'Formato','Se detectó tabulación dentro de un campo','');

        if (cols.length !== 80 && $('kind').value === 'RCE')
          addIssue(row,'Cantidad de campos','RCE: revisar cantidad de campos. El archivo de reemplazo RCE contempla hasta 80 posiciones según el Anexo 11.',cols.length);

        if (cols.length !== 33 && $('kind').value === 'RVIE')
          addIssue(row,'Cantidad de campos','RVIE: revisar estructura. La versión de reemplazo RVIE contempla campos de datos y posiciones automáticas según Anexo 3; verifica la estructura vigente.',cols.length);

        if (cols.length > 0 && cols[0] && !/^\d{11}$/.test(cols[0]))
          addIssue(row,'RUC','Debe contener 11 dígitos',cols[0]);
      });
    }

    const bad = new Set(validationIssues.map(x => x.row)).size;
    $('totalRows').textContent = total;
    $('badRows').textContent = bad;
    $('errorCount').textContent = validationIssues.length;
    $('resultState').textContent = validationIssues.length ? 'OBSERVADO' : 'SIN ERRORES BÁSICOS';
    $('resultState').className = validationIssues.length ? 'bad' : 'ok';

    $('issues').innerHTML = validationIssues.length
      ? validationIssues.map(x => `<tr>
          <td>${x.row}</td>
          <td>${escapeHtml(x.field)}</td>
          <td>${escapeHtml(x.message)}</td>
          <td>${escapeHtml(x.value)}</td>
        </tr>`).join('')
      : '<tr><td colspan="4" class="ok">No se detectaron inconsistencias en los controles básicos aplicados.</td></tr>';

    $('validationNote').textContent =
      `Archivos revisados: ${files.length}. Filas con contenido: ${total}. La revisión es preliminar y no verifica catálogos completos, reglas cruzadas, identidad del emisor ni aceptación SUNAT.`;

    $('validationResults').classList.remove('hidden');
    window.lastValidation = {
      kind:$('kind').value, ruc, period,
      issues:validationIssues, total
    };

  } catch(err) {
    alert('No se pudo leer el archivo: ' + err.message);
  }
});

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&':'&amp;', '<':'&lt;', '>':'&gt;',
    '"':'&quot;', "'":'&#39;'
  }[c]));
}

$('downloadReport').addEventListener('click', () => {
  const rows = [
    ['FILA','CAMPO','OBSERVACION','VALOR'],
    ...validationIssues.map(x => [x.row,x.field,x.message,x.value])
  ];

  const csv = '\uFEFF' + rows.map(r =>
    r.map(v => '"' + String(v ?? '').replace(/"/g,'""') + '"').join(';')
  ).join('\r\n');

  downloadBlob(csv,'REPORTE_PREVALIDACION_SIRE.csv','text/csv;charset=utf-8');
});

$('resetValidate').addEventListener('click', () => {
  $('txtFile').value = '';
  $('ruc').value = '';
  $('validationResults').classList.add('hidden');
  $('fileInfo').textContent = 'Aún no has seleccionado un archivo.';
  validationIssues = [];
});

function downloadBlob(content,name,type) {
  const a = document.createElement('a');
  const url = URL.createObjectURL(new Blob([content],{type}));
  a.href = url;
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url),1000);
}

$('loadExcel').addEventListener('click', async () => {
  const file = $('excelFile').files[0];

  if (!file) {
    alert('Selecciona un archivo Excel o CSV.');
    return;
  }

  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf,{type:'array',cellDates:true});
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true});

    let header = -1;

    for (let i=0; i<Math.min(10,rows.length); i++) {
      const h = rows[i].map(x => norm(x).toUpperCase());
      if (h.includes('RUC') && h.includes('TIPO') &&
          h.includes('SERIE') && h.includes('NUMERO') &&
          (h.includes('TOTAL') || h.includes('IMPORTE'))) {
        header = i;
        break;
      }
    }

    let indexes;

    if (header >= 0) {
      const h = rows[header].map(x => norm(x).toUpperCase());
      const idx = k => h.findIndex(x => x === k);

      indexes = [
        idx('RUC'), idx('TIPO'), idx('SERIE'),
        idx('NUMERO'), idx('FECHA'),
        h.findIndex(x => ['TOTAL','IMPORTE','IMPORTE TOTAL'].includes(x))
      ];
    } else {
      header = 0;
      indexes = [0,1,2,3,4,5];
    }

    if (indexes.some(i => i < 0))
      throw Error('No se encontraron las seis columnas requeridas (RUC, TIPO, SERIE, NUMERO, FECHA y TOTAL).');

    excelData = [];

    for (let i=header+1; i<rows.length; i++) {
      const r = rows[i], v = indexes.map(ix => r[ix]);
      if (v.every(x => norm(x) === '')) continue;

      excelData.push({
        ruc:norm(v[0]),
        tipo:norm(v[1]).padStart(2,'0'),
        serie:norm(v[2]),
        numero:norm(v[3]),
        fecha:parseDate(v[4]),
        total:norm(v[5]),
        sourceRow:i+1
      });
    }

    renderExcel();
    $('excelInfo').textContent =
      `Archivo: ${file.name}. Encabezados detectados en fila ${header+1}; registros leídos: ${excelData.length}.`;
    $('excelResults').classList.remove('hidden');

  } catch(e) {
    alert('Error al leer Excel: ' + e.message);
  }
});

function renderExcel() {
  $('excelCount').textContent = excelData.length + ' filas';

  $('excelRows').innerHTML = excelData.slice(0,300).map(x => {
    const issues = [];
    if (!/^\d{11}$/.test(x.ruc)) issues.push('RUC');
    if (!x.tipo) issues.push('Tipo');
    if (!x.serie) issues.push('Serie');
    if (!x.numero) issues.push('Número');
    if (!dateValid(x.fecha)) issues.push('Fecha');
    if (!/^\d+(\.\d{1,2})?$/.test(x.total.replace(',','.'))) issues.push('Total');

    return `<tr>
      <td>${escapeHtml(x.ruc)}</td>
      <td>${escapeHtml(x.tipo)}</td>
      <td>${escapeHtml(x.serie)}</td>
      <td>${escapeHtml(x.numero)}</td>
      <td>${escapeHtml(x.fecha)}</td>
      <td>${escapeHtml(x.total)}</td>
      <td class="${issues.length?'bad':'ok'}">
        ${issues.length?'Revisar: '+issues.join(', '):'Básico OK'}
      </td>
    </tr>`;
  }).join('');
}

$('exportTxt').addEventListener('click', async () => {
  if (!excelData.length) return;

  const zip = new JSZip();
  const kind = $('excelKind').value;
  const month = $('month').value;
  const year = $('year').value;

  let valid = excelData.filter(x =>
    /^\d{11}$/.test(x.ruc) &&
    x.tipo && x.serie && x.numero &&
    dateValid(x.fecha) &&
    /^\d+(\.\d{1,2})?$/.test(x.total.replace(',','.'))
  );

  if (!valid.length) {
    alert('No hay filas que superen las validaciones básicas.');
    return;
  }

  for (let i=0; i<valid.length; i+=100) {
    const n = Math.floor(i/100)+1;
    const txt = valid.slice(i,i+100).map(x =>
      [x.ruc,x.tipo,x.serie,x.numero,x.fecha,x.total.replace(',','.')].join('|')
    ).join('\r\n');

    zip.file(`${kind} ${month} ${n}.txt`,txt);
  }

  const blob = await zip.generateAsync({type:'blob'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `TXT ${kind} ${month} ${year}.zip`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href),1000);
});

$('exportXlsx').addEventListener('click', () => {
  const rows = [
    ['RUC','TIPO','SERIE','NUMERO','FECHA','IMPORTE TOTAL'],
    ...excelData.map(x => [x.ruc,x.tipo,x.serie,x.numero,x.fecha,x.total])
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),'Registros');
  XLSX.writeFile(wb,`REGISTROS ${$('excelKind').value} ${$('month').value} ${$('year').value}.xlsx`);
});

$('clearExcel').addEventListener('click', () => {
  $('excelFile').value = '';
  excelData = [];
  $('excelResults').classList.add('hidden');
  $('excelInfo').textContent = 'Carga la plantilla o un archivo con los mismos campos.';
});
