const express = require('express');
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// ─── CONFIGURA AQUÍ TU BASE DE DATOS ───────────────────────────────────────
const pool = new Pool({
  host:     'localhost',
  port:     5432,
  database: 'control_activos_occ',
  user:     'postgres',
  password: 'Insta2025',
});
// ───────────────────────────────────────────────────────────────────────────

// ─── SUBIDA DE IMÁGENES (multer) ───────────────────────────────────────────
const UPLOADS_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOADS_DIR)) fs.mkdirSync(UPLOADS_DIR, { recursive: true });

// Servir las imágenes guardadas en /uploads
app.use('/uploads', express.static(UPLOADS_DIR));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    const unico = Date.now() + '-' + Math.round(Math.random() * 1e9);
    cb(null, `activo-${unico}${ext}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 5 * 1024 * 1024 }, // máx 5 MB
  fileFilter: (req, file, cb) => {
    const permitidos = ['.jpg', '.jpeg', '.png', '.gif', '.webp'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (permitidos.includes(ext)) cb(null, true);
    else cb(new Error('Solo se permiten imágenes (jpg, png, gif, webp)'));
  }
});

// Borra el archivo físico de una imagen si existe
function borrarImagen(rutaImagen) {
  if (!rutaImagen) return;
  const archivo = path.join(UPLOADS_DIR, path.basename(rutaImagen));
  fs.unlink(archivo, () => {}); // ignorar error si no existe
}
// ───────────────────────────────────────────────────────────────────────────

// ── Helpers HTML ──────────────────────────────────────────────────────────
function layout(title, body) {
  return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title} – Control Activos OCC</title>
<style>
  :root {
    --bg: #0f1117;
    --surface: #1a1d27;
    --card: #22263a;
    --border: #2e3354;
    --accent: #4f7cff;
    --accent2: #7c5cff;
    --success: #22c55e;
    --danger: #ef4444;
    --warn: #f59e0b;
    --text: #e8eaf6;
    --muted: #8890b5;
    --radius: 10px;
    --font: 'Segoe UI', system-ui, sans-serif;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: var(--bg); color: var(--text); font-family: var(--font); min-height: 100vh; }

  /* NAV */
  nav {
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 0 2rem;
    display: flex;
    align-items: center;
    gap: 2rem;
    height: 56px;
    position: sticky; top: 0; z-index: 100;
  }
  .nav-brand { font-weight: 700; font-size: 1.1rem; color: var(--accent); letter-spacing: .5px; white-space: nowrap; }
  .nav-links { display: flex; gap: .25rem; flex-wrap: wrap; }
  .nav-links a {
    color: var(--muted); text-decoration: none; padding: .4rem .8rem;
    border-radius: 6px; font-size: .85rem; transition: all .15s;
  }
  .nav-links a:hover, .nav-links a.active { background: var(--card); color: var(--text); }

  /* LAYOUT */
  .container { max-width: 2000px; margin: 0 auto; padding: 2rem 1.5rem; }
  h1 { font-size: 1.5rem; margin-bottom: 1.5rem; color: var(--text); }
  h2 { font-size: 1.1rem; margin-bottom: 1rem; color: var(--muted); font-weight: 500; }

  /* CARDS */
  .card { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius); padding: 1.5rem; margin-bottom: 1.5rem; }

  /* ALERTS */
  .alert { padding: .85rem 1rem; border-radius: var(--radius); margin-bottom: 1rem; font-size: .9rem; }
  .alert-success { background: #14532d44; border: 1px solid var(--success); color: var(--success); }
  .alert-error   { background: #7f1d1d44; border: 1px solid var(--danger);  color: var(--danger); }

  /* FORMS */
  .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; }
  .grid-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 1rem; }
  .field { display: flex; flex-direction: column; gap: .35rem; }
  .field label { font-size: .8rem; color: var(--muted); font-weight: 500; letter-spacing: .3px; }
  .field input, .field select, .field textarea {
    background: var(--surface); border: 1px solid var(--border); border-radius: 7px;
    color: var(--text); padding: .55rem .8rem; font-size: .9rem; font-family: var(--font);
    transition: border-color .15s;
  }
  .field input:focus, .field select:focus, .field textarea:focus {
    outline: none; border-color: var(--accent);
  }
  .field textarea { resize: vertical; min-height: 80px; }
  select option { background: var(--surface); }
  .field-full { grid-column: 1 / -1; }
  .field input[type=file] { padding: .4rem; }

  /* IMÁGENES */
  .thumb { width: 46px; height: 46px; object-fit: cover; border-radius: 6px; border: 1px solid var(--border); display: block; }
  .thumb-placeholder { width: 46px; height: 46px; border-radius: 6px; border: 1px dashed var(--border);
    display: flex; align-items: center; justify-content: center; color: var(--muted); font-size: 1.1rem; }
  .img-preview { max-width: 220px; max-height: 220px; object-fit: contain; border-radius: 8px;
    border: 1px solid var(--border); margin-top: .5rem; background: var(--surface); }

  /* BUTTONS */
  .btn { display: inline-block; padding: .6rem 1.4rem; border-radius: 7px; border: none;
    cursor: pointer; font-size: .9rem; font-weight: 600; text-decoration: none; transition: opacity .15s; }
  .btn-primary { background: var(--accent); color: #fff; }
  .btn-danger  { background: var(--danger); color: #fff; }
  .btn-ghost   { background: var(--border); color: var(--text); }
  .btn-success { background: var(--success); color: #fff; }
  .btn:hover   { opacity: .85; }
  .btn-sm      { padding: .3rem .8rem; font-size: .8rem; }

  /* TABLES */
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: .87rem; }
  th { background: var(--surface); color: var(--muted); font-weight: 600; font-size: .78rem;
       letter-spacing: .4px; text-transform: uppercase; padding: .7rem 1rem; text-align: left; border-bottom: 1px solid var(--border); }
  td { padding: .7rem 1rem; border-bottom: 1px solid var(--border); vertical-align: middle; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #ffffff06; }

  /* BADGES */
  .badge { display: inline-block; padding: .2rem .6rem; border-radius: 20px; font-size: .75rem; font-weight: 600; }
  .badge-blue   { background: #1e3a8a55; color: #60a5fa; }
  .badge-green  { background: #14532d55; color: #4ade80; }
  .badge-yellow { background: #78350f55; color: #fbbf24; }
  .badge-red    { background: #7f1d1d55; color: #f87171; }
  .badge-gray   { background: #374151;   color: #9ca3af; }

  /* DASHBOARD */
  .stat-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 1rem; margin-bottom: 2rem; }
  .stat-card { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius);
               padding: 1.25rem; display: flex; flex-direction: column; gap: .4rem; }
  .stat-label { font-size: .78rem; color: var(--muted); font-weight: 500; letter-spacing: .3px; text-transform: uppercase; }
  .stat-value { font-size: 2rem; font-weight: 700; color: var(--text); line-height: 1; }
  .stat-sub   { font-size: .8rem; color: var(--muted); }

  /* QUICK LINKS */
  .quick-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); gap: .75rem; margin-bottom: 2rem; }
  .quick-link { background: var(--card); border: 1px solid var(--border); border-radius: var(--radius);
                padding: 1rem; text-align: center; text-decoration: none; color: var(--text);
                transition: border-color .15s, transform .15s; display: flex; flex-direction: column; gap: .5rem; align-items: center; }
  .quick-link:hover { border-color: var(--accent); transform: translateY(-2px); }
  .quick-link .icon { font-size: 1.6rem; }
  .quick-link span  { font-size: .82rem; font-weight: 500; color: var(--muted); }

  /* MISC */
  .actions { display: flex; gap: .5rem; }
  .row-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: .5rem; }
  .section-divider { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }
  @media(max-width: 700px) { .grid-2, .grid-3 { grid-template-columns: 1fr; } }
</style>
</head>
<body>
<nav>
  <div class="nav-brand">⚙ OCC Activos</div>
  <div class="nav-links">
    <a href="/">Inicio</a>
    <a href="/activos">Activos</a>
    <a href="/activos/nuevo">+ Activo</a>
    <a href="/asignaciones">Asignaciones</a>
    <a href="/empleados">Empleados</a>
    <a href="/catalogos">Catálogos</a>
  </div>
</nav>
<div class="container">
${body}
</div>
</body>
</html>`;
}

function badge(text, color) {
  const map = { 'Activo':'green','En uso':'blue','Dado de baja':'red','En reparación':'yellow','Devuelto':'gray' };
  const c = color || map[text] || 'gray';
  return `<span class="badge badge-${c}">${text}</span>`;
}

// Miniatura de imagen (o placeholder si no hay)
function thumbHtml(ruta, nombre) {
  return ruta
    ? `<a href="${ruta}" target="_blank" title="Ver imagen completa"><img class="thumb" src="${ruta}" alt="${nombre||''}"></a>`
    : `<div class="thumb-placeholder">📷</div>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// INICIO – Dashboard
// ─────────────────────────────────────────────────────────────────────────────
app.get('/', async (req, res) => {
  const stats = await pool.query(`
    SELECT
      (SELECT COUNT(*) FROM activo) AS total_activos,
      (SELECT COUNT(*) FROM activo WHERE fecha_baja IS NULL) AS activos_activos,
      (SELECT COUNT(*) FROM asignacion WHERE estado='Activo') AS asignaciones_activas,
      (SELECT COUNT(*) FROM empleado) AS total_empleados
  `);
  const s = stats.rows[0];

  const recientes = await pool.query(`
    SELECT a.codigo_inventario, a.nombre, a.ruta_imagen, c.nombre AS categoria, e.nombre_estado, a.fecha_alta
    FROM activo a
    LEFT JOIN categoria c ON a.id_categoria=c.id_categoria
    LEFT JOIN estado_activo e ON a.id_estado=e.id_estado
    ORDER BY a.id_activo DESC LIMIT 6
  `);

  const rows = recientes.rows.map(r => `
    <tr>
      <td>${thumbHtml(r.ruta_imagen, r.nombre)}</td>
      <td><code style="color:var(--accent)">${r.codigo_inventario}</code></td>
      <td>${r.nombre}</td>
      <td>${r.categoria || '—'}</td>
      <td>${badge(r.nombre_estado || '—')}</td>
      <td style="color:var(--muted)">${r.fecha_alta ? new Date(r.fecha_alta).toLocaleDateString('es-HN') : '—'}</td>
    </tr>`).join('');

  res.send(layout('Inicio', `
    <h1>Panel de Control</h1>
    <div class="stat-grid">
      <div class="stat-card">
        <div class="stat-label">Total Activos</div>
        <div class="stat-value">${s.total_activos}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Activos Operativos</div>
        <div class="stat-value" style="color:var(--success)">${s.activos_activos}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Asignaciones Activas</div>
        <div class="stat-value" style="color:var(--accent)">${s.asignaciones_activas}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Empleados</div>
        <div class="stat-value">${s.total_empleados}</div>
      </div>
    </div>

    <h2>Accesos Rápidos</h2>
    <div class="quick-grid">
      <a class="quick-link" href="/activos/nuevo"><div class="icon">📦</div><span>Nuevo Activo</span></a>
      <a class="quick-link" href="/asignaciones/nueva"><div class="icon">🔗</div><span>Nueva Asignación</span></a>
      <a class="quick-link" href="/empleados/nuevo"><div class="icon">👤</div><span>Nuevo Empleado</span></a>
      <a class="quick-link" href="/catalogos"><div class="icon">📋</div><span>Catálogos</span></a>
      <a class="quick-link" href="/activos"><div class="icon">🔍</div><span>Ver Activos</span></a>
      <a class="quick-link" href="/asignaciones"><div class="icon">📊</div><span>Asignaciones</span></a>
    </div>

    <div class="card">
      <div class="row-top"><h2 style="margin:0">Activos Recientes</h2><a class="btn btn-ghost btn-sm" href="/activos">Ver todos →</a></div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Img</th><th>Código</th><th>Nombre</th><th>Categoría</th><th>Estado</th><th>Fecha Alta</th></tr></thead>
          <tbody>${rows || '<tr><td colspan="6" style="text-align:center;color:var(--muted)">Sin activos registrados</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `));
});

// ─────────────────────────────────────────────────────────────────────────────
// ACTIVOS – Lista + Exportar Excel
// ─────────────────────────────────────────────────────────────────────────────
app.get('/activos', async (req, res) => {
  const { q, cat, est } = req.query;
  let where = ['1=1']; let params = [];
  if (q)   { params.push(`%${q}%`); where.push(`(a.nombre ILIKE $${params.length} OR a.codigo_inventario ILIKE $${params.length})`); }
  if (cat) { params.push(cat); where.push(`a.id_categoria=$${params.length}`); }
  if (est) { params.push(est); where.push(`a.id_estado=$${params.length}`); }

  const [activos, cats, estados] = await Promise.all([
    pool.query(`
      SELECT a.*, c.nombre AS cat, u.edificio, u.oficina, m.nombre AS marca,
             d.nombre AS depto, e.nombre_estado
      FROM activo a
      LEFT JOIN categoria c ON a.id_categoria=c.id_categoria
      LEFT JOIN ubicacion u ON a.id_ubicacion=u.id_ubicacion
      LEFT JOIN marca m ON a.id_marca=m.id_marca
      LEFT JOIN departamento d ON a.id_departamento=d.id_departamento
      LEFT JOIN estado_activo e ON a.id_estado=e.id_estado
      WHERE ${where.join(' AND ')}
      ORDER BY a.id_activo DESC`, params),
    pool.query('SELECT * FROM categoria ORDER BY nombre'),
    pool.query('SELECT * FROM estado_activo ORDER BY nombre_estado'),
  ]);

  const catOpts = cats.rows.map(c => `<option value="${c.id_categoria}" ${cat==c.id_categoria?'selected':''}>${c.nombre}</option>`).join('');
  const estOpts = estados.rows.map(e => `<option value="${e.id_estado}" ${est==e.id_estado?'selected':''}>${e.nombre_estado}</option>`).join('');

  // Build query string to pass to export
  const qs = new URLSearchParams({ ...(q&&{q}), ...(cat&&{cat}), ...(est&&{est}) }).toString();

  const rows = activos.rows.map(a => `
    <tr>
      <td>${thumbHtml(a.ruta_imagen, a.nombre)}</td>
      <td><code style="color:var(--accent);font-size:.8rem">${a.codigo_inventario}</code></td>
      <td><strong>${a.nombre}</strong>${a.modelo?`<br><span style="color:var(--muted);font-size:.100rem">${a.modelo}</span>`:''}</td>
      <td>${a.cat||'—'}</td>
      <td>${a.marca||'—'}</td>
      <td>${a.depto||'—'}</td>
     <td>${a.edificio ? `${a.edificio}${a.oficina?'<br>'+a.oficina:''}` : '—'}</td>
      <td style="max-width:220px;color:var(--muted);font-size:.82rem;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${a.descripcion||''}">${a.descripcion||'—'}</td>
      <td>${badge(a.nombre_estado||'—')}</td>
      <td>
        <div class="actions">
          <a class="btn btn-ghost btn-sm" href="/activos/${a.id_activo}/editar">✏ Editar</a>
          <form method="POST" action="/activos/${a.id_activo}/eliminar" style="display:inline" onsubmit="return confirm('¿Eliminar este activo?')">
            <button class="btn btn-danger btn-sm" type="submit">🗑</button>
          </form>
        </div>
      </td>
    </tr>`).join('');

  res.send(layout('Activos', `
    <div class="row-top">
      <h1 style="margin:0">Activos (${activos.rows.length})</h1>
      <div class="actions">
        <a class="btn btn-success" href="/activos/exportar${qs?'?'+qs:''}">📥 Exportar Excel</a>
        <a class="btn btn-primary" href="/activos/nuevo">+ Nuevo Activo</a>
      </div>
    </div>
    <div class="card">
      <form method="GET" style="display:flex;gap:.75rem;flex-wrap:wrap;align-items:flex-end">
        <div class="field"><label>Buscar</label><input name="q" value="${q||''}" placeholder="Nombre o código…" style="min-width:200px"></div>
        <div class="field"><label>Categoría</label><select name="cat"><option value="">Todas</option>${catOpts}</select></div>
        <div class="field"><label>Estado</label><select name="est"><option value="">Todos</option>${estOpts}</select></div>
        <button class="btn btn-primary" type="submit">Filtrar</button>
        <a class="btn btn-ghost" href="/activos">Limpiar</a>
      </form>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table>
        <thead><tr><th>Img</th><th>Código</th><th>Nombre / Modelo</th><th>Categoría</th><th>Marca</th><th>Departamento</th><th>Ubicación</th><th>Descripción</th><th>Estado</th><th>Acciones</th></tr></thead>
          <tbody>${rows||'<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:2rem">No se encontraron activos</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `));
});

// ─────────────────────────────────────────────────────────────────────────────
// ACTIVOS – Exportar a Excel (con empleado asignado)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/activos/exportar', async (req, res) => {
  const { q, cat, est } = req.query;
  let where = ['1=1']; let params = [];
  if (q)   { params.push(`%${q}%`); where.push(`(a.nombre ILIKE $${params.length} OR a.codigo_inventario ILIKE $${params.length})`); }
  if (cat) { params.push(cat); where.push(`a.id_categoria=$${params.length}`); }
  if (est) { params.push(est); where.push(`a.id_estado=$${params.length}`); }

  const result = await pool.query(`
    SELECT
      a.codigo_inventario,
      a.nombre,
      a.modelo,
      a.numero_serie,
      c.nombre        AS categoria,
      m.nombre        AS marca,
      d.nombre        AS departamento,
      u.edificio,
      u.oficina,
      ea.nombre_estado AS estado,
      a.fecha_alta,
      a.fecha_baja,
      a.descripcion,
      a.observaciones,
      a.ruta_imagen,
      -- Empleado con asignación activa más reciente
      emp.nombre      AS empleado_asignado,
      emp.puesto      AS puesto_empleado,
      emp.email       AS email_empleado,
      asig.fecha_asignacion,
      asig.estado     AS estado_asignacion
    FROM activo a
    LEFT JOIN categoria c    ON a.id_categoria   = c.id_categoria
    LEFT JOIN ubicacion u    ON a.id_ubicacion   = u.id_ubicacion
    LEFT JOIN marca m        ON a.id_marca       = m.id_marca
    LEFT JOIN departamento d ON a.id_departamento = d.id_departamento
    LEFT JOIN estado_activo ea ON a.id_estado    = ea.id_estado
    -- última asignación activa
    LEFT JOIN LATERAL (
      SELECT id_empleado, fecha_asignacion, estado
      FROM asignacion
      WHERE id_activo = a.id_activo AND estado = 'Activo'
      ORDER BY fecha_asignacion DESC
      LIMIT 1
    ) asig ON TRUE
    LEFT JOIN empleado emp ON asig.id_empleado = emp.id_empleado
    WHERE ${where.join(' AND ')}
    ORDER BY a.id_activo DESC
  `, params);

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Control Activos OCC';
  workbook.created = new Date();

  const ws = workbook.addWorksheet('Activos', {
    pageSetup: { paperSize: 9, orientation: 'landscape', fitToPage: true }
  });

  // Header style
  const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A8A' } };
  const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  const border = {
    top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
    left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
    bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
    right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
  };

  ws.columns = [
    { header: 'Código Inventario', key: 'codigo_inventario', width: 18 },
    { header: 'Nombre',            key: 'nombre',            width: 28 },
    { header: 'Modelo',            key: 'modelo',            width: 20 },
    { header: 'Número Serie',      key: 'numero_serie',      width: 18 },
    { header: 'Categoría',         key: 'categoria',         width: 16 },
    { header: 'Marca',             key: 'marca',             width: 14 },
    { header: 'Departamento',      key: 'departamento',      width: 18 },
    { header: 'Edificio',          key: 'edificio',          width: 14 },
    { header: 'Oficina',           key: 'oficina',           width: 14 },
    { header: 'Estado',            key: 'estado',            width: 14 },
    { header: 'Fecha Alta',        key: 'fecha_alta',        width: 14 },
    { header: 'Fecha Baja',        key: 'fecha_baja',        width: 14 },
    { header: 'Empleado Asignado', key: 'empleado_asignado', width: 24 },
    { header: 'Puesto',            key: 'puesto_empleado',   width: 18 },
    { header: 'Email Empleado',    key: 'email_empleado',    width: 24 },
    { header: 'F. Asignación',     key: 'fecha_asignacion',  width: 14 },
    { header: 'Estado Asignación', key: 'estado_asignacion', width: 16 },
    { header: 'Imagen',            key: 'ruta_imagen',       width: 24 },
    { header: 'Descripción',       key: 'descripcion',       width: 30 },
    { header: 'Observaciones',     key: 'observaciones',     width: 30 },
  ];

  // Style header row
  ws.getRow(1).eachCell(cell => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.border = border;
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  });
  ws.getRow(1).height = 32;

  // Add data rows
  result.rows.forEach((row, i) => {
    const dataRow = ws.addRow({
      codigo_inventario: row.codigo_inventario,
      nombre:            row.nombre,
      modelo:            row.modelo || '',
      numero_serie:      row.numero_serie || '',
      categoria:         row.categoria || '',
      marca:             row.marca || '',
      departamento:      row.departamento || '',
      edificio:          row.edificio || '',
      oficina:           row.oficina || '',
      estado:            row.estado || '',
      fecha_alta:        row.fecha_alta ? new Date(row.fecha_alta).toLocaleDateString('es-HN') : '',
      fecha_baja:        row.fecha_baja ? new Date(row.fecha_baja).toLocaleDateString('es-HN') : '',
      empleado_asignado: row.empleado_asignado || 'Sin asignar',
      puesto_empleado:   row.puesto_empleado || '',
      email_empleado:    row.email_empleado || '',
      fecha_asignacion:  row.fecha_asignacion ? new Date(row.fecha_asignacion).toLocaleDateString('es-HN') : '',
      estado_asignacion: row.estado_asignacion || '',
      ruta_imagen:       row.ruta_imagen || '',
      descripcion:       row.descripcion || '',
      observaciones:     row.observaciones || '',
    });

    const rowBg = i % 2 === 0 ? 'FFFAFAFA' : 'FFF0F4FF';
    dataRow.eachCell(cell => {
      cell.border = border;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBg } };
      cell.font = { size: 9, color: { argb: 'FF111827' } };
      cell.alignment = { vertical: 'middle', wrapText: true };
    });

    // Highlight "Sin asignar" in orange
    if (!row.empleado_asignado) {
      const empCell = dataRow.getCell('empleado_asignado');
      empCell.font = { size: 9, color: { argb: 'FFD97706' }, italic: true };
    }

    dataRow.height = 20;
  });

  // Freeze header
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  // Auto-filter
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: ws.columns.length } };

  // Summary sheet
  const wsSummary = workbook.addWorksheet('Resumen');
  wsSummary.columns = [
    { header: 'Métrica', key: 'metrica', width: 30 },
    { header: 'Valor',   key: 'valor',   width: 20 },
  ];
  wsSummary.getRow(1).eachCell(cell => {
    cell.fill = headerFill; cell.font = headerFont; cell.border = border;
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });

  const totalActivos   = result.rows.length;
  const conAsignacion  = result.rows.filter(r => r.empleado_asignado).length;
  const sinAsignacion  = totalActivos - conAsignacion;
  const generado       = new Date().toLocaleString('es-HN');

  [[`Total de activos exportados`, totalActivos],
   [`Con empleado asignado`,       conAsignacion],
   [`Sin asignación activa`,       sinAsignacion],
   [`Generado el`,                 generado],
  ].forEach(([metrica, valor], i) => {
    const r = wsSummary.addRow({ metrica, valor });
    r.eachCell(cell => {
      cell.border = border;
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: i%2===0?'FFFAFAFA':'FFF0F4FF' } };
      cell.font = { size: 10, color: { argb: 'FF111827' } };
    });
  });

  // Send file
  const fecha = new Date().toISOString().slice(0,10);
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="activos_occ_${fecha}.xlsx"`);
  await workbook.xlsx.write(res);
  res.end();
});

// ─────────────────────────────────────────────────────────────────────────────
// ACTIVOS – Nuevo / Editar (formulario compartido, ahora con imagen)
// ─────────────────────────────────────────────────────────────────────────────
async function formActivo(activo, msg) {
  const [cats, ubics, marcas, deptos, estados] = await Promise.all([
    pool.query('SELECT * FROM categoria ORDER BY nombre'),
    pool.query('SELECT * FROM ubicacion ORDER BY edificio'),
    pool.query('SELECT * FROM marca ORDER BY nombre'),
    pool.query('SELECT * FROM departamento ORDER BY nombre'),
    pool.query('SELECT * FROM estado_activo ORDER BY nombre_estado'),
  ]);

  const sel = (list, idField, valField, cur) =>
    list.map(r => `<option value="${r[idField]}" ${cur==r[idField]?'selected':''}>${r[valField]}</option>`).join('');

  const a = activo || {};
  const alert = msg ? `<div class="alert ${msg.ok?'alert-success':'alert-error'}">${msg.text}</div>` : '';
  const action = a.id_activo ? `/activos/${a.id_activo}/editar` : '/activos/nuevo';
  const title  = a.id_activo ? `Editar: ${a.nombre}` : 'Nuevo Activo';

  // fecha_alta / fecha_baja pueden venir como Date (de la BD) o string (del form)
  const fmtDate = (v, def='') => {
    if (!v) return def;
    if (v instanceof Date) return v.toISOString().slice(0,10);
    return String(v).slice(0,10);
  };

  return layout(title, `
    <div class="row-top">
      <h1 style="margin:0">${title}</h1>
      <a class="btn btn-ghost" href="/activos">← Volver</a>
    </div>
    ${alert}
    <form method="POST" action="${action}" enctype="multipart/form-data">
      <div class="card">
        <h2>Información Principal</h2>
        <div class="grid-3">
          <div class="field"><label>Código de Inventario *</label><input name="codigo_inventario" value="${a.codigo_inventario||''}" required placeholder="OCC-2024-001"></div>
          <div class="field field-full" style="grid-column:span 2"><label>Nombre del Activo *</label><input name="nombre" value="${a.nombre||''}" required placeholder="Laptop HP EliteBook"></div>
          <div class="field"><label>Modelo</label><input name="modelo" value="${a.modelo||''}" placeholder="EliteBook 840 G9"></div>
          <div class="field"><label>Número de Serie</label><input name="numero_serie" value="${a.numero_serie||''}" placeholder="CNU12345678"></div>
          <div class="field"><label>Fecha de Alta</label><input type="date" name="fecha_alta" value="${fmtDate(a.fecha_alta, new Date().toISOString().slice(0,10))}" required></div>
          <div class="field"><label>Fecha de Baja</label><input type="date" name="fecha_baja" value="${fmtDate(a.fecha_baja)}"></div>
        </div>
        <hr class="section-divider">
        <h2>Imagen del Activo</h2>
        <div class="grid-2">
          <div class="field">
            <label>Foto (jpg, png, gif, webp – máx 5 MB)</label>
            <input type="file" name="imagen" accept="image/*" onchange="previewImg(this)">
            <img id="img-preview" class="img-preview" src="${a.ruta_imagen||''}" alt="" style="${a.ruta_imagen?'':'display:none'}">
          </div>
          ${a.ruta_imagen ? `
          <div class="field">
            <label>Imagen actual</label>
            <label style="display:flex;align-items:center;gap:.5rem;color:var(--danger);font-size:.85rem;cursor:pointer">
              <input type="checkbox" name="quitar_imagen" value="1" style="width:auto"> Quitar la imagen actual
            </label>
          </div>` : ''}
        </div>
        <script>
          function previewImg(input) {
            const img = document.getElementById('img-preview');
            if (input.files && input.files[0]) {
              img.src = URL.createObjectURL(input.files[0]);
              img.style.display = 'block';
            }
          }
        </script>
        <hr class="section-divider">
        <h2>Clasificación</h2>
        <div class="grid-3">
          <div class="field"><label>Categoría</label><select name="id_categoria"><option value="">— Sin categoría —</option>${sel(cats.rows,'id_categoria','nombre',a.id_categoria)}</select></div>
          <div class="field"><label>Marca</label><select name="id_marca"><option value="">— Sin marca —</option>${sel(marcas.rows,'id_marca','nombre',a.id_marca)}</select></div>
          <div class="field"><label>Departamento</label><select name="id_departamento"><option value="">— Sin departamento —</option>${sel(deptos.rows,'id_departamento','nombre',a.id_departamento)}</select></div>
          <div class="field"><label>Ubicación</label><select name="id_ubicacion"><option value="">— Sin ubicación —</option>${ubics.rows.map(u=>`<option value="${u.id_ubicacion}" ${a.id_ubicacion==u.id_ubicacion?'selected':''}>${u.edificio||''} – ${u.oficina||u.piso||''}</option>`).join('')}</select></div>
          <div class="field"><label>Estado</label><select name="id_estado"><option value="">— Sin estado —</option>${sel(estados.rows,'id_estado','nombre_estado',a.id_estado)}</select></div>
        </div>
        <hr class="section-divider">
        <h2>Notas</h2>
        <div class="grid-2">
          <div class="field field-full"><label>Descripción</label><textarea name="descripcion">${a.descripcion||''}</textarea></div>
          <div class="field field-full"><label>Observaciones</label><textarea name="observaciones">${a.observaciones||''}</textarea></div>
        </div>
      </div>
      <div style="display:flex;gap:.75rem">
        <button class="btn btn-primary" type="submit">${a.id_activo?'💾 Guardar Cambios':'➕ Registrar Activo'}</button>
        <a class="btn btn-ghost" href="/activos">Cancelar</a>
      </div>
    </form>
  `);
}

app.get('/activos/nuevo', async (req, res) => res.send(await formActivo()));

app.post('/activos/nuevo', upload.single('imagen'), async (req, res) => {
  const b = req.body;
  const rutaImagen = req.file ? '/uploads/' + req.file.filename : null;
  try {
    await pool.query(`
      INSERT INTO activo (id_categoria,id_ubicacion,id_marca,id_departamento,id_estado,
        codigo_inventario,nombre,descripcion,modelo,numero_serie,fecha_alta,fecha_baja,observaciones,ruta_imagen)
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`,
      [b.id_categoria||null, b.id_ubicacion||null, b.id_marca||null, b.id_departamento||null,
       b.id_estado||null, b.codigo_inventario, b.nombre, b.descripcion||null,
       b.modelo||null, b.numero_serie||null, b.fecha_alta, b.fecha_baja||null,
       b.observaciones||null, rutaImagen]);
    res.redirect('/activos?ok=1');
  } catch(e) {
    // Si falló el INSERT pero se subió el archivo, borrarlo para no dejar huérfanos
    borrarImagen(rutaImagen);
    res.send(await formActivo(b, { ok:false, text: 'Error: ' + e.message }));
  }
});

app.get('/activos/:id/editar', async (req, res) => {
  const r = await pool.query('SELECT * FROM activo WHERE id_activo=$1', [req.params.id]);
  if (!r.rows.length) return res.redirect('/activos');
  res.send(await formActivo(r.rows[0]));
});

app.post('/activos/:id/editar', upload.single('imagen'), async (req, res) => {
  const b = req.body; const id = req.params.id;
  try {
    // Obtener la imagen actual en BD
    const actual = await pool.query('SELECT ruta_imagen FROM activo WHERE id_activo=$1', [id]);
    const rutaActual = actual.rows.length ? actual.rows[0].ruta_imagen : null;

    // Determinar la nueva ruta:
    //  - si subieron archivo nuevo → usarlo (y borrar el anterior)
    //  - si marcaron "quitar imagen" → null (y borrar el anterior)
    //  - si no → conservar la actual
    let rutaImagen = rutaActual;
    if (req.file) {
      rutaImagen = '/uploads/' + req.file.filename;
      borrarImagen(rutaActual);
    } else if (b.quitar_imagen) {
      rutaImagen = null;
      borrarImagen(rutaActual);
    }

    await pool.query(`
      UPDATE activo SET id_categoria=$1,id_ubicacion=$2,id_marca=$3,id_departamento=$4,id_estado=$5,
        codigo_inventario=$6,nombre=$7,descripcion=$8,modelo=$9,numero_serie=$10,
        fecha_alta=$11,fecha_baja=$12,observaciones=$13,ruta_imagen=$14 WHERE id_activo=$15`,
      [b.id_categoria||null, b.id_ubicacion||null, b.id_marca||null, b.id_departamento||null,
       b.id_estado||null, b.codigo_inventario, b.nombre, b.descripcion||null,
       b.modelo||null, b.numero_serie||null, b.fecha_alta, b.fecha_baja||null,
       b.observaciones||null, rutaImagen, id]);
    res.redirect('/activos');
  } catch(e) {
    const a = { ...b, id_activo: id };
    res.send(await formActivo(a, { ok:false, text: 'Error: ' + e.message }));
  }
});

app.post('/activos/:id/eliminar', async (req, res) => {
  try {
    // Borrar también el archivo de imagen asociado
    const r = await pool.query('SELECT ruta_imagen FROM activo WHERE id_activo=$1', [req.params.id]);
    await pool.query('DELETE FROM activo WHERE id_activo=$1', [req.params.id]);
    if (r.rows.length) borrarImagen(r.rows[0].ruta_imagen);
  }
  catch(e) { /* ignorar si tiene referencias */ }
  res.redirect('/activos');
});

// Manejo de errores de multer (archivo muy grande, tipo no permitido)
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError || (err && err.message && err.message.includes('imágenes'))) {
    return res.send(layout('Error', `
      <div class="alert alert-error">Error al subir la imagen: ${err.message}</div>
      <a class="btn btn-ghost" href="javascript:history.back()">← Volver</a>
    `));
  }
  next(err);
});

// ─────────────────────────────────────────────────────────────────────────────
// EMPLEADOS
// ─────────────────────────────────────────────────────────────────────────────
app.get('/empleados', async (req, res) => {
  const emp = await pool.query('SELECT * FROM empleado ORDER BY nombre');
  const rows = emp.rows.map(e => `
    <tr>
      <td>${e.id_empleado}</td>
      <td><strong>${e.nombre}</strong></td>
      <td>${e.puesto||'—'}</td>
      <td>${e.email||'—'}</td>
      <td>
        <div class="actions">
          <a class="btn btn-ghost btn-sm" href="/empleados/${e.id_empleado}/editar">✏ Editar</a>
          <form method="POST" action="/empleados/${e.id_empleado}/eliminar" style="display:inline" onsubmit="return confirm('¿Eliminar empleado?')">
            <button class="btn btn-danger btn-sm">🗑</button>
          </form>
        </div>
      </td>
    </tr>`).join('');

  res.send(layout('Empleados', `
    <div class="row-top">
      <h1 style="margin:0">Empleados (${emp.rows.length})</h1>
      <a class="btn btn-primary" href="/empleados/nuevo">+ Nuevo Empleado</a>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table>
          <thead><tr><th>#</th><th>Nombre</th><th>Puesto</th><th>Email</th><th>Acciones</th></tr></thead>
          <tbody>${rows||'<tr><td colspan="5" style="text-align:center;color:var(--muted)">Sin empleados</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `));
});

function formEmpleado(e, msg) {
  const a = e || {};
  const alert = msg ? `<div class="alert ${msg.ok?'alert-success':'alert-error'}">${msg.text}</div>` : '';
  const action = a.id_empleado ? `/empleados/${a.id_empleado}/editar` : '/empleados/nuevo';
  const title  = a.id_empleado ? `Editar: ${a.nombre}` : 'Nuevo Empleado';
  return layout(title, `
    <div class="row-top"><h1 style="margin:0">${title}</h1><a class="btn btn-ghost" href="/empleados">← Volver</a></div>
    ${alert}
    <form method="POST" action="${action}">
      <div class="card">
        <div class="grid-2">
          <div class="field field-full"><label>Nombre Completo *</label><input name="nombre" value="${a.nombre||''}" required></div>
          <div class="field"><label>Puesto</label><input name="puesto" value="${a.puesto||''}"></div>
          <div class="field"><label>Email</label><input type="email" name="email" value="${a.email||''}"></div>
        </div>
      </div>
      <div style="display:flex;gap:.75rem">
        <button class="btn btn-primary" type="submit">${a.id_empleado?'💾 Guardar':'➕ Registrar'}</button>
        <a class="btn btn-ghost" href="/empleados">Cancelar</a>
      </div>
    </form>
  `);
}

app.get('/empleados/nuevo', (req, res) => res.send(formEmpleado()));
app.post('/empleados/nuevo', async (req, res) => {
  const b = req.body;
  try {
    await pool.query('INSERT INTO empleado (nombre,puesto,email) VALUES ($1,$2,$3)', [b.nombre, b.puesto||null, b.email||null]);
    res.redirect('/empleados');
  } catch(e) { res.send(formEmpleado(b, { ok:false, text: e.message })); }
});
app.get('/empleados/:id/editar', async (req, res) => {
  const r = await pool.query('SELECT * FROM empleado WHERE id_empleado=$1', [req.params.id]);
  res.send(formEmpleado(r.rows[0]));
});
app.post('/empleados/:id/editar', async (req, res) => {
  const b = req.body;
  try {
    await pool.query('UPDATE empleado SET nombre=$1,puesto=$2,email=$3 WHERE id_empleado=$4', [b.nombre, b.puesto||null, b.email||null, req.params.id]);
    res.redirect('/empleados');
  } catch(e) { res.send(formEmpleado({...b,id_empleado:req.params.id},{ok:false,text:e.message})); }
});
app.post('/empleados/:id/eliminar', async (req, res) => {
  try { await pool.query('DELETE FROM empleado WHERE id_empleado=$1', [req.params.id]); } catch(e) {}
  res.redirect('/empleados');
});

// ─────────────────────────────────────────────────────────────────────────────
// ASIGNACIONES – Lista
// ─────────────────────────────────────────────────────────────────────────────
app.get('/asignaciones', async (req, res) => {
  const { emp, est_asig } = req.query;
  let where = ['1=1']; let params = [];
  if (emp)     { params.push(emp);     where.push(`as2.id_empleado=$${params.length}`); }
  if (est_asig){ params.push(est_asig);where.push(`as2.estado=$${params.length}`); }

  const [rows, empleados] = await Promise.all([
    pool.query(`
      SELECT as2.*, a.nombre AS activo, a.codigo_inventario, e.nombre AS empleado
      FROM asignacion as2
      JOIN activo a ON as2.id_activo=a.id_activo
      JOIN empleado e ON as2.id_empleado=e.id_empleado
      WHERE ${where.join(' AND ')}
      ORDER BY as2.id_asignacion DESC
    `, params),
    pool.query('SELECT * FROM empleado ORDER BY nombre'),
  ]);

  const empOpts = empleados.rows.map(e => `<option value="${e.id_empleado}" ${emp==e.id_empleado?'selected':''}>${e.nombre}</option>`).join('');
  const estadoOpts = ['Activo','Devuelto'].map(s => `<option value="${s}" ${est_asig===s?'selected':''}>${s}</option>`).join('');

  const trs = rows.rows.map(r => `
    <tr>
      <td><code style="color:var(--accent);font-size:.8rem">${r.codigo_inventario}</code></td>
      <td>${r.activo}</td>
      <td>${r.empleado}</td>
      <td>${r.fecha_asignacion ? new Date(r.fecha_asignacion).toLocaleDateString('es-HN') : '—'}</td>
      <td>${r.fecha_devolucion ? new Date(r.fecha_devolucion).toLocaleDateString('es-HN') : '—'}</td>
      <td>${badge(r.estado)}</td>
      <td>
        <div class="actions">
          <a class="btn btn-ghost btn-sm" href="/asignaciones/${r.id_asignacion}/editar">✏ Editar</a>
          ${r.estado==='Activo' ? `
          <form method="POST" action="/asignaciones/${r.id_asignacion}/devolver" style="display:inline">
            <button class="btn btn-success btn-sm" type="submit">✅ Devolver</button>
          </form>` : ''}
          <form method="POST" action="/asignaciones/${r.id_asignacion}/eliminar" style="display:inline" onsubmit="return confirm('¿Eliminar esta asignación?')">
            <button class="btn btn-danger btn-sm" type="submit">🗑</button>
          </form>
        </div>
      </td>
    </tr>`).join('');

  res.send(layout('Asignaciones', `
    <div class="row-top">
      <h1 style="margin:0">Asignaciones (${rows.rows.length})</h1>
      <a class="btn btn-primary" href="/asignaciones/nueva">+ Nueva Asignación</a>
    </div>
    <div class="card">
      <form method="GET" style="display:flex;gap:.75rem;flex-wrap:wrap;align-items:flex-end">
        <div class="field"><label>Empleado</label><select name="emp"><option value="">Todos</option>${empOpts}</select></div>
        <div class="field"><label>Estado</label><select name="est_asig"><option value="">Todos</option>${estadoOpts}</select></div>
        <button class="btn btn-primary" type="submit">Filtrar</button>
        <a class="btn btn-ghost" href="/asignaciones">Limpiar</a>
      </form>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table>
          <thead><tr><th>Código</th><th>Activo</th><th>Empleado</th><th>Asignado</th><th>Devuelto</th><th>Estado</th><th>Acciones</th></tr></thead>
          <tbody>${trs||'<tr><td colspan="7" style="text-align:center;color:var(--muted)">Sin asignaciones</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `));
});

// ─────────────────────────────────────────────────────────────────────────────
// ASIGNACIONES – Nueva
// ─────────────────────────────────────────────────────────────────────────────
app.get('/asignaciones/nueva', async (req, res) => {
  const [activos, empleados] = await Promise.all([
    pool.query("SELECT a.id_activo, a.codigo_inventario, a.nombre FROM activo a WHERE a.fecha_baja IS NULL ORDER BY a.nombre"),
    pool.query('SELECT * FROM empleado ORDER BY nombre'),
  ]);
  const aOpts = activos.rows.map(a => `<option value="${a.id_activo}">[${a.codigo_inventario}] ${a.nombre}</option>`).join('');
  const eOpts = empleados.rows.map(e => `<option value="${e.id_empleado}">${e.nombre}</option>`).join('');

  res.send(layout('Nueva Asignación', `
    <div class="row-top"><h1 style="margin:0">Nueva Asignación</h1><a class="btn btn-ghost" href="/asignaciones">← Volver</a></div>
    <form method="POST" action="/asignaciones/nueva">
      <div class="card">
        <div class="grid-2">
          <div class="field"><label>Activo *</label><select name="id_activo" required><option value="">— Seleccionar —</option>${aOpts}</select></div>
          <div class="field"><label>Empleado *</label><select name="id_empleado" required><option value="">— Seleccionar —</option>${eOpts}</select></div>
          <div class="field"><label>Fecha Asignación</label><input type="date" name="fecha_asignacion" value="${new Date().toISOString().slice(0,10)}"></div>
          <div class="field"><label>Estado</label>
            <select name="estado">
              <option value="Activo">Activo</option>
              <option value="Devuelto">Devuelto</option>
            </select>
          </div>
          <div class="field field-full"><label>Observaciones</label><textarea name="observaciones"></textarea></div>
        </div>
      </div>
      <div style="display:flex;gap:.75rem">
        <button class="btn btn-primary" type="submit">🔗 Registrar Asignación</button>
        <a class="btn btn-ghost" href="/asignaciones">Cancelar</a>
      </div>
    </form>
  `));
});

app.post('/asignaciones/nueva', async (req, res) => {
  const b = req.body;
  try {
    await pool.query(
      'INSERT INTO asignacion (id_activo,id_empleado,fecha_asignacion,estado,observaciones) VALUES ($1,$2,$3,$4,$5)',
      [b.id_activo, b.id_empleado, b.fecha_asignacion, b.estado||'Activo', b.observaciones||null]);
    res.redirect('/asignaciones');
  } catch(e) { res.send('<p style="color:red;padding:2rem">Error: '+e.message+'</p>'); }
});

// ─────────────────────────────────────────────────────────────────────────────
// ASIGNACIONES – Editar
// ─────────────────────────────────────────────────────────────────────────────
app.get('/asignaciones/:id/editar', async (req, res) => {
  const [asig, activos, empleados] = await Promise.all([
    pool.query('SELECT * FROM asignacion WHERE id_asignacion=$1', [req.params.id]),
    pool.query("SELECT id_activo, codigo_inventario, nombre FROM activo WHERE fecha_baja IS NULL ORDER BY nombre"),
    pool.query('SELECT * FROM empleado ORDER BY nombre'),
  ]);

  if (!asig.rows.length) return res.redirect('/asignaciones');
  const a = asig.rows[0];

  const aOpts = activos.rows.map(r => `<option value="${r.id_activo}" ${a.id_activo==r.id_activo?'selected':''}>[${r.codigo_inventario}] ${r.nombre}</option>`).join('');
  const eOpts = empleados.rows.map(r => `<option value="${r.id_empleado}" ${a.id_empleado==r.id_empleado?'selected':''}>${r.nombre}</option>`).join('');
  const estadoOpts = ['Activo','Devuelto'].map(s => `<option value="${s}" ${a.estado===s?'selected':''}>${s}</option>`).join('');

  res.send(layout('Editar Asignación', `
    <div class="row-top"><h1 style="margin:0">✏ Editar Asignación #${a.id_asignacion}</h1><a class="btn btn-ghost" href="/asignaciones">← Volver</a></div>
    <form method="POST" action="/asignaciones/${a.id_asignacion}/editar">
      <div class="card">
        <div class="grid-2">
          <div class="field"><label>Activo *</label><select name="id_activo" required><option value="">— Seleccionar —</option>${aOpts}</select></div>
          <div class="field"><label>Empleado *</label><select name="id_empleado" required><option value="">— Seleccionar —</option>${eOpts}</select></div>
          <div class="field"><label>Fecha Asignación</label><input type="date" name="fecha_asignacion" value="${a.fecha_asignacion?a.fecha_asignacion.toISOString().slice(0,10):''}"></div>
          <div class="field"><label>Fecha Devolución</label><input type="date" name="fecha_devolucion" value="${a.fecha_devolucion?a.fecha_devolucion.toISOString().slice(0,10):''}"></div>
          <div class="field"><label>Estado</label><select name="estado">${estadoOpts}</select></div>
          <div class="field field-full"><label>Observaciones</label><textarea name="observaciones">${a.observaciones||''}</textarea></div>
        </div>
      </div>
      <div style="display:flex;gap:.75rem">
        <button class="btn btn-primary" type="submit">💾 Guardar Cambios</button>
        <a class="btn btn-ghost" href="/asignaciones">Cancelar</a>
      </div>
    </form>
  `));
});

app.post('/asignaciones/:id/editar', async (req, res) => {
  const b = req.body;
  try {
    await pool.query(`
      UPDATE asignacion
      SET id_activo=$1, id_empleado=$2, fecha_asignacion=$3, fecha_devolucion=$4,
          estado=$5, observaciones=$6
      WHERE id_asignacion=$7`,
      [b.id_activo, b.id_empleado, b.fecha_asignacion||null, b.fecha_devolucion||null,
       b.estado, b.observaciones||null, req.params.id]);
    res.redirect('/asignaciones');
  } catch(e) {
    res.send('<p style="color:red;padding:2rem">Error: '+e.message+'</p>');
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// ASIGNACIONES – Devolver y Eliminar
// ─────────────────────────────────────────────────────────────────────────────
app.post('/asignaciones/:id/devolver', async (req, res) => {
  await pool.query(
    "UPDATE asignacion SET estado='Devuelto', fecha_devolucion=CURRENT_DATE WHERE id_asignacion=$1",
    [req.params.id]);
  res.redirect('/asignaciones');
});

app.post('/asignaciones/:id/eliminar', async (req, res) => {
  try { await pool.query('DELETE FROM asignacion WHERE id_asignacion=$1', [req.params.id]); } catch(e) {}
  res.redirect('/asignaciones');
});

// ─────────────────────────────────────────────────────────────────────────────
// CATÁLOGOS (categoría, ubicación, marca, departamento, estado)
// ─────────────────────────────────────────────────────────────────────────────
app.get('/catalogos', async (req, res) => {
  const [cats, ubics, marcas, deptos, estados] = await Promise.all([
    pool.query('SELECT * FROM categoria ORDER BY nombre'),
    pool.query('SELECT * FROM ubicacion ORDER BY edificio'),
    pool.query('SELECT * FROM marca ORDER BY nombre'),
    pool.query('SELECT * FROM departamento ORDER BY nombre'),
    pool.query('SELECT * FROM estado_activo ORDER BY nombre_estado'),
  ]);

  function tableSimple(rows, fields, editBase, label) {
    const ths = fields.map(f => `<th>${f.label}</th>`).join('') + '<th>Acciones</th>';
    const trs = rows.map(r => `<tr>${fields.map(f=>`<td>${r[f.key]||'—'}</td>`).join('')}
      <td><div class="actions">
        <a class="btn btn-ghost btn-sm" href="${editBase}/${r[fields[0].id]}/editar">✏</a>
        <form method="POST" action="${editBase}/${r[fields[0].id]}/eliminar" style="display:inline" onsubmit="return confirm('¿Eliminar?')">
          <button class="btn btn-danger btn-sm">🗑</button>
        </form>
      </div></td></tr>`).join('');
    return `<table><thead><tr>${ths}</tr></thead><tbody>${trs||'<tr><td colspan="'+(fields.length+1)+'" style="color:var(--muted);text-align:center">Sin registros</td></tr>'}</tbody></table>`;
  }

  res.send(layout('Catálogos', `
    <h1>Catálogos</h1>

    <!-- CATEGORÍAS -->
    <div class="card">
      <div class="row-top"><h2 style="margin:0">📁 Categorías</h2><a class="btn btn-primary btn-sm" href="/catalogos/categorias/nueva">+ Agregar</a></div>
      <div class="table-wrap">${tableSimple(cats.rows,[{key:'id_categoria',id:'id_categoria',label:'#'},{key:'nombre',id:'id_categoria',label:'Nombre'},{key:'descripcion',id:'id_categoria',label:'Descripción'}],'/catalogos/categorias','cat')}</div>
    </div>

    <!-- MARCAS -->
    <div class="card">
      <div class="row-top"><h2 style="margin:0">🏷 Marcas</h2><a class="btn btn-primary btn-sm" href="/catalogos/marcas/nueva">+ Agregar</a></div>
      <div class="table-wrap">${tableSimple(marcas.rows,[{key:'id_marca',id:'id_marca',label:'#'},{key:'nombre',id:'id_marca',label:'Nombre'}],'/catalogos/marcas','marca')}</div>
    </div>

    <!-- DEPARTAMENTOS -->
    <div class="card">
      <div class="row-top"><h2 style="margin:0">🏢 Departamentos</h2><a class="btn btn-primary btn-sm" href="/catalogos/departamentos/nuevo">+ Agregar</a></div>
      <div class="table-wrap">${tableSimple(deptos.rows,[{key:'id_departamento',id:'id_departamento',label:'#'},{key:'nombre',id:'id_departamento',label:'Nombre'},{key:'descripcion',id:'id_departamento',label:'Descripción'}],'/catalogos/departamentos','depto')}</div>
    </div>

    <!-- UBICACIONES -->
    <div class="card">
      <div class="row-top"><h2 style="margin:0">📍 Ubicaciones</h2><a class="btn btn-primary btn-sm" href="/catalogos/ubicaciones/nueva">+ Agregar</a></div>
      <div class="table-wrap">${tableSimple(ubics.rows,[{key:'id_ubicacion',id:'id_ubicacion',label:'#'},{key:'edificio',id:'id_ubicacion',label:'Edificio'},{key:'piso',id:'id_ubicacion',label:'Piso'},{key:'oficina',id:'id_ubicacion',label:'Oficina'}],'/catalogos/ubicaciones','ubic')}</div>
    </div>

    <!-- ESTADOS -->
    <div class="card">
      <div class="row-top"><h2 style="margin:0">🔵 Estados</h2><a class="btn btn-primary btn-sm" href="/catalogos/estados/nuevo">+ Agregar</a></div>
      <div class="table-wrap">${tableSimple(estados.rows,[{key:'id_estado',id:'id_estado',label:'#'},{key:'nombre_estado',id:'id_estado',label:'Nombre'},{key:'descripcion',id:'id_estado',label:'Descripción'}],'/catalogos/estados','est')}</div>
    </div>
  `));
});

// ── Catálogo helpers ──────────────────────────────────────────────────────────
function crudCatalogo({ tabla, idCol, campos, ruta, titulo, emoji }) {
  function formPage(vals, msg) {
    const a = vals || {};
    const isEdit = !!a[idCol];
    const alert  = msg ? `<div class="alert ${msg.ok?'alert-success':'alert-error'}">${msg.text}</div>` : '';
    const fields  = campos.map(c => `
      <div class="field ${c.full?'field-full':''}">
        <label>${c.label}${c.req?' *':''}</label>
        ${c.textarea
          ? `<textarea name="${c.name}">${a[c.name]||''}</textarea>`
          : `<input name="${c.name}" value="${a[c.name]||''}" ${c.req?'required':''}>`}
      </div>`).join('');
    const action = isEdit ? `${ruta}/${a[idCol]}/editar` : `${ruta}/nueva`;
    return layout((isEdit ? 'Editar' : 'Nuevo') + ' ' + titulo, `
      <div class="row-top">
        <h1 style="margin:0">${emoji} ${isEdit?'Editar':'Nuevo'} ${titulo}</h1>
        <a class="btn btn-ghost" href="/catalogos">← Catálogos</a>
      </div>
      ${alert}
      <form method="POST" action="${action}">
        <div class="card"><div class="grid-2">${fields}</div></div>
        <button class="btn btn-primary" type="submit">${isEdit?'💾 Guardar':'➕ Agregar'}</button>
      </form>
    `);
  }

  app.get([`${ruta}/nueva`, `${ruta}/nuevo`], (req, res) => res.send(formPage()));

  app.post([`${ruta}/nueva`, `${ruta}/nuevo`], async (req, res) => {
    const b = req.body;
    const keys = campos.map(c => c.name);
    const vals = keys.map(k => b[k]||null);
    const phs  = keys.map((_,i)=>`$${i+1}`);
    try {
      await pool.query(`INSERT INTO ${tabla} (${keys.join(',')}) VALUES (${phs})`, vals);
      res.redirect('/catalogos');
    } catch(e) { res.send(formPage(b, {ok:false, text:e.message})); }
  });

  app.get(`${ruta}/:id/editar`, async (req, res) => {
    const r = await pool.query(`SELECT * FROM ${tabla} WHERE ${idCol}=$1`, [req.params.id]);
    res.send(formPage(r.rows[0]));
  });

  app.post(`${ruta}/:id/editar`, async (req, res) => {
    const b = req.body; const id = req.params.id;
    const keys = campos.map(c => c.name);
    const vals = [...keys.map(k => b[k]||null), id];
    const sets = keys.map((k,i)=>`${k}=$${i+1}`);
    try {
      await pool.query(`UPDATE ${tabla} SET ${sets.join(',')} WHERE ${idCol}=$${vals.length}`, vals);
      res.redirect('/catalogos');
    } catch(e) { res.send(formPage({...b,[idCol]:id},{ok:false,text:e.message})); }
  });

  app.post(`${ruta}/:id/eliminar`, async (req, res) => {
    try { await pool.query(`DELETE FROM ${tabla} WHERE ${idCol}=$1`, [req.params.id]); } catch(e){}
    res.redirect('/catalogos');
  });
}

crudCatalogo({ tabla:'categoria', idCol:'id_categoria', ruta:'/catalogos/categorias', titulo:'Categoría', emoji:'📁',
  campos:[{name:'nombre',label:'Nombre',req:true},{name:'descripcion',label:'Descripción',textarea:true,full:true}] });

crudCatalogo({ tabla:'marca', idCol:'id_marca', ruta:'/catalogos/marcas', titulo:'Marca', emoji:'🏷',
  campos:[{name:'nombre',label:'Nombre',req:true}] });

crudCatalogo({ tabla:'departamento', idCol:'id_departamento', ruta:'/catalogos/departamentos', titulo:'Departamento', emoji:'🏢',
  campos:[{name:'nombre',label:'Nombre',req:true},{name:'descripcion',label:'Descripción',textarea:true,full:true}] });

crudCatalogo({ tabla:'ubicacion', idCol:'id_ubicacion', ruta:'/catalogos/ubicaciones', titulo:'Ubicación', emoji:'📍',
  campos:[{name:'edificio',label:'Edificio',req:true},{name:'piso',label:'Piso'},{name:'oficina',label:'Oficina'}] });

crudCatalogo({ tabla:'estado_activo', idCol:'id_estado', ruta:'/catalogos/estados', titulo:'Estado', emoji:'🔵',
  campos:[{name:'nombre_estado',label:'Nombre',req:true},{name:'descripcion',label:'Descripción',textarea:true,full:true}] });

// ─────────────────────────────────────────────────────────────────────────────
const PORT = 4000;
app.listen(PORT, () => {
  console.log(`\n✅  Control Activos OCC corriendo en: http://localhost:${PORT}\n`);
});