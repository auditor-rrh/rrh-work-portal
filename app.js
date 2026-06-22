// ============================================================
// R R H & ASSOCIATES LLP — Main Application JS
// Google Sheets backend via Apps Script Web App
// ============================================================

// ── CONFIG ──────────────────────────────────────────────────
let SCRIPT_URL = localStorage.getItem('rrh_script_url') || '';
let currentUser = null;

// In-memory cache (fallback when no script URL set)
let DB = {
  users: [
    {id:'U-001',name:'CA Rajesh Kumar',username:'rajesh',password:'rrhlogin',role:'Partner',email:'rajesh@rrh.com',mobile:'9876543210',reportsTo:'—'},
    {id:'U-002',name:'CA Ramesh Hegde',username:'ramesh',password:'rrhlogin',role:'Partner',email:'ramesh@rrh.com',mobile:'9876543211',reportsTo:'—'},
    {id:'U-003',name:'CA Harish Patel',username:'harish',password:'rrhlogin',role:'Partner',email:'harish@rrh.com',mobile:'9876543212',reportsTo:'—'},
    {id:'U-004',name:'Suresh Nair',username:'suresh',password:'rrhlogin',role:'Asst. Manager',email:'suresh@rrh.com',mobile:'9123456780',reportsTo:'CA Rajesh Kumar'},
    {id:'U-005',name:'Priya Sharma',username:'priya',password:'rrhlogin',role:'Asst. Manager',email:'priya@rrh.com',mobile:'9123456781',reportsTo:'CA Ramesh Hegde'},
    {id:'U-006',name:'Ananya Reddy',username:'ananya',password:'rrhlogin',role:'Team Lead',email:'ananya@rrh.com',mobile:'9988776650',reportsTo:'Suresh Nair'},
    {id:'U-007',name:'Vikram Singh',username:'vikram',password:'rrhlogin',role:'Team Lead',email:'vikram@rrh.com',mobile:'9988776651',reportsTo:'Suresh Nair'},
    {id:'U-008',name:'Kavya Menon',username:'kavya',password:'rrhlogin',role:'Team Lead',email:'kavya@rrh.com',mobile:'9988776652',reportsTo:'Priya Sharma'},
    {id:'U-009',name:'Arjun Mehta',username:'arjun',password:'rrhlogin',role:'Article',email:'arjun@rrh.com',mobile:'9000000001',reportsTo:'Ananya Reddy'},
    {id:'U-010',name:'Sneha Pillai',username:'sneha',password:'rrhlogin',role:'Article',email:'sneha@rrh.com',mobile:'9000000002',reportsTo:'Ananya Reddy'},
    {id:'U-011',name:'Rahul Joshi',username:'rahul',password:'rrhlogin',role:'Article',email:'rahul@rrh.com',mobile:'9000000003',reportsTo:'Vikram Singh'},
    {id:'U-012',name:'Divya Nair',username:'divya',password:'rrhlogin',role:'Article',email:'divya@rrh.com',mobile:'9000000004',reportsTo:'Vikram Singh'},
    {id:'U-013',name:'Kiran Rao',username:'kiran',password:'rrhlogin',role:'Article',email:'kiran@rrh.com',mobile:'9000000005',reportsTo:'Kavya Menon'},
    {id:'U-014',name:'Pooja Iyer',username:'pooja',password:'rrhlogin',role:'Article',email:'pooja@rrh.com',mobile:'9000000006',reportsTo:'Kavya Menon'},
    {id:'U-015',name:'Nikhil Verma',username:'nikhil',password:'rrhlogin',role:'Article',email:'nikhil@rrh.com',mobile:'9000000007',reportsTo:'Ananya Reddy'},
    {id:'U-016',name:'Meera Das',username:'meera',password:'rrhlogin',role:'Article',email:'meera@rrh.com',mobile:'9000000008',reportsTo:'Vikram Singh'},
    {id:'U-017',name:'Siddharth Bose',username:'siddharth',password:'rrhlogin',role:'Article',email:'siddharth@rrh.com',mobile:'9000000009',reportsTo:'Kavya Menon'},
    {id:'U-018',name:'Tanvi Kulkarni',username:'tanvi',password:'rrhlogin',role:'Article',email:'tanvi@rrh.com',mobile:'9000000010',reportsTo:'Ananya Reddy'},
  ],
  tasks: [
    {id:'T-001',name:'GST Return Q2',group:'GST',client:'Mehta Traders',head:'CA Rajesh Kumar',assignedTo:'Arjun Mehta',reviewedBy:'Suresh Nair',priority:'High',status:'Review',dueDate:'2024-11-20',description:'GSTR-3B for Oct 2024',billId:'INV-001',subTaskOf:'',createdAt:'2024-11-01'},
    {id:'T-002',name:'Statutory Audit 2023-24',group:'Audit',client:'Sunrise Pvt Ltd',head:'CA Ramesh Hegde',assignedTo:'Sneha Pillai',reviewedBy:'Priya Sharma',priority:'High',status:'Done',dueDate:'2024-11-12',description:'Annual statutory audit',billId:'INV-002',subTaskOf:'',createdAt:'2024-10-15'},
    {id:'T-003',name:'Balance Sheet Finalization',group:'Accounts',client:'KL Industries',head:'CA Harish Patel',assignedTo:'Rahul Joshi',reviewedBy:'Ananya Reddy',priority:'Medium',status:'In Progress',dueDate:'2024-11-14',description:'Final accounts preparation',billId:'',subTaskOf:'',createdAt:'2024-11-02'},
    {id:'T-004',name:'TDS Return Q2',group:'Income Tax',client:'ABC Enterprises',head:'CA Rajesh Kumar',assignedTo:'Divya Nair',reviewedBy:'Vikram Singh',priority:'High',status:'Pending',dueDate:'2024-11-30',description:'Form 26Q filing',billId:'',subTaskOf:'',createdAt:'2024-11-05'},
    {id:'T-005',name:'ITR Filing FY2023-24',group:'Income Tax',client:'Mehta Traders',head:'CA Ramesh Hegde',assignedTo:'Kiran Rao',reviewedBy:'Priya Sharma',priority:'Medium',status:'Done',dueDate:'2024-11-08',description:'Income tax return filing',billId:'INV-003',subTaskOf:'',createdAt:'2024-10-20'},
    {id:'T-006',name:'ROC Annual Filing',group:'ROC',client:'Sunrise Pvt Ltd',head:'CA Harish Patel',assignedTo:'Pooja Iyer',reviewedBy:'Ananya Reddy',priority:'Low',status:'Pending',dueDate:'2024-12-31',description:'MGT-7 and AOC-4',billId:'',subTaskOf:'',createdAt:'2024-11-01'},
  ],
  clients: [
    {id:'C-001',name:'Mehta Traders',type:'Partnership',gstin:'27AABCM1234A1Z5',pan:'AABCM1234A',contact:'Ramesh Mehta',mobile:'9876543210',email:'mehta@traders.com',partner:'CA Rajesh Kumar',activeTasks:'4',outstanding:'45000',createdAt:'2023-04-01'},
    {id:'C-002',name:'Sunrise Pvt Ltd',type:'Company',gstin:'29AABCS5678B2Z3',pan:'AABCS5678B',contact:'Sunil Jain',mobile:'9123456780',email:'info@sunrise.com',partner:'CA Ramesh Hegde',activeTasks:'6',outstanding:'0',createdAt:'2023-04-01'},
    {id:'C-003',name:'KL Industries',type:'Company',gstin:'36AABCK9012C3Z1',pan:'AABCK9012C',contact:'Krishna Lal',mobile:'9988776655',email:'kl@industries.com',partner:'CA Harish Patel',activeTasks:'3',outstanding:'22500',createdAt:'2023-06-15'},
    {id:'C-004',name:'ABC Enterprises',type:'LLP',gstin:'27AABCA3456D4Z8',pan:'AABCA3456D',contact:'Ankit Bose',mobile:'9876501234',email:'abc@enterprises.com',partner:'CA Rajesh Kumar',activeTasks:'2',outstanding:'78000',createdAt:'2023-08-10'},
  ],
  invoices: [
    {id:'INV-001',client:'Mehta Traders',task:'GST Return Q2',date:'2024-11-15',dueDate:'2024-11-30',description:'GST Return filing Oct 2024',amount:'30000',gstRate:'18',gstAmount:'5400',total:'35400',status:'Unpaid',notes:'Payment due within 15 days',createdAt:'2024-11-15'},
    {id:'INV-002',client:'Sunrise Pvt Ltd',task:'Statutory Audit 2023-24',date:'2024-11-12',dueDate:'2024-11-27',description:'Statutory Audit 2023-24',amount:'75000',gstRate:'18',gstAmount:'13500',total:'88500',status:'Paid',notes:'',createdAt:'2024-11-12'},
    {id:'INV-003',client:'Mehta Traders',task:'ITR Filing FY2023-24',date:'2024-11-08',dueDate:'2024-11-23',description:'Income Tax Return FY 2023-24',amount:'15000',gstRate:'18',gstAmount:'2700',total:'17700',status:'Partial',notes:'Partial payment received ₹10,000',createdAt:'2024-11-08'},
  ],
  timesheets: [
    {id:'TS-001',date:'2024-11-18',staffName:'Arjun Mehta',client:'Mehta Traders',task:'GST Return Q2',hours:'3.5',description:'Prepared GSTR-3B data',createdAt:'2024-11-18'},
    {id:'TS-002',date:'2024-11-18',staffName:'Sneha Pillai',client:'Sunrise Pvt Ltd',task:'Statutory Audit 2023-24',hours:'6',description:'Vouching of purchase ledger',createdAt:'2024-11-18'},
    {id:'TS-003',date:'2024-11-17',staffName:'Rahul Joshi',client:'KL Industries',task:'Balance Sheet Finalization',hours:'4',description:'Finalization of P&L statement',createdAt:'2024-11-17'},
  ],
  documents: [
    {id:'D-001',name:'GSTR-3B_Oct2024_Mehta.pdf',category:'GST',client:'Mehta Traders',uploadedBy:'Arjun Mehta',date:'2024-11-15',url:'',createdAt:'2024-11-15'},
    {id:'D-002',name:'AuditReport_Sunrise_2324.pdf',category:'Audit',client:'Sunrise Pvt Ltd',uploadedBy:'Sneha Pillai',date:'2024-11-12',url:'',createdAt:'2024-11-12'},
  ]
};

// Invoice template settings
let invoiceTemplate = JSON.parse(localStorage.getItem('rrh_invoice_template') || 'null') || {
  firmName: 'R R H & ASSOCIATES LLP',
  tagline: 'Chartered Accountants',
  address: '123, CA Tower, Banjara Hills, Hyderabad – 500034, Telangana',
  gstin: '36AABCR1234F1Z9',
  pan: 'AABCR1234F',
  bank: 'HDFC Bank',
  account: '50100123456789',
  ifsc: 'HDFC0001234',
  logo: '',
  footer: 'Subject to Hyderabad jurisdiction. Payment due within 15 days of invoice date.',
};

// ── API LAYER ────────────────────────────────────────────────
async function api(params) {
  if (!SCRIPT_URL) return localFallback(params);
  try {
    const url = SCRIPT_URL + '?' + new URLSearchParams(params).toString();
    const res = await fetch(url);
    const data = await res.json();
    if (data.error) throw new Error(data.error);
    return data;
  } catch (e) {
    console.warn('API error, using local cache:', e.message);
    return localFallback(params);
  }
}

function localFallback(params) {
  const { action } = params;
  const sheetMap = {
    getTasks: 'tasks', getClients: 'clients', getInvoices: 'invoices',
    getTimesheets: 'timesheets', getUsers: 'users', getDocuments: 'documents',
  };
  if (sheetMap[action]) return { rows: DB[sheetMap[action]] };

  if (action === 'login') {
    const u = DB.users.find(u => u.username === params.username && u.password === params.password);
    return u ? { success: true, user: u } : { success: false, message: 'Invalid credentials' };
  }
  if (action === 'addTask')      { addLocalRow('tasks', params); return { success: true }; }
  if (action === 'addClient')    { addLocalRow('clients', params); return { success: true }; }
  if (action === 'addInvoice')   { addLocalRow('invoices', params); return { success: true }; }
  if (action === 'addTimesheet') { addLocalRow('timesheets', params); return { success: true }; }
  if (action === 'addUser')      { addLocalRow('users', params); return { success: true }; }
  if (action === 'addDocument')  { addLocalRow('documents', params); return { success: true }; }
  if (action === 'updateTask')   { updateLocalRow('tasks', params); return { success: true }; }
  if (action === 'updateInvoice'){ updateLocalRow('invoices', params); return { success: true }; }
  if (action === 'deleteTask')   { deleteLocalRow('tasks', params.id); return { success: true }; }
  return { success: false, message: 'Unknown action' };
}

function addLocalRow(sheet, params) {
  const prefixMap = {tasks:'T',clients:'C',invoices:'INV',timesheets:'TS',users:'U',documents:'D'};
  const prefix = prefixMap[sheet] || 'X';
  params.id = prefix + '-' + String(DB[sheet].length + 1).padStart(3, '0');
  params.createdAt = today();
  DB[sheet].unshift({ ...params });
}
function updateLocalRow(sheet, params) {
  const idx = DB[sheet].findIndex(r => r.id === params.id);
  if (idx >= 0) DB[sheet][idx] = { ...DB[sheet][idx], ...params };
}
function deleteLocalRow(sheet, id) {
  DB[sheet] = DB[sheet].filter(r => r.id !== id);
}

// ── HELPERS ──────────────────────────────────────────────────
function today() { return new Date().toISOString().split('T')[0]; }
function fmt(n) { return '₹' + Number(n || 0).toLocaleString('en-IN'); }
function initials(name) { return name.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase(); }
function avatarColor(role) {
  return { Partner:'#d4960a', 'Asst. Manager':'#2563a8', 'Team Lead':'#7c3aed', Article:'#16a34a' }[role] || '#6b7280';
}
function roleBadgeClass(role) {
  return { Partner:'badge-partner', 'Asst. Manager':'badge-am', 'Team Lead':'badge-tl', Article:'badge-article' }[role] || 'badge-article';
}
function statusBadge(s) {
  const m = { Pending:'badge-pending', 'In Progress':'badge-progress', Review:'badge-review', Done:'badge-done' };
  return `<span class="badge ${m[s]||'badge-pending'}">${s}</span>`;
}
function priorityBadge(p) {
  const m = { High:'badge-high', Medium:'badge-medium', Low:'badge-low' };
  const dot = { High:'#dc2626', Medium:'#ca8a04', Low:'#16a34a' };
  return `<span class="badge ${m[p]||'badge-medium'}"><span style="width:7px;height:7px;border-radius:50%;background:${dot[p]||'#6b7280'};display:inline-block;margin-right:4px"></span>${p}</span>`;
}
function payBadge(s) {
  const m = { Paid:'badge-paid', Unpaid:'badge-unpaid', Partial:'badge-partial' };
  return `<span class="badge ${m[s]||'badge-unpaid'}">${s}</span>`;
}
function isOverdue(dueDate) {
  if (!dueDate) return false;
  return new Date(dueDate) < new Date() && true;
}

function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast show ' + type;
  setTimeout(() => { t.className = 'toast'; }, 3000);
}

// ── AUTH ─────────────────────────────────────────────────────
async function doLogin() {
  const username = document.getElementById('loginUser').value.trim();
  const password = document.getElementById('loginPass').value.trim();
  const err = document.getElementById('loginError');
  const btn = document.getElementById('loginBtnText');
  if (!username || !password) { err.style.display = 'block'; err.textContent = 'Enter username and password.'; return; }
  btn.textContent = 'Signing in…';
  const res = await api({ action: 'login', username, password });
  btn.textContent = 'Sign In';
  if (res.success) {
    currentUser = res.user;
    err.style.display = 'none';
    startApp();
  } else {
    err.style.display = 'block';
    err.textContent = res.message || 'Invalid credentials.';
  }
}

function startApp() {
  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('appScreen').style.display = 'flex';
  const av = initials(currentUser.name);
  const col = avatarColor(currentUser.role);
  document.getElementById('sidebarAvatar').textContent = av;
  document.getElementById('sidebarAvatar').style.background = col;
  document.getElementById('topbarAvatar').textContent = av;
  document.getElementById('topbarAvatar').style.background = col;
  document.getElementById('sidebarName').textContent = currentUser.name;
  document.getElementById('sidebarRole').textContent = currentUser.role;
  // Set today in timesheet date
  const tsDate = document.getElementById('ts_date');
  if (tsDate) tsDate.value = today();
  loadAll();
}

function logout() {
  currentUser = null;
  document.getElementById('loginScreen').style.display = 'flex';
  document.getElementById('appScreen').style.display = 'none';
  document.getElementById('loginUser').value = '';
  document.getElementById('loginPass').value = '';
}

// ── LOAD ALL DATA ────────────────────────────────────────────
async function loadAll() {
  await Promise.all([loadTasks(), loadClients(), loadInvoices(), loadTimesheets(), loadUsers(), loadDocuments()]);
  updateDashboard();
}

async function loadTasks() {
  const res = await api({ action: 'getTasks' });
  DB.tasks = res.rows || DB.tasks;
  renderTasks();
  populateTaskSelects();
  document.getElementById('navTaskCount').textContent = DB.tasks.filter(t => t.status !== 'Done').length;
}
async function loadClients() {
  const res = await api({ action: 'getClients' });
  DB.clients = res.rows || DB.clients;
  renderClients();
  populateClientSelects();
}
async function loadInvoices() {
  const res = await api({ action: 'getInvoices' });
  DB.invoices = res.rows || DB.invoices;
  renderInvoices();
  renderPayments();
}
async function loadTimesheets() {
  const res = await api({ action: 'getTimesheets' });
  DB.timesheets = res.rows || DB.timesheets;
  renderTimesheets();
}
async function loadUsers() {
  const res = await api({ action: 'getUsers' });
  DB.users = res.rows || DB.users;
  renderTeam();
  populateUserSelects();
}
async function loadDocuments() {
  const res = await api({ action: 'getDocuments' });
  DB.documents = res.rows || DB.documents;
  renderDocuments();
}

// ── DASHBOARD ────────────────────────────────────────────────
function updateDashboard() {
  const overdue = DB.tasks.filter(t => t.status !== 'Done' && isOverdue(t.dueDate));
  const inReview = DB.tasks.filter(t => t.status === 'Review');
  const unpaidInvoices = DB.invoices.filter(i => i.status !== 'Paid');

  document.getElementById('ds-tasks').textContent = DB.tasks.length;
  document.getElementById('ds-clients').textContent = DB.clients.length;
  document.getElementById('ds-pending').textContent = unpaidInvoices.length;
  document.getElementById('ds-overdue').textContent = overdue.length;
  document.getElementById('ds-review').textContent = inReview.length;
  document.getElementById('ds-team').textContent = DB.users.length;

  // Overdue table
  const ot = document.querySelector('#dashOverdueTable tbody');
  ot.innerHTML = overdue.length ? overdue.slice(0, 5).map(t =>
    `<tr><td><b>${t.name}</b></td><td>${t.client}</td><td>${t.assignedTo}</td><td style="color:var(--danger);font-weight:700">${t.dueDate}</td></tr>`
  ).join('') : `<tr><td colspan="4" class="empty-row">✅ No overdue tasks</td></tr>`;

  // Review table
  const rt = document.querySelector('#dashReviewTable tbody');
  rt.innerHTML = inReview.length ? inReview.slice(0, 5).map(t =>
    `<tr><td><b>${t.name}</b></td><td>${t.assignedTo}</td><td>${t.reviewedBy}</td></tr>`
  ).join('') : `<tr><td colspan="3" class="empty-row">✅ Nothing in review</td></tr>`;
}

// ── TASKS ─────────────────────────────────────────────────────
function renderTasks() {
  const tbody = document.getElementById('taskTableBody');
  if (!tbody) return;
  let list = [...DB.tasks];
  const g = document.getElementById('filterGroup')?.value;
  const s = document.getElementById('filterStatus')?.value;
  const p = document.getElementById('filterPriority')?.value;
  if (g) list = list.filter(t => t.group === g);
  if (s) list = list.filter(t => t.status === s);
  if (p) list = list.filter(t => t.priority === p);

  tbody.innerHTML = list.length ? list.map(t => `
    <tr>
      <td><span class="text-muted">${t.id}</span></td>
      <td><b>${t.name}</b>${t.subTaskOf ? `<br><span class="text-muted" style="font-size:11px">↳ ${t.subTaskOf}</span>` : ''}</td>
      <td><span class="chip">${t.group}</span></td>
      <td>${t.client}</td>
      <td style="font-size:12px">${t.head}</td>
      <td>${t.assignedTo}</td>
      <td style="font-size:12px">${t.reviewedBy}</td>
      <td>${priorityBadge(t.priority)}</td>
      <td>${statusBadge(t.status)}</td>
      <td style="${isOverdue(t.dueDate) && t.status!=='Done' ? 'color:var(--danger);font-weight:700' : ''}">${t.dueDate || '—'}</td>
      <td>${t.billId ? `<span class="text-muted" style="font-size:12px">${t.billId}</span>` : '<span class="text-muted">—</span>'}</td>
      <td>
        <div class="action-btns">
          <button class="btn btn-secondary btn-sm" title="Edit" onclick="editTask('${t.id}')"><i class="ti ti-edit"></i></button>
          <button class="btn btn-secondary btn-sm" title="Generate Bill" onclick="prefillInvoice('${t.id}')"><i class="ti ti-receipt"></i></button>
          <button class="btn btn-danger btn-sm" title="Delete" onclick="deleteTask('${t.id}')"><i class="ti ti-trash"></i></button>
        </div>
      </td>
    </tr>`).join('') : `<tr><td colspan="12" class="empty-row">No tasks found. <a href="#" onclick="openModal('addTaskModal')">Add one?</a></td></tr>`;
}

async function saveTask() {
  const name = document.getElementById('t_name').value.trim();
  if (!name) { showToast('Task name is required', 'error'); return; }
  const params = {
    action: 'addTask',
    name,
    group: document.getElementById('t_group').value,
    client: document.getElementById('t_client').value,
    head: document.getElementById('t_head').value,
    assignedTo: document.getElementById('t_assign').value,
    reviewedBy: document.getElementById('t_review').value,
    priority: document.getElementById('t_priority').value,
    status: document.getElementById('t_status').value,
    dueDate: document.getElementById('t_due').value,
    description: document.getElementById('t_desc').value,
    subTaskOf: document.getElementById('t_parent').value,
    billId: '',
  };
  await api(params);
  closeModal('addTaskModal');
  document.getElementById('t_name').value = '';
  document.getElementById('t_desc').value = '';
  await loadTasks();
  updateDashboard();
  showToast('Task created!', 'success');
}

function editTask(id) {
  const t = DB.tasks.find(t => t.id === id);
  if (!t) return;
  document.getElementById('t_name').value = t.name;
  document.getElementById('t_group').value = t.group;
  document.getElementById('t_client').value = t.client;
  document.getElementById('t_head').value = t.head;
  document.getElementById('t_assign').value = t.assignedTo;
  document.getElementById('t_review').value = t.reviewedBy;
  document.getElementById('t_priority').value = t.priority;
  document.getElementById('t_status').value = t.status;
  document.getElementById('t_due').value = t.dueDate;
  document.getElementById('t_desc').value = t.description || '';
  openModal('addTaskModal');
}

async function deleteTask(id) {
  if (!confirm('Delete this task?')) return;
  await api({ action: 'deleteTask', id });
  await loadTasks();
  updateDashboard();
  showToast('Task deleted');
}

// ── CLIENTS ───────────────────────────────────────────────────
function renderClients() {
  const tbody = document.getElementById('clientTableBody');
  if (!tbody) return;
  let list = [...DB.clients];
  const f = document.getElementById('filterClientType')?.value;
  if (f) list = list.filter(c => c.type === f);
  tbody.innerHTML = list.length ? list.map(c => `
    <tr>
      <td><span class="text-muted">${c.id}</span></td>
      <td><b>${c.name}</b></td>
      <td><span class="badge badge-company" style="font-size:11px">${c.type}</span></td>
      <td style="font-size:12px;font-family:monospace">${c.gstin || '—'}</td>
      <td style="font-size:12px;font-family:monospace">${c.pan || '—'}</td>
      <td>${c.mobile || '—'}</td>
      <td style="font-size:12px">${c.partner}</td>
      <td style="text-align:center">${c.activeTasks || 0}</td>
      <td style="font-weight:700;color:${Number(c.outstanding)>0?'var(--danger)':'var(--success)'}">${fmt(c.outstanding)}</td>
      <td><div class="action-btns">
        <button class="btn btn-secondary btn-sm" onclick="editClient('${c.id}')"><i class="ti ti-edit"></i></button>
        <button class="btn btn-secondary btn-sm" onclick="filterByClient('${c.name}')"><i class="ti ti-checklist"></i></button>
      </div></td>
    </tr>`).join('') : `<tr><td colspan="10" class="empty-row">No clients. <a href="#" onclick="openModal('addClientModal')">Add one?</a></td></tr>`;
}

async function saveClient() {
  const name = document.getElementById('c_name').value.trim();
  if (!name) { showToast('Client name required', 'error'); return; }
  await api({
    action: 'addClient', name,
    type: document.getElementById('c_type').value,
    gstin: document.getElementById('c_gstin').value.toUpperCase(),
    pan: document.getElementById('c_pan').value.toUpperCase(),
    contact: document.getElementById('c_contact').value,
    mobile: document.getElementById('c_mobile').value,
    email: document.getElementById('c_email').value,
    partner: document.getElementById('c_partner').value,
    activeTasks: '0', outstanding: '0',
  });
  closeModal('addClientModal');
  ['c_name','c_gstin','c_pan','c_contact','c_mobile','c_email'].forEach(id => document.getElementById(id).value = '');
  await loadClients();
  showToast('Client added!', 'success');
}

function editClient(id) {
  const c = DB.clients.find(c => c.id === id);
  if (!c) return;
  document.getElementById('c_name').value = c.name;
  document.getElementById('c_type').value = c.type;
  document.getElementById('c_gstin').value = c.gstin || '';
  document.getElementById('c_pan').value = c.pan || '';
  document.getElementById('c_contact').value = c.contact || '';
  document.getElementById('c_mobile').value = c.mobile || '';
  document.getElementById('c_email').value = c.email || '';
  document.getElementById('c_partner').value = c.partner || '';
  openModal('addClientModal');
}

function filterByClient(name) {
  showPage('tasks', document.querySelector('.nav-item:nth-child(2)'));
  setTimeout(() => {
    document.getElementById('filterGroup').value = '';
    document.getElementById('filterStatus').value = '';
    // filter manually
    const tbody = document.getElementById('taskTableBody');
    const list = DB.tasks.filter(t => t.client === name);
    tbody.innerHTML = list.length ? list.map(t => `<tr><td>${t.id}</td><td><b>${t.name}</b></td><td><span class="chip">${t.group}</span></td><td>${t.client}</td><td>${t.head}</td><td>${t.assignedTo}</td><td>${t.reviewedBy}</td><td>${priorityBadge(t.priority)}</td><td>${statusBadge(t.status)}</td><td>${t.dueDate||'—'}</td><td>${t.billId||'—'}</td><td></td></tr>`).join('') : `<tr><td colspan="12" class="empty-row">No tasks for ${name}</td></tr>`;
  }, 100);
}

// ── INVOICES ──────────────────────────────────────────────────
function renderInvoices() {
  const tbody = document.getElementById('invoiceTableBody');
  if (!tbody) return;
  tbody.innerHTML = DB.invoices.length ? DB.invoices.map(inv => `
    <tr>
      <td><b>${inv.id}</b></td>
      <td>${inv.client}</td>
      <td style="font-size:12px">${inv.task || '—'}</td>
      <td>${inv.date}</td>
      <td>${fmt(inv.amount)}</td>
      <td>${fmt(inv.gstAmount)}</td>
      <td style="font-weight:700">${fmt(inv.total)}</td>
      <td>${payBadge(inv.status)}</td>
      <td><div class="action-btns">
        <button class="btn btn-secondary btn-sm" onclick="previewInvoice('${inv.id}')"><i class="ti ti-eye"></i></button>
        <button class="btn btn-secondary btn-sm" onclick="window.print()"><i class="ti ti-download"></i></button>
      </div></td>
    </tr>`).join('') : `<tr><td colspan="9" class="empty-row">No invoices yet.</td></tr>`;
}

function calcInvoice() {
  const amt = parseFloat(document.getElementById('inv_amount').value) || 0;
  const rate = parseFloat(document.getElementById('inv_gstrate').value) || 0;
  const gst = +(amt * rate / 100).toFixed(2);
  document.getElementById('inv_gst').value = gst ? fmt(gst).replace('₹', '') : '';
  document.getElementById('inv_total').textContent = fmt(amt + gst);
}

async function saveInvoice() {
  const client = document.getElementById('inv_client').value;
  const amount = document.getElementById('inv_amount').value;
  if (!client || !amount) { showToast('Client and amount are required', 'error'); return; }
  const gstRate = document.getElementById('inv_gstrate').value;
  const gstAmount = +(parseFloat(amount) * parseFloat(gstRate) / 100).toFixed(2);
  const total = +(parseFloat(amount) + gstAmount).toFixed(2);
  await api({
    action: 'addInvoice', client,
    task: document.getElementById('inv_task').value,
    date: document.getElementById('inv_date').value || today(),
    dueDate: document.getElementById('inv_due').value,
    description: document.getElementById('inv_desc').value,
    amount, gstRate, gstAmount: String(gstAmount), total: String(total),
    status: document.getElementById('inv_status').value,
    notes: document.getElementById('inv_notes').value,
  });
  await loadInvoices();
  showBillingTab('list');
  showToast('Invoice created!', 'success');
}

function previewInvoice(id) {
  const inv = DB.invoices.find(i => i.id === id);
  if (!inv) return;
  const half = +(parseFloat(inv.gstAmount) / 2).toFixed(2);
  const tmpl = invoiceTemplate;
  document.getElementById('inv_firm_name').textContent = tmpl.firmName;
  document.getElementById('inv_firm_sub').textContent = tmpl.tagline;
  document.getElementById('inv_firm_addr').textContent = tmpl.address;
  document.getElementById('inv_firm_gstin').textContent = 'GSTIN: ' + tmpl.gstin;
  document.getElementById('inv_firm_bank').textContent = `Bank: ${tmpl.bank} · A/C: ${tmpl.account} · IFSC: ${tmpl.ifsc}`;
  document.getElementById('prev_invno').textContent = inv.id;
  document.getElementById('prev_date').textContent = 'Date: ' + inv.date;
  document.getElementById('prev_duedate').textContent = 'Due: ' + (inv.dueDate || '—');
  const cli = DB.clients.find(c => c.name === inv.client);
  document.getElementById('prev_client').textContent = inv.client;
  document.getElementById('prev_gstin').textContent = cli ? 'GSTIN: ' + (cli.gstin || '—') : '';
  document.getElementById('prev_desc').textContent = inv.description;
  document.getElementById('prev_amt').textContent = fmt(inv.amount);
  document.getElementById('prev_cgst').textContent = fmt(half);
  document.getElementById('prev_sgst').textContent = fmt(half);
  document.getElementById('prev_total').textContent = fmt(inv.total);
  document.getElementById('prev_grand').textContent = 'Grand Total: ' + fmt(inv.total);
  showBillingTab('preview');
}

function prefillInvoice(taskId) {
  const t = DB.tasks.find(t => t.id === taskId);
  if (!t) return;
  showPage('billing', null);
  showBillingTab('create');
  setTimeout(() => {
    document.getElementById('inv_client').value = t.client;
    document.getElementById('inv_task').value = t.name;
    document.getElementById('inv_date').value = today();
    document.getElementById('inv_desc').value = t.name + ' – ' + t.client;
  }, 100);
}

// ── PAYMENTS ─────────────────────────────────────────────────
function renderPayments() {
  const tbody = document.getElementById('paymentTableBody');
  if (!tbody) return;
  let totalBilled = 0, totalPaid = 0;
  tbody.innerHTML = DB.invoices.map(inv => {
    const total = parseFloat(inv.total) || 0;
    const paid = inv.status === 'Paid' ? total : inv.status === 'Partial' ? total * 0.5 : 0;
    totalBilled += total; totalPaid += paid;
    const balance = total - paid;
    return `<tr>
      <td><b>${inv.id}</b></td>
      <td>${inv.client}</td>
      <td>${inv.date}</td>
      <td>${fmt(total)}</td>
      <td style="color:var(--success)">${fmt(paid)}</td>
      <td style="font-weight:700;color:${balance>0?'var(--danger)':'var(--success)'}">${fmt(balance)}</td>
      <td>${payBadge(inv.status)}</td>
      <td>${inv.status !== 'Paid' ?
        `<button class="btn btn-secondary btn-sm" onclick="markPaid('${inv.id}')"><i class="ti ti-check"></i> Mark Paid</button>` :
        '<span class="text-muted">—</span>'}</td>
    </tr>`;
  }).join('') || `<tr><td colspan="8" class="empty-row">No invoices yet.</td></tr>`;
  document.getElementById('pay_total').textContent = fmt(totalBilled);
  document.getElementById('pay_received').textContent = fmt(totalPaid);
  document.getElementById('pay_outstanding').textContent = fmt(totalBilled - totalPaid);
  document.getElementById('pay_overdue').textContent = fmt(
    DB.invoices.filter(i => i.status !== 'Paid' && isOverdue(i.dueDate)).reduce((a, i) => a + parseFloat(i.total||0), 0)
  );
}

async function markPaid(id) {
  await api({ action: 'updateInvoice', id, status: 'Paid' });
  await loadInvoices();
  showToast('Marked as Paid!', 'success');
}

// ── TEAM ──────────────────────────────────────────────────────
function renderTeam() {
  const grid = document.getElementById('teamGrid');
  if (!grid) return;
  let list = [...DB.users];
  const f = document.getElementById('filterRole')?.value;
  if (f) list = list.filter(u => u.role === f);
  grid.innerHTML = list.map(u => `
    <div class="team-card">
      <div class="team-card-top">
        <div class="avatar" style="background:${avatarColor(u.role)};width:42px;height:42px;font-size:14px">${initials(u.name)}</div>
        <div>
          <div style="font-weight:700;font-size:14px">${u.name}</div>
          <div style="margin-top:3px"><span class="badge ${roleBadgeClass(u.role)}">${u.role}</span></div>
          <div style="font-size:11px;color:var(--text-muted);margin-top:3px">Reports to: ${u.reportsTo || '—'}</div>
        </div>
      </div>
      <div class="team-card-actions">
        <a href="mailto:${u.email}" class="btn btn-secondary btn-sm"><i class="ti ti-mail"></i></a>
        <a href="tel:${u.mobile}" class="btn btn-secondary btn-sm"><i class="ti ti-phone"></i></a>
        <button class="btn btn-secondary btn-sm" style="margin-left:auto" onclick="filterByStaff('${u.name}')"><i class="ti ti-checklist"></i> Tasks</button>
      </div>
    </div>`).join('');
}

function filterByStaff(name) {
  showPage('tasks', null);
  setTimeout(() => {
    const list = DB.tasks.filter(t => t.assignedTo === name);
    document.getElementById('taskTableBody').innerHTML = list.length
      ? list.map(t => `<tr><td>${t.id}</td><td><b>${t.name}</b></td><td><span class="chip">${t.group}</span></td><td>${t.client}</td><td>${t.head}</td><td>${t.assignedTo}</td><td>${t.reviewedBy}</td><td>${priorityBadge(t.priority)}</td><td>${statusBadge(t.status)}</td><td>${t.dueDate||'—'}</td><td>—</td><td></td></tr>`).join('')
      : `<tr><td colspan="12" class="empty-row">No tasks for ${name}</td></tr>`;
  }, 100);
}

async function saveUser() {
  const name = document.getElementById('u_name').value.trim();
  const username = document.getElementById('u_username').value.trim();
  if (!name || !username) { showToast('Name and username required', 'error'); return; }
  await api({
    action: 'addUser', name, username,
    password: document.getElementById('u_password').value || 'rrhlogin',
    role: document.getElementById('u_role').value,
    email: document.getElementById('u_email').value,
    mobile: document.getElementById('u_mobile').value,
    reportsTo: document.getElementById('u_reports').value,
  });
  closeModal('addUserModal');
  await loadUsers();
  showToast('Team member added!', 'success');
}

// ── TIMESHEETS ────────────────────────────────────────────────
function renderTimesheets() {
  const tbody = document.getElementById('timesheetTableBody');
  if (!tbody) return;
  const list = [...DB.timesheets];
  tbody.innerHTML = list.length ? list.map(ts => `
    <tr>
      <td>${ts.date}</td>
      <td>${ts.staffName}</td>
      <td>${ts.client}</td>
      <td>${ts.task}</td>
      <td style="font-weight:700">${ts.hours}h</td>
      <td>${ts.description || '—'}</td>
      <td><button class="btn btn-danger btn-sm" onclick="deleteTS('${ts.id}')"><i class="ti ti-trash"></i></button></td>
    </tr>`).join('') : `<tr><td colspan="7" class="empty-row">No timesheets logged.</td></tr>`;
}

async function saveTimesheet() {
  const hours = document.getElementById('ts_hours').value;
  if (!hours) { showToast('Enter hours', 'error'); return; }
  await api({
    action: 'addTimesheet',
    date: document.getElementById('ts_date').value || today(),
    staffName: currentUser.name,
    client: document.getElementById('ts_client').value,
    task: document.getElementById('ts_task').value,
    hours,
    description: document.getElementById('ts_desc').value,
  });
  closeModal('logTimeModal');
  document.getElementById('ts_hours').value = '';
  document.getElementById('ts_desc').value = '';
  await loadTimesheets();
  showToast('Time logged!', 'success');
}

function deleteTS(id) {
  DB.timesheets = DB.timesheets.filter(t => t.id !== id);
  renderTimesheets();
}

// ── DOCUMENTS ─────────────────────────────────────────────────
function renderDocuments() {
  const tbody = document.getElementById('docTableBody');
  if (!tbody) return;
  const f = document.getElementById('filterDocCat')?.value;
  const list = f ? DB.documents.filter(d => d.category === f) : DB.documents;
  tbody.innerHTML = list.length ? list.map(d => `
    <tr>
      <td><i class="ti ti-file-text" style="color:var(--info);margin-right:6px"></i><b>${d.name}</b></td>
      <td><span class="chip">${d.category}</span></td>
      <td>${d.client}</td>
      <td>${d.uploadedBy}</td>
      <td>${d.date}</td>
      <td><div class="action-btns">
        ${d.url ? `<a href="${d.url}" target="_blank" class="btn btn-secondary btn-sm"><i class="ti ti-external-link"></i></a>` : ''}
        <button class="btn btn-danger btn-sm" onclick="deleteDoc('${d.id}')"><i class="ti ti-trash"></i></button>
      </div></td>
    </tr>`).join('') : `<tr><td colspan="6" class="empty-row">No documents.</td></tr>`;
}

async function saveDocument() {
  const name = document.getElementById('doc_name').value.trim();
  if (!name) { showToast('Document name required', 'error'); return; }
  await api({
    action: 'addDocument', name,
    category: document.getElementById('doc_cat').value,
    client: document.getElementById('doc_client').value,
    uploadedBy: currentUser.name,
    date: today(),
    url: document.getElementById('doc_url').value,
  });
  closeModal('addDocModal');
  document.getElementById('doc_name').value = '';
  document.getElementById('doc_url').value = '';
  await loadDocuments();
  showToast('Document saved!', 'success');
}

function deleteDoc(id) {
  DB.documents = DB.documents.filter(d => d.id !== id);
  renderDocuments();
}

// ── BILLING TABS ──────────────────────────────────────────────
function showBillingTab(tab) {
  ['list','create','preview','template'].forEach(t => {
    const el = document.getElementById('billing-' + t);
    if (el) el.style.display = t === tab ? 'block' : 'none';
  });
}

// ── INVOICE TEMPLATE ──────────────────────────────────────────
function saveTemplate() {
  invoiceTemplate = {
    firmName: document.getElementById('tmpl_name').value,
    tagline: document.getElementById('tmpl_tagline').value,
    address: document.getElementById('tmpl_addr').value,
    gstin: document.getElementById('tmpl_gstin').value,
    pan: document.getElementById('tmpl_pan').value,
    bank: document.getElementById('tmpl_bank').value,
    account: document.getElementById('tmpl_acc').value,
    ifsc: document.getElementById('tmpl_ifsc').value,
    logo: document.getElementById('tmpl_logo').value,
    footer: document.getElementById('tmpl_footer').value,
  };
  localStorage.setItem('rrh_invoice_template', JSON.stringify(invoiceTemplate));
  showToast('Template saved!', 'success');
}

// ── REPORTS ───────────────────────────────────────────────────
function generateReport(type) {
  const el = document.getElementById('reportOutput');
  el.style.display = 'block';
  if (type === 'task-summary') {
    const byStatus = {};
    DB.tasks.forEach(t => { byStatus[t.status] = (byStatus[t.status] || 0) + 1; });
    el.innerHTML = `<div class="card"><div class="card-header"><h3>Task Summary</h3></div><div class="card-body">
      ${Object.entries(byStatus).map(([s,c]) => `<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid var(--border)"><span>${statusBadge(s)}</span><b>${c} tasks</b></div>`).join('')}
    </div></div>`;
  } else if (type === 'billing') {
    const total = DB.invoices.reduce((a, i) => a + parseFloat(i.total||0), 0);
    const paid = DB.invoices.filter(i=>i.status==='Paid').reduce((a,i)=>a+parseFloat(i.total||0),0);
    el.innerHTML = `<div class="card"><div class="card-header"><h3>Billing Report</h3></div><div class="card-body">
      <div style="display:flex;gap:16px">
        <div class="stat-card" style="flex:1"><div class="stat-label">Total Billed</div><div class="stat-value">${fmt(total)}</div></div>
        <div class="stat-card" style="flex:1"><div class="stat-label">Collected</div><div class="stat-value success">${fmt(paid)}</div></div>
        <div class="stat-card" style="flex:1"><div class="stat-label">Outstanding</div><div class="stat-value danger">${fmt(total-paid)}</div></div>
      </div>
    </div></div>`;
  } else if (type === 'performance') {
    const byStaff = {};
    DB.tasks.forEach(t => { if (!byStaff[t.assignedTo]) byStaff[t.assignedTo] = {total:0,done:0}; byStaff[t.assignedTo].total++; if(t.status==='Done') byStaff[t.assignedTo].done++; });
    el.innerHTML = `<div class="card"><div class="card-header"><h3>Staff Performance</h3></div><div class="table-wrap"><table>
      <thead><tr><th>Staff</th><th>Total Tasks</th><th>Completed</th><th>Completion Rate</th></tr></thead>
      <tbody>${Object.entries(byStaff).map(([name, d]) => `<tr><td>${name}</td><td>${d.total}</td><td>${d.done}</td><td>${Math.round(d.done/d.total*100)}%</td></tr>`).join('')}</tbody>
    </table></div></div>`;
  } else if (type === 'outstanding') {
    const unpaid = DB.invoices.filter(i => i.status !== 'Paid');
    el.innerHTML = `<div class="card"><div class="card-header"><h3>Outstanding Fees</h3></div><div class="table-wrap"><table>
      <thead><tr><th>Invoice</th><th>Client</th><th>Amount</th><th>Status</th></tr></thead>
      <tbody>${unpaid.map(i=>`<tr><td>${i.id}</td><td>${i.client}</td><td style="font-weight:700;color:var(--danger)">${fmt(i.total)}</td><td>${payBadge(i.status)}</td></tr>`).join('')}</tbody>
    </table></div></div>`;
  } else {
    el.innerHTML = `<div class="card"><div class="card-body"><p style="color:var(--text-muted)">Report generation for "<b>${type}</b>" — coming soon. All data is available in Google Sheets.</p></div></div>`;
  }
}

// ── POPULATE SELECTS ─────────────────────────────────────────
function populateClientSelects() {
  const opts = DB.clients.map(c => `<option value="${c.name}">${c.name}</option>`).join('');
  ['t_client','inv_client','ts_client','doc_client'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<option value="">Select client…</option>' + opts;
  });
}
function populateUserSelects() {
  const partners = DB.users.filter(u => u.role === 'Partner').map(u => `<option>${u.name}</option>`).join('');
  const reviewers = DB.users.filter(u => ['Partner','Asst. Manager','Team Lead'].includes(u.role)).map(u => `<option>${u.name}</option>`).join('');
  const staff = DB.users.map(u => `<option>${u.name}</option>`).join('');
  const el_head = document.getElementById('t_head');
  const el_assign = document.getElementById('t_assign');
  const el_review = document.getElementById('t_review');
  const el_partner = document.getElementById('c_partner');
  const el_reports = document.getElementById('u_reports');
  if (el_head) el_head.innerHTML = partners;
  if (el_assign) el_assign.innerHTML = staff;
  if (el_review) el_review.innerHTML = reviewers;
  if (el_partner) el_partner.innerHTML = partners;
  if (el_reports) el_reports.innerHTML = reviewers;
}
function populateTaskSelects() {
  const opts = DB.tasks.map(t => `<option value="${t.name}">${t.id} – ${t.name}</option>`).join('');
  const inv_task = document.getElementById('inv_task');
  const ts_task = document.getElementById('ts_task');
  const t_parent = document.getElementById('t_parent');
  if (inv_task) inv_task.innerHTML = '<option value="">Select task…</option>' + opts;
  if (ts_task) ts_task.innerHTML = '<option value="">Select task…</option>' + opts;
  if (t_parent) t_parent.innerHTML = '<option value="">— Main Task —</option>' + opts;
}

// ── SEARCH ───────────────────────────────────────────────────
function globalSearch(q) {
  if (!q || q.length < 2) { renderTasks(); return; }
  q = q.toLowerCase();
  const filtered = DB.tasks.filter(t =>
    t.name.toLowerCase().includes(q) ||
    t.client.toLowerCase().includes(q) ||
    t.assignedTo.toLowerCase().includes(q) ||
    t.group.toLowerCase().includes(q)
  );
  const tbody = document.getElementById('taskTableBody');
  if (tbody) {
    tbody.innerHTML = filtered.length ? filtered.map(t => `
      <tr><td>${t.id}</td><td><b>${t.name}</b></td><td><span class="chip">${t.group}</span></td>
      <td>${t.client}</td><td>${t.head}</td><td>${t.assignedTo}</td><td>${t.reviewedBy}</td>
      <td>${priorityBadge(t.priority)}</td><td>${statusBadge(t.status)}</td>
      <td>${t.dueDate||'—'}</td><td>—</td><td></td></tr>`).join('')
      : `<tr><td colspan="12" class="empty-row">No results for "${q}"</td></tr>`;
  }
}

// ── NAV / UI ──────────────────────────────────────────────────
function showPage(id, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  const pg = document.getElementById('page-' + id);
  if (pg) pg.classList.add('active');
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  if (el) el.classList.add('active');
  const titles = {
    dashboard:'Dashboard', tasks:'Tasks', clients:'Clients',
    billing:'Billing & Invoices', payments:'Payments', team:'Team',
    timesheet:'Timesheets', documents:'Documents',
    deadlines:'Due Dates / Statutory Calendar', reports:'Reports', settings:'Settings'
  };
  document.getElementById('topbarTitle').textContent = titles[id] || id;
  // Close sidebar on mobile
  if (window.innerWidth < 768) document.getElementById('sidebar').classList.remove('open');
}

function switchTab(el) {
  el.closest('.tabs').querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
}

function openModal(id) {
  document.getElementById(id).classList.add('open');
}
function closeModal(id) {
  document.getElementById(id).classList.remove('open');
}

function toggleSidebar() {
  document.getElementById('sidebar').classList.toggle('open');
}

// ── SETTINGS ─────────────────────────────────────────────────
function toggleSetup() {
  const el = document.getElementById('loginSetup');
  el.style.display = el.style.display === 'none' ? 'block' : 'none';
}
function saveScriptUrl() {
  const url = document.getElementById('scriptUrlInput').value.trim();
  if (!url) return;
  SCRIPT_URL = url;
  localStorage.setItem('rrh_script_url', url);
  document.getElementById('loginSetup').style.display = 'none';
  showToast('API URL saved!', 'success');
}
function saveScriptUrlFromSettings() {
  const url = document.getElementById('settingsScriptUrl').value.trim();
  if (!url) { showToast('Enter a URL first', 'error'); return; }
  SCRIPT_URL = url;
  localStorage.setItem('rrh_script_url', url);
  showToast('Saved!', 'success');
}
async function testConnection() {
  const status = document.getElementById('connectionStatus');
  status.textContent = '⏳ Testing…';
  try {
    const res = await fetch(SCRIPT_URL + '?action=getUsers');
    const data = await res.json();
    status.innerHTML = data.rows ? '✅ <span style="color:var(--success)">Connected! Google Sheets is working.</span>' : '⚠️ Connected but unexpected response.';
  } catch (e) {
    status.innerHTML = '❌ <span style="color:var(--danger)">Connection failed. Check URL.</span>';
  }
}

// Populate settings URL field on load
window.addEventListener('DOMContentLoaded', () => {
  const s = document.getElementById('settingsScriptUrl');
  if (s && SCRIPT_URL) s.value = SCRIPT_URL;
  // Pre-fill invoice date
  const invDate = document.getElementById('inv_date');
  if (invDate) invDate.value = today();
});

// Close modals on overlay click
window.addEventListener('click', e => {
  document.querySelectorAll('.modal-overlay').forEach(m => {
    if (e.target === m) m.classList.remove('open');
  });
});
