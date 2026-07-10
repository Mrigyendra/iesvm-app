import { supabase } from './supabase.js';
import Chart from 'chart.js/auto';
import * as XLSX from 'xlsx';

// Constant settings
const CUR = 'Rs '; // currency prefix
const today = () => new Date().toISOString().slice(0, 10);
const monthKey = d => d ? d.slice(0, 7) : '';
const money = n => CUR + (Number(n) || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 });

// Application State Store
const store = {
  vehicles: [],
  fuel: [],
  maint: [],
  accounts: [], // Admin only: user profiles list
  user: null,
  role: 'staff' // defaults to read-only staff
};

// Cached Chart Instances
let charts = {};

// Helpers
const vName = id => {
  const v = store.vehicles.find(x => x.id === id);
  return v ? v.no : '—';
};

// ---- Authentication Setup ----
let isSignUpMode = false;
let pickedRole = 'admin';

document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  // Bind Auth UI Event Listeners
  const btnSubmitAuth = document.getElementById('btnSubmitAuth');
  const toggleAuthMode = document.getElementById('toggleAuthMode');
  const loginEmail = document.getElementById('loginEmail');
  const loginPass = document.getElementById('loginPass');
  const tabAdmin = document.getElementById('tabAdmin');
  const tabStaff = document.getElementById('tabStaff');

  if (tabAdmin && tabStaff) {
    tabAdmin.addEventListener('click', () => {
      pickedRole = 'admin';
      tabAdmin.classList.add('active');
      tabStaff.classList.remove('active');
      document.getElementById('loginErr').textContent = '';
    });
    tabStaff.addEventListener('click', () => {
      pickedRole = 'staff';
      tabStaff.classList.add('active');
      tabAdmin.classList.remove('active');
      document.getElementById('loginErr').textContent = '';
    });
  }

  btnSubmitAuth.addEventListener('click', handleAuthSubmit);
  toggleAuthMode.addEventListener('click', (e) => {
    e.preventDefault();
    isSignUpMode = !isSignUpMode;
    document.getElementById('loginTitle').textContent = isSignUpMode ? 'Create IES VM Account' : 'IES VM';
    document.getElementById('loginSub').textContent = isSignUpMode ? 'Register a new profile (defaults to Staff)' : 'Vehicle Management · Innovative Engineering Services';
    btnSubmitAuth.textContent = isSignUpMode ? 'Sign Up' : 'Sign In';
    toggleAuthMode.textContent = isSignUpMode ? 'Already have an account? Sign In' : "Don't have an account? Sign Up";
    
    // Hide role tabs during sign up
    const roletabs = document.querySelector('.roletabs');
    if (roletabs) {
      roletabs.style.display = isSignUpMode ? 'none' : 'flex';
    }
    document.getElementById('loginErr').textContent = '';
  });

  // Allow enter key submission
  [loginEmail, loginPass].forEach(el => {
    el.addEventListener('keydown', e => {
      if (e.key === 'Enter') handleAuthSubmit();
    });
  });

  // Toggle Password Visibility
  const togglePassVisibility = document.getElementById('togglePassVisibility');
  if (togglePassVisibility) {
    togglePassVisibility.addEventListener('click', () => {
      const isPass = loginPass.type === 'password';
      loginPass.type = isPass ? 'text' : 'password';
      togglePassVisibility.textContent = isPass ? '🙈' : '👁️';
    });
  }

  // App Navigation Listeners
  document.querySelectorAll('nav button').forEach(b => {
    b.addEventListener('click', () => {
      document.querySelectorAll('nav button').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.view').forEach(x => x.classList.remove('active'));
      b.classList.add('active');
      document.getElementById(b.dataset.view).classList.add('active');
      renderAll();
    });
  });

  // Setup other UI controls
  setupControls();

  // Listen for Supabase auth state change
  supabase.auth.onAuthStateChange(async (event, session) => {
    if (session) {
      store.user = session.user;
      await fetchUserRole(session.user.id);
      
      // Validate role if user is logging in (exclude sign up mode)
      const loginScreenVisible = document.getElementById('loginScreen').style.display !== 'none';
      if (loginScreenVisible && !isSignUpMode && store.role !== pickedRole) {
        document.getElementById('loginErr').textContent = `Access denied: Account is registered as ${store.role.toUpperCase()}, but you selected ${pickedRole.toUpperCase()} login.`;
        if (btnSubmitAuth) btnSubmitAuth.disabled = false;
        await supabase.auth.signOut();
        return;
      }
      
      showAppScreen();
      await fetchAllData();
    } else {
      store.user = null;
      store.role = 'staff';
      showLoginScreen();
    }
  });

  // Register Service Worker for PWA
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('/sw.js')
        .then(reg => console.log('Service Worker registered', reg.scope))
        .catch(err => console.error('Service Worker registration failed', err));
    });
  }
}

// Fetch user role from public.profiles
async function fetchUserRole(userId) {
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('role')
      .eq('id', userId)
      .single();

    if (error) {
      console.warn('Error reading user profile, defaulting to staff:', error.message);
      store.role = 'staff';
    } else if (data) {
      store.role = data.role;
    }
  } catch (err) {
    console.error('Failed to get user profile', err);
    store.role = 'staff';
  }
}

function showAppScreen() {
  document.body.classList.toggle('role-staff', store.role === 'staff');
  const badge = document.getElementById('roleBadge');
  badge.textContent = store.role === 'admin' ? 'Admin' : 'Staff';
  badge.className = 'rolebadge ' + (store.role === 'admin' ? 'admin' : 'staff');

  // Staff banner (allows adding entries, but hides edit/delete buttons)
  const note = store.role === 'staff'
    ? '<div class="viewonly-note">👁 Staff view — you can view and add entries, but you cannot edit or delete them once saved.</div>' 
    : '';
  document.getElementById('fuelViewNote').innerHTML = note;
  document.getElementById('maintViewNote').innerHTML = note;

  document.getElementById('loginScreen').style.display = 'none';
  document.getElementById('app').classList.add('on');
}

function showLoginScreen() {
  document.getElementById('app').classList.remove('on');
  document.body.classList.remove('role-staff');
  document.getElementById('loginScreen').style.display = 'grid';
  document.getElementById('loginEmail').value = '';
  document.getElementById('loginPass').value = '';
  document.getElementById('loginErr').textContent = '';
}

// ---- Database CRUD Operations ----

async function fetchAllData() {
  try {
    // 1. Fetch Vehicles
    const { data: vData, error: vErr } = await supabase
      .from('vehicles')
      .select('*')
      .order('no');
    if (vErr) throw vErr;
    store.vehicles = vData || [];

    // 2. Fetch Fuel
    const { data: fData, error: fErr } = await supabase
      .from('fuel')
      .select('*')
      .order('date', { ascending: false });
    if (fErr) throw fErr;
    store.fuel = (fData || []).map(r => ({
      id: r.id,
      veh: r.veh,
      date: r.date,
      km: Number(r.km),
      prevOdo: Number(r.prev_odo),
      odo: Number(r.odo),
      lit: Number(r.lit),
      cost: Number(r.cost)
    }));

    // 3. Fetch Maintenance
    const { data: mData, error: mErr } = await supabase
      .from('maint')
      .select('*')
      .order('date', { ascending: false });
    if (mErr) throw mErr;
    store.maint = (mData || []).map(r => ({
      id: r.id,
      veh: r.veh,
      date: r.date,
      cat: r.cat,
      cost: Number(r.cost),
      odo: r.odo ? Number(r.odo) : null,
      vendor: r.vendor,
      desc: r.description,
      nextDate: r.next_date,
      nextKm: r.next_km ? Number(r.next_km) : null
    }));

    if (store.role === 'admin') {
      await fetchAccounts();
    }
    renderAll();
  } catch (err) {
    console.error('Error fetching database records:', err.message);
    alert('Failed to synchronize database records: ' + err.message);
  }
}

// Check admin role guard
function requireAdmin() {
  if (store.role !== 'admin') {
    alert('Access denied: Staff accounts cannot edit or delete records.');
    return false;
  }
  return true;
}

// Bind Button Listeners and Dialog controls
function setupControls() {
  const vehDlg = document.getElementById('vehDlg');
  const passDlg = document.getElementById('passDlg');

  // Vehicles Dialog Action
  document.getElementById('btnAddVehicle').addEventListener('click', () => {
    if (!requireAdmin()) return;
    document.getElementById('vId').value = '';
    document.getElementById('vNo').value = '';
    document.getElementById('vModel').value = '';
    document.getElementById('vDriver').value = '';
    document.getElementById('vYear').value = '';
    document.getElementById('vType').value = 'Car';
    document.getElementById('vFuel').value = 'Diesel';
    document.getElementById('vehDlgTitle').textContent = 'Add Vehicle';
    vehDlg.showModal();
  });

  document.getElementById('btnCloseVehDlg').addEventListener('click', () => vehDlg.close());
  document.getElementById('btnCancelVehDlg').addEventListener('click', () => vehDlg.close());
  document.getElementById('btnSaveVehicle').addEventListener('click', saveVehicle);

  // Fuel Form Action
  document.getElementById('btnSaveFuel').addEventListener('click', addFuel);
  document.getElementById('fVeh').addEventListener('change', fuelPreview);
  document.getElementById('fOdo').addEventListener('input', fuelPreview);
  document.getElementById('fLit').addEventListener('input', fuelPreview);
  document.getElementById('fPrevOdo').addEventListener('input', fuelPreview);
  document.getElementById('fFilter').addEventListener('change', renderFuel);

  // Maintenance Form Action
  document.getElementById('btnSaveMaint').addEventListener('click', addMaint);
  document.getElementById('mFilter').addEventListener('change', renderMaint);

  // Backup & Restore
  document.getElementById('btnExportData').addEventListener('click', exportData);
  document.getElementById('btnRestoreDataTrigger').addEventListener('click', () => {
    if (!requireAdmin()) return;
    document.getElementById('importFile').click();
  });
  document.getElementById('importFile').addEventListener('change', importData);
  document.getElementById('btnLoadSample').addEventListener('click', loadSample);

  // Excel & Reports Action
  document.getElementById('btnExportMonthExcel').addEventListener('click', exportMonthExcel);
  document.getElementById('btnExportAllExcel').addEventListener('click', exportAllExcel);
  document.getElementById('rVeh').addEventListener('change', renderReport);
  document.getElementById('btnPrintReport').addEventListener('click', () => window.print());

  // Password Update modal
  document.getElementById('btnChangePass').addEventListener('click', () => {
    document.getElementById('newPass').value = '';
    document.getElementById('confirmPass').value = '';
    document.getElementById('passErr').textContent = '';
    passDlg.showModal();
  });
  document.getElementById('btnClosePassDlg').addEventListener('click', () => passDlg.close());
  document.getElementById('btnCancelPassDlg').addEventListener('click', () => passDlg.close());
  document.getElementById('btnConfirmChangePass').addEventListener('click', changePassword);

  // Logout Action
  document.getElementById('btnLogout').addEventListener('click', async () => {
    const { error } = await supabase.auth.signOut();
    if (error) alert('Error logging out: ' + error.message);
  });

  // Admin Password Reset Modal Action
  const adminPassDlg = document.getElementById('adminPassDlg');
  if (adminPassDlg) {
    document.getElementById('btnCloseAdminPassDlg').addEventListener('click', () => adminPassDlg.close());
    document.getElementById('btnCancelAdminPassDlg').addEventListener('click', () => adminPassDlg.close());
    document.getElementById('btnConfirmAdminChangePass').addEventListener('click', resetUserPassword);
  }
}

// Authentication Submit
async function handleAuthSubmit() {
  const email = document.getElementById('loginEmail').value.trim();
  const password = document.getElementById('loginPass').value;
  const errDiv = document.getElementById('loginErr');
  errDiv.textContent = '';

  if (!email || !password) {
    errDiv.textContent = 'Please enter both email and password.';
    return;
  }

  const submitBtn = document.getElementById('btnSubmitAuth');
  submitBtn.disabled = true;
  submitBtn.textContent = isSignUpMode ? 'Registering...' : 'Signing In...';

  try {
    if (isSignUpMode) {
      const { data, error } = await supabase.auth.signUp({ email, password });
      if (error) throw error;
      
      if (data.user && data.session === null) {
        errDiv.innerHTML = '<span style="color:var(--green)">Registration successful! Please check your email to confirm registration before logging in.</span>';
      } else {
        alert('Registration complete! Logging in...');
      }
    } else {
      const { error } = await supabase.auth.signInWithPassword({ email, password });
      if (error) throw error;
    }
  } catch (err) {
    errDiv.textContent = err.message;
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = isSignUpMode ? 'Sign Up' : 'Sign In';
  }
}

// User password updates
async function changePassword() {
  const np = document.getElementById('newPass').value;
  const cp = document.getElementById('confirmPass').value;
  const errDiv = document.getElementById('passErr');
  errDiv.textContent = '';

  if (np.length < 6) {
    errDiv.textContent = 'Password must be at least 6 characters.';
    return;
  }
  if (np !== cp) {
    errDiv.textContent = 'Passwords do not match.';
    return;
  }

  const btn = document.getElementById('btnConfirmChangePass');
  btn.disabled = true;

  try {
    const { error } = await supabase.auth.updateUser({ password: np });
    if (error) throw error;
    alert('Password updated successfully!');
    document.getElementById('passDlg').close();
  } catch (err) {
    errDiv.textContent = err.message;
  } finally {
    btn.disabled = false;
  }
}

// ---- Vehicles Section Operations ----

async function saveVehicle() {
  const idVal = document.getElementById('vId').value;
  if (idVal) {
    if (!requireAdmin()) return;
  } else {
    if (store.role !== 'admin' && store.role !== 'staff') {
      alert('Access denied: Unauthorized role.');
      return;
    }
  }
  const vNoVal = document.getElementById('vNo').value.trim();
  const vModelVal = document.getElementById('vModel').value.trim();
  const vDriverVal = document.getElementById('vDriver').value.trim();
  const vYearVal = document.getElementById('vYear').value.trim();
  const vTypeVal = document.getElementById('vType').value;
  const vFuelVal = document.getElementById('vFuel').value;

  if (!vNoVal) {
    alert('Vehicle number is required');
    return;
  }

  const payload = {
    no: vNoVal,
    model: vModelVal,
    driver: vDriverVal,
    year: vYearVal,
    type: vTypeVal,
    fuel: vFuelVal
  };

  try {
    if (idVal) {
      // Update
      const { error } = await supabase
        .from('vehicles')
        .update(payload)
        .eq('id', idVal);
      if (error) throw error;
    } else {
      // Insert
      const { error } = await supabase
        .from('vehicles')
        .insert([payload]);
      if (error) throw error;
    }
    
    document.getElementById('vehDlg').close();
    await fetchAllData();
  } catch (err) {
    alert('Error saving vehicle: ' + err.message);
  }
}

// Global hook for vehicle delete (defined on window to let HTML onclick access it)
window.delVehicle = async function(id) {
  if (!requireAdmin()) return;
  if (!confirm('Are you sure you want to delete this vehicle? ALL fuel logs and maintenance records will be deleted as well.')) return;

  try {
    const { error } = await supabase
      .from('vehicles')
      .delete()
      .eq('id', id);
    if (error) throw error;
    await fetchAllData();
  } catch (err) {
    alert('Error deleting vehicle: ' + err.message);
  }
};

window.openVehicle = function(id) {
  if (!requireAdmin()) return;
  const v = store.vehicles.find(x => x.id === id);
  if (!v) return;

  document.getElementById('vId').value = v.id;
  document.getElementById('vNo').value = v.no;
  document.getElementById('vModel').value = v.model || '';
  document.getElementById('vDriver').value = v.driver || '';
  document.getElementById('vYear').value = v.year || '';
  document.getElementById('vType').value = v.type || 'Car';
  document.getElementById('vFuel').value = v.fuel || 'Diesel';
  document.getElementById('vehDlgTitle').textContent = 'Edit Vehicle';
  document.getElementById('vehDlg').showModal();
};

function vehStats(id) {
  const f = store.fuel.filter(x => x.veh === id);
  const km = f.reduce((s, x) => s + (x.km || 0), 0);
  const lit = f.reduce((s, x) => s + (x.lit || 0), 0);
  const odoSorted = f.filter(x => x.odo).sort((a, b) => b.odo - a.odo);
  const lastOdoVal = odoSorted.length ? odoSorted[0].odo : null;
  return {
    logs: f.length,
    km,
    lit,
    avg: lit ? km / lit : 0,
    odo: lastOdoVal
  };
}

function renderVehicles() {
  const b = document.getElementById('vehBody');
  b.innerHTML = '';
  document.getElementById('vehEmpty').style.display = store.vehicles.length ? 'none' : 'block';

  store.vehicles.forEach(v => {
    const s = vehStats(v.id);
    b.insertAdjacentHTML('beforeend', `<tr>
      <td class="vehchip">${v.no}</td>
      <td>${v.type || '—'}</td>
      <td>${v.model || '—'}</td>
      <td>${v.driver || '—'}</td>
      <td>${v.fuel || '—'}</td>
      <td>${s.logs}</td>
      <td>${s.odo != null ? s.odo.toLocaleString() : '—'}</td>
      <td><b>${s.avg ? s.avg.toFixed(2) : '—'}</b></td>
      <td style="white-space:nowrap">
        <button class="btn ghost sm admin-only" onclick="openVehicle('${v.id}')">Edit</button>
        <button class="btn danger sm admin-only" onclick="delVehicle('${v.id}')">Del</button>
      </td>
    </tr>`);
  });
}

// ---- Fuel Log Operations ----

function lastOdo(vehId) {
  const rows = store.fuel.filter(x => x.veh === vehId && x.odo != null).sort((a, b) => b.odo - a.odo);
  return rows.length ? rows[0].odo : null;
}

function autoPrevOdo() {
  const selectVal = document.getElementById('fVeh').value;
  const lo = lastOdo(selectVal);
  const prevOdoInput = document.getElementById('fPrevOdo');
  if (lo != null && !prevOdoInput.value) {
    prevOdoInput.placeholder = lo + ' (last recorded)';
    return lo;
  }
  return prevOdoInput.value ? Number(prevOdoInput.value) : (lo != null ? lo : null);
}

function fuelPreview() {
  const prev = autoPrevOdo();
  const cur = Number(document.getElementById('fOdo').value);
  const lit = Number(document.getElementById('fLit').value);
  const previewDiv = document.getElementById('fPreview');

  if (cur && prev != null) {
    const km = cur - prev;
    if (km <= 0) {
      previewDiv.innerHTML = '<span style="color:var(--danger)">⚠ Current odometer must be greater than previous.</span>';
      return;
    }
    previewDiv.textContent = `→ Distance: ${km} km  ·  Mileage: ${lit ? (km / lit).toFixed(2) + ' km/L' : 'enter litres'}`;
  } else {
    previewDiv.textContent = cur ? 'Enter previous odometer (or log one entry first).' : '';
  }
}

async function addFuel() {
  if (store.role !== 'admin' && store.role !== 'staff') {
    alert('Access denied: Unauthorized role.');
    return;
  }
  const veh = document.getElementById('fVeh').value;
  const date = document.getElementById('fDate').value || today();
  const cur = Number(document.getElementById('fOdo').value);
  const lit = Number(document.getElementById('fLit').value);
  const cost = Number(document.getElementById('fCost').value) || 0;

  if (!veh) {
    alert('Add and select a vehicle first');
    return;
  }

  const prev = document.getElementById('fPrevOdo').value ? Number(document.getElementById('fPrevOdo').value) : lastOdo(veh);

  if (!cur) {
    alert('Enter the current odometer reading');
    return;
  }
  if (prev == null) {
    alert('Enter the previous odometer reading (the first entry for this vehicle requires it).');
    return;
  }
  if (cur - prev <= 0) {
    alert('Current odometer must be greater than previous odometer.');
    return;
  }
  if (!lit) {
    alert('Enter fuel inserted (L)');
    return;
  }

  const payload = {
    veh,
    date,
    km: cur - prev,
    prev_odo: prev,
    odo: cur,
    lit,
    cost
  };

  try {
    const { error } = await supabase
      .from('fuel')
      .insert([payload]);
    if (error) throw error;

    // Reset inputs
    document.getElementById('fOdo').value = '';
    document.getElementById('fLit').value = '';
    document.getElementById('fCost').value = '';
    document.getElementById('fPrevOdo').value = '';
    document.getElementById('fPrevOdo').placeholder = 'auto';
    document.getElementById('fPreview').textContent = '';

    await fetchAllData();
  } catch (err) {
    alert('Error logging fuel entry: ' + err.message);
  }
}

window.delFuel = async function(id) {
  if (!requireAdmin()) return;
  if (!confirm('Are you sure you want to delete this fuel record?')) return;

  try {
    const { error } = await supabase
      .from('fuel')
      .delete()
      .eq('id', id);
    if (error) throw error;
    await fetchAllData();
  } catch (err) {
    alert('Error deleting fuel log: ' + err.message);
  }
};

function renderFuel() {
  const filt = document.getElementById('fFilter').value;
  const rows = store.fuel.filter(x => !filt || x.veh === filt).sort((a, b) => b.date.localeCompare(a.date));
  const b = document.getElementById('fuelBody');
  b.innerHTML = '';
  document.getElementById('fuelEmpty').style.display = rows.length ? 'none' : 'block';

  rows.forEach(r => {
    b.insertAdjacentHTML('beforeend', `<tr>
      <td>${r.date}</td>
      <td class="vehchip">${vName(r.veh)}</td>
      <td>${r.prevOdo != null ? r.prevOdo.toLocaleString() : '—'}</td>
      <td>${r.odo != null ? r.odo.toLocaleString() : '—'}</td>
      <td>${r.km}</td>
      <td>${r.lit}</td>
      <td><b>${(r.km / r.lit).toFixed(2)}</b></td>
      <td>${r.cost ? money(r.cost) : '—'}</td>
      <td>${r.cost && r.lit ? money(r.cost / r.lit) : '—'}</td>
      <td><button class="btn danger sm admin-only" onclick="delFuel('${r.id}')">Del</button></td>
    </tr>`);
  });
}

// ---- Maintenance Section Operations ----

async function addMaint() {
  if (store.role !== 'admin' && store.role !== 'staff') {
    alert('Access denied: Unauthorized role.');
    return;
  }
  const veh = document.getElementById('mVeh').value;
  const date = document.getElementById('mDate').value || today();
  const cat = document.getElementById('mCat').value;
  const cost = Number(document.getElementById('mCost').value);
  const odo = document.getElementById('mOdo').value ? Number(document.getElementById('mOdo').value) : null;
  const vendor = document.getElementById('mVendor').value.trim();
  const desc = document.getElementById('mDesc').value.trim();
  const nextDate = document.getElementById('mNextDate').value || null;
  const nextKm = document.getElementById('mNextKm').value ? Number(document.getElementById('mNextKm').value) : null;

  if (!veh) {
    alert('Add and select a vehicle first');
    return;
  }
  if (!cost) {
    alert('Enter the maintenance cost');
    return;
  }

  const payload = {
    veh,
    date,
    cat,
    cost,
    odo,
    vendor,
    description: desc,
    next_date: nextDate,
    next_km: nextKm
  };

  try {
    const { error } = await supabase
      .from('maint')
      .insert([payload]);
    if (error) throw error;

    // Reset fields
    document.getElementById('mCost').value = '';
    document.getElementById('mOdo').value = '';
    document.getElementById('mVendor').value = '';
    document.getElementById('mDesc').value = '';
    document.getElementById('mNextKm').value = '';
    document.getElementById('mNextDate').value = '';

    await fetchAllData();
  } catch (err) {
    alert('Error saving maintenance record: ' + err.message);
  }
}

window.delMaint = async function(id) {
  if (!requireAdmin()) return;
  if (!confirm('Are you sure you want to delete this maintenance record?')) return;

  try {
    const { error } = await supabase
      .from('maint')
      .delete()
      .eq('id', id);
    if (error) throw error;
    await fetchAllData();
  } catch (err) {
    alert('Error deleting maintenance record: ' + err.message);
  }
};

function renderMaint() {
  const filt = document.getElementById('mFilter').value;
  const rows = store.maint.filter(x => !filt || x.veh === filt).sort((a, b) => b.date.localeCompare(a.date));
  const b = document.getElementById('maintBody');
  b.innerHTML = '';
  document.getElementById('maintEmpty').style.display = rows.length ? 'none' : 'block';

  rows.forEach(r => {
    const nd = r.nextDate ? `${r.nextDate}` : (r.nextKm ? `${r.nextKm} km` : '—');
    const isMajor = r.cat === 'Major Repair';
    b.insertAdjacentHTML('beforeend', `<tr>
      <td>${r.date}</td>
      <td class="vehchip">${vName(r.veh)}</td>
      <td>${isMajor ? '<span class="pill bad">' + r.cat + '</span>' : r.cat}</td>
      <td>${r.desc || '—'}${r.odo ? '<br><span class="muted">@' + r.odo.toLocaleString() + ' km</span>' : ''}</td>
      <td>${r.vendor || '—'}</td>
      <td><b>${money(r.cost)}</b></td>
      <td>${nd}</td>
      <td><button class="btn danger sm admin-only" onclick="delMaint('${r.id}')">Del</button></td>
    </tr>`);
  });
}

// ---- Populate Selectors ----

function fillSelectors() {
  const opts = store.vehicles.map(v => `<option value="${v.id}">${v.no}${v.model ? ' — ' + v.model : ''}</option>`).join('');
  document.getElementById('fVeh').innerHTML = opts || '<option value="">No vehicles</option>';
  document.getElementById('mVeh').innerHTML = opts || '<option value="">No vehicles</option>';
  document.getElementById('rVeh').innerHTML = opts || '<option value="">No vehicles</option>';

  const fo = '<option value="">All vehicles</option>' + opts;
  const keepF = document.getElementById('fFilter').value;
  const keepM = document.getElementById('mFilter').value;

  document.getElementById('fFilter').innerHTML = fo;
  document.getElementById('mFilter').innerHTML = fo;
  document.getElementById('fFilter').value = keepF;
  document.getElementById('mFilter').value = keepM;

  if (!document.getElementById('fDate').value) document.getElementById('fDate').value = today();
  if (!document.getElementById('mDate').value) document.getElementById('mDate').value = today();
}

// ---- Dashboard Render Stats & Charts ----

function renderStats() {
  const totKm = store.fuel.reduce((s, x) => s + (x.km || 0), 0);
  const totLit = store.fuel.reduce((s, x) => s + (x.lit || 0), 0);
  const fuelCost = store.fuel.reduce((s, x) => s + (x.cost || 0), 0);
  const maintCost = store.maint.reduce((s, x) => s + (x.cost || 0), 0);
  const avg = totLit ? totKm / totLit : 0;
  
  const statsList = [
    ['Vehicles', store.vehicles.length, ''],
    ['Distance Logged', totKm.toLocaleString(), 'km'],
    ['Fuel Consumed', totLit.toFixed(1), 'L'],
    ['Fleet Avg Mileage', avg.toFixed(2), 'km/L'],
    ['Fuel Spend', money(fuelCost), ''],
    ['Maintenance Spend', money(maintCost), ''],
  ];

  document.getElementById('dashStats').innerHTML = statsList.map(([l, n, u]) => `
    <div class="card">
      <div class="label">${l}</div>
      <div class="num">${n} <span class="unit">${u}</span></div>
    </div>
  `).join('');

  // Check upcoming maintenance alerts (due in <= 14 days or overdue)
  const alerts = [];
  store.maint.forEach(m => {
    if (m.nextDate) {
      const days = Math.ceil((new Date(m.nextDate) - new Date()) / 864e5);
      if (days <= 14) {
        alerts.push(`${vName(m.veh)}: service due ${days < 0 ? Math.abs(days) + ' days overdue' : 'in ' + days + ' days'} (${m.nextDate})`);
      }
    }
  });
  document.getElementById('dashAlerts').innerHTML = alerts.length 
    ? alerts.map(a => `<div class="alertbar">⚠ ${a}</div>`).join('') 
    : '';
}

function mkChart(id, cfg) {
  if (charts[id]) charts[id].destroy();
  const el = document.getElementById(id);
  if (el) {
    charts[id] = new Chart(el, cfg);
  }
}

const C = {
  diesel: '#c65330',
  green: '#256a4b',
  amber: '#b8861e',
  blue: '#3a6ea2',
  line: '#e2dfd5'
};

function baseOpts(unit) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: { legend: { display: false } },
    scales: {
      y: { beginAtZero: true, grid: { color: '#f0ede4' }, title: { display: true, text: unit } },
      x: { grid: { display: false } }
    }
  };
}

function renderCharts() {
  const vs = store.vehicles;
  const labels = vs.map(v => v.no);
  const avgData = vs.map(v => {
    const s = vehStats(v.id);
    return Number(s.avg.toFixed(2));
  });
  const litData = vs.map(v => Number(vehStats(v.id).lit.toFixed(1)));

  mkChart('chMileage', {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'km/L', data: avgData, backgroundColor: C.green, borderRadius: 6 }]
    },
    options: baseOpts('km/L')
  });

  mkChart('chFuel', {
    type: 'bar',
    data: {
      labels,
      datasets: [{ label: 'Litres', data: litData, backgroundColor: C.diesel, borderRadius: 6 }]
    },
    options: baseOpts('Litres')
  });

  // Monthly repair cost
  const months = [...new Set(store.maint.map(m => monthKey(m.date)))].sort();
  const monthCost = months.map(mo => store.maint.filter(m => monthKey(m.date) === mo).reduce((s, x) => s + (x.cost || 0), 0));
  
  mkChart('chRepairMonth', {
    type: 'bar',
    data: {
      labels: months,
      datasets: [{ label: 'Cost', data: monthCost, backgroundColor: C.amber, borderRadius: 6 }]
    },
    options: baseOpts(CUR.trim())
  });

  // Cost split
  const fuelCost = store.fuel.reduce((s, x) => s + (x.cost || 0), 0);
  const maintCost = store.maint.reduce((s, x) => s + (x.cost || 0), 0);
  
  mkChart('chCostSplit', {
    type: 'doughnut',
    data: {
      labels: ['Fuel', 'Maintenance'],
      datasets: [{ data: [fuelCost, maintCost], backgroundColor: [C.diesel, C.amber], borderWidth: 0 }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'bottom' } }
    }
  });
}

// ---- Per-Vehicle Report ----

function renderReport() {
  const id = document.getElementById('rVeh').value;
  const area = document.getElementById('reportArea');
  if (!id) {
    area.innerHTML = '<div class="empty">No vehicle selected.</div>';
    return;
  }

  const v = store.vehicles.find(x => x.id === id);
  const s = vehStats(id);
  const f = store.fuel.filter(x => x.veh === id).sort((a, b) => a.date.localeCompare(b.date));
  const m = store.maint.filter(x => x.veh === id).sort((a, b) => a.date.localeCompare(b.date));
  const mCost = m.reduce((a, x) => a + (x.cost || 0), 0);
  const fCost = f.reduce((a, x) => a + (x.cost || 0), 0);
  const costPerKm = s.km ? ((fCost + mCost) / s.km) : 0;

  area.innerHTML = `
    <div class="card" style="margin-bottom:20px; border-color: var(--diesel)">
      <h3 style="margin-bottom:12px; font-family: var(--font-display); font-size:18px;">
        ${v.no} <span class="muted">${v.model || ''} · ${v.type || ''} · ${v.fuel || ''}${v.driver ? ' · Driver: ' + v.driver : ''}</span>
      </h3>
      <div class="grid stat">
        <div class="card"><div class="label">Total Distance</div><div class="num">${s.km.toLocaleString()} <span class="unit">km</span></div></div>
        <div class="card"><div class="label">Fuel Used</div><div class="num">${s.lit.toFixed(1)} <span class="unit">L</span></div></div>
        <div class="card"><div class="label">Avg Mileage</div><div class="num">${s.avg.toFixed(2)} <span class="unit">km/L</span></div></div>
        <div class="card"><div class="label">Fuel Cost</div><div class="num" style="font-size:20px">${money(fCost)}</div></div>
        <div class="card"><div class="label">Maintenance Cost</div><div class="num" style="font-size:20px">${money(mCost)}</div></div>
        <div class="card"><div class="label">Running Cost / km</div><div class="num" style="font-size:20px">${money(costPerKm)}</div></div>
      </div>
    </div>
    <div class="chartbox">
      <h3>Mileage Trend (km/L per entry)</h3>
      <div class="cwrap"><canvas id="chRepTrend"></canvas></div>
    </div>
    <div class="row2">
      <div class="tablewrap">
        <table>
          <thead><tr><th>Date</th><th>km</th><th>L</th><th>km/L</th></tr></thead>
          <tbody>
            ${f.length ? f.map(r => `<tr><td>${r.date}</td><td>${r.km}</td><td>${r.lit}</td><td>${(r.km / r.lit).toFixed(2)}</td></tr>`).join('') : '<tr><td colspan=4 class="muted">No fuel logs</td></tr>'}
          </tbody>
        </table>
      </div>
      <div class="tablewrap">
        <table>
          <thead><tr><th>Date</th><th>Category</th><th>Cost</th></tr></thead>
          <tbody>
            ${m.length ? m.map(r => `<tr><td>${r.date}</td><td>${r.cat}</td><td>${money(r.cost)}</td></tr>`).join('') : '<tr><td colspan=3 class="muted">No maintenance records</td></tr>'}
          </tbody>
        </table>
      </div>
    </div>`;

  const tl = f.map(r => r.date);
  const td = f.map(r => Number((r.km / r.lit).toFixed(2)));

  mkChart('chRepTrend', {
    type: 'line',
    data: {
      labels: tl,
      datasets: [{
        label: 'km/L',
        data: td,
        borderColor: C.green,
        backgroundColor: 'rgba(37, 106, 75, 0.08)',
        fill: true,
        tension: 0.25,
        pointRadius: 4
      }]
    },
    options: baseOpts('km/L')
  });
}

// ---- Data Import / Export Operations ----

function exportData() {
  const backup = {
    vehicles: store.vehicles,
    fuel: store.fuel,
    maint: store.maint
  };
  const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `iesvm-supabase-backup-${today()}.json`;
  a.click();
}

async function importData(e) {
  if (!requireAdmin()) return;
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async () => {
    try {
      const d = JSON.parse(reader.result);
      if (d.vehicles && d.fuel && d.maint) {
        if (!confirm('This action will delete all existing vehicles, fuel logs, and maintenance logs in your Supabase database and replace them with this backup. Do you want to proceed?')) {
          return;
        }

        // Batch deletion
        const { error: dMaintErr } = await supabase.from('maint').delete().neq('id', '00000000-0000-0000-0000-000000000000');
        if (dMaintErr) throw dMaintErr;
        const { error: dFuelErr } = await supabase.from('fuel').delete().neq('id', '00000000-0000-0000-0000-000000000000');
        if (dFuelErr) throw dFuelErr;
        const { error: dVehErr } = await supabase.from('vehicles').delete().neq('id', '00000000-0000-0000-0000-000000000000');
        if (dVehErr) throw dVehErr;

        // Re-insert vehicles first
        const vPayloads = d.vehicles.map(v => ({
          id: v.id,
          no: v.no,
          type: v.type,
          model: v.model,
          fuel: v.fuel,
          driver: v.driver,
          year: v.year
        }));

        const { error: vInsErr } = await supabase.from('vehicles').insert(vPayloads);
        if (vInsErr) throw vInsErr;

        // Re-insert fuel
        if (d.fuel.length) {
          const fPayloads = d.fuel.map(f => ({
            id: f.id,
            veh: f.veh,
            date: f.date,
            km: f.km,
            prev_odo: f.prevOdo,
            odo: f.odo,
            lit: f.lit,
            cost: f.cost
          }));
          const { error: fInsErr } = await supabase.from('fuel').insert(fPayloads);
          if (fInsErr) throw fInsErr;
        }

        // Re-insert maintenance
        if (d.maint.length) {
          const mPayloads = d.maint.map(m => ({
            id: m.id,
            veh: m.veh,
            date: m.date,
            cat: m.cat,
            cost: m.cost,
            odo: m.odo,
            vendor: m.vendor,
            description: m.desc,
            next_date: m.nextDate,
            next_km: m.nextKm
          }));
          const { error: mInsErr } = await supabase.from('maint').insert(mPayloads);
          if (mInsErr) throw mInsErr;
        }

        alert('Database restore complete!');
        await fetchAllData();
      } else {
        alert('Invalid backup structure.');
      }
    } catch (err) {
      alert('Could not restore database backup: ' + err.message);
    }
  };
  reader.readAsText(file);
  e.target.value = ''; // clear input selection
}

async function loadSample() {
  if (!requireAdmin()) return;
  if (store.vehicles.length && !confirm('Clear current database tables and load sample records?')) return;

  try {
    // Empty tables
    await supabase.from('maint').delete().neq('id', '00000000-0000-0000-0000-000000000000');
    await supabase.from('fuel').delete().neq('id', '00000000-0000-0000-0000-000000000000');
    await supabase.from('vehicles').delete().neq('id', '00000000-0000-0000-0000-000000000000');

    // Add vehicles
    const vs = [
      { no: 'BA 2 KHA 1234', type: 'SUV / Jeep', model: 'Toyota Hilux', fuel: 'Diesel', driver: 'Ram B.', year: '2021' },
      { no: 'BA 5 CHA 8890', type: 'Car', model: 'Suzuki Dzire', fuel: 'Petrol', driver: 'Sita K.', year: '2022' },
      { no: 'BA 1 PA 4521', type: 'Pickup', model: 'Mahindra Bolero', fuel: 'Diesel', driver: 'Hari T.', year: '2019' }
    ];

    const { data: vInserted, error: vErr } = await supabase.from('vehicles').insert(vs).select();
    if (vErr) throw vErr;

    const v1 = vInserted.find(x => x.no === 'BA 2 KHA 1234').id;
    const v2 = vInserted.find(x => x.no === 'BA 5 CHA 8890').id;
    const v3 = vInserted.find(x => x.no === 'BA 1 PA 4521').id;

    // Fuel Logs
    const fs = [
      { veh: v1, date: '2026-03-05', km: 380, prev_odo: 44820, odo: 45200, lit: 32, cost: 4800 },
      { veh: v1, date: '2026-04-02', km: 410, prev_odo: 45200, odo: 45610, lit: 35, cost: 5250 },
      { veh: v1, date: '2026-05-06', km: 395, prev_odo: 45610, odo: 46005, lit: 34, cost: 5100 },
      { veh: v1, date: '2026-06-04', km: 420, prev_odo: 46005, odo: 46425, lit: 36, cost: 5400 },
      { veh: v2, date: '2026-03-08', km: 300, prev_odo: 27800, odo: 28100, lit: 20, cost: 3200 },
      { veh: v2, date: '2026-04-10', km: 320, prev_odo: 28100, odo: 28420, lit: 21, cost: 3360 },
      { veh: v2, date: '2026-05-12', km: 310, prev_odo: 28420, odo: 28730, lit: 20.5, cost: 3280 },
      { veh: v2, date: '2026-06-15', km: 330, prev_odo: 28730, odo: 29060, lit: 21.5, cost: 3440 },
      { veh: v3, date: '2026-03-15', km: 280, prev_odo: 60920, odo: 61200, lit: 30, cost: 4500 },
      { veh: v3, date: '2026-04-18', km: 300, prev_odo: 61200, odo: 61500, lit: 32, cost: 4800 },
      { veh: v3, date: '2026-05-20', km: 290, prev_odo: 61500, odo: 61790, lit: 31, cost: 4650 }
    ];

    const { error: fErr } = await supabase.from('fuel').insert(fs);
    if (fErr) throw fErr;

    // Maintenance Records
    const ms = [
      { veh: v1, date: '2026-03-20', cat: 'Routine Service', cost: 6500, description: 'Oil + filter change', vendor: 'City Motors', next_date: '2026-06-20' },
      { veh: v1, date: '2026-05-11', cat: 'Major Repair', cost: 28000, description: 'Clutch plate replacement', vendor: 'City Motors' },
      { veh: v2, date: '2026-04-05', cat: 'Tyre', cost: 12000, description: '2 front tyres', vendor: 'Tyre House', next_date: '2026-07-25' },
      { veh: v3, date: '2026-04-22', cat: 'Engine', cost: 35000, description: 'Injector overhaul', vendor: 'Diesel Works' },
      { veh: v3, date: '2026-06-01', cat: 'Routine Service', cost: 5500, description: 'General service', vendor: 'Diesel Works', next_date: '2026-07-15' },
      { veh: v2, date: '2026-06-10', cat: 'Brakes', cost: 8000, description: 'Brake pads', vendor: 'City Motors' }
    ];

    const { error: mErr } = await supabase.from('maint').insert(ms);
    if (mErr) throw mErr;

    alert('Sample records successfully loaded into Supabase!');
    await fetchAllData();
  } catch (err) {
    alert('Error loading sample data: ' + err.message);
  }
}

// ---- SheetJS Excel Exports ----

function monthLabel(mk) {
  if (!mk) return '';
  const [y, m] = mk.split('-');
  return new Date(y, m - 1, 1).toLocaleString('en', { month: 'long', year: 'numeric' });
}

function buildSheets(fuelRows, maintRows, periodLabel, vehLabel) {
  const totKm = fuelRows.reduce((s, x) => s + (x.km || 0), 0);
  const totLit = fuelRows.reduce((s, x) => s + (x.lit || 0), 0);
  const fuelCost = fuelRows.reduce((s, x) => s + (x.cost || 0), 0);
  const maintCost = maintRows.reduce((s, x) => s + (x.cost || 0), 0);

  const summary = [
    ['IES VM — Monthly Vehicle Report'],
    ['Period', periodLabel || 'All time'],
    ['Vehicle', vehLabel || 'All vehicles'],
    ['Generated', new Date().toLocaleString()],
    [],
    ['Total distance (km)', totKm],
    ['Total fuel (L)', Number(totLit.toFixed(2))],
    ['Average mileage (km/L)', totLit ? Number((totKm / totLit).toFixed(2)) : 0],
    ['Fuel cost', fuelCost],
    ['Maintenance cost', maintCost],
    ['Total cost', fuelCost + maintCost],
    ['Running Cost per km', totKm ? Number(((fuelCost + maintCost) / totKm).toFixed(2)) : 0],
  ];

  // Per-vehicle summary rows
  const perVeh = {};
  fuelRows.forEach(r => {
    const k = vName(r.veh);
    perVeh[k] = perVeh[k] || { km: 0, lit: 0, fuel: 0, maint: 0 };
    perVeh[k].km += r.km || 0;
    perVeh[k].lit += r.lit || 0;
    perVeh[k].fuel += r.cost || 0;
  });

  maintRows.forEach(r => {
    const k = vName(r.veh);
    perVeh[k] = perVeh[k] || { km: 0, lit: 0, fuel: 0, maint: 0 };
    perVeh[k].maint += r.cost || 0;
  });

  summary.push([], ['Per-vehicle breakdown']);
  summary.push(['Vehicle', 'Distance (km)', 'Fuel (L)', 'Avg km/L', 'Fuel cost', 'Maint cost']);
  
  Object.entries(perVeh).forEach(([k, v]) => {
    summary.push([
      k, 
      v.km, 
      Number(v.lit.toFixed(2)), 
      v.lit ? Number((v.km / v.lit).toFixed(2)) : 0, 
      v.fuel, 
      v.maint
    ]);
  });

  // Fuel list sheet
  const fuelSheet = [['Date', 'Vehicle', 'Prev Odometer', 'Current Odometer', 'Distance (km)', 'Fuel (L)', 'Mileage (km/L)', 'Fuel Cost', 'Rate/L']];
  fuelRows.sort((a, b) => a.date.localeCompare(b.date)).forEach(r => {
    fuelSheet.push([
      r.date,
      vName(r.veh),
      r.prevOdo ?? '',
      r.odo ?? '',
      r.km,
      r.lit,
      Number((r.km / r.lit).toFixed(2)),
      r.cost || 0,
      r.cost && r.lit ? Number((r.cost / r.lit).toFixed(2)) : ''
    ]);
  });

  // Maintenance list sheet
  const maintSheet = [['Date', 'Vehicle', 'Category', 'Description', 'Vendor', 'Odometer', 'Cost', 'Next Due Date', 'Next Due km']];
  maintRows.sort((a, b) => a.date.localeCompare(b.date)).forEach(r => {
    maintSheet.push([
      r.date,
      vName(r.veh),
      r.cat,
      r.desc || '',
      r.vendor || '',
      r.odo ?? '',
      r.cost || 0,
      r.nextDate || '',
      r.nextKm ?? ''
    ]);
  });

  const wb = XLSX.utils.book_new();
  const s1 = XLSX.utils.aoa_to_sheet(summary);
  s1['!cols'] = [{ wch: 24 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }];

  const s2 = XLSX.utils.aoa_to_sheet(fuelSheet);
  s2['!cols'] = [{ wch: 12 }, { wch: 18 }, { wch: 15 }, { wch: 16 }, { wch: 13 }, { wch: 10 }, { wch: 14 }, { wch: 12 }, { wch: 10 }];

  const s3 = XLSX.utils.aoa_to_sheet(maintSheet);
  s3['!cols'] = [{ wch: 12 }, { wch: 18 }, { wch: 16 }, { wch: 30 }, { wch: 16 }, { wch: 12 }, { wch: 12 }, { wch: 14 }, { wch: 12 }];

  XLSX.utils.book_append_sheet(wb, s1, 'Summary');
  XLSX.utils.book_append_sheet(wb, s2, 'Fuel & Mileage');
  XLSX.utils.book_append_sheet(wb, s3, 'Maintenance');
  return wb;
}

function exportMonthExcel() {
  const mk = document.getElementById('xlMonth').value;
  if (!mk) {
    alert('Pick a month first.');
    return;
  }
  const veh = document.getElementById('xlVeh').value;
  const fuelRows = store.fuel.filter(r => monthKey(r.date) === mk && (!veh || r.veh === veh));
  const maintRows = store.maint.filter(r => monthKey(r.date) === mk && (!veh || r.veh === veh));

  if (!fuelRows.length && !maintRows.length) {
    alert('No records found for that month/vehicle.');
    return;
  }

  const vehLabel = veh ? vName(veh) : 'All vehicles';
  const wb = buildSheets(fuelRows, maintRows, monthLabel(mk), vehLabel);
  XLSX.writeFile(wb, `iesvm-report-${mk}${veh ? '-' + vName(veh).replace(/\s+/g, '') : ''}.xlsx`);
}

function exportAllExcel() {
  const wb = buildSheets([...store.fuel], [...store.maint], 'All time', 'All vehicles');
  XLSX.writeFile(wb, `iesvm-report-full-${today()}.xlsx`);
}

// ---- Render All views ----

function renderAll() {
  fillSelectors();
  renderVehicles();
  renderFuel();
  renderMaint();
  renderStats();
  renderCharts();

  // Excel vehicle filter list update
  const xv = document.getElementById('xlVeh');
  if (xv) {
    const keep = xv.value;
    xv.innerHTML = '<option value="">All vehicles</option>' + store.vehicles.map(v => `<option value="${v.id}">${v.no}</option>`).join('');
    xv.value = keep;
  }

  if (document.getElementById('xlMonth') && !document.getElementById('xlMonth').value) {
    document.getElementById('xlMonth').value = today().slice(0, 7);
  }

  if (document.getElementById('reports').classList.contains('active')) {
    renderReport();
  }

  if (document.getElementById('accounts') && document.getElementById('accounts').classList.contains('active')) {
    renderAccounts();
  }
}

// ================= ADMIN USER MANAGEMENT OPERATIONS =================

async function fetchAccounts() {
  if (store.role !== 'admin') return;
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .order('email');
    if (error) throw error;
    store.accounts = data || [];
    renderAccounts();
  } catch (err) {
    console.error('Error fetching profiles:', err.message);
  }
}

function renderAccounts() {
  const b = document.getElementById('accountsBody');
  if (!b) return;
  b.innerHTML = '';
  
  store.accounts.forEach(a => {
    const isSelf = a.id === store.user.id;
    const roleSelect = isSelf 
      ? `<span class="rolebadge ${a.role}">${a.role}</span>`
      : `<select class="role-select" data-id="${a.id}" style="width:auto; padding: 5px 10px; border-radius: var(--radius-sm);">
           <option value="staff" ${a.role === 'staff' ? 'selected' : ''}>Staff</option>
           <option value="admin" ${a.role === 'admin' ? 'selected' : ''}>Admin</option>
         </select>`;
         
    const deleteBtn = isSelf 
      ? `<span class="muted">(Current Account)</span>`
      : `<button class="btn danger sm" onclick="deleteAccount('${a.id}')">Delete</button>`;
      
    const passBtn = isSelf 
      ? '' 
      : `<button class="btn ghost sm" onclick="openResetPassDlg('${a.id}', '${a.email}')">Reset Password</button>`;

    b.insertAdjacentHTML('beforeend', `<tr>
      <td><b>${a.email}</b></td>
      <td>${roleSelect}</td>
      <td>
        <div style="display:flex; gap:8px; align-items:center;">
          ${passBtn}
          ${deleteBtn}
        </div>
      </td>
    </tr>`);
  });

  // Bind change event to role selectors
  document.querySelectorAll('.role-select').forEach(sel => {
    sel.addEventListener('change', async (e) => {
      const userId = e.target.dataset.id;
      const newRole = e.target.value;
      if (!confirm(`Are you sure you want to change this user's role to ${newRole.toUpperCase()}?`)) {
        e.target.value = newRole === 'admin' ? 'staff' : 'admin'; // revert
        return;
      }
      try {
        const { error } = await supabase
          .from('profiles')
          .update({ role: newRole })
          .eq('id', userId);
        if (error) throw error;
        alert('User role updated successfully!');
        await fetchAccounts();
      } catch (err) {
        alert('Failed to update role: ' + err.message);
        await fetchAccounts(); // reload to sync
      }
    });
  });
}

window.deleteAccount = async function(id) {
  if (!requireAdmin()) return;
  if (!confirm('Are you sure you want to permanently delete this account? They will lose access immediately.')) return;
  try {
    const { error } = await supabase.rpc('admin_delete_user', { target_user_id: id });
    if (error) throw error;
    alert('User account deleted successfully!');
    await fetchAccounts();
  } catch (err) {
    alert('Failed to delete account: ' + err.message);
  }
};

window.openResetPassDlg = function(id, email) {
  if (!requireAdmin()) return;
  document.getElementById('resetUserId').value = id;
  document.getElementById('resetUserEmail').value = email;
  document.getElementById('adminNewPass').value = '';
  document.getElementById('adminPassErr').textContent = '';
  document.getElementById('adminPassDlg').showModal();
};

async function resetUserPassword() {
  if (!requireAdmin()) return;
  const id = document.getElementById('resetUserId').value;
  const newPass = document.getElementById('adminNewPass').value;
  const errDiv = document.getElementById('adminPassErr');
  errDiv.textContent = '';

  if (newPass.length < 6) {
    errDiv.textContent = 'Password must be at least 6 characters.';
    return;
  }

  const btn = document.getElementById('btnConfirmAdminChangePass');
  btn.disabled = true;

  try {
    const { error } = await supabase.rpc('admin_update_password', { 
      target_user_id: id, 
      new_password: newPass 
    });
    if (error) throw error;
    alert('User password updated successfully!');
    document.getElementById('adminPassDlg').close();
  } catch (err) {
    errDiv.textContent = err.message;
  } finally {
    btn.disabled = false;
  }
}
