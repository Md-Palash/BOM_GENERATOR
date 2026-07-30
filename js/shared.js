/* ============================================================
   SHARED — Home/Brand Navigation + Password Gate
   Used by: index.html (loaded first, before any engine script)
   ============================================================ */
/* ============================================================
   HOME / BRAND NAVIGATION
   ============================================================ */
/* ============================================================
   PASSWORD GATE — separate password per engine
   ============================================================ */
const DECATHLON_PASSWORD = 'Decathlon#2026';
const HADDAD_PASSWORD = 'Haddad#2026';
const MALACCA_PASSWORD = 'MalaccaMalacca';
let unlockedHaddad = false;
let unlockedDecathlon = false;
let unlockedMalacca = false;
let pendingUnlock = null; // { which: 'haddad' | 'decathlon' | 'malacca', onSuccess }

const pwOverlay = document.getElementById('pwOverlay');
const pwInput = document.getElementById('pwInput');
const pwError = document.getElementById('pwError');
const pwTitle = document.getElementById('pwTitle');
const pwSub = document.getElementById('pwSub');
const pwSubmit = document.getElementById('pwSubmit');
const pwCancel = document.getElementById('pwCancel');

function requestUnlock(which, onSuccess) {
  const already = which === 'haddad' ? unlockedHaddad : which === 'decathlon' ? unlockedDecathlon : unlockedMalacca;
  if (already) { onSuccess(); return; }
  pendingUnlock = { which, onSuccess };
  const label = which === 'haddad' ? 'Haddad' : which === 'decathlon' ? 'Decathlon' : 'Malacca';
  pwTitle.textContent = label + ' Tool';
  pwSub.textContent = 'This tool is password protected. Enter the password to continue.';
  pwInput.value = '';
  pwError.textContent = '';
  pwOverlay.hidden = false;
  setTimeout(() => pwInput.focus(), 0);
}

function closePwOverlay() {
  pwOverlay.hidden = true;
  pendingUnlock = null;
}

function tryPwSubmit() {
  if (!pendingUnlock) return;
  const correct = pendingUnlock.which === 'haddad' ? HADDAD_PASSWORD : pendingUnlock.which === 'decathlon' ? DECATHLON_PASSWORD : MALACCA_PASSWORD;
  if (pwInput.value === correct) {
    if (pendingUnlock.which === 'haddad') unlockedHaddad = true;
    else if (pendingUnlock.which === 'decathlon') unlockedDecathlon = true;
    else unlockedMalacca = true;
    const cb = pendingUnlock.onSuccess;
    closePwOverlay();
    cb();
  } else {
    pwError.textContent = 'Incorrect password. Try again.';
    pwInput.value = '';
    pwInput.focus();
  }
}

pwSubmit.addEventListener('click', tryPwSubmit);
pwCancel.addEventListener('click', closePwOverlay);
pwInput.addEventListener('keydown', e => { if (e.key === 'Enter') tryPwSubmit(); });
pwOverlay.addEventListener('click', e => { if (e.target === pwOverlay) closePwOverlay(); });

