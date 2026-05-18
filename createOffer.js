import XLSX from 'xlsx';

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const CONFIG = {
  baseUrl:       'https://north-america.api.capillarytech.com',
  authorization: 'Basic ZGV2LWFkbWluOjMxZjg1ZjkwMDI3ZGY1MTljNmM2NjkwMzQ2NWE4Yzhm',
  brandId:       27,
  excelFile:     'Offer details.xlsx',
  offerIdSheet:  'Offer ID',
};

// All 5 required locales with their display names (used when creating missing metadata)
const REQUIRED_LOCALES = {
  'en':     'English',
  'en-qc':  'English QC',
  'en-roc': 'English ROC',
  'fr-qc':  'French QC',
  'fr-roc': 'French ROC',
};

// Priority order for choosing a clone source when a locale is missing on a reward
const CLONE_SOURCE_PRIORITY = ['en-roc', 'en-qc', 'en'];

// ─── API HELPERS ─────────────────────────────────────────────────────────────
const baseHeaders = {
  Authorization:  CONFIG.authorization,
  'Content-Type': 'application/json',
  Accept:         'application/json',
};

async function apiGet(path) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { headers: baseHeaders });
  const body = await res.json();
  if (!res.ok || body.status?.success === false) {
    throw new Error(`GET ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  }
  return body;
}

async function apiPost(path, data) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { method: 'POST', headers: baseHeaders, body: JSON.stringify(data) });
  const body = await res.json();
  if (!res.ok || body.status?.success === false) {
    throw new Error(`POST ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  }
  return body;
}

async function apiPut(path, data) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { method: 'PUT', headers: baseHeaders, body: JSON.stringify(data) });
  const body = await res.json();
  if (!res.ok || body.status?.success === false) {
    throw new Error(`PUT ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  }
  return body;
}

// ─── EXCEL READER ────────────────────────────────────────────────────────────
function readOfferIds() {
  const wb   = XLSX.readFile(CONFIG.excelFile);
  const ws   = wb.Sheets[CONFIG.offerIdSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.offerIdSheet}" not found in ${CONFIG.excelFile}`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).map(r => r[0]).filter(v => v != null);
}

// ─── STEP 1: PREFLIGHT — ensure all 5 locales exist in org metadata ──────────
async function ensureLanguageMetadata() {
  console.log('[PREFLIGHT] Checking language metadata...');
  const body = await apiGet('/api_gateway/rewards/core/v1/metadata/languages');

  const registeredCodes = new Set(
    (body.languageList ?? []).map(l => l.code.toLowerCase())
  );

  const missing = Object.keys(REQUIRED_LOCALES).filter(
    code => !registeredCodes.has(code.toLowerCase())
  );

  if (missing.length === 0) {
    console.log(`[PREFLIGHT] ✓ All 5 locales registered: ${Object.keys(REQUIRED_LOCALES).join(', ')}\n`);
    return;
  }

  console.log(`[PREFLIGHT] Missing from org metadata: ${missing.join(', ')}`);

  for (const code of missing) {
    const name = REQUIRED_LOCALES[code];
    try {
      await apiPost('/api_gateway/rewards/core/v1/metadata/language/create', { code, name });
      console.log(`[PREFLIGHT] ✓ Created language: ${code} (${name})`);
    } catch (err) {
      // Language might already exist under a different case — not fatal
      console.log(`[PREFLIGHT] ⚠ Could not create ${code}: ${err.message}`);
    }
  }
  console.log();
}

// ─── STEP 2: BUILD PUT PAYLOAD ────────────────────────────────────────────────
// Existing locales → kept from GET response
// Missing locales  → cloned from first available in CLONE_SOURCE_PRIORITY
function buildPutPayload(localeMap) {
  const cloneSource = CLONE_SOURCE_PRIORITY.find(c => localeMap[c]);
  if (!cloneSource) throw new Error('No clone source locale found — reward has no usable locale data');

  const cloneLog = [];

  const languageSpecificInfo = Object.keys(REQUIRED_LOCALES).map(code => {
    if (localeMap[code]) {
      const { rewardId: _rid, ...rest } = localeMap[code];
      return rest;
    }
    const { rewardId: _rid, languageCode: _lc, ...rest } = localeMap[cloneSource];
    cloneLog.push(`${code} ← ${cloneSource}`);
    return { ...rest, languageCode: code };
  });

  return { payload: { languageSpecificInfo }, cloneLog };
}

// ─── STEP 3: GET + PUT EACH OFFER ─────────────────────────────────────────────
async function processOffer(id) {
  console.log(`──────────────────────────────`);
  console.log(`Offer ID: ${id}`);

  // GET current reward
  let reward;
  try {
    const body = await apiGet(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`);
    reward = body.reward;
    const name = reward.languageSpecificInfo?.find(l => l.languageCode === 'en')?.name
              ?? reward.languageSpecificInfo?.[0]?.name ?? '(no name)';
    console.log(`  [GET] ✓ "${name}" | type=${reward.type} | enabled=${reward.enabled}`);
  } catch (err) {
    console.log(`  [GET] ✗ ${err.message}`);
    return { id, getStatus: 'FAILED', putStatus: 'SKIPPED', error: err.message };
  }

  // Map locales, report what's present / missing
  const localeMap = Object.fromEntries(
    (reward.languageSpecificInfo ?? []).map(l => [l.languageCode, l])
  );
  const existing = Object.keys(localeMap);
  const missing  = Object.keys(REQUIRED_LOCALES).filter(l => !localeMap[l]);

  console.log(`  [GET] Present : ${existing.join(', ') || 'none'}`);
  if (missing.length > 0) {
    console.log(`  [GET] Missing : ${missing.join(', ')} → will be cloned`);
  }

  // Build payload
  let payload, cloneLog;
  try {
    ({ payload, cloneLog } = buildPutPayload(localeMap));
  } catch (err) {
    console.log(`  [BUILD] ✗ ${err.message}`);
    return { id, getStatus: 'OK', putStatus: 'SKIPPED', error: err.message };
  }

  if (cloneLog.length > 0) {
    console.log(`  [CLONE] ${cloneLog.join(', ')}`);
  }

  // PUT update
  console.log(`  [PUT] Updating all 5 locales...`);
  try {
    const result = await apiPut(
      `/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`,
      payload
    );
    console.log(`  [PUT] ✓ ${result.status?.message ?? 'Updated successfully'}`);
  } catch (err) {
    console.log(`  [PUT] ✗ ${err.message}`);
    return { id, getStatus: 'OK', putStatus: 'FAILED', error: err.message };
  }

  // VERIFY — GET again and confirm all 5 locales are present
  console.log(`  [VERIFY] Re-fetching to confirm...`);
  try {
    const verified = await apiGet(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`);
    const verifiedLocales = (verified.reward.languageSpecificInfo ?? []).map(l => l.languageCode);
    const requiredList    = Object.keys(REQUIRED_LOCALES);
    const stillMissing    = requiredList.filter(l => !verifiedLocales.includes(l));

    if (stillMissing.length === 0) {
      console.log(`  [VERIFY] ✓ All 5 locales confirmed: ${verifiedLocales.join(', ')}`);
      return { id, getStatus: 'OK', putStatus: 'OK', verifyStatus: 'PASS', cloneLog };
    } else {
      console.log(`  [VERIFY] ✗ Still missing after update: ${stillMissing.join(', ')}`);
      return { id, getStatus: 'OK', putStatus: 'OK', verifyStatus: 'FAIL', cloneLog, error: `Missing after PUT: ${stillMissing.join(', ')}` };
    }
  } catch (err) {
    console.log(`  [VERIFY] ✗ ${err.message}`);
    return { id, getStatus: 'OK', putStatus: 'OK', verifyStatus: 'ERROR', cloneLog, error: err.message };
  }
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  console.log('\n=== Capillary Offer Update Automation (Stage) ===');
  console.log(`Host     : ${CONFIG.baseUrl}`);
  console.log(`Brand ID : ${CONFIG.brandId}`);
  console.log(`Locales  : ${Object.keys(REQUIRED_LOCALES).join(', ')}`);
  console.log(`Excel    : ${CONFIG.excelFile}\n`);

  // Preflight: register missing languages at org level
  await ensureLanguageMetadata();

  const offerIds = readOfferIds();
  if (offerIds.length === 0) throw new Error(`No offer IDs found in sheet "${CONFIG.offerIdSheet}"`);
  console.log(`Found ${offerIds.length} offer ID(s): ${offerIds.join(', ')}\n`);

  const summary = [];
  for (const id of offerIds) {
    const result = await processOffer(id);
    summary.push(result);
    console.log();
  }

  // Final summary
  console.log('══════════════════════════════');
  console.log('SUMMARY');
  console.log('══════════════════════════════');
  for (const s of summary) {
    const icon   = s.putStatus === 'OK' && s.verifyStatus === 'PASS' ? '✓' : '✗';
    const verify = s.verifyStatus ? `  VERIFY:${s.verifyStatus}` : '';
    const clone  = s.cloneLog?.length ? `  cloned: ${s.cloneLog.join(', ')}` : '';
    const err    = s.error ? `  → ${s.error}` : '';
    console.log(`${icon}  ID ${s.id}  GET:${s.getStatus}  PUT:${s.putStatus}${verify}${clone}${err}`);
  }
  const passed = summary.filter(s => s.putStatus === 'OK').length;
  const failed = summary.filter(s => s.putStatus !== 'OK').length;
  console.log(`\n  ${passed} updated, ${failed} failed\n`);
}

main().catch((err) => {
  console.error(`\n✗ ${err.message}`);
  process.exit(1);
});
