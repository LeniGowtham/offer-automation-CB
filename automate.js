import XLSX from 'xlsx';

// ─── CONFIG ──────────────────────────────────────────────────────────────────
const CONFIG = {
  baseUrl:         'https://north-america.api.capillarytech.com',
  authorization:   'Basic ZGV2LWFkbWluOjMxZjg1ZjkwMDI3ZGY1MTljNmM2NjkwMzQ2NWE4Yzhm',
  brandId:         27,
  excelFile:       'Offer details.xlsx',
  offerIdSheet:    'Offer ID',
  attributesSheet: 'Attributes to be updated',
  contentSheet:    'Content',
};

const REQUIRED_LOCALES = {
  'en':     'English',
  'en-qc':  'English QC',
  'en-roc': 'English ROC',
  'fr-qc':  'French QC',
  'fr-roc': 'French ROC',
};

const CLONE_SOURCE_PRIORITY = ['en-roc', 'en-qc', 'en'];

const TARGET_LOCALES  = ['en', 'en-qc', 'en-roc', 'fr-qc', 'fr-roc'];
const ENGLISH_CONTENT = ['en', 'en-qc', 'en-roc'];
const FRENCH_CONTENT  = ['fr-qc', 'fr-roc'];

const IMAGE_NAME_MAP = {
  'merchantimage': 'MERCHANT_IMAGE',
  'small image':   'OFFER_IMAGE_SMALL',
  'medium image':  'OFFER_IMAGE_MEDIUM',
  'large image':   'OFFER_IMAGE_LARGE',
};

const RICH_CONTENT_ATTRS = new Set(['keyTerms']);

// ─── API HELPERS ─────────────────────────────────────────────────────────────
const baseHeaders = {
  Authorization:  CONFIG.authorization,
  'Content-Type': 'application/json',
  Accept:         'application/json',
};

async function apiGet(path) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { headers: baseHeaders });
  const body = await res.json();
  if (!res.ok || body.status?.success === false)
    throw new Error(`GET ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  return body;
}

async function apiPost(path, data) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { method: 'POST', headers: baseHeaders, body: JSON.stringify(data) });
  const body = await res.json();
  if (!res.ok || body.status?.success === false)
    throw new Error(`POST ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  return body;
}

async function apiPut(path, data) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { method: 'PUT', headers: baseHeaders, body: JSON.stringify(data) });
  const body = await res.json();
  if (!res.ok || body.status?.success === false)
    throw new Error(`PUT ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  return body;
}

// ─── EXCEL READERS ────────────────────────────────────────────────────────────
function readOfferIds() {
  const wb   = XLSX.readFile(CONFIG.excelFile);
  const ws   = wb.Sheets[CONFIG.offerIdSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.offerIdSheet}" not found in ${CONFIG.excelFile}`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return rows.slice(1).map(r => r[0]).filter(v => v != null);
}

function readAttributes() {
  const wb = XLSX.readFile(CONFIG.excelFile);
  const ws = wb.Sheets[CONFIG.attributesSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.attributesSheet}" not found`);
  return XLSX.utils.sheet_to_json(ws, { header: 1 })
    .slice(1)
    .filter(r => r[0] != null && r[1] != null)
    .reduce((acc, [attr, val]) => {
      const strVal   = val.toString().trim();
      const key      = attr.toString().trim();
      const lowerKey = key.toLowerCase();
      const storeKey = lowerKey in IMAGE_NAME_MAP ? lowerKey : key;
      acc[storeKey]  = strVal.toLowerCase() === 'null' ? '' : strVal;
      return acc;
    }, {});
}

function readContent() {
  const wb = XLSX.readFile(CONFIG.excelFile);
  const ws = wb.Sheets[CONFIG.contentSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.contentSheet}" not found`);
  return XLSX.utils.sheet_to_json(ws, { header: 1 })
    .slice(1)
    .filter(r => r[0] != null)
    .reduce((acc, [field, english, french]) => {
      acc[field.toString().toLowerCase().trim()] = { english: english ?? '', french: french ?? '' };
      return acc;
    }, {});
}

// ─── PHASE 1: LOCALE MANAGEMENT ──────────────────────────────────────────────
async function ensureLanguageMetadata() {
  console.log('[PREFLIGHT] Checking language metadata...');
  const body = await apiGet('/api_gateway/rewards/core/v1/metadata/languages');

  const registered = new Set((body.languageList ?? []).map(l => l.code.toLowerCase()));
  const missing    = Object.keys(REQUIRED_LOCALES).filter(c => !registered.has(c.toLowerCase()));

  if (missing.length === 0) {
    console.log(`[PREFLIGHT] ✓ All 5 locales registered: ${Object.keys(REQUIRED_LOCALES).join(', ')}\n`);
    return;
  }

  console.log(`[PREFLIGHT] Missing from org metadata: ${missing.join(', ')}`);
  for (const code of missing) {
    try {
      await apiPost('/api_gateway/rewards/core/v1/metadata/language/create', { code, name: REQUIRED_LOCALES[code] });
      console.log(`[PREFLIGHT] ✓ Created language: ${code}`);
    } catch (err) {
      console.log(`[PREFLIGHT] ⚠ Could not create ${code}: ${err.message}`);
    }
  }
  console.log();
}

async function phase1EnsureLocales(id) {
  console.log(`  ── Phase 1: Locale Management ──`);

  let reward;
  try {
    const body = await apiGet(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`);
    reward     = body.reward;
    const name = reward.languageSpecificInfo?.find(l => l.languageCode === 'en')?.name
              ?? reward.languageSpecificInfo?.[0]?.name ?? '(no name)';
    console.log(`  [GET] ✓ "${name}"`);
  } catch (err) {
    console.log(`  [GET] ✗ ${err.message}`);
    return { phase1: 'FAILED', error: err.message };
  }

  const localeMap = Object.fromEntries(
    (reward.languageSpecificInfo ?? []).map(l => [l.languageCode, l])
  );
  const missing = Object.keys(REQUIRED_LOCALES).filter(l => !localeMap[l]);

  console.log(`  [GET] Present: ${Object.keys(localeMap).join(', ') || 'none'}`);

  if (missing.length === 0) {
    console.log(`  [SKIP] All 5 locales already present`);
    return { phase1: 'OK' };
  }

  console.log(`  [GET] Missing: ${missing.join(', ')} → will be cloned`);

  const cloneSource = CLONE_SOURCE_PRIORITY.find(c => localeMap[c]);
  if (!cloneSource) {
    const err = 'No clone source locale found — reward has no usable locale data';
    console.log(`  [BUILD] ✗ ${err}`);
    return { phase1: 'FAILED', error: err };
  }

  const cloneLog = [];
  const languageSpecificInfo = Object.keys(REQUIRED_LOCALES).map(code => {
    if (localeMap[code]) {
      const { rewardId: _r, ...rest } = localeMap[code];
      return rest;
    }
    const { rewardId: _r, languageCode: _lc, ...rest } = localeMap[cloneSource];
    cloneLog.push(`${code} ← ${cloneSource}`);
    return { ...rest, languageCode: code };
  });

  console.log(`  [CLONE] ${cloneLog.join(', ')}`);

  try {
    await apiPut(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`, { languageSpecificInfo });
  } catch (err) {
    console.log(`  [PUT] ✗ ${err.message}`);
    return { phase1: 'FAILED', error: err.message };
  }

  try {
    const verified      = await apiGet(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`);
    const verifiedCodes = (verified.reward.languageSpecificInfo ?? []).map(l => l.languageCode);
    const stillMissing  = Object.keys(REQUIRED_LOCALES).filter(l => !verifiedCodes.includes(l));

    if (stillMissing.length === 0) {
      console.log(`  [VERIFY] ✓ All 5 locales confirmed`);
      return { phase1: 'OK', cloneLog };
    } else {
      const err = `Missing after PUT: ${stillMissing.join(', ')}`;
      console.log(`  [VERIFY] ✗ ${err}`);
      return { phase1: 'FAILED', error: err };
    }
  } catch (err) {
    console.log(`  [VERIFY] ✗ ${err.message}`);
    return { phase1: 'FAILED', error: err.message };
  }
}

// ─── PHASE 2: ATTRIBUTE & CONTENT UPDATE ─────────────────────────────────────
function applyReplacements(text, vendorName, grossCashbackValue, replacePercent = true) {
  if (!text) return text;
  let out = text.replace(/"Vendorname"/gi, vendorName);
  if (replacePercent) out = out.replace(/"cashbackValue"/gi, `${grossCashbackValue}%`);
  return out;
}

function applyUpdates(locale, attrs, content, vendorName, grossCashbackValue, richContentKeys = new Set()) {
  const code       = locale.languageCode;
  const useEnglish = ENGLISH_CONTENT.includes(code);
  const useFrench  = FRENCH_CONTENT.includes(code);
  const lang       = useEnglish ? 'english' : 'french';

  const { rewardId: _r, ...base }          = locale;
  const { images: _imgOrig, ...baseNoImg } = base;

  let images = _imgOrig ?? [];
  if (TARGET_LOCALES.includes(code)) {
    images = images.map(img => {
      const attrKey = Object.keys(IMAGE_NAME_MAP).find(k => IMAGE_NAME_MAP[k] === img.name);
      if (!attrKey || !(attrKey in attrs)) return img;
      const val = attrs[attrKey];
      if (val.startsWith('http')) {
        const { id: _id, ...rest } = img;
        return { ...rest, url: val, isExternal: true };
      } else {
        const { url: _url, ...rest } = img;
        return { ...rest, id: val, isExternal: false };
      }
    });
  }

  let name         = base.name;
  let customFields = { ...base.customFields };
  let richContent  = { ...base.richContentRO };

  // Apply non-image attrs — route to richContentRO if key appears in any locale's richContentRO, else customFields
  if (TARGET_LOCALES.includes(code)) {
    for (const [key, val] of Object.entries(attrs)) {
      if (key.toLowerCase() in IMAGE_NAME_MAP) continue;
      if (RICH_CONTENT_ATTRS.has(key) || richContentKeys.has(key)) {
        richContent[key] = { ...(richContent[key] ?? { isEnabled: true }), content: val };
      } else {
        customFields[key] = val;
      }
    }
    // Always write vendorName (from GET) and grossCashbackValue to all target locales
    if (vendorName)         customFields.vendorName         = vendorName;
    if (grossCashbackValue) customFields.grossCashbackValue = grossCashbackValue;
  }

  if (useEnglish || useFrench) {
    name = applyReplacements(content['name']?.[lang] ?? base.name, vendorName, grossCashbackValue);
    customFields.shortTitle = applyReplacements(
      content['shorttitle']?.[lang] ?? customFields.shortTitle,
      vendorName, grossCashbackValue
    );
    richContent = {
      ...richContent,
      description: {
        content:   applyReplacements(content['description']?.[lang] ?? richContent?.description?.content, vendorName, grossCashbackValue, true),
        isEnabled: richContent?.description?.isEnabled ?? true,
      },
      termsAndConditions: {
        content:   applyReplacements(content['termsandconditions']?.[lang] ?? richContent?.termsAndConditions?.content, vendorName, grossCashbackValue, false),
        isEnabled: richContent?.termsAndConditions?.isEnabled ?? true,
      },
    };
  }

  return { ...baseNoImg, name, images, customFields, richContentRO: richContent };
}

async function phase2UpdateContent(id, attrs, content) {
  console.log(`  ── Phase 2: Attributes & Content ──`);

  let reward;
  try {
    const body = await apiGet(`/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`);
    reward     = body.reward;
  } catch (err) {
    console.log(`  [GET] ✗ ${err.message}`);
    return { phase2: 'FAILED', error: err.message };
  }

  const localeWithData = ['en-qc', 'en-roc', 'fr-qc', 'fr-roc', 'en']
    .map(code => reward.languageSpecificInfo?.find(l => l.languageCode === code))
    .find(l => l?.customFields?.vendorName);

  const vendorName         = localeWithData?.customFields?.vendorName         ?? reward.customFields?.vendorName         ?? '';
  const grossCashbackValue = attrs.grossCashbackValue ?? localeWithData?.customFields?.grossCashbackValue ?? reward.customFields?.grossCashbackValue ?? '2';
  console.log(`  [GET] ✓ vendorName="${vendorName}"  grossCashbackValue="${grossCashbackValue}"`);

  const richContentKeys = new Set(
    (reward.languageSpecificInfo ?? []).flatMap(l => Object.keys(l.richContentRO ?? {}))
  );

  const languageSpecificInfo = (reward.languageSpecificInfo ?? []).map(locale =>
    applyUpdates(locale, attrs, content, vendorName, grossCashbackValue, richContentKeys)
  );

  try {
    const result = await apiPut(
      `/api_gateway/rewards/core/v1/reward/${id}/brand/${CONFIG.brandId}`,
      { languageSpecificInfo }
    );
    console.log(`  [PUT] ✓ ${result.status?.message ?? 'Updated successfully'}`);
    return { phase2: 'OK' };
  } catch (err) {
    console.log(`  [PUT] ✗ ${err.message}`);
    return { phase2: 'FAILED', error: err.message };
  }
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  console.log('\n=== Capillary Offer Automation (Stage) ===');
  console.log(`Host     : ${CONFIG.baseUrl}`);
  console.log(`Brand ID : ${CONFIG.brandId}`);
  console.log(`Locales  : ${Object.keys(REQUIRED_LOCALES).join(', ')}`);
  console.log(`Excel    : ${CONFIG.excelFile}\n`);

  await ensureLanguageMetadata();

  const offerIds = readOfferIds();
  if (offerIds.length === 0) throw new Error(`No offer IDs found in sheet "${CONFIG.offerIdSheet}"`);
  console.log(`Found ${offerIds.length} offer ID(s): ${offerIds.join(', ')}\n`);

  const attrs   = readAttributes();
  const content = readContent();
  console.log(`[READ] Attributes : ${Object.keys(attrs).join(', ') || '(none)'}`);
  console.log(`[READ] Content    : ${Object.keys(content).join(', ') || '(none)'}\n`);

  const summary = [];
  for (const id of offerIds) {
    console.log(`══════════════════════════════`);
    console.log(`Offer ID: ${id}`);

    const p1 = await phase1EnsureLocales(id);

    let p2 = { phase2: 'SKIPPED' };
    if (p1.phase1 === 'OK') {
      p2 = await phase2UpdateContent(id, attrs, content);
    } else {
      console.log(`  [SKIP] Phase 2 skipped — Phase 1 failed`);
    }

    summary.push({ id, ...p1, ...p2 });
    console.log();
  }

  console.log('══════════════════════════════');
  console.log('SUMMARY');
  console.log('══════════════════════════════');
  for (const s of summary) {
    const ok    = s.phase1 === 'OK' && s.phase2 === 'OK';
    const icon  = ok ? '✓' : '✗';
    const clone = s.cloneLog?.length ? `  cloned:${s.cloneLog.join(', ')}` : '';
    const err   = s.error ? `  → ${s.error}` : '';
    console.log(`${icon}  ID ${s.id}  P1:${s.phase1}  P2:${s.phase2}${clone}${err}`);
  }
  const passed = summary.filter(s => s.phase2 === 'OK').length;
  const failed = summary.length - passed;
  console.log(`\n  ${passed} updated, ${failed} failed\n`);
}

main().catch(err => {
  console.error(`\n✗ ${err.message}`);
  process.exit(1);
});
