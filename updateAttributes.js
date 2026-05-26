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

// Locales that receive image attribute updates
const TARGET_LOCALES  = ['en', 'en-qc', 'en-roc', 'fr-qc', 'fr-roc'];
// Content language routing
const ENGLISH_CONTENT = ['en', 'en-qc', 'en-roc'];
const FRENCH_CONTENT  = ['fr-qc', 'fr-roc'];

// Excel attribute name → API image name
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

async function apiPut(path, data) {
  const res  = await fetch(`${CONFIG.baseUrl}${path}`, { method: 'PUT', headers: baseHeaders, body: JSON.stringify(data) });
  const body = await res.json();
  if (!res.ok || body.status?.success === false)
    throw new Error(`PUT ${path} → [${body.status?.code ?? res.status}] ${body.status?.message ?? res.statusText}`);
  return body;
}

// ─── EXCEL READERS ────────────────────────────────────────────────────────────
function readOfferId() {
  const wb = XLSX.readFile(CONFIG.excelFile);
  const ws = wb.Sheets[CONFIG.offerIdSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.offerIdSheet}" not found`);
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const id = rows[1]?.[0];
  if (id == null) throw new Error(`No offer ID found in sheet "${CONFIG.offerIdSheet}"`);
  return id;
}

function readAttributes() {
  const wb = XLSX.readFile(CONFIG.excelFile);
  const ws = wb.Sheets[CONFIG.attributesSheet];
  if (!ws) throw new Error(`Sheet "${CONFIG.attributesSheet}" not found`);
  return XLSX.utils.sheet_to_json(ws, { header: 1 })
    .slice(1)
    .filter(r => r[0] != null && r[1] != null)
    .reduce((acc, [attr, val]) => {
      const strVal = val.toString().trim();
      // Image keys are lowercased for IMAGE_NAME_MAP matching; all others keep original case
      const key        = attr.toString().trim();
      const lowerKey   = key.toLowerCase();
      const storeKey   = lowerKey in IMAGE_NAME_MAP ? lowerKey : key;
      acc[storeKey] = strVal.toLowerCase() === 'null' ? '' : strVal;
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

// ─── REPLACEMENT HELPERS ──────────────────────────────────────────────────────
// replacePercent: true for name/shortTitle/description, false for termsAndConditions
function applyReplacements(text, vendorName, grossCashbackValue, replacePercent = true) {
  if (!text) return text;
  let out = text.replace(/"Vendorname"/gi, vendorName);
  if (replacePercent) out = out.replace(/"cashbackValue"/gi, `${grossCashbackValue}%`);
  return out;
}

// ─── PER-LOCALE UPDATE ────────────────────────────────────────────────────────
function applyUpdates(locale, attrs, content, vendorName, grossCashbackValue, richContentKeys = new Set()) {
  const code        = locale.languageCode;
  const isTarget    = TARGET_LOCALES.includes(code);
  const useEnglish  = ENGLISH_CONTENT.includes(code);
  const useFrench   = FRENCH_CONTENT.includes(code);
  const useContent  = useEnglish || useFrench;
  const lang        = useEnglish ? 'english' : 'french';

  const { rewardId: _rid, ...base } = locale;

  // Step 1 — image attributes (4 target locales only)
  const { images: _imgOrig, ...baseNoImages } = base;
  let images = _imgOrig ?? [];
  if (isTarget) {
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

  // Step 2 — content (en-qc/en-roc get English, fr-qc/fr-roc get French)
  let name         = base.name;
  let customFields = { ...base.customFields };
  let richContent  = { ...base.richContentRO };

  // Apply non-image attrs — route to richContentRO if key appears in any locale's richContentRO, else customFields
  if (isTarget) {
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

  if (useContent) {
    name = applyReplacements(content['name']?.[lang]           ?? base.name,                              vendorName, grossCashbackValue);
    customFields.shortTitle = applyReplacements(content['shorttitle']?.[lang] ?? customFields.shortTitle, vendorName, grossCashbackValue);

    richContent = {
      ...richContent,
      description: {
        content:   applyReplacements(content['description']?.[lang]       ?? richContent?.description?.content,       vendorName, grossCashbackValue, true),
        isEnabled: richContent?.description?.isEnabled ?? true,
      },
      termsAndConditions: {
        content:   applyReplacements(content['termsandconditions']?.[lang] ?? richContent?.termsAndConditions?.content, vendorName, grossCashbackValue, false),
        isEnabled: richContent?.termsAndConditions?.isEnabled ?? true,
      },
    };
  }

  return { ...baseNoImages, name, images, customFields, richContentRO: richContent };
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
async function main() {
  const offerId = readOfferId();
  console.log('\n=== Offer Attribute & Content Update ===');
  console.log(`Offer ID : ${offerId}  |  Brand : ${CONFIG.brandId}\n`);

  const attrs  = readAttributes();
  const content = readContent();

  console.log('[READ] Attributes to be updated:');
  for (const [k, v] of Object.entries(attrs)) console.log(`  ${k} → ${v}`);
  console.log(`[READ] Content fields: ${Object.keys(content).join(', ')}\n`);

  // GET current reward
  console.log('[GET] Fetching current reward...');
  const { reward } = await apiGet(`/api_gateway/rewards/core/v1/reward/${offerId}/brand/${CONFIG.brandId}`);

  // vendorName and grossCashbackValue live in locale-level customFields, not top-level.
  // Find the first locale that has them (en-qc / en-roc are most reliable sources).
  const localeWithData = ['en-qc', 'en-roc', 'fr-qc', 'fr-roc', 'en']
    .map(code => reward.languageSpecificInfo?.find(l => l.languageCode === code))
    .find(l => l?.customFields?.vendorName);

  const vendorName         = localeWithData?.customFields?.vendorName         ?? reward.customFields?.vendorName         ?? '';
  const grossCashbackValue = attrs.grossCashbackValue ?? localeWithData?.customFields?.grossCashbackValue ?? reward.customFields?.grossCashbackValue ?? '2';
  console.log(`[GET] ✓  vendorName="${vendorName}"  grossCashbackValue="${grossCashbackValue}"`);
  console.log(`[GET]    Locales present: ${reward.languageSpecificInfo?.map(l => l.languageCode).join(', ')}\n`);

  // Collect all richContentRO field names across all locales for correct routing
  const richContentKeys = new Set(
    (reward.languageSpecificInfo ?? []).flatMap(l => Object.keys(l.richContentRO ?? {}))
  );

  // Apply updates to each locale
  const languageSpecificInfo = (reward.languageSpecificInfo ?? []).map(locale =>
    applyUpdates(locale, attrs, content, vendorName, grossCashbackValue, richContentKeys)
  );

  const changed = languageSpecificInfo.map(l => l.languageCode).filter(c =>
    TARGET_LOCALES.includes(c) || ENGLISH_CONTENT.includes(c) || FRENCH_CONTENT.includes(c)
  );
  console.log(`[PUT] Applying changes to: ${[...new Set(changed)].join(', ')}`);
  console.log(`[PUT] Sending all ${languageSpecificInfo.length} locales in payload...\n`);

  const result = await apiPut(
    `/api_gateway/rewards/core/v1/reward/${offerId}/brand/${CONFIG.brandId}`,
    { languageSpecificInfo }
  );
  console.log(`[PUT] ✓ ${result.status?.message ?? 'Updated successfully'}`);
}

main().catch(err => {
  console.error(`\n✗ ${err.message}`);
  process.exit(1);
});
