// ─────────────────────────────────────────────────────────────
//  Stashd — Excel Custom Functions
//  Registers =STASHD.FETCH(url) which returns a single row
//  of profile values spilled horizontally.
// ─────────────────────────────────────────────────────────────

const ACTOR_ID    = 'LpVuK3Zozwuipa5bp';
const BASE        = 'https://api.apify.com/v2';
const DEFAULT_TOKEN = 'apify_api_fufuf8tg8g';
const STORAGE_KEY = 'stashd_token';

// ── Fields returned (in order) ─────────────────────────────────
const FIELDS = [
  'name', 'linkedinUrl', 'headline', 'about', 'location',
  'currentCompany', 'currentJobTitle', 'currentCompanyStart', 'currentCompanyTenure',
  'connections', 'followers', 'premium', 'openToWork', 'hiring', 'verified',
  'currentEmploymentType', 'currentWorkplaceType', 'currentRoleLocation',
  'prevCompany', 'prevJobTitle', 'totalExperience', 'numRoles',
  'eduSchool', 'eduDegree', 'skills', 'languages',
];

// ── Helpers (mirrors index.html logic) ─────────────────────────
const MONTH_MAP = { Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12 };

function toMonthNum(m) {
  if (!m) return 0;
  return typeof m === 'number' ? m : (MONTH_MAP[String(m).trim()] || 0);
}

function calcTenure(startDate) {
  if (!startDate?.year) return '';
  const m     = toMonthNum(startDate.month) || 1;
  const start = new Date(startDate.year, m - 1, 1);
  const now   = new Date();
  let years  = now.getFullYear() - start.getFullYear();
  let months = now.getMonth()    - start.getMonth();
  if (months < 0) { years--; months += 12; }
  if (years === 0 && months === 0) return '< 1 mo';
  const parts = [];
  if (years  > 0) parts.push(`${years} yr${years  !== 1 ? 's' : ''}`);
  if (months > 0) parts.push(`${months} mo${months !== 1 ? 's' : ''}`);
  return parts.join(' ');
}

function formatDate(d) {
  if (!d?.year) return d?.text || '';
  const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const mNum   = toMonthNum(d.month) || 1;
  return `${MONTHS[mNum - 1]} ${d.year}`;
}

function deriveCurrentRole(p) {
  const exp = p.experience || [];
  if (!exp.length) return { currentCompany:'', currentJobTitle:'', currentCompanyStartDateObj:null, currentCompanyTenure:'' };
  const isCurrent  = e => e.endDate?.text === 'Present' || (!e.endDate?.month && !e.endDate?.year);
  const latestRole = exp.find(isCurrent) || exp[0];
  const companyId  = latestRole.companyId;
  const companyName = latestRole.companyName || '';
  const allAtCo = exp.filter(e => companyId ? e.companyId === companyId : e.companyName === companyName);
  allAtCo.sort((a, b) => {
    const dy = (a.startDate?.year || 0) - (b.startDate?.year || 0);
    return dy !== 0 ? dy : toMonthNum(a.startDate?.month) - toMonthNum(b.startDate?.month);
  });
  const earliest = allAtCo[0];
  return {
    currentCompany:             companyName,
    currentJobTitle:            latestRole.position || '',
    currentCompanyStartDateObj: earliest?.startDate || null,
    currentCompanyTenure:       calcTenure(earliest?.startDate),
  };
}

function profileToRow(p) {
  const name     = [p.firstName, p.lastName].filter(Boolean).join(' ');
  const location = p.location?.parsed?.text || p.location?.linkedinText || '';
  const { currentCompany, currentJobTitle, currentCompanyStartDateObj, currentCompanyTenure } = deriveCurrentRole(p);
  const exp = p.experience || [];
  const isCurrent = e => e.endDate?.text === 'Present' || (!e.endDate?.month && !e.endDate?.year);
  const latestRole  = exp.find(isCurrent) || exp[0] || {};
  const curCompanyId   = latestRole.companyId;
  const curCompanyName = latestRole.companyName;
  const prevRole = exp.find(e => curCompanyId ? e.companyId !== curCompanyId : e.companyName !== curCompanyName) || null;
  const edu0 = p.education?.[0] || {};

  const map = {
    name,
    linkedinUrl:           p.linkedinUrl          || '',
    headline:              p.headline             || '',
    about:                 p.about                || '',
    location,
    currentCompany,
    currentJobTitle,
    currentCompanyStart:   formatDate(currentCompanyStartDateObj),
    currentCompanyTenure,
    connections:           p.connectionsCount != null ? String(p.connectionsCount) : '',
    followers:             p.followerCount    != null ? String(p.followerCount)    : '',
    premium:               p.premium    ? 'Yes' : '',
    openToWork:            p.openToWork ? 'Yes' : '',
    hiring:                p.hiring     ? 'Yes' : '',
    verified:              p.verified   ? 'Yes' : '',
    currentEmploymentType: latestRole.employmentType || '',
    currentWorkplaceType:  latestRole.workplaceType  || '',
    currentRoleLocation:   latestRole.location       || '',
    prevCompany:           prevRole?.companyName      || '',
    prevJobTitle:          prevRole?.position         || '',
    totalExperience:       (() => {
      const expWithStart = exp.filter(e => e.startDate?.year).sort((a, b) => {
        const dy = (a.startDate.year || 0) - (b.startDate.year || 0);
        return dy !== 0 ? dy : toMonthNum(a.startDate.month) - toMonthNum(b.startDate.month);
      });
      return calcTenure(expWithStart[0]?.startDate);
    })(),
    numRoles:              exp.length ? String(exp.length) : '',
    eduSchool:             edu0.schoolName || '',
    eduDegree:             [edu0.degree, edu0.fieldOfStudy].filter(Boolean).join(', '),
    skills:                (p.skills    || []).map(s => s.name).join(', '),
    languages:             (p.languages || []).map(l => l.name + (l.proficiency ? ` (${l.proficiency})` : '')).join(', '),
  };

  return FIELDS.map(f => map[f] ?? '');
}

// ── Apify API ───────────────────────────────────────────────────
async function getToken() {
  try {
    const t = await OfficeRuntime.storage.getItem(STORAGE_KEY);
    return t || DEFAULT_TOKEN;
  } catch {
    return DEFAULT_TOKEN;
  }
}

async function apiFetch(path, options) {
  const res = await fetch(`${BASE}${path}`, options);
  if (!res.ok) { const t = await res.text(); throw new Error(`Apify ${res.status}: ${t}`); }
  return res.json();
}

async function fetchProfile(url) {
  const token = await getToken();

  // Start run
  const { data: run } = await apiFetch(`/acts/${ACTOR_ID}/runs?token=${token}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      profileScraperMode: 'Profile details no email ($4 per 1k)',
      queries: [url],
    }),
  });

  // Poll until done
  const TERMINAL = ['SUCCEEDED', 'FAILED', 'ABORTED', 'TIMED-OUT'];
  for (let i = 0; i < 120; i++) {
    await new Promise(r => setTimeout(r, 3000));
    const { data } = await apiFetch(`/actor-runs/${run.id}?token=${token}`);
    if (data.status === 'SUCCEEDED') {
      const items = await apiFetch(`/datasets/${data.defaultDatasetId}/items?token=${token}`);
      const profile = Array.isArray(items) ? items[0] : items;
      if (!profile) throw new Error('No profile data returned.');
      return profile;
    }
    if (TERMINAL.includes(data.status)) throw new Error(`Run ended: ${data.status}`);
  }
  throw new Error('Timed out.');
}

// ── Custom Function ─────────────────────────────────────────────
/**
 * Fetch a LinkedIn profile and return a row of values.
 * @customfunction
 * @param {string} url LinkedIn profile URL
 * @returns {Promise<string[][]>}
 */
async function FETCH(url) {
  if (!url || !url.includes('linkedin.com/in/')) {
    throw new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      'Please provide a valid LinkedIn profile URL.'
    );
  }
  const profile = await fetchProfile(url.trim());
  return [profileToRow(profile)]; // 1 row × N columns
}

CustomFunctions.associate('FETCH', FETCH);
