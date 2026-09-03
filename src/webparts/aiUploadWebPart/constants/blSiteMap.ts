export const BL_SITE_MAP: { [name: string]: string } = {
  'Architecture': '/sites/edsp_ARC',
  'Building Engineering': '/sites/edsp_BEG',
  'Environment': '/sites/edsp_ENV',
  'Geotechnical': '/sites/edsp_GEO',
  'Digital': '/sites/edsp_ISD',
  'Land Supply and Municipal': '/sites/edsp_LSM',
  'MEP': '/sites/edsp_MEP',
  'Project and Construction Management': '/sites/edsp_PCM',
  'Program, Cost and Consultancy': '/sites/edsp_PCC',
  'Transportation': '/sites/edsp_TRA',
  'Unclassified': '/sites/edsp_UNC',
  'Urbanism and Planning': '/sites/edsp_UAP',
  'Water': '/sites/edsp_WAT'
};

export const LEADING_BL_OPTIONS: string[] = Object.keys(BL_SITE_MAP);

export const LEADING_BL_LABEL: string = 'Leading BL';

export function isLeadingBlField(label: string): boolean {
  return (label || '').trim().toLowerCase() === LEADING_BL_LABEL.toLowerCase();
}

export function canonicalLeadingBl(value: string): string {
  const match = matchLeadingBl(value);
  return match ? match.name : (value || '').trim();
}

export function resolveLeadingBlSite(value: string, tenantUrl: string): {
  name: string;
  siteUrl: string;
} | undefined {
  const match = matchLeadingBl(value);
  if (!match) {
    return undefined;
  }
  return {
    name: match.name,
    siteUrl: `${trimSlash(tenantUrl)}${match.path}`
  };
}

function normalizeKey(value: string): string {
  return (value || '')
    .trim()
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function trimSlash(value: string): string {
  return (value || '').replace(/\/+$/, '');
}

function buildLookup(): { [key: string]: string } {
  const lookup: { [key: string]: string } = {};
  Object.keys(BL_SITE_MAP).forEach((name) => {
    lookup[normalizeKey(name)] = name;
    const path = BL_SITE_MAP[name];
    const code = path.substring(path.lastIndexOf('_') + 1);
    if (code) {
      lookup[normalizeKey(code)] = name;
      lookup[normalizeKey(`edsp_${code}`)] = name;
      lookup[normalizeKey(path)] = name;
    }
  });
  return lookup;
}

const BL_LOOKUP: { [key: string]: string } = buildLookup();

function matchLeadingBl(value: string): { name: string; path: string } | undefined {
  const key = normalizeKey(value);
  if (!key) {
    return undefined;
  }
  const name = BL_LOOKUP[key];
  if (!name) {
    return undefined;
  }
  return {
    name,
    path: BL_SITE_MAP[name]
  };
}
