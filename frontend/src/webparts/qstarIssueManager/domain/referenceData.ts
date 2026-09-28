/** Canonical Region choices used by the production form and connection diagnostics. */
export const REGIONS = [
  "Americas (Miami)", "Asia Pacific", "China (Shanghai)", "Eastern Europe (Vienna)",
  "Head Office (Neu-Isenburg)", "Western Europe (Amsterdam)",
];

const LEGACY_REGIONS: Record<string, string> = {
  Germany: "Western Europe (Amsterdam)",
  Americas: "Americas (Miami)",
  China: "China (Shanghai)",
  "Eastern Europe": "Eastern Europe (Vienna)",
  "Head Office": "Head Office (Neu-Isenburg)",
  "Western Europe": "Western Europe (Amsterdam)",
};

export function normalizeRegion(value: string): string {
  return LEGACY_REGIONS[value] || value;
}
