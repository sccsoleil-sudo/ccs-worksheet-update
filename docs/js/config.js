/** Shared mappings — keep in sync with ccs_tool/config.py */

export const DIV_MAPPING = {
  "02AA": "CPD",
  "02AB": "PPD",
  "02AC": "LPD",
  "02AD": "LDB",
};

export const BRAND_AXE_MAPPING = [
  { acronym: "LOP0", brand: "L'Oreal Paris", axe: "*General*", div: "CPD" },
  { acronym: "LOP1", brand: "L'Oreal Paris", axe: "HAIR", div: "CPD" },
  { acronym: "LOP2", brand: "L'Oreal Paris", axe: "MAKE UP", div: "CPD" },
  { acronym: "LOP3", brand: "L'Oreal Paris", axe: "SKIN CARE", div: "CPD" },
  { acronym: "GAR0", brand: "Garnier", axe: "*General*", div: "CPD" },
  { acronym: "GAR1", brand: "Garnier", axe: "HAIR", div: "CPD" },
  { acronym: "GAR3", brand: "Garnier", axe: "SKIN CARE", div: "CPD" },
  { acronym: "MNY0", brand: "Maybelline", axe: "MAKE UP", div: "CPD" },
  { acronym: "ESS0", brand: "Essie", axe: "MAKE UP", div: "CPD" },
  { acronym: "NYX0", brand: "NYX Prof. Make-up", axe: "MAKE UP", div: "CPD" },
  { acronym: "CDA0", brand: "Carol's Daughter", axe: "*General*", div: "CPD" },
  { acronym: "THA0", brand: "Thayers", axe: "*General*", div: "CPD" },
  { acronym: "VIC0", brand: "Vichy", axe: "SKIN CARE", div: "LDB" },
  { acronym: "LRP0", brand: "La Roche Posay", axe: "SKIN CARE", div: "LDB" },
  { acronym: "SKC0", brand: "Skinceuticals", axe: "SKIN CARE", div: "LDB" },
  { acronym: "CER0", brand: "CeraVe", axe: "SKIN CARE", div: "LDB" },
  { acronym: "SKB0", brand: "Skinbetter Science", axe: "SKIN CARE", div: "LDB" },
  { acronym: "LAN0", brand: "Lancome", axe: "*General*", div: "LPD" },
  { acronym: "LAN2", brand: "Lancome", axe: "MAKE UP", div: "LPD" },
  { acronym: "LAN3", brand: "Lancome", axe: "SKIN CARE", div: "LPD" },
  { acronym: "LAN5", brand: "Lancome", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "BIO0", brand: "Biotherm", axe: "*General*", div: "LPD" },
  { acronym: "RPL0", brand: "Ralph Lauren", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "ARM0", brand: "Armani", axe: "*General*", div: "LPD" },
  { acronym: "ARM2", brand: "Armani", axe: "MAKE UP", div: "LPD" },
  { acronym: "ARM5", brand: "Armani", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "VKR0", brand: "Viktor & Rolf", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "SHU0", brand: "Shu Uemura", axe: "*General*", div: "LPD" },
  { acronym: "KIE0", brand: "Kiehl's", axe: "*General*", div: "LPD" },
  { acronym: "MMM0", brand: "Maison Margiela", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "YSL0", brand: "Yves Saint Laurent", axe: "*General*", div: "LPD" },
  { acronym: "YSL2", brand: "Yves Saint Laurent", axe: "MAKE UP", div: "LPD" },
  { acronym: "YSL5", brand: "Yves Saint Laurent", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "UDY0", brand: "Urban Decay", axe: "*General*", div: "LPD" },
  { acronym: "ITC0", brand: "IT Cosmetics", axe: "*General*", div: "LPD" },
  { acronym: "ITC2", brand: "IT Cosmetics", axe: "MAKE UP", div: "LPD" },
  { acronym: "ITC3", brand: "IT Cosmetics", axe: "SKIN CARE", div: "LPD" },
  { acronym: "VAL0", brand: "Valentino", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "AZZ0", brand: "Azzaro", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "MUG0", brand: "Mugler", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "PRA0", brand: "Prada", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "YTP0", brand: "Youth to the People", axe: "*General*", div: "LPD" },
  { acronym: "FRA5", brand: "Fragrances", axe: "FRAGRANCE", div: "LPD" },
  { acronym: "LOR0", brand: "L'Oreal Professionnel", axe: "HAIR", div: "PPD" },
  { acronym: "KER0", brand: "Kerastase", axe: "HAIR", div: "PPD" },
  { acronym: "MAT0", brand: "Matrix", axe: "HAIR", div: "PPD" },
  { acronym: "RKN0", brand: "Redken", axe: "HAIR", div: "PPD" },
  { acronym: "SHU0", brand: "Shu Uemura Pro", axe: "HAIR", div: "PPD" },
  { acronym: "PUL0", brand: "Pulp Riot", axe: "HAIR", div: "PPD" },
  { acronym: "PUR0", brand: "Pureology", axe: "HAIR", div: "PPD" },
  { acronym: "FRA0", brand: "Fragrances", axe: "FRAGRANCE", div: "PPD" },
];

export const CPD_CUSTOMER_GROUPS = {
  LCL: ["10118351"],
  SDM: ["10118010"],
  AMZ: ["10118507"],
  WMT: ["10118025"],
  PJC: ["10118028"],
  MCKESSON: [
    "10118012", "10118037", "10118039", "10104785", "10118071",
    "10118254", "10118502", "10118048", "10118108", "10118563", "10118188",
  ],
  OTHER: [
    "10112015", "10118567", "10118566", "10118563", "10118032",
    "10118083", "10118017", "10118546", "10118124", "10118094", "10118113",
  ],
};

export const NON_CPD_DIVS = ["LDB", "LPD", "PPD"];
export const ALL_DIVS = ["CPD", "LDB", "LPD", "PPD"];
export const AMOUNT_COL = "Amount (CoCode Crcy)";
export const AMOUNT_ALIASES = ["SUBI $", "Amount (CoCode Crcy)", "Amount"];

export const WORKSHEET_PRESETS = {
  // Example: "1a. new LCL - CPD CCS Deductions Worksheet.xlsm"
  "Advance V2": {
    label: "Advance V2",
    sheets: ["KAMs", "KAMs - To correct", "Pricing", "Credit", "Cleared"],
    appendSheet: "KAMs",
    discAppendSheet: "KAMs",
  },
  // Example: "0. CPD Percentage Deduction - EXR CAR Worksheet.xlsx"
  // Append default = first tab of the uploaded file (resolved at runtime)
  "Version 1": {
    label: "Version 1",
    sheets: ["CPD"],
    appendSheet: null, // first sheet of workbook
    discAppendSheet: null, // first sheet of workbook
  },
};
