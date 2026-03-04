// officeMapping.js - Office location to address and language mapping
// Ported from browser-extension/lib/officeMapping.js + popup.js

const OFFICE_ADDRESS_MAP = {
  'bergeijk': 'Stokskesweg 9, 5571TJ Bergeijk',
  'braunschweig': 'Berliner Str. 52 j, 38104 Braunschweig, Germany',
  'bremen': 'Im Hollergrund 3, 28357 Bremen, Germany',
  'dänemark': 'Königsberger Str. 115, 47809 Krefeld, Germany',
  'dortmund': 'Rodenbergstraße 47, 44287 Dortmund, Germany',
  'krefeld': 'Königsberger Str. 115, 47809 Krefeld, Germany',
  'schweiz': 'acadon (Schweiz) GmbH, General-Guisan-Str. 6, CH-6300 Zug',
  'österreich': 'acadon GmbH, Am Euro Platz 2, AT-1120 Wien'
};

const DEFAULT_ADDRESS = 'Königsberger Str. 115, 47809 Krefeld, Germany';

const OFFICE_TO_LANG = {
  'krefeld': 'DE',
  'braunschweig': 'DE',
  'bremen': 'DE',
  'dortmund': 'DE',
  'dänemark': 'DK',
  'bergeijk': 'NL',
  'schweiz': 'CH',
  'österreich': 'DE'
};

function resolveAddress(officeLocation) {
  if (!officeLocation) return DEFAULT_ADDRESS;
  const key = officeLocation.toLowerCase().trim();
  return OFFICE_ADDRESS_MAP[key] || DEFAULT_ADDRESS;
}

function resolveLanguage(officeLocation) {
  if (!officeLocation) return 'DE';
  const key = officeLocation.toLowerCase().trim();
  return OFFICE_TO_LANG[key] || 'DE';
}

const LANGUAGE_TO_GREETING = {
  'DE': 'Mit freundlichen Grüßen',
  'EN': 'Best regards',
  'CH': 'Freundliche Grüsse',
  'DK': 'Med venlig hilsen',
  'FR': 'Cordialement',
  'NL': 'Met vriendelijke groeten'
};

function resolveGreeting(language) {
  return LANGUAGE_TO_GREETING[language] || LANGUAGE_TO_GREETING['DE'];
}
