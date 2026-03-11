// templateEngine.js - Template placeholder replacement
// Ported from browser-extension/lib/templateEngine.js

const PLACEHOLDER_REGEX = /@@(\w+)@@/g;

const PLACEHOLDER_MAP = {
  'FirstName': 'givenName',
  'LastName': 'surname',
  'JobTitle': 'jobTitle',
  'TelefonNumber': 'phone',
  'MailAddress': 'mail',
  'Address': 'address',
  'CompanyName': 'companyName',
  'WebsiteUrl': 'websiteUrl'
};

function applyPlaceholders(template, userData) {
  return template.replace(PLACEHOLDER_REGEX, (match, key) => {
    const fieldName = PLACEHOLDER_MAP[key];
    if (fieldName && userData[fieldName] != null && userData[fieldName] !== '') {
      return userData[fieldName];
    }
    return match;
  });
}

function fixCharsetMeta(htmlContent) {
  return htmlContent.replace(/charset=windows-1252/gi, 'charset=utf-8');
}
