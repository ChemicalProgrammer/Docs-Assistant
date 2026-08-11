const GEMINI_KEY_PROPERTY = 'GEMINI_API_KEY';

function saveGeminiApiKey(apiKey) {
  apiKey = String(apiKey || '').trim();
  if (!apiKey) throw new Error('Enter a Gemini API key.');
  PropertiesService.getUserProperties().setProperty(GEMINI_KEY_PROPERTY, apiKey);
  return true;
}

function hasGeminiApiKey() {
  return !!PropertiesService.getUserProperties().getProperty(GEMINI_KEY_PROPERTY);
}

function deleteGeminiApiKey() {
  PropertiesService.getUserProperties().deleteProperty(GEMINI_KEY_PROPERTY);
  return true;
}

function getMaskedGeminiApiKey() {
  const key = PropertiesService.getUserProperties().getProperty(GEMINI_KEY_PROPERTY);
  if (!key) return '';
  if (key.length <= 8) return '••••••••';
  return key.slice(0, 4) + '••••••••' + key.slice(-4);
}

function getGeminiApiKey_() {
  const key = PropertiesService.getUserProperties().getProperty(GEMINI_KEY_PROPERTY);
  if (!key) throw new Error('Gemini API key is not configured. Open Settings.');
  return key;
}
