/**
 * Debug and management functions for SaltEdge setup
 * Run these functions from Apps Script console to check/reset stored credentials
 */

/**
 * Check what SaltEdge credentials are currently stored
 */
function checkSaltEdgeCredentials() {
  const props = PropertiesService.getScriptProperties();
  const privateKey = props.getProperty('SALTEDGE_PRIVATE_KEY');
  const appId = props.getProperty('SALTEDGE_APP_ID');
  const secret = props.getProperty('SALTEDGE_SECRET');
  
  console.log('=== SaltEdge Credentials Status ===');
  console.log('Private Key stored:', privateKey ? 'YES (' + privateKey.substring(0, 50) + '...)' : 'NO');
  console.log('App ID stored:', appId || 'NO');
  console.log('Secret stored:', secret ? 'YES (' + secret.substring(0, 10) + '...)' : 'NO');
  
  return {
    hasPrivateKey: !!privateKey,
    hasAppId: !!appId,
    hasSecret: !!secret
  };
}

/**
 * Clear all SaltEdge credentials (use if you want to re-run complete setup)
 */
function clearSaltEdgeCredentials() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SALTEDGE_PRIVATE_KEY');
  props.deleteProperty('SALTEDGE_APP_ID');
  props.deleteProperty('SALTEDGE_SECRET');
  
  console.log('All SaltEdge credentials cleared. Run Setup SaltEdge again.');
  return 'Credentials cleared successfully';
}

/**
 * Force re-setup of private key only
 */
function setupPrivateKeyOnly() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const privateKeyResponse = ui.prompt(
    'SaltEdge Setup - RSA Private Key',
    'Paste your RSA private key (complete PEM format including -----BEGIN/END----- lines):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (privateKeyResponse.getSelectedButton() !== ui.Button.OK) {
    return 'Setup cancelled by user';
  }
  
  const privateKey = privateKeyResponse.getResponseText().trim();
  if (!privateKey.includes('-----BEGIN') || !privateKey.includes('-----END')) {
    throw new Error('Invalid RSA private key format. Must include -----BEGIN/END----- lines.');
  }
  
  props.setProperty('SALTEDGE_PRIVATE_KEY', privateKey);
  console.log('RSA private key updated successfully');
  return 'Private key updated successfully';
}
