#!/usr/bin/env node

/**
 * eBay User Token Generator
 * 
 * This script handles the eBay OAuth authorization flow:
 * 1. Generates a consent URL for the user to authorize the application
 * 2. Exchanges the authorization code for access and refresh tokens
 * 3. Tests the token validity with the eBay API
 * 4. Saves tokens and expiry information to .env file
 */

require('dotenv').config();
const https = require('https');
const fs = require('fs');
const path = require('path');
const { Buffer } = require('buffer');

const {
  EBAY_CLIENT_ID,
  EBAY_CLIENT_SECRET,
  EBAY_ENVIRONMENT = 'sandbox',
  EBAY_RUNAME, // Using RuName instead of REDIRECT_URI
  EBAY_AUTH_CODE,
  EBAY_CUSTOM_STATE = 'auth-state',
  EBAY_SCOPES,
  EBAY_SANDBOX_AUTH_URL = 'https://auth.sandbox.ebay.com/oauth2/authorize',
  EBAY_PRODUCTION_AUTH_URL = 'https://auth.ebay.com/oauth2/authorize'
} = process.env;

/**
 * Generate a consent URL for the user to authorize the application
 * @param {string} clientId - eBay Client ID
 * @param {string} environment - eBay environment (sandbox or production)
 * @returns {string} - Consent URL
 */
function generateConsentUrl(clientId, environment) {
  // Use URLs from .env file if available
  const authUrl = environment === 'sandbox' 
    ? EBAY_SANDBOX_AUTH_URL 
    : EBAY_PRODUCTION_AUTH_URL;
  
  // Use scopes from .env if available, otherwise use defaults
  const scopes = EBAY_SCOPES || [
    'https://api.ebay.com/oauth/api_scope',
    'https://api.ebay.com/oauth/api_scope/sell.inventory.readonly',
    'https://api.ebay.com/oauth/api_scope/sell.inventory'
  ].join(' ');
  
  // Use RuName instead of traditional redirect_uri
  const params = new URLSearchParams({
    client_id: clientId,
    response_type: 'code',
    redirect_uri: EBAY_RUNAME, // Using RuName here instead of redirect URI
    scope: scopes,
    state: EBAY_CUSTOM_STATE
  });
  
  return `${authUrl}?${params.toString()}`;
}

/**
 * Exchange authorization code for access and refresh tokens
 * @param {string} clientId - eBay Client ID
 * @param {string} clientSecret - eBay Client Secret
 * @param {string} authCode - Authorization code from eBay
 * @param {string} redirectUri - Redirect URI registered with eBay
 * @param {string} environment - eBay environment (sandbox or production)
 * @returns {Promise<Object>} - Token response data
 */
async function exchangeCodeForToken(clientId, clientSecret, authCode, ruName, environment) {
  // Use API URLs from environment variables if available
  const baseUrl = environment === 'sandbox' 
    ? process.env.EBAY_SANDBOX_API_URL || 'api.sandbox.ebay.com'
    : process.env.EBAY_PRODUCTION_API_URL || 'api.ebay.com';
  
  const tokenPath = process.env.EBAY_OAUTH_TOKEN_URL || '/identity/v1/oauth2/token';
  
  const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
  
  const postData = new URLSearchParams({
    grant_type: 'authorization_code',
    code: authCode,
    redirect_uri: ruName  // Use RuName for token exchange
  }).toString();
  
  const options = {
    hostname: baseUrl.replace(/^https?:\/\//, ''), // Remove protocol if present
    port: 443,
    path: tokenPath,
    method: 'POST',
    headers: {
      'Authorization': `Basic ${credentials}`,
      'Content-Type': 'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(postData)
    }
  };

  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          const responseData = JSON.parse(data);
          if (res.statusCode === 200) {
            resolve(responseData);
          } else {
            reject(new Error(`HTTP ${res.statusCode}: ${responseData.error_description || data}`));
          }
        } catch (parseError) {
          reject(new Error(`Failed to parse response: ${data}`));
        }
      });
    });
    
    req.on('error', reject);
    req.write(postData);
    req.end();
  });
}

/**
 * Update the .env file with token information
 * @param {Object} tokenData - Token response data from eBay
 * @returns {Promise<void>}
 */
async function updateEnvFile(tokenData) {
  const envPath = path.resolve(process.cwd(), '.env');
  
  try {
    // Read current .env file
    let envContent = fs.readFileSync(envPath, 'utf8');
    
    // Calculate expiry timestamp
    const expiryTimestamp = Math.floor(Date.now() / 1000) + tokenData.expires_in;
    
    // Update access token
    envContent = envContent.replace(
      /EBAY_ACCESS_TOKEN=.*/,
      `EBAY_ACCESS_TOKEN=${tokenData.access_token}`
    );
    
    // Update expiry timestamp
    envContent = envContent.replace(
      /EBAY_TOKEN_EXPIRY=.*/,
      `EBAY_TOKEN_EXPIRY=${expiryTimestamp}`
    );
    
    // Update refresh token if provided
    if (tokenData.refresh_token) {
      if (envContent.includes('EBAY_REFRESH_TOKEN=')) {
        envContent = envContent.replace(
          /EBAY_REFRESH_TOKEN=.*/,
          `EBAY_REFRESH_TOKEN=${tokenData.refresh_token}`
        );
      } else {
        // Add refresh token if it doesn't exist in the file
        envContent += `\nEBAY_REFRESH_TOKEN=${tokenData.refresh_token}`;
      }
    }
    
    // Write back to .env file
    fs.writeFileSync(envPath, envContent);
    
    console.log('Updated .env file with token information');
    
    // Force reload of environment variables
    require('dotenv').config();
    
  } catch (error) {
    console.error('Error updating .env file:', error.message);
    throw error;
  }
}

/**
 * Test the access token by making a simple API call
 * @param {string} accessToken - Access token to test
 * @param {string} environment - eBay environment (sandbox or production)
 * @returns {Promise<Object>} - API response
 */
async function testUserToken(accessToken, environment) {
  // Use API URLs from environment variables if available
  const baseUrl = environment === 'sandbox' 
    ? process.env.EBAY_SANDBOX_API_URL || 'api.sandbox.ebay.com'
    : process.env.EBAY_PRODUCTION_API_URL || 'api.ebay.com';
  
  const options = {
    hostname: baseUrl.replace(/^https?:\/\//, ''), // Remove protocol if present
    port: 443,
    path: '/sell/inventory/v1/inventory_item?limit=1',
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  };

  return new Promise((resolve, reject) => {
    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', (chunk) => data += chunk);
      res.on('end', () => {
        try {
          const responseData = data ? JSON.parse(data) : {};
          resolve({
            status: res.statusCode,
            data: responseData
          });
        } catch (parseError) {
          resolve({
            status: res.statusCode,
            data: data
          });
        }
      });
    });
    
    req.on('error', reject);
    req.end();
  });
}

/**
 * Main function that runs the eBay OAuth flow
 */
async function main() {
  // Validate required environment variables
  if (!EBAY_CLIENT_ID || !EBAY_CLIENT_SECRET) {
    console.error('Error: EBAY_CLIENT_ID and EBAY_CLIENT_SECRET environment variables required');
    console.error('Please update your .env file with your eBay API credentials');
    process.exit(1);
  }
  
  // Step 1: If no auth code, generate consent URL
  if (!EBAY_AUTH_CODE) {
    const consentUrl = generateConsentUrl(EBAY_CLIENT_ID, EBAY_ENVIRONMENT);
    console.log('\n=== eBay OAuth - Step 1: User Authorization ===');
    console.log('\nSetup required in eBay Developer Console:');
    console.log('');
    console.log('1. Go to this URL in your browser to authorize the application:');
    console.log('\n' + consentUrl);
    console.log('');
    console.log('2. After authorizing, you will be redirected to your redirect URI');
    console.log('3. Copy the "code" parameter from the redirect URL');
    console.log('4. Update your .env file with:');
    console.log('   EBAY_AUTH_CODE=<the code from the URL>');
    console.log('');
    console.log('5. Run this script again to complete the OAuth flow');
    return;
  }

  // Prevent common mistake - don't paste the entire URL as the auth code
  if (EBAY_AUTH_CODE.startsWith('http')) {
    console.error('Error: EBAY_AUTH_CODE should not be the entire URL');
    console.error('Please extract only the "code" parameter from the redirect URL');
    console.error('Example: If redirected to https://localhost:3000/callback?code=ABC123&state=auth-state');
    console.error('Set EBAY_AUTH_CODE=ABC123 in your .env file');
    process.exit(1);
  }
  
  console.log('\n=== eBay OAuth - Step 2: Token Exchange ===');
  
  try {
    // Step 2: Exchange auth code for tokens
    console.log('Exchanging authorization code for tokens...');
    const tokenData = await exchangeCodeForToken(
      EBAY_CLIENT_ID, 
      EBAY_CLIENT_SECRET, 
      EBAY_AUTH_CODE, 
      EBAY_RUNAME, 
      EBAY_ENVIRONMENT
    );
    
    // Step 3: Save tokens to .env file
    await updateEnvFile(tokenData);
    
    console.log('\n=== Token Information ===');
    console.log(`Token Type: ${tokenData.token_type}`);
    console.log(`Expires In: ${tokenData.expires_in} seconds`);
    console.log(`Scope: ${tokenData.scope}`);
    console.log('');
    console.log('Access Token (first 20 chars):');
    console.log(`${tokenData.access_token.substring(0, 20)}...`);
    
    if (tokenData.refresh_token) {
      console.log('');
      console.log('Refresh Token (first 20 chars):');
      console.log(`${tokenData.refresh_token.substring(0, 20)}...`);
      console.log('\nToken will auto-refresh when expired using ebay-refresh-token.js');
    }
    
    // Step 4: Test the token
    console.log('\n=== Testing Token ===');
    console.log('Making test API call to eBay...');
    const testResult = await testUserToken(tokenData.access_token, EBAY_ENVIRONMENT);
    
    if (testResult.status === 200 || testResult.status === 204) {
      console.log('\nToken test: SUCCESS');
      console.log('\nYou can now use ebay-listings-exporter.js to export your listings');
    } else {
      console.log('\nToken test failed');
      console.log(`Status: ${testResult.status}`);
      console.log('Response:');
      console.log(JSON.stringify(testResult.data, null, 2));
      console.log('\nCheck your eBay developer account for any permission issues');
    }
    
  } catch (error) {
    console.error('\nToken exchange failed:', error.message);
    
    // Provide helpful error messages
    if (error.message.includes('invalid_grant')) {
      console.error('\nThis error often occurs when:');
      console.error('1. The authorization code has expired (valid for ~5 minutes)');
      console.error('2. The code has already been used');
      console.error('3. The redirect URI does not match what was registered with eBay');
      console.error('\nPlease start the process again by:');
      console.error('1. Clearing EBAY_AUTH_CODE in your .env file');
      console.error('2. Running this script again to get a new consent URL');
    }
    
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}