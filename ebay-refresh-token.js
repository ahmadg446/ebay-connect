#!/usr/bin/env node

/**
 * eBay Token Refresh Script
 * 
 * This script checks if the stored access token is expired and
 * automatically refreshes it using the refresh token.
 * 
 * Usage:
 * - Run directly: node ebay-refresh-token.js
 * - Import as module: const { refreshTokenIfNeeded } = require('./ebay-refresh-token');
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
  EBAY_REFRESH_TOKEN,
  EBAY_TOKEN_EXPIRY
} = process.env;

/**
 * Check if the current token is expired or will expire soon
 * @param {number} expiryTimestamp - Token expiry timestamp
 * @param {number} bufferSeconds - Buffer time in seconds before expiry (default: 300 = 5 minutes)
 * @returns {boolean} - True if token is expired or will expire soon
 */
function isTokenExpired(expiryTimestamp, bufferSeconds = 300) {
  if (!expiryTimestamp) return true;
  
  const currentTime = Math.floor(Date.now() / 1000);
  return currentTime + bufferSeconds > parseInt(expiryTimestamp, 10);
}

/**
 * Refresh the access token using the refresh token
 * @param {string} clientId - eBay client ID
 * @param {string} clientSecret - eBay client secret
 * @param {string} refreshToken - Refresh token
 * @param {string} environment - eBay environment (sandbox or production)
 * @returns {Promise<Object>} - Token response data
 */
async function refreshAccessToken(clientId, clientSecret, refreshToken, environment) {
  const baseUrl = environment === 'sandbox' 
    ? 'api.sandbox.ebay.com' 
    : 'api.ebay.com';
  
  const credentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
  
  const postData = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken
  }).toString();
  
  const options = {
    hostname: baseUrl,
    port: 443,
    path: '/identity/v1/oauth2/token',
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
 * Update the .env file with new token information
 * @param {string} accessToken - New access token
 * @param {string} refreshToken - New refresh token (if provided)
 * @param {number} expiresIn - Token expiration in seconds
 * @returns {Promise<void>}
 */
async function updateEnvFile(accessToken, refreshToken, expiresIn) {
  const envPath = path.resolve(process.cwd(), '.env');
  
  try {
    // Read current .env file
    let envContent = fs.readFileSync(envPath, 'utf8');
    
    // Calculate expiry timestamp
    const expiryTimestamp = Math.floor(Date.now() / 1000) + expiresIn;
    
    // Update access token
    envContent = envContent.replace(
      /EBAY_ACCESS_TOKEN=.*/,
      `EBAY_ACCESS_TOKEN=${accessToken}`
    );
    
    // Update expiry timestamp
    envContent = envContent.replace(
      /EBAY_TOKEN_EXPIRY=.*/,
      `EBAY_TOKEN_EXPIRY=${expiryTimestamp}`
    );
    
    // Update refresh token if provided
    if (refreshToken) {
      envContent = envContent.replace(
        /EBAY_REFRESH_TOKEN=.*/,
        `EBAY_REFRESH_TOKEN=${refreshToken}`
      );
    }
    
    // Write back to .env file
    fs.writeFileSync(envPath, envContent);
    
    console.log('Updated .env file with new token information');
    
    // Force reload of environment variables
    require('dotenv').config();
    
  } catch (error) {
    console.error('Error updating .env file:', error.message);
    throw error;
  }
}

/**
 * Check if token needs refresh and refresh it if needed
 * @returns {Promise<{accessToken: string, refreshed: boolean}>}
 */
async function refreshTokenIfNeeded() {
  // Check if we have a refresh token
  if (!EBAY_REFRESH_TOKEN) {
    console.error('Error: No refresh token found in .env file');
    console.error('Run ebay-user-token.js first to get initial tokens');
    return { accessToken: process.env.EBAY_ACCESS_TOKEN, refreshed: false };
  }
  
  // Check if token is expired
  if (!isTokenExpired(EBAY_TOKEN_EXPIRY)) {
    console.log('Token is still valid, no refresh needed');
    return { accessToken: process.env.EBAY_ACCESS_TOKEN, refreshed: false };
  }
  
  console.log('Token expired or will expire soon, refreshing...');
  
  try {
    // Refresh the token
    const tokenData = await refreshAccessToken(
      EBAY_CLIENT_ID,
      EBAY_CLIENT_SECRET,
      EBAY_REFRESH_TOKEN,
      EBAY_ENVIRONMENT
    );
    
    console.log('Token refreshed successfully');
    console.log(`New token expires in: ${tokenData.expires_in} seconds`);
    
    // Update the .env file
    await updateEnvFile(
      tokenData.access_token,
      tokenData.refresh_token,
      tokenData.expires_in
    );
    
    return { accessToken: tokenData.access_token, refreshed: true };
    
  } catch (error) {
    console.error('Token refresh failed:', error.message);
    console.error('You may need to re-authorize using ebay-user-token.js');
    throw error;
  }
}

// Run if executed directly
async function main() {
  if (!EBAY_CLIENT_ID || !EBAY_CLIENT_SECRET) {
    console.error('Error: EBAY_CLIENT_ID and EBAY_CLIENT_SECRET environment variables required');
    process.exit(1);
  }
  
  try {
    const { accessToken, refreshed } = await refreshTokenIfNeeded();
    
    if (refreshed) {
      console.log('Access token has been refreshed and saved');
    } else {
      console.log('Using existing valid access token');
    }
    
    // Just show token info
    console.log('\nAccess token (first few characters):');
    console.log(`${accessToken.substring(0, 20)}...`);
    console.log(`Expires at: ${new Date(parseInt(process.env.EBAY_TOKEN_EXPIRY, 10) * 1000).toLocaleString()}`);
    
  } catch (error) {
    console.error('Error refreshing token:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

// Export for use as a module
module.exports = {
  refreshTokenIfNeeded,
  isTokenExpired
};
