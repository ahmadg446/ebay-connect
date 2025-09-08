#!/usr/bin/env node

/**
 * eBay Discount Manager - 48 Hour Expiry Monitor
 * 
 * This script monitors eBay Marketing API discounts that are ending within 48 hours.
 * Uses OAuth 2.0 + Marketing API to fetch and display expiring promotions.
 */

require('dotenv').config();
const https = require('https');
const readline = require('readline');

// ==================== CONFIGURATION ====================
const CONFIG = {
  MARKETING_API: {
    SANDBOX_URL: 'api.sandbox.ebay.com',
    PRODUCTION_URL: 'api.ebay.com',
    VERSION: 'v1',
    REQUESTS_PER_SECOND: 5  // Marketing API is more generous
  },
  
  ALERT_WINDOW_HOURS: 48  // Alert for discounts ending within 48 hours
};

// ==================== LOGGER ====================
class Logger {
  constructor() {
    this.startTime = Date.now();
    this.requestCount = 0;
    this.errorCount = 0;
  }

  log(level, message, data = null) {
    const timestamp = new Date().toISOString();
    const elapsed = ((Date.now() - this.startTime) / 1000).toFixed(1);
    
    console.log(`[${timestamp}] [${elapsed}s] [${level.padEnd(5)}] ${message}`);
    
    if (data && level === 'ERROR') {
      console.log(JSON.stringify(data, null, 2));
    }
  }

  info(message, data = null) { this.log('INFO', message, data); }
  warn(message, data = null) { this.log('WARN', message, data); }
  error(message, data = null) { 
    this.errorCount++;
    this.log('ERROR', message, data); 
  }
  
  incrementRequest() { this.requestCount++; }
  
  getStats() {
    return {
      totalRequests: this.requestCount,
      totalErrors: this.errorCount,
      elapsedSeconds: ((Date.now() - this.startTime) / 1000).toFixed(1)
    };
  }
}

// ==================== RATE LIMITER ====================
class RateLimiter {
  constructor() {
    this.requestTimes = [];
    this.requestQueue = [];
    this.processing = false;
  }

  async waitForSlot() {
    return new Promise((resolve) => {
      this.requestQueue.push(resolve);
      this.processQueue();
    });
  }

  processQueue() {
    if (this.processing || this.requestQueue.length === 0) return;
    
    this.processing = true;
    
    const now = Date.now();
    const minInterval = 1000 / CONFIG.MARKETING_API.REQUESTS_PER_SECOND;
    
    // Clean old requests (older than 1 second)
    this.requestTimes = this.requestTimes.filter(time => 
      now - time < 1000
    );
    
    // Check if we can make a request
    if (this.requestTimes.length < CONFIG.MARKETING_API.REQUESTS_PER_SECOND) {
      const resolve = this.requestQueue.shift();
      this.requestTimes.push(now);
      this.processing = false;
      
      resolve();
      
      // Process next request after minimum interval
      if (this.requestQueue.length > 0) {
        setTimeout(() => this.processQueue(), minInterval);
      }
    } else {
      // Wait and try again
      setTimeout(() => {
        this.processing = false;
        this.processQueue();
      }, minInterval);
    }
  }
}

// ==================== MARKETING API CLIENT ====================
class MarketingApiClient {
  constructor(logger, rateLimiter) {
    this.logger = logger;
    this.rateLimiter = rateLimiter;
    this.accessToken = process.env.EBAY_ACCESS_TOKEN;
    this.environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
  }

  async executeRequest(endpoint, method, body) {
    const hostname = this.environment === 'sandbox' 
      ? CONFIG.MARKETING_API.SANDBOX_URL 
      : CONFIG.MARKETING_API.PRODUCTION_URL;

    const options = {
      hostname: hostname,
      port: 443,
      path: `/sell/marketing/${CONFIG.MARKETING_API.VERSION}${endpoint}`,
      method: method,
      headers: {
        'Authorization': `Bearer ${this.accessToken}`,
        'Content-Type': 'application/json',
        'Accept': 'application/json'
      }
    };

    if (body) {
      const bodyString = JSON.stringify(body);
      options.headers['Content-Length'] = Buffer.byteLength(bodyString);
    }

    this.logger.info(`Making API request: ${method} ${options.path}`);
    if (body) {
      this.logger.info(`Request body:`, body);
    }

    return new Promise((resolve, reject) => {
      const req = https.request(options, (res) => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => {
          try {
            const responseData = data ? JSON.parse(data) : {};
            this.logger.info(`Response status: ${res.statusCode}`);
            
            // Show error details for failed requests
            if (res.statusCode >= 400) {
              this.logger.info(`Response error: ${data}`);
            } else {
              this.logger.info(`Response data:`, responseData);
            }

            if (res.statusCode >= 200 && res.statusCode < 300) {
              resolve(responseData);
            } else {
              // Create more detailed error message
              let errorMessage = `HTTP ${res.statusCode}`;
              if (responseData.errors) {
                errorMessage += ': ' + responseData.errors.map(e => e.message).join('; ');
              } else if (data) {
                errorMessage += `: ${data}`;
              }
              reject(new Error(errorMessage));
            }
          } catch (parseError) {
            this.logger.error(`Failed to parse response: ${data}`);
            reject(new Error(`Failed to parse response: ${data}`));
          }
        });
      });

      req.on('error', (error) => {
        this.logger.error(`Request error: ${error.message}`);
        reject(error);
      });

      if (body) {
        req.write(JSON.stringify(body));
      }

      req.end();
    });
  }

  async makeRequest(endpoint, method = 'GET', body = null, attempt = 1) {
    // Wait for rate limit slot
    await this.rateLimiter.waitForSlot();

    try {
      this.logger.incrementRequest();
      this.logger.info(`Attempt ${attempt}: Sending request to ${endpoint}`);

      const result = await this.executeRequest(endpoint, method, body);

      if (attempt > 1) {
        this.logger.info(`Request succeeded on attempt ${attempt}: ${endpoint}`);
      }

      return result;

    } catch (error) {
      this.logger.error(`Request failed on attempt ${attempt}: ${error.message}`);

      if (attempt < 3 && this.isRetryableError(error)) {
        const waitTime = Math.pow(2, attempt) * 1000;
        this.logger.warn(`Retrying in ${waitTime}ms...`);

        await this.sleep(waitTime);
        return this.makeRequest(endpoint, method, body, attempt + 1);
      } else {
        this.logger.error(`Request failed permanently: ${endpoint}`, error);
        throw error;
      }
    }
  }

  isRetryableError(error) {
    // Retry on network errors and some HTTP errors
    return error.message.includes('ECONNRESET') ||
           error.message.includes('ETIMEDOUT') ||
           error.message.includes('HTTP 429') ||
           error.message.includes('HTTP 5');
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
  
  async getItemPromotion(promotionId) {
    const endpoint = `/item_promotion/${promotionId}`;
    return this.makeRequest(endpoint, 'GET');
  }
  
  async getItemPriceMarkdownPromotion(promotionId) {
    const endpoint = `/item_price_markdown/${promotionId}`;
    return this.makeRequest(endpoint, 'GET');
  }
  
  async updateItemPromotion(promotionId, promotionData) {
    const endpoint = `/item_promotion/${promotionId}`;
    return this.makeRequest(endpoint, 'PUT', promotionData);
  }
  
  async updateItemPriceMarkdownPromotion(promotionId, promotionData) {
    const endpoint = `/item_price_markdown/${promotionId}`;
    return this.makeRequest(endpoint, 'PUT', promotionData);
  }
}

// ==================== DISCOUNT FETCHER ====================
class DiscountFetcher {
  constructor(apiClient, logger) {
    this.apiClient = apiClient;
    this.logger = logger;
  }

  async fetchExpiringDiscounts() {
    this.logger.info('Fetching all active discounts...');
    
    const now = new Date();
    const alertThreshold = new Date(now.getTime() + (CONFIG.ALERT_WINDOW_HOURS * 60 * 60 * 1000));
    
    this.logger.info(`Looking for discounts ending before: ${alertThreshold.toISOString()}`);
    
    // Fetch all discounts using the correct endpoint
    const allDiscounts = await this.fetchAllDiscounts();
    const expiringDiscounts = this.filterExpiringDiscounts(allDiscounts, alertThreshold);
    
    this.logger.info(`Found ${allDiscounts.length} total active discounts, ${expiringDiscounts.length} expiring within ${CONFIG.ALERT_WINDOW_HOURS} hours`);
    
    return expiringDiscounts;
  }

  async fetchAllDiscounts() {
    this.logger.info('Fetching all discounts...');
    
    try {
      let allDiscounts = [];
      let offset = 0;
      const limit = 50; // API limit
      let hasMore = true;
      
      // Set the marketplace ID based on environment
      const marketplaceId = this.apiClient.environment === 'sandbox' ? 'EBAY_AT' : 'EBAY_US';
      
      while (hasMore) {
        // Use the correct eBay API endpoint
        const endpoint = `/promotion?limit=${limit}&offset=${offset}&marketplace_id=${marketplaceId}`;
        const response = await this.apiClient.makeRequest(endpoint);
        
        if (response.promotions) {
          // Add all discount types to the array
          const discounts = response.promotions.map(discount => ({
            ...discount,
            discountType: discount.promotionType || 'Unknown'
          }));
          
          allDiscounts.push(...discounts);
          this.logger.info(`Fetched ${discounts.length} discounts (offset ${offset})`);
          
          // Check if there are more results
          hasMore = discounts.length === limit;
          offset += limit;
        } else {
          hasMore = false;
        }
      }
      
      return allDiscounts;
      
    } catch (error) {
      this.logger.warn(`Error fetching all discounts: ${error.message}`);
      return [];
    }
  }

  filterExpiringDiscounts(discounts, alertThreshold) {
    return discounts.filter(discount => {
      const endDate = new Date(discount.endDate);
      const status = discount.promotionStatus;
      
      // Include RUNNING and SCHEDULED discounts that end within our threshold
      // Also include PAUSED discounts if needed
      return (status === 'RUNNING' || status === 'SCHEDULED' || status === 'PAUSED') && 
             endDate <= alertThreshold;
    }).sort((a, b) => new Date(a.endDate) - new Date(b.endDate)); // Sort by end date
  }
  
  async extendDiscountEndDate(discount) {
    try {
      // Use the full promotion ID including marketplace
      const promotionId = discount.promotionId;
      
      let fullDiscount;
      let updatedDiscount;
      let newEndDate; // Declare newEndDate at function scope
      
      // Handle different discount types
      if (discount.discountType === 'MARKDOWN_SALE') {
        // For markdown discounts, use the markdown-specific endpoint
        fullDiscount = await this.apiClient.getItemPriceMarkdownPromotion(promotionId);
        
        // Calculate new end date (2 weeks from current end date)
        const currentEndDate = new Date(fullDiscount.endDate);
        newEndDate = new Date(currentEndDate.getTime() + (14 * 24 * 60 * 60 * 1000)); // Add 2 weeks
        
        // Create updated discount object with only writable fields
        updatedDiscount = {
          name: fullDiscount.name,
          description: `Auto-generated, ${new Date().toISOString().slice(0, -5)}Z`, // Fixed: 36 chars max
          startDate: fullDiscount.startDate,
          endDate: newEndDate.toISOString(),
          marketplaceId: fullDiscount.marketplaceId,
          promotionType: fullDiscount.promotionType,
          promotionStatus: 'SCHEDULED', // FIXED: Set status to SCHEDULED for markdown sales
          selectedInventoryDiscounts: fullDiscount.selectedInventoryDiscounts,
          inventoryCriterion: fullDiscount.inventoryCriterion
        };
        
        // Add optional fields if they exist
        if (fullDiscount.promotionImageUrl) updatedDiscount.promotionImageUrl = fullDiscount.promotionImageUrl;
        
      } else {
        // For threshold discounts, use the item promotion endpoint
        fullDiscount = await this.apiClient.getItemPromotion(promotionId);
        
        // Calculate new end date (2 weeks from current end date)
        const currentEndDate = new Date(fullDiscount.endDate);
        newEndDate = new Date(currentEndDate.getTime() + (14 * 24 * 60 * 60 * 1000)); // Add 2 weeks
        
        // Create updated discount object with only writable fields
        updatedDiscount = {
          name: fullDiscount.name,
          description: `Auto-generated, ${new Date().toISOString().slice(0, -5)}Z`, // Fixed: 36 chars max
          startDate: fullDiscount.startDate,
          endDate: newEndDate.toISOString(),
          marketplaceId: fullDiscount.marketplaceId,
          promotionType: fullDiscount.promotionType,
          promotionStatus: 'SCHEDULED', // FIXED: Set status to SCHEDULED for volume discounts
          discountRules: fullDiscount.discountRules,
          inventoryCriterion: fullDiscount.inventoryCriterion
        };
        
        // Add optional fields if they exist
        if (fullDiscount.promotionImageUrl) updatedDiscount.promotionImageUrl = fullDiscount.promotionImageUrl;
      }
      
      // Add optional fields if they exist
      if (fullDiscount.priority) updatedDiscount.priority = fullDiscount.priority;
      
      // Update the discount using the appropriate method
      let result;
      if (discount.discountType === 'MARKDOWN_SALE') {
        result = await this.apiClient.updateItemPriceMarkdownPromotion(promotionId, updatedDiscount);
      } else {
        result = await this.apiClient.updateItemPromotion(promotionId, updatedDiscount);
      }
      
      this.logger.info(`Successfully extended discount ${discount.promotionId} to ${newEndDate.toISOString()}`);
      return result;
    } catch (error) {
      this.logger.error(`Failed to extend discount ${discount.promotionId}: ${error.message}`);
      throw error;
    }
  }
}

// ==================== DISCOUNT DISPLAY ====================
class DiscountDisplay {
  constructor(logger) {
    this.logger = logger;
  }

  displayExpiringDiscounts(discounts) {
    if (discounts.length === 0) {
      this.logger.info('🎉 No discounts expiring within the next 48 hours!');
      return;
    }

    console.log('\n' + '='.repeat(100));
    console.log('DISCOUNTS EXPIRING WITHIN 48 HOURS');
    console.log('='.repeat(100));

    discounts.forEach((discount, index) => {
      this.displaySingleDiscount(discount, index + 1);
      if (index < discounts.length - 1) {
        console.log('-'.repeat(100));
      }
    });

    console.log('='.repeat(100));
    console.log(`Total expiring discounts: ${discounts.length}`);
    console.log('='.repeat(100));
  }

  displaySingleDiscount(discount, index) {
    const now = new Date();
    const endDate = new Date(discount.endDate);
    const timeLeft = this.calculateTimeLeft(now, endDate);
    const urgencyLevel = this.getUrgencyLevel(timeLeft.totalHours);

    console.log(`
${urgencyLevel} DISCOUNT #${index}`);
    console.log(`Promotion ID: ${discount.promotionId || 'N/A'}`);
    console.log(`Type: ${discount.discountType}`);
    console.log(`Status: ${discount.promotionStatus}`);
    console.log(`Name: ${discount.name || 'Unnamed Promotion'}`);
    
    if (discount.description) {
      console.log(`Description: ${discount.description}`);
    }

    // Display discount details
    this.displayDiscountDetails(discount);

    console.log(`Start Date: ${new Date(discount.startDate).toLocaleString()}`);
    console.log(`End Date: ${endDate.toLocaleString()}`);
    console.log(`Time Remaining: ${timeLeft.display}`);
    
    if (timeLeft.totalHours <= 12) {
      console.log('URGENT: Less than 12 hours remaining!');
    } else if (timeLeft.totalHours <= 24) {
      console.log('WARNING: Less than 24 hours remaining!');
    }
  }

  displayDiscountDetails(discount) {
    if (discount.discountType === 'Markdown') {
      // Markdown discount details
      if (discount.selectedInventoryDiscounts) {
        discount.selectedInventoryDiscounts.forEach(item => {
          if (item.discountPercentage) {
            console.log(`Discount: ${item.discountPercentage}% off`);
          } else if (item.discountAmount) {
            console.log(`Discount: $${item.discountAmount.value} off`);
          }
        });
      }
    } else if (discount.discountType === 'Threshold') {
      // Threshold discount details
      if (discount.discountBenefit) {
        const benefit = discount.discountBenefit;
        if (benefit.percentageOffOrder) {
          console.log(`Benefit: ${benefit.percentageOffOrder}% off order`);
        } else if (benefit.amountOffOrder) {
          console.log(`Benefit: $${benefit.amountOffOrder.value} off order`);
        } else if (benefit.percentageOffItem) {
          console.log(`Benefit: ${benefit.percentageOffItem}% off items`);
        } else if (benefit.amountOffItem) {
          console.log(`Benefit: $${benefit.amountOffItem.value} off items`);
        }
      }

      if (discount.discountSpecification) {
        const spec = discount.discountSpecification;
        if (spec.minimumPurchaseAmount) {
          console.log(`Minimum Purchase: $${spec.minimumPurchaseAmount.value}`);
        }
        if (spec.minimumQuantity) {
          console.log(`Minimum Quantity: ${spec.minimumQuantity}`);
        }
      }
    }

    // Display item count if available
    if (discount.inventoryCriterion) {
      const criterion = discount.inventoryCriterion;
      if (criterion.inventoryItems) {
        console.log(`Items: ${criterion.inventoryItems.length} specific items`);
      } else if (criterion.listingIds) {
        console.log(`Items: ${criterion.listingIds.length} specific listings`);
      } else if (criterion.ruleCriteria) {
        console.log(`Items: Rule-based selection`);
      }
    }
  }

  calculateTimeLeft(now, endDate) {
    const totalMs = endDate - now;
    const totalHours = totalMs / (1000 * 60 * 60);
    const days = Math.floor(totalHours / 24);
    const hours = Math.floor(totalHours % 24);
    const minutes = Math.floor((totalMs % (1000 * 60 * 60)) / (1000 * 60));

    let display = '';
    if (days > 0) {
      display += `${days}d `;
    }
    if (hours > 0) {
      display += `${hours}h `;
    }
    display += `${minutes}m`;

    return {
      totalHours,
      days,
      hours,
      minutes,
      display: display.trim()
    };
  }

  getUrgencyLevel(hoursLeft) {
    if (hoursLeft <= 6) return '[!!!]';
    if (hoursLeft <= 12) return '[!!] ';
    if (hoursLeft <= 24) return '[!]';
    return '[*]';
  }
}

// ==================== MAIN APPLICATION ====================
class DiscountManager {
  constructor() {
    this.logger = new Logger();
    this.rateLimiter = new RateLimiter();
    this.apiClient = new MarketingApiClient(this.logger, this.rateLimiter);
    this.discountFetcher = new DiscountFetcher(this.apiClient, this.logger);
    this.discountDisplay = new DiscountDisplay(this.logger);
  }

  async validateEnvironment() {
    const requiredFields = [
      'EBAY_ACCESS_TOKEN'
    ];
    
    const missing = requiredFields.filter(key => !process.env[key]);
    
    if (missing.length > 0) {
      throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
    }

    this.logger.info('Environment validation passed');
  }

  async run() {
    this.logger.info('Connecting to eBay API...');
    const startTime = Date.now();
    
    try {
      console.log('='.repeat(80));
      console.log('eBay Discount Manager - 48 Hour Expiry Monitor');
      console.log('='.repeat(80));
      
      await this.validateEnvironment();
      
      const environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
      this.logger.info(`Environment: ${environment}`);
      this.logger.info(`Alert window: ${CONFIG.ALERT_WINDOW_HOURS} hours`);
      
      // Fetch expiring discounts
      const expiringDiscounts = await this.discountFetcher.fetchExpiringDiscounts();
      
      // Display results
      this.discountDisplay.displayExpiringDiscounts(expiringDiscounts);
      
      // Ask user if they want to extend expiring discounts
      if (expiringDiscounts.length > 0) {
        const readline = require('readline');
        const rl = readline.createInterface({
          input: process.stdin,
          output: process.stdout
        });
        
        const answer = await new Promise(resolve => {
          rl.question('\nWould you like to extend all expiring discounts by 2 weeks? (y/n): ', resolve);
        });
        rl.close();
        
        if (answer.toLowerCase() === 'y' || answer.toLowerCase() === 'yes') {
          this.logger.info('Extending expiring discounts...');
          let successCount = 0;
          let errorCount = 0;
          
          for (const discount of expiringDiscounts) {
            try {
              await this.discountFetcher.extendDiscountEndDate(discount);
              successCount++;
            } catch (error) {
              errorCount++;
              // Log different messages based on error type
              if (error.message.includes('404')) {
                this.logger.warn(`Discount ${discount.promotionId} not found - may have been deleted`);
              } else {
                this.logger.error(`Failed to extend discount ${discount.promotionId}: ${error.message}`);
              }
            }
          }
          
          this.logger.info(`Finished extending discounts: ${successCount} successful, ${errorCount} errors`);
        }
      }
      
      // Final summary
      const stats = this.logger.getStats();
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      this.logger.info('='.repeat(80));
      this.logger.info('DISCOUNT MONITORING COMPLETED');
      this.logger.info('='.repeat(80));
      this.logger.info(`Duration: ${duration} seconds`);
      this.logger.info(`Total API requests: ${stats.totalRequests}`);
      this.logger.info(`Expiring discounts found: ${expiringDiscounts.length}`);
      this.logger.info('='.repeat(80));
      
      return {
        success: true,
        expiringDiscounts: expiringDiscounts.length,
        stats: {
          duration: parseFloat(duration),
          apiRequests: stats.totalRequests,
          errors: stats.totalErrors
        }
      };
      
    } catch (error) {
      const stats = this.logger.getStats();
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      this.logger.error('='.repeat(80));
      this.logger.error('DISCOUNT MONITORING FAILED');
      this.logger.error('='.repeat(80));
      this.logger.error(`Error: ${error.message}`);
      this.logger.error(`Duration: ${duration} seconds`);
      this.logger.error(`API requests made: ${stats.totalRequests}`);
      this.logger.error('='.repeat(80));
      
      throw error;
    }
  }
}

// ==================== ENTRY POINT ====================
async function main() {
  if (require.main === module) {
    const manager = new DiscountManager();
    
    try {
      await manager.run();
      process.exit(0);
    } catch (error) {
      console.error('\nFatal error:', error.message);
      
      // Provide helpful error messages
      if (error.message.includes('access token') || error.message.includes('401')) {
        console.error('\nToken issues detected. Try:');
        console.error('1. Run ebay-refresh-token.js to refresh your token');
        console.error('2. Or run ebay-user-token.js to get a new token');
      }
      
      process.exit(1);
    }
  }
}

// Export for use as module
module.exports = DiscountManager;

// Run if called directly
main();
