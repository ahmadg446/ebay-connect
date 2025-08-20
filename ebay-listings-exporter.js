#!/usr/bin/env node

// DEPENDENCIES: npm install xlsx xml2js

require('dotenv').config();
const https = require('https');
const fs = require('fs').promises;
const xml2js = require('xml2js');

let XLSX;
try {
  XLSX = require('xlsx');
} catch (error) {
  console.error('XLSX library not found. Install with: npm install xlsx xml2js');
  process.exit(1);
}

// ==================== CONFIGURATION ====================
const CONFIG = {
  TRADING_API: {
    SANDBOX_URL: 'api.sandbox.ebay.com',
    PRODUCTION_URL: 'api.ebay.com',
    ENDPOINT_PATH: '/ws/api.dll',
    VERSION: '1291', // Current Trading API version
    CALL_RATE_LIMIT: 5000, // calls per day
    REQUESTS_PER_SECOND: 2  // Conservative for Trading API
  },
  
  PAGINATION: {
    DEFAULT_ENTRIES_PER_PAGE: 100,
    MAX_ENTRIES_PER_PAGE: 200
  }
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
    const minInterval = 1000 / CONFIG.TRADING_API.REQUESTS_PER_SECOND;
    
    // Clean old requests (older than 1 second)
    this.requestTimes = this.requestTimes.filter(time => 
      now - time < 1000
    );
    
    // Check if we can make a request
    if (this.requestTimes.length < CONFIG.TRADING_API.REQUESTS_PER_SECOND) {
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

// ==================== TRADING API CLIENT ====================
class TradingApiClient {
  constructor(logger, rateLimiter) {
    this.logger = logger;
    this.rateLimiter = rateLimiter;
    this.appId = process.env.EBAY_APP_ID || process.env.EBAY_CLIENT_ID;
    this.devId = process.env.EBAY_DEV_ID;
    this.certId = process.env.EBAY_CERT_ID;
    // Use OAuth token instead of old Auth'n'Auth token
    this.oauthToken = process.env.EBAY_ACCESS_TOKEN; // Your v^1.1#... token
    this.siteId = process.env.EBAY_SITE_ID || '0'; // 0 = US
    this.environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
    this.parser = new xml2js.Parser({ explicitArray: false });
  }

  async makeRequest(callName, requestBody, attempt = 1) {
    // Wait for rate limit slot
    await this.rateLimiter.waitForSlot();
    
    try {
      this.logger.incrementRequest();
      
      const xmlRequest = this.buildXmlRequest(callName, requestBody);
      const result = await this.executeRequest(callName, xmlRequest);
      
      if (attempt > 1) {
        this.logger.info(`Request succeeded on attempt ${attempt}: ${callName}`);
      }
      
      return result;
      
    } catch (error) {
      if (attempt < 3) {
        const waitTime = Math.pow(2, attempt) * 1000;
        this.logger.warn(`Request failed (attempt ${attempt}/3): ${error.message}. Retrying in ${waitTime}ms...`);
        
        await this.sleep(waitTime);
        return this.makeRequest(callName, requestBody, attempt + 1);
      } else {
        this.logger.error(`Request failed permanently: ${callName}`, error);
        throw error;
      }
    }
  }

  buildXmlRequest(callName, requestBody) {
    // Modern OAuth approach - no RequesterCredentials needed
    return `<?xml version="1.0" encoding="utf-8"?>
<${callName}Request xmlns="urn:ebay:apis:eBLBaseComponents">
  <Version>${CONFIG.TRADING_API.VERSION}</Version>
  ${requestBody}
</${callName}Request>`;
  }

  async executeRequest(callName, xmlRequest) {
    const hostname = this.environment === 'sandbox' 
      ? CONFIG.TRADING_API.SANDBOX_URL 
      : CONFIG.TRADING_API.PRODUCTION_URL;

    const options = {
      hostname: hostname,
      port: 443,
      path: CONFIG.TRADING_API.ENDPOINT_PATH,
      method: 'POST',
      headers: {
        'Content-Type': 'text/xml',
        'Content-Length': Buffer.byteLength(xmlRequest),
        'X-EBAY-API-CALL-NAME': callName,
        'X-EBAY-API-SITEID': this.siteId,
        'X-EBAY-API-APP-NAME': this.appId,
        'X-EBAY-API-VERSION': CONFIG.TRADING_API.VERSION,
        'X-EBAY-API-COMPATIBILITY-LEVEL': CONFIG.TRADING_API.VERSION, 
        'X-EBAY-API-REQUEST-ENCODING': 'XML',
        'X-EBAY-API-IAF-TOKEN': this.oauthToken
      }
    };

    return new Promise((resolve, reject) => {
      const req = https.request(options, (res) => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', async () => {
          try {
            const parsedResponse = await this.parser.parseStringPromise(data);
            const responseRoot = parsedResponse[`${callName}Response`];
            
            if (!responseRoot) {
              throw new Error(`Invalid response format for ${callName}`);
            }

            if (responseRoot.Ack === 'Failure' || responseRoot.Ack === 'PartialFailure') {
              const errors = Array.isArray(responseRoot.Errors) 
                ? responseRoot.Errors 
                : [responseRoot.Errors];
              
              const errorMessages = errors.map(err => 
                `${err.ErrorCode}: ${err.LongMessage || err.ShortMessage}`
              ).join('; ');
              
              throw new Error(`eBay API Error: ${errorMessages}`);
            }

            resolve(responseRoot);
          } catch (parseError) {
            reject(new Error(`Failed to parse response: ${parseError.message}`));
          }
        });
      });

      req.on('error', reject);
      req.write(xmlRequest);
      req.end();
    });
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// ==================== DATA FETCHER ====================
class TradingDataFetcher {
  constructor(apiClient, logger) {
    this.apiClient = apiClient;
    this.logger = logger;
  }

  async fetchAllSellerListings() {
    this.logger.info('Starting complete seller listings fetch using Trading API...');
    
    let allItems = [];
    
    // First, try to get active listings (no date restrictions)
    try {
      this.logger.info('Fetching active listings using GetMyeBaySelling...');
      const myeBayData = await this.fetchMyeBaySellingData();
      
      // Combine all active, sold, and unsold items
      const combinedItems = [
        ...myeBayData.active,
        ...myeBayData.sold,
        ...myeBayData.unsold
      ];
      
      if (combinedItems.length > 0) {
        this.logger.info(`Found ${combinedItems.length} listings via GetMyeBaySelling`);
        allItems = combinedItems;
      }
    } catch (error) {
      this.logger.warn(`GetMyeBaySelling failed: ${error.message}`);
    }
    
    // If we didn't get many results, try historical search with time windows
    if (allItems.length < 10) {
      this.logger.info('Searching historical listings using time windows...');
      const historicalItems = await this.fetchHistoricalListings();
      allItems = [...allItems, ...historicalItems];
    }

    this.logger.info(`Seller listings fetch completed: ${allItems.length} items`);
    return allItems;
  }

  async fetchHistoricalListings() {
    const allItems = [];
    const maxDays = 120; // Stay under 121-day limit
    let currentEndDate = new Date();
    let searchAttempts = 0;
    const maxSearchAttempts = 50; // Prevent infinite loops, covers ~16 years
    
    this.logger.info('Searching historical listings in 120-day windows...');
    
    while (searchAttempts < maxSearchAttempts) {
      const startDate = new Date(currentEndDate);
      startDate.setDate(startDate.getDate() - maxDays);
      
      const startTimeISO = startDate.toISOString();
      const endTimeISO = currentEndDate.toISOString();
      
      this.logger.info(`Searching window ${searchAttempts + 1}: ${startTimeISO.substring(0, 10)} to ${endTimeISO.substring(0, 10)}`);
      
      try {
        const windowItems = await this.fetchListingsInTimeWindow(startTimeISO, endTimeISO);
        
        if (windowItems.length > 0) {
          allItems.push(...windowItems);
          this.logger.info(`Found ${windowItems.length} listings in this window (total: ${allItems.length})`);
        } else {
          this.logger.info('No listings found in this window');
          
          // If we haven't found any listings in the last 3 windows, probably safe to stop
          if (searchAttempts >= 3) {
            this.logger.info('No recent listings found, stopping historical search');
            break;
          }
        }
        
        // Move to the next time window (go further back)
        currentEndDate = new Date(startDate);
        currentEndDate.setDate(currentEndDate.getDate() - 1); // Avoid overlap
        
        searchAttempts++;
        
        // Add delay between windows to be respectful
        await this.apiClient.sleep(1000);
        
      } catch (error) {
        this.logger.warn(`Error in time window ${searchAttempts + 1}: ${error.message}`);
        
        // If it's a date-related error, we might have gone too far back
        if (error.message.includes('time') || error.message.includes('date')) {
          this.logger.info('Reached the limit of searchable history');
          break;
        }
        
        // Move to next window anyway
        currentEndDate = new Date(startDate);
        currentEndDate.setDate(currentEndDate.getDate() - 1);
        searchAttempts++;
      }
    }
    
    return allItems;
  }

  async fetchListingsInTimeWindow(startTimeISO, endTimeISO) {
    const allItems = [];
    let pageNumber = 1;
    let hasMoreItems = true;
    const entriesPerPage = CONFIG.PAGINATION.DEFAULT_ENTRIES_PER_PAGE;

    while (hasMoreItems) {
      try {
        const requestBody = `
          <DetailLevel>ReturnAll</DetailLevel>
          <Pagination>
            <EntriesPerPage>${entriesPerPage}</EntriesPerPage>
            <PageNumber>${pageNumber}</PageNumber>
          </Pagination>
          <StartTimeFrom>${startTimeISO}</StartTimeFrom>
          <StartTimeTo>${endTimeISO}</StartTimeTo>
          <GranularityLevel>Coarse</GranularityLevel>
        `;

        const response = await this.apiClient.makeRequest('GetSellerList', requestBody);
        
        if (response.ItemArray && response.ItemArray.Item) {
          const items = Array.isArray(response.ItemArray.Item) 
            ? response.ItemArray.Item 
            : [response.ItemArray.Item];
          
          allItems.push(...items);
          
          // Check if there are more pages
          hasMoreItems = response.HasMoreItems === 'true';
          pageNumber++;
          
          // Add delay between pages
          if (hasMoreItems) {
            await this.apiClient.sleep(300);
          }
        } else {
          hasMoreItems = false;
        }
        
      } catch (error) {
        if (error.message.includes('Invalid page number')) {
          break;
        } else {
          throw error;
        }
      }
    }

    return allItems;
  }

  async fetchMyeBaySellingData() {
    this.logger.info('Fetching My eBay active listings...');
    
    const results = {
      active: [],
      sold: [],
      unsold: []
    };
    
    results.active = await this.fetchMyeBayCategory('ActiveList', 'active listings');

    const totalItems = results.active.length;
    this.logger.info(`My eBay data: ${results.active.length} active listings`);
    
    return results;
  }

  async fetchMyeBayCategory(listType, description) {
    this.logger.info(`Fetching ${description}...`);
    
    const allItems = [];
    let pageNumber = 1;
    let hasMoreItems = true;
    const entriesPerPage = 200; // Max allowed by eBay
    
    while (hasMoreItems) {
      try {
        const requestBody = `
          <${listType}>
            <Include>true</Include>
            <Sort>TimeLeft</Sort>
            <Pagination>
              <EntriesPerPage>${entriesPerPage}</EntriesPerPage>
              <PageNumber>${pageNumber}</PageNumber>
            </Pagination>
          </${listType}>
        `;

        const response = await this.apiClient.makeRequest('GetMyeBaySelling', requestBody);
        
        const listContainer = response[listType];
        const items = this.extractItems(listContainer);
        
        if (items.length > 0) {
          allItems.push(...items);
          this.logger.info(`${description} page ${pageNumber}: ${items.length} items (total: ${allItems.length})`);
          
          // Check if there are more items
          // eBay returns HasMoreItems or we can check if we got fewer than requested
          const hasMoreFromResponse = listContainer?.HasMoreItems === 'true';
          const hasMoreFromCount = items.length === entriesPerPage;
          hasMoreItems = hasMoreFromResponse || hasMoreFromCount;
          
          pageNumber++;
          
          // Add delay between pages
          if (hasMoreItems) {
            await this.apiClient.sleep(500);
          }
        } else {
          hasMoreItems = false;
          this.logger.info(`No more ${description} found`);
        }
        
      } catch (error) {
        this.logger.warn(`Error fetching ${description} page ${pageNumber}: ${error.message}`);
        hasMoreItems = false;
      }
    }
    
    return allItems;
  }

  extractItems(listContainer) {
    if (!listContainer || !listContainer.ItemArray) {
      return [];
    }
    
    const items = listContainer.ItemArray.Item;
    return Array.isArray(items) ? items : [items];
  }
}

// ==================== DATA PROCESSOR ====================
class TradingDataProcessor {
  constructor(logger) {
    this.logger = logger;
  }

  processSellerListings(items) {
    this.logger.info('Processing active listings...');
    
    return items.map(item => ({
      // Basic Item Info
      'Item ID': item.ItemID || '',
      'SKU': item.SKU || '',
      'Title': item.Title || '',
      'Subtitle': item.SubTitle || '',
      'Description': this.cleanDescription(item.Description) || '',
      
      // Category & Classification
      'Category ID': item.PrimaryCategory?.CategoryID || '',
      'Category Name': item.PrimaryCategory?.CategoryName || '',
      'Secondary Category ID': item.SecondaryCategory?.CategoryID || '',
      'Secondary Category Name': item.SecondaryCategory?.CategoryName || '',
      
      // Condition
      'Condition ID': item.ConditionID || '',
      'Condition Name': item.ConditionDisplayName || '',
      'Condition Description': item.ConditionDescription || '',
      
      // Listing Format & Duration
      'Listing Type': item.ListingType || '',
      'Listing Duration': item.ListingDuration || '',
      'Listing Format': item.ListingDetails?.ListingType || '',
      
      // Pricing
      'Start Price': item.StartPrice?._ || '',
      'Current Price': item.SellingStatus?.CurrentPrice?._ || '',
      'Buy It Now Price': item.BuyItNowPrice?._ || '',
      'Reserve Price': item.ReservePrice?._ || '',
      'Currency': item.SellingStatus?.CurrentPrice?.currencyID || '',
      
      // Quantity & Sales
      'Quantity': item.Quantity || '',
      'Quantity Sold': item.SellingStatus?.QuantitySold || '',
      'Quantity Available': item.QuantityAvailable || '',
      'Min Qty Per Buyer': item.QuantityInfo?.MinimumRemnantSet || '',
      
      // Bidding & Offers
      'Bid Count': item.SellingStatus?.BidCount || '',
      'High Bidder': item.SellingStatus?.HighBidder?.UserID || '',
      'Best Offer Enabled': item.BestOfferDetails?.BestOfferEnabled || 'false',
      'Auto Accept Price': item.BestOfferDetails?.BestOfferAutoAcceptPrice?._ || '',
      'Min Accept Price': item.BestOfferDetails?.BestOfferAutoDeclinePrice?._ || '',
      
      // Status & Timing
      'Listing Status': item.SellingStatus?.ListingStatus || '',
      'Time Left': item.SellingStatus?.TimeLeft || '',
      'Start Time': item.ListingDetails?.StartTime || '',
      'End Time': item.ListingDetails?.EndTime || '',
      'Time Left (Days)': this.calculateDaysLeft(item.SellingStatus?.TimeLeft),
      
      // Location & Shipping
      'Site': item.Site || '',
      'Country': item.Country || '',
      'Location': item.Location || '',
      'Postal Code': item.PostalCode || '',
      'Shipping Type': item.ShippingDetails?.ShippingType || '',
      'Shipping Cost': item.ShippingDetails?.ShippingServiceOptions?.[0]?.ShippingServiceCost?._ || '',
      'Free Shipping': item.ShippingDetails?.ShippingServiceOptions?.[0]?.FreeShipping || 'false',
      'Fast Handling': item.ShippingDetails?.FastAndFree || 'false',
      
      // Payment & Returns
      'Payment Methods': this.extractPaymentMethods(item.PaymentMethods),
      'PayPal Email': item.PayPalEmailAddress || '',
      'Returns Accepted': item.ReturnPolicy?.ReturnsAcceptedOption || '',
      'Return Period': item.ReturnPolicy?.ReturnsWithinOption || '',
      'Return Policy Description': item.ReturnPolicy?.Description || '',
      
      // Images & Media
      'Gallery URL': item.GalleryURL || '',
      'Gallery Type': item.GalleryType || '',
      'Picture Count': item.PictureDetails?.PhotoDisplay?.length || '0',
      'Has Pictures': item.PictureDetails ? 'true' : 'false',
      
      // URLs & Links
      'View Item URL': item.ListingDetails?.ViewItemURL || '',
      'View Item URL For Natural Search': item.ListingDetails?.ViewItemURLForNaturalSearch || '',
      
      // Performance Metrics
      'Watch Count': item.WatchCount || '',
      'Hit Count': item.HitCount || '',
      'Question Count': item.QuestionCount || '',
      
      // Listing Features
      'Private Listing': item.PrivateListing || 'false',
      'Bold Title': item.ListingEnhancement?.includes('Bold') || 'false',
      'Featured': item.ListingEnhancement?.includes('Featured') || 'false',
      'Highlight': item.ListingEnhancement?.includes('Highlight') || 'false',
      'Gallery Plus': item.ListingEnhancement?.includes('GalleryPlus') || 'false',
      
      // Business Policies
      'Payment Policy ID': item.SellerProfiles?.SellerPaymentProfile?.PaymentProfileID || '',
      'Shipping Policy ID': item.SellerProfiles?.SellerShippingProfile?.ShippingProfileID || '',
      'Return Policy ID': item.SellerProfiles?.SellerReturnProfile?.ReturnProfileID || '',
      
      // Item Specifics
      'Brand': this.extractItemSpecific(item.ItemSpecifics, 'Brand'),
      'Model': this.extractItemSpecific(item.ItemSpecifics, 'Model'),
      'Size': this.extractItemSpecific(item.ItemSpecifics, 'Size'),
      'Color': this.extractItemSpecific(item.ItemSpecifics, 'Color'),
      'Material': this.extractItemSpecific(item.ItemSpecifics, 'Material'),
      
      // Seller Info
      'Seller ID': item.Seller?.UserID || '',
      'Feedback Score': item.Seller?.FeedbackScore || '',
      'Positive Feedback %': item.Seller?.PositiveFeedbackPercent || '',
      
      // Fees & Costs
      'Listing Fee': item.ListingDetails?.ListingFee?._ || '',
      'Final Value Fee': item.SellingStatus?.FinalValueFee?._ || '',
      'Total Fees': this.calculateTotalFees(item),
      
      // Technical Details
      'Revision': item.ReviseStatus?.ItemRevised || 'false',
      'UUID': item.UUID || '',
      'Application Data': item.ApplicationData || '',
      
      // Timestamps (formatted for Excel)
      'Created Date': this.formatDate(item.ListingDetails?.StartTime),
      'End Date': this.formatDate(item.ListingDetails?.EndTime),
      'Last Modified': this.formatDate(item.TimeLeft)
    }));
  }

  processMyeBayData(myeBayData) {
    this.logger.info('Processing My eBay selling data...');
    
    const processedData = [];
    
    // Process active items
    myeBayData.active.forEach(item => {
      processedData.push({
        ...this.processSellerListings([item])[0],
        'Listing Status Category': 'Active'
      });
    });

    // Process sold items
    myeBayData.sold.forEach(item => {
      processedData.push({
        ...this.processSellerListings([item])[0],
        'Listing Status Category': 'Sold'
      });
    });

    // Process unsold items
    myeBayData.unsold.forEach(item => {
      processedData.push({
        ...this.processSellerListings([item])[0],
        'Listing Status Category': 'Unsold'
      });
    });

    return processedData;
  }

  cleanDescription(description) {
    if (!description) return '';
    
    return description.replace(/<[^>]*>/g, '').trim().substring(0, 500);
  }

  extractPaymentMethods(paymentMethods) {
    if (!paymentMethods) return '';
    
    if (typeof paymentMethods === 'string') return paymentMethods;
    if (Array.isArray(paymentMethods)) return paymentMethods.join(', ');
    if (paymentMethods.Payment) {
      const payments = Array.isArray(paymentMethods.Payment) 
        ? paymentMethods.Payment 
        : [paymentMethods.Payment];
      return payments.join(', ');
    }
    
    return '';
  }

  extractItemSpecific(itemSpecifics, name) {
    if (!itemSpecifics || !itemSpecifics.NameValueList) return '';
    
    const nameValueList = Array.isArray(itemSpecifics.NameValueList) 
      ? itemSpecifics.NameValueList 
      : [itemSpecifics.NameValueList];
    
    const specific = nameValueList.find(item => 
      item.Name && item.Name.toLowerCase() === name.toLowerCase()
    );
    
    if (specific && specific.Value) {
      return Array.isArray(specific.Value) ? specific.Value.join(', ') : specific.Value;
    }
    
    return '';
  }

  calculateDaysLeft(timeLeft) {
    if (!timeLeft) return '';
    
    // Parse eBay time format like "P5DT12H30M45S" (5 days, 12 hours, 30 minutes, 45 seconds)
    const match = timeLeft.match(/P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
    if (match) {
      const days = parseInt(match[1]) || 0;
      const hours = parseInt(match[2]) || 0;
      return days + Math.round(hours / 24 * 10) / 10; // Include fractional days
    }
    
    return timeLeft;
  }

  calculateTotalFees(item) {
    const listingFee = parseFloat(item.ListingDetails?.ListingFee?._ || 0);
    const finalValueFee = parseFloat(item.SellingStatus?.FinalValueFee?._ || 0);
    const total = listingFee + finalValueFee;
    return total > 0 ? total.toFixed(2) : '';
  }

  formatDate(dateString) {
    if (!dateString) return '';
    
    try {
      const date = new Date(dateString);
      return date.toISOString().split('T')[0]; // YYYY-MM-DD format for Excel
    } catch (e) {
      return dateString;
    }
  }
}

// ==================== EXCEL EXPORTER ====================
class ExcelExporter {
  constructor(logger) {
    this.logger = logger;
  }

  async exportToExcel(data, filename) {
    this.logger.info(`Creating Excel file: ${filename}`);
    
    if (!data || data.length === 0) {
      throw new Error('No data to export');
    }

    try {
      const workbook = XLSX.utils.book_new();
      
      // Create main listings sheet
      const worksheet = XLSX.utils.json_to_sheet(data);
      
      // Auto-size columns
      const columnWidths = this.calculateColumnWidths(data);
      worksheet['!cols'] = columnWidths;
      
      // Add freeze panes and filters
      worksheet['!freeze'] = { xSplit: 1, ySplit: 1 };
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
      
      XLSX.utils.book_append_sheet(workbook, worksheet, 'All Listings');
      
      // Create summary sheet
      const summary = this.createSummary(data);
      const summarySheet = XLSX.utils.json_to_sheet(summary);
      XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
      
      // Write file
      XLSX.writeFile(workbook, filename);
      
      const stats = await fs.stat(filename);
      this.logger.info(`Excel file created: ${filename} (${this.formatFileSize(stats.size)})`);
      
      return { filename, recordCount: data.length, fileSize: stats.size };
      
    } catch (error) {
      this.logger.error(`Error creating Excel file: ${error.message}`);
      throw error;
    }
  }

  createSummary(data) {
    const statusCounts = {};
    const typeCounts = {};
    const categoryCount = {};
    let totalValue = 0;
    let totalWatchers = 0;
    let totalQuantity = 0;

    data.forEach(item => {
      const status = item['Listing Status'] || 'Unknown';
      const type = item['Listing Type'] || 'Unknown';
      const category = item['Category Name'] || 'Unknown';
      const price = parseFloat(item['Current Price']) || 0;
      const watchers = parseInt(item['Watch Count']) || 0;
      const quantity = parseInt(item['Quantity Available']) || 0;

      statusCounts[status] = (statusCounts[status] || 0) + 1;
      typeCounts[type] = (typeCounts[type] || 0) + 1;
      categoryCount[category] = (categoryCount[category] || 0) + 1;
      totalValue += price;
      totalWatchers += watchers;
      totalQuantity += quantity;
    });

    const summary = [
      { 'Metric': 'ACTIVE LISTINGS SUMMARY', 'Value': '' },
      { 'Metric': '', 'Value': '' },
      { 'Metric': 'Total Active Listings', 'Value': data.length },
      { 'Metric': 'Total Inventory Value', 'Value': `${totalValue.toFixed(2)}` },
      { 'Metric': 'Total Items Available', 'Value': totalQuantity },
      { 'Metric': 'Total Watchers', 'Value': totalWatchers },
      { 'Metric': 'Average Price per Item', 'Value': data.length > 0 ? `${(totalValue / data.length).toFixed(2)}` : '$0.00' },
      { 'Metric': '', 'Value': '' },
      { 'Metric': 'BY LISTING STATUS:', 'Value': '' }
    ];

    Object.entries(statusCounts).forEach(([status, count]) => {
      summary.push({ 'Metric': `  ${status}`, 'Value': count });
    });

    summary.push({ 'Metric': '', 'Value': '' });
    summary.push({ 'Metric': 'BY LISTING TYPE:', 'Value': '' });

    Object.entries(typeCounts).forEach(([type, count]) => {
      summary.push({ 'Metric': `  ${type}`, 'Value': count });
    });

    summary.push({ 'Metric': '', 'Value': '' });
    summary.push({ 'Metric': 'TOP CATEGORIES:', 'Value': '' });

    // Show top 10 categories
    const topCategories = Object.entries(categoryCount)
      .sort(([,a], [,b]) => b - a)
      .slice(0, 10);

    topCategories.forEach(([category, count]) => {
      summary.push({ 'Metric': `  ${category}`, 'Value': count });
    });

    return summary;
  }

  calculateColumnWidths(data) {
    const columnWidths = [];
    
    if (data.length > 0) {
      Object.keys(data[0]).forEach((key, index) => {
        const maxLength = Math.max(
          key.length,
          ...data.slice(0, 100).map(row => {
            const value = row[key];
            return value ? String(value).length : 0;
          })
        );
        
        const width = Math.max(8, Math.min(maxLength + 2, 50));
        columnWidths[index] = { width };
      });
    }
    
    return columnWidths;
  }

  formatFileSize(bytes) {
    const units = ['B', 'KB', 'MB', 'GB'];
    let size = bytes;
    let unitIndex = 0;
    
    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }
    
    return `${size.toFixed(1)} ${units[unitIndex]}`;
  }
}

// ==================== MAIN APPLICATION ====================
class TradingApiExporter {
  constructor() {
    this.logger = new Logger();
    this.rateLimiter = new RateLimiter();
    this.apiClient = new TradingApiClient(this.logger, this.rateLimiter);
    this.dataFetcher = new TradingDataFetcher(this.apiClient, this.logger);
    this.dataProcessor = new TradingDataProcessor(this.logger);
    this.excelExporter = new ExcelExporter(this.logger);
  }

  async validateEnvironment() {
    const requiredFields = [
      'EBAY_ACCESS_TOKEN',  
      'EBAY_CLIENT_ID'      
    ];
    
    const missing = requiredFields.filter(key => !process.env[key]);
    
    if (missing.length > 0) {
      throw new Error(`Missing required environment variables: ${missing.join(', ')}`);
    }

    this.logger.info('Environment validation passed');
  }

  async run() {
    const startTime = Date.now();
    
    try {
      this.logger.info('='.repeat(80));
      this.logger.info('eBay Trading API - Active Listings Exporter v2.1');
      this.logger.info('Uses OAuth 2.0 + Trading API to export ALL active listings');
      this.logger.info('(Both web-created + API-created active listings)');
      this.logger.info('='.repeat(80));
      
      await this.validateEnvironment();
      
      const environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
      const outputFile = process.env.OUTPUT_FILE || `ebay_all_listings_${Date.now()}.xlsx`;
      
      this.logger.info(`Environment: ${environment}`);
      this.logger.info(`Output file: ${outputFile}`);
      
      // Fetch all seller listings
      const allListings = await this.dataFetcher.fetchAllSellerListings();
      
      if (allListings.length === 0) {
        this.logger.warn('No listings found');
        return { success: true, recordCount: 0 };
      }
      
      // Process data
      const processedData = this.dataProcessor.processSellerListings(allListings);
      
      // Export to Excel
      const exportResult = await this.excelExporter.exportToExcel(processedData, outputFile);
      
      // Final summary
      const stats = this.logger.getStats();
      const duration = ((Date.now() - startTime) / 1000).toFixed(1);
      
      this.logger.info('='.repeat(80));
      this.logger.info('ACTIVE LISTINGS EXPORT COMPLETED SUCCESSFULLY');
      this.logger.info('='.repeat(80));
      this.logger.info(`Duration: ${duration} seconds`);
      this.logger.info(`Total API requests: ${stats.totalRequests}`);
      this.logger.info(`Active listings exported: ${exportResult.recordCount}`);
      this.logger.info(`Output file: ${exportResult.filename}`);
      this.logger.info('='.repeat(80));
      
      return {
        success: true,
        ...exportResult,
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
      this.logger.error('EXPORT FAILED');
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
    const exporter = new TradingApiExporter();
    
    try {
      await exporter.run();
      process.exit(0);
    } catch (error) {
      console.error('\nFatal error:', error.message);
      process.exit(1);
    }
  }
}

// Export for use as module
module.exports = TradingApiExporter;

// Run if called directly
main();