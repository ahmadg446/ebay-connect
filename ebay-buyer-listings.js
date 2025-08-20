#!/usr/bin/env node

// DEPENDENCIES: npm install xlsx papaparse ebay-oauth-nodejs-client

require('dotenv').config();
const https = require('https');
const fs = require('fs').promises;
const fss = require('fs'); // for streaming writes
const Papa = require('papaparse');
const XLSX = require('xlsx');
const EbayAuthToken = require('ebay-oauth-nodejs-client');

// ==================== CONFIGURATION ====================
const CONFIG = {
  BUY_API: {
    SANDBOX_URL: 'api.sandbox.ebay.com',
    PRODUCTION_URL: 'api.ebay.com',
    BROWSE_PATH: '/buy/browse/v1/item_summary/search',
    ITEM_PATH: '/buy/browse/v1/item',
    REQUESTS_PER_SECOND: 5
  },
  
  SEARCH: {
    ITEMS_PER_PAGE: 200,
    MAX_RESULTS_PER_ITEM: 50,  // Get top 10 competitors per item for detailed analysis
    // Minimum seller quality thresholds (can be overridden via environment variables)
    MIN_SELLER_FEEDBACK_PERCENT: parseFloat(process.env.MIN_SELLER_FEEDBACK_PERCENT) || 98.0,
    MIN_SELLER_FEEDBACK_SCORE: parseInt(process.env.MIN_SELLER_FEEDBACK_SCORE) || 10000,
    // Minimum number of target tokens that must appear in a competitor title (default: 2)
    MIN_TYPE_TOKEN_MATCHES: parseInt(process.env.MIN_TYPE_TOKEN_MATCHES) || 3
  },
  SCALE: {
    CONCURRENCY: parseInt(process.env.CONCURRENCY) || 50,
    BATCH_SIZE: parseInt(process.env.BATCH_SIZE) || 500,
    OUTPUT_MODE: (process.env.OUTPUT_MODE || 'excel').toLowerCase() // 'excel' or 'csv'
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
    const minInterval = 1000 / CONFIG.BUY_API.REQUESTS_PER_SECOND;
    
    this.requestTimes = this.requestTimes.filter(time => 
      now - time < 1000
    );
    
    if (this.requestTimes.length < CONFIG.BUY_API.REQUESTS_PER_SECOND) {
      const resolve = this.requestQueue.shift();
      this.requestTimes.push(now);
      this.processing = false;
      
      resolve();
      
      if (this.requestQueue.length > 0) {
        setTimeout(() => this.processQueue(), minInterval);
      }
    } else {
      setTimeout(() => {
        this.processing = false;
        this.processQueue();
      }, minInterval);
    }
  }
}

// ==================== BUY API CLIENT ====================
class BuyApiClient {
  constructor(logger, rateLimiter) {
    this.logger = logger;
    this.rateLimiter = rateLimiter;
    this.environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
    this.ebayAuth = new EbayAuthToken({
      clientId: process.env.EBAY_CLIENT_ID,
      clientSecret: process.env.EBAY_CLIENT_SECRET,
    });
    this.accessToken = null;
  }

  async getAccessToken() {
    if (!this.accessToken) {
      try {
        this.logger.info('Getting application access token...');
        const token = await this.ebayAuth.getApplicationToken(
          this.environment.toUpperCase()
        );
        this.accessToken = token;
        this.logger.info('Access token obtained successfully');
      } catch (error) {
        this.logger.error('Failed to get access token', error);
        throw error;
      }
    }
    return this.accessToken;
  }

  async searchItems(searchQuery, categoryId = null, priceRange = null) {
    await this.rateLimiter.waitForSlot();
    
    try {
      this.logger.incrementRequest();
      
      const params = new URLSearchParams();
      params.append('q', searchQuery);
      params.append('limit', CONFIG.SEARCH.ITEMS_PER_PAGE);
      
      if (categoryId) {
        params.append('category_ids', categoryId);
      }
      
      const filters = [];
      if (priceRange) {
        filters.push(`price:[${priceRange.min}..${priceRange.max}]`);
      }
      filters.push('buyingOptions:{FIXED_PRICE}');
      filters.push('conditions:{NEW}');
      
      filters.forEach(filter => params.append('filter', filter));
      
      params.append('sort', 'price');
      params.append('fieldgroups', 'EXTENDED');
      
      const hostname = this.environment === 'sandbox' 
        ? CONFIG.BUY_API.SANDBOX_URL 
        : CONFIG.BUY_API.PRODUCTION_URL;
      
      const token = await this.getAccessToken();
      
      const options = {
        hostname: hostname,
        port: 443,
        path: `${CONFIG.BUY_API.BROWSE_PATH}?${params.toString()}`,
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'X-EBAY-C-MARKETPLACE-ID': 'EBAY_US',
          'Content-Type': 'application/json'
        }
      };

      return new Promise((resolve, reject) => {
        const req = https.request(options, (res) => {
          let data = '';
          res.on('data', chunk => data += chunk);
          res.on('end', () => {
            try {
              const response = JSON.parse(data);
              
              if (res.statusCode !== 200) {
                reject(new Error(`API Error: ${response.message || data}`));
              } else {
                resolve(response);
              }
            } catch (parseError) {
              reject(new Error(`Failed to parse response: ${parseError.message}`));
            }
          });
        });

        req.on('error', reject);
        req.end();
      });
      
    } catch (error) {
      this.logger.error(`Search failed for query: ${searchQuery}`, error);
      throw error;
    }
  }

  async getItemDetails(itemIds) {
    await this.rateLimiter.waitForSlot();
    
    try {
      this.logger.incrementRequest();
      
      // Buy API allows up to 20 items per request
      const params = new URLSearchParams();
      params.append('item_ids', itemIds.slice(0, 20).join(','));
      
      const hostname = this.environment === 'sandbox' 
        ? CONFIG.BUY_API.SANDBOX_URL 
        : CONFIG.BUY_API.PRODUCTION_URL;
      
      const token = await this.getAccessToken();
      
      const options = {
        hostname: hostname,
        port: 443,
        path: `${CONFIG.BUY_API.ITEM_PATH}?${params.toString()}`,
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'X-EBAY-C-MARKETPLACE-ID': 'EBAY_US',
          'Content-Type': 'application/json'
        }
      };

      return new Promise((resolve, reject) => {
        const req = https.request(options, (res) => {
          let data = '';
          res.on('data', chunk => data += chunk);
          res.on('end', () => {
            try {
              const response = JSON.parse(data);
              
              if (res.statusCode !== 200) {
                this.logger.warn(`Could not get details for items: ${response.message || 'Unknown error'}`);
                resolve({ items: [] }); // Return empty instead of failing
              } else {
                resolve(response);
              }
            } catch (parseError) {
              this.logger.warn('Failed to parse item details response');
              resolve({ items: [] });
            }
          });
        });

        req.on('error', (err) => {
          this.logger.warn(`Failed to get item details: ${err.message}`);
          resolve({ items: [] });
        });
        req.end();
      });
      
    } catch (error) {
      this.logger.warn(`Item details request failed: ${error.message}`);
      return { items: [] };
    }
  }

  sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
}

// ==================== MY LISTINGS LOADER ====================
class MyListingsLoader {
  constructor(logger) {
    this.logger = logger;
  }

  async loadMyListings(filepath = 'my_listings.xlsx') {
    try {
      this.logger.info(`Loading your listings from: ${filepath}`);
      
      const workbook = XLSX.readFile(filepath);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const myListings = XLSX.utils.sheet_to_json(worksheet);
      
      this.logger.info(`Loaded ${myListings.length} of your listings`);
      
      // Create a map for easy lookup by title/SKU
      const listingsMap = new Map();
      
      myListings.forEach(listing => {
        // Try to match by title (simplified)
        const simplifiedTitle = this.simplifyTitle(listing['Title'] || '');
        if (simplifiedTitle) {
          listingsMap.set(simplifiedTitle, listing);
        }
        
        // Also map by SKU if available
        if (listing['SKU']) {
          listingsMap.set(listing['SKU'], listing);
        }
      });
      
      return { listings: myListings, map: listingsMap };
      
    } catch (error) {
      this.logger.warn(`Could not load your listings: ${error.message}`);
      this.logger.warn('Proceeding without your listing data');
      return { listings: [], map: new Map() };
    }
  }

  simplifyTitle(title) {
    // Remove common words and normalize for matching
    return title.toLowerCase()
      .replace(/[^\w\s]/g, '')
      .replace(/\b(the|and|or|for|with|set|pc|piece|pieces)\b/g, '')
      .trim()
      .substring(0, 50);
  }

  findMyListing(productName, listingsMap) {
    // Try to find matching listing
    const simplified = this.simplifyTitle(productName);
    
    // Direct match
    if (listingsMap.has(simplified)) {
      return listingsMap.get(simplified);
    }
    
    // Partial match
    for (const [key, listing] of listingsMap) {
      if (key.includes(simplified.substring(0, 20)) || 
          simplified.includes(key.substring(0, 20))) {
        return listing;
      }
    }
    
    return null;
  }
}

// ==================== CSV PROCESSOR ====================
class CsvProcessor {
  constructor(logger) {
    this.logger = logger;
  }

  async loadItemsFromCsv(filepath) {
    this.logger.info(`Loading items from CSV: ${filepath}`);
    
    try {
      const csvContent = await fs.readFile(filepath, 'utf8');
      const parsed = Papa.parse(csvContent, {
        header: true,
        dynamicTyping: true,
        skipEmptyLines: true
      });
      
      if (parsed.errors.length > 0) {
        this.logger.warn('CSV parsing warnings:', parsed.errors);
      }
      
      this.logger.info(`Loaded ${parsed.data.length} items from CSV`);
      return parsed.data;
      
    } catch (error) {
      this.logger.error('Failed to load CSV', error);
      throw error;
    }
  }

  extractSearchTerms(productName) {
    const stopWords = ['luxury', 'hotel', 'the', 'and', 'or', 'with', 'for'];
    const brandPattern = /^[A-Z][a-z]+\s/;
    
    let cleaned = productName.replace(brandPattern, '');
    
    const importantTerms = [];
    
    const pieceMatch = cleaned.match(/(\d+)[\s-]?(piece|pc|pcs)/i);
    if (pieceMatch) importantTerms.push(pieceMatch[0]);
    
    const threadMatch = cleaned.match(/(\d+)\s?(thread count|TC)/i);
    if (threadMatch) importantTerms.push(threadMatch[1] + ' thread count');
    
    const materials = ['microfiber', 'cotton', 'bamboo', 'velvet', 'linen'];
    materials.forEach(material => {
      if (cleaned.toLowerCase().includes(material)) {
        importantTerms.push(material);
      }
    });
    
    const types = ['sheet set', 'duvet cover', 'coverlet', 'bedspread', 'pillowcase', 
                   'curtain', 'valance', 'quilt', 'comforter', 'blackut curtain', 'sheer valance',];
    types.forEach(type => {
      if (cleaned.toLowerCase().includes(type)) {
        importantTerms.push(type);
      }
    });
    
    const sizes = ['king', 'queen', 'full', 'twin', 'california king'];
    sizes.forEach(size => {
      if (cleaned.toLowerCase().includes(size)) {
        importantTerms.push(size);
      }
    });
    
    return importantTerms.join(' ').substring(0, 100);
  }

  // Return the product type keywords likely to describe the item (pillowcase, sheet set, etc.)
  getTargetTypes(productName) {
    const types = ['sheet set', 'duvet cover', 'coverlet', 'bedspread', 'pillowcase',
                   'pillow case', 'pillowcase set', 'pillow sham', 'fitted sheet',
                   'flat sheet', 'pillowcase', 'sham', 'curtain', 'valance', 'quilt', 'comforter'];
    const lc = (productName || '').toLowerCase();
    const matches = types.filter(type => {
      const t = type.toLowerCase();
      if (lc.includes(t)) return true;
      // also check first word/plural variants (e.g., "pillow" -> "pillowcases", "sheets")
      const first = t.split(' ')[0];
      if (first && (lc.includes(first) || lc.includes(first + 's') || lc.includes(first + 'case'))) return true;
      return false;
    });
    // return unique matches
    return [...new Set(matches)];
  }

  // Return deduped tokens extracted from the product name (materials, types, sizes, counts)
  getTargetTokens(productName) {
    // reuse extractSearchTerms to collect the important terms, then split into tokens
    const phrase = this.extractSearchTerms(productName || '');
    const tokens = (phrase || '')
      .toLowerCase()
      .replace(/[^\w\s]/g, ' ')
      .split(/\s+/)
      .filter(t => t && t.length > 1);

    // also include explicit type tokens (e.g., "pillowcase", "sheet")
    const explicit = this.getTargetTypes(productName || '').flatMap(t => t.split(' '));

    const all = [...tokens, ...explicit].map(t => t.trim()).filter(Boolean);
    return [...new Set(all)];
  }
}

// ==================== COMPETITOR ANALYZER ====================
class CompetitorAnalyzer {
  constructor(apiClient, logger) {
    this.apiClient = apiClient;
    this.logger = logger;
    this.listingsLoader = new MyListingsLoader(logger);
  }

  async findCompetitors(myItem, myListingData = null) {
    const productName = myItem['Product Name'] || myItem['Title'] || '';
    const category = myItem['Category'] || '';
    
    this.logger.info(`Finding competitors for: ${productName}`);
    
    const csvProcessor = new CsvProcessor(this.logger);
    let searchQuery = csvProcessor.extractSearchTerms(productName);

    // Build tokens and append the top two tokens to narrow the search (helps relevancy)
    const targetTokens = csvProcessor.getTargetTokens(productName);
    if (targetTokens.length > 0) {
      const tokensToAppend = targetTokens.slice(0, 2).join(' ');
      searchQuery = `${searchQuery} ${tokensToAppend}`.trim();
    }
    
    if (!searchQuery) {
      this.logger.warn(`Could not extract search terms from: ${productName}`);
      return [];
    }
    
    this.logger.info(`Search query: ${searchQuery}`);
    
    // Calculate price range if we have your price
    let priceRange = null;
    if (myListingData && myListingData['Current Price']) {
      const myPrice = parseFloat(myListingData['Current Price']);
      if (myPrice > 0) {
        priceRange = {
          min: Math.max(1, myPrice * 0.5),
          max: myPrice * 2
        };
      }
    }
    
    try {
      const searchResults = await this.apiClient.searchItems(searchQuery, null, priceRange);
      
      if (!searchResults.itemSummaries) {
        this.logger.warn('No competitors found');
        return [];
      }
      
      // Get top competitors
      const allSummaries = searchResults.itemSummaries || [];
      const minPercent = CONFIG.SEARCH.MIN_SELLER_FEEDBACK_PERCENT;
      const minScore = CONFIG.SEARCH.MIN_SELLER_FEEDBACK_SCORE;
  
      // Filter by seller quality: feedback percentage and feedback score
      let filteredSummaries = allSummaries.filter(s => {
        const percent = parseFloat(s.seller?.feedbackPercentage || 0);
        const score = parseInt(s.seller?.feedbackScore || 0);
        return percent >= minPercent && score >= minScore;
      });
  
      if (filteredSummaries.length === 0) {
        // If nothing meets the strict thresholds, fall back to the top results
        this.logger.info(`No competitors met seller thresholds (>= ${minPercent}% & >= ${minScore}). Falling back to top ${CONFIG.SEARCH.MAX_RESULTS_PER_ITEM} results.`);
        filteredSummaries = allSummaries;
      } else {
        this.logger.info(`Filtered competitors to ${filteredSummaries.length} items meeting seller thresholds (>= ${minPercent}% & >= ${minScore}).`);
      }
  
      // Further filter by product-type keywords (only keep listings that mention pillowcase/sheet/etc. in title)
      if (targetTokens.length > 0) {
        const minMatches = CONFIG.SEARCH.MIN_TYPE_TOKEN_MATCHES;
        const tokenFiltered = filteredSummaries.filter(s => {
          const title = (s.title || '').toLowerCase();
          let count = 0;
          for (const t of targetTokens) {
            if (title.includes(t)) count++;
            if (count >= minMatches) return true;
          }
          return false;
        });

        if (tokenFiltered.length > 0) {
          filteredSummaries = tokenFiltered;
          this.logger.info(`Filtered competitors to ${filteredSummaries.length} items with >= ${minMatches} token matches from: ${targetTokens.slice(0,6).join(', ')}`);
        } else {
          this.logger.info(`No competitors matched >= ${minMatches} tokens (${targetTokens.slice(0,6).join(', ')}). Keeping previous filtered set (${filteredSummaries.length}).`);
        }
      }
  
      const topCompetitors = filteredSummaries.slice(0, CONFIG.SEARCH.MAX_RESULTS_PER_ITEM);
      
      // Get detailed info for top competitors
      const itemIds = topCompetitors.map(item => item.itemId);
      const detailsResponse = await this.apiClient.getItemDetails(itemIds);
      
      // Merge summary and detailed data
      const competitors = topCompetitors.map(item => {
        // Find detailed data if available
        const detailed = detailsResponse.items?.find(d => d.itemId === item.itemId) || {};
        
        const competitorData = {
          // === MY ITEM INFO ===
          'My Product': productName,
          'My Category': category,
          'My Item ID': myListingData?.['Item ID'] || '',
          'My SKU': myListingData?.['SKU'] || '',
          'My Price': myListingData?.['Current Price'] || '',
          'My Quantity': myListingData?.['Quantity Available'] || '',
          'My Watchers': myListingData?.['Watch Count'] || '',
          'My Views': myListingData?.['Hit Count'] || '',
          'My Sold': myListingData?.['Quantity Sold'] || '',
          'My URL': myListingData?.['View Item URL'] || '',
          
          // === COMPETITOR INFO ===
          'Comp Title': item.title || '',
          'Comp Price': item.price?.value || '',
          'Comp Shipping': item.shippingOptions?.[0]?.shippingCost?.value || '0',
          'Comp Total Price': this.calculateTotalPrice(item),
          'Comp Seller': item.seller?.username || '',
          'Comp Feedback %': item.seller?.feedbackPercentage || '',
          'Comp Feedback Score': item.seller?.feedbackScore || '',
          'Comp Location': item.itemLocation?.country || '',
          'Comp Available': detailed.estimatedAvailabilities?.[0]?.availabilityThreshold || '',
          'Comp Sold': detailed.quantitySold || '',
          'Comp URL': item.itemWebUrl || '',
          'Comp Item ID': item.itemId || '',
          
          // === COMPARISON METRICS ===
          'Price Diff ($)': this.calculatePriceDiff(myListingData, item),
          'Price Diff (%)': this.calculatePriceDiffPercent(myListingData, item),
          'Price Position': this.getPricePosition(myListingData, item),
          'Feedback Advantage': this.compareFeedback(myListingData, item),
          'Search Rank': searchResults.itemSummaries.indexOf(item) + 1,
          
          // === SELLING SIGNALS ===
          'Fast Selling': detailed.quantitySold > 10 ? 'YES' : 'NO',
          'Low Stock': detailed.estimatedAvailabilities?.[0]?.availabilityThreshold < 5 ? 'YES' : 'NO',
          'Top Rated': item.seller?.feedbackPercentage >= 99.5 ? 'YES' : 'NO',
          'Free Shipping': item.shippingOptions?.[0]?.shippingCost?.value === '0' ? 'YES' : 'NO'
        };
        
        return competitorData;
      });
      
      this.logger.info(`Found ${competitors.length} competitors with details`);
      return competitors;
      
    } catch (error) {
      this.logger.error(`Failed to find competitors for: ${productName}`, error);
      return [];
    }
  }
  
  calculateTotalPrice(item) {
    const price = parseFloat(item.price?.value || 0);
    const shipping = parseFloat(item.shippingOptions?.[0]?.shippingCost?.value || 0);
    return (price + shipping).toFixed(2);
  }
  
  calculatePriceDiff(myListing, competitor) {
    if (!myListing || !myListing['Current Price']) return '';
    const myPrice = parseFloat(myListing['Current Price']);
    const compPrice = parseFloat(competitor.price?.value || 0);
    if (myPrice && compPrice) {
      return (compPrice - myPrice).toFixed(2);
    }
    return '';
  }
  
  calculatePriceDiffPercent(myListing, competitor) {
    if (!myListing || !myListing['Current Price']) return '';
    const myPrice = parseFloat(myListing['Current Price']);
    const compPrice = parseFloat(competitor.price?.value || 0);
    if (myPrice && compPrice) {
      const percent = ((compPrice - myPrice) / myPrice * 100).toFixed(1);
      return percent + '%';
    }
    return '';
  }
  
  getPricePosition(myListing, competitor) {
    if (!myListing || !myListing['Current Price']) return '';
    const myPrice = parseFloat(myListing['Current Price']);
    const compPrice = parseFloat(competitor.price?.value || 0);
    if (myPrice && compPrice) {
      if (compPrice > myPrice * 1.1) return 'UNDERPRICED';
      if (compPrice < myPrice * 0.9) return 'OVERPRICED';
      return 'COMPETITIVE';
    }
    return '';
  }
  
  compareFeedback(myListing, competitor) {
    const myScore = parseInt(myListing?.['Feedback Score'] || 0);
    const compScore = parseInt(competitor.seller?.feedbackScore || 0);
    if (myScore > compScore * 2) return 'STRONG';
    if (myScore > compScore) return 'BETTER';
    if (myScore < compScore / 2) return 'WEAK';
    return 'SIMILAR';
  }


  // Parallel analyzer that streams results via onRow callback. Uses configurable concurrency.
  async analyzeAllItemsParallel(items, myListingsMap, onRowCallback) {
    const concurrency = CONFIG.SCALE.CONCURRENCY;
    let index = 0;
    let active = 0;
    let processed = 0;

    return new Promise((resolve, reject) => {
      const runNext = async () => {
        if (index >= items.length && active === 0) {
          return resolve();
        }
        while (active < concurrency && index < items.length) {
          const currentIndex = index++;
          const item = items[currentIndex];
          active++;

          (async () => {
            try {
              this.logger.info(`Processing ${currentIndex + 1}/${items.length}`);
              const myListing = this.listingsLoader.findMyListing(item['Product Name'], myListingsMap);
              const competitors = await this.findCompetitors(item, myListing);
              // callback for each found competitor (stream out)
              if (competitors && competitors.length > 0) {
                await onRowCallback(competitors);
              }
              processed++;
            } catch (err) {
              this.logger.warn(`Error processing item index ${currentIndex}: ${err.message || err}`);
            } finally {
              active--;
              // slight delay to avoid bursts
              setTimeout(runNext, 50);
            }
          })();
        }
      };

      runNext();
    });
  }
 
  async analyzeAllItems(items, myListingsMap) {
    this.logger.info(`Starting competitor analysis for ${items.length} items...`);
    
    const allCompetitors = [];
    
    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      this.logger.info(`Processing item ${i + 1}/${items.length}`);
      
      // Try to find your listing data for this item
      const myListing = this.listingsLoader.findMyListing(
        item['Product Name'], 
        myListingsMap
      );
      
      if (myListing) {
        this.logger.info(`Matched with your listing: ${myListing['Title']}`);
      }
      
      const competitors = await this.findCompetitors(item, myListing);
      allCompetitors.push(...competitors);
      
      if (i < items.length - 1) {
        await this.apiClient.sleep(500);
      }
    }
    
    return allCompetitors;
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
      
      // Main comparison sheet
      const worksheet = XLSX.utils.json_to_sheet(data);
      
      // Format columns
      const columnWidths = [
        { width: 30 }, // My Product
        { width: 15 }, // My Category
        { width: 15 }, // My Item ID
        { width: 10 }, // My SKU
        { width: 10 }, // My Price
        { width: 10 }, // My Quantity
        { width: 10 }, // My Watchers
        { width: 10 }, // My Views
        { width: 10 }, // My Sold
        { width: 40 }, // My URL
        { width: 30 }, // Comp Title
        { width: 10 }, // Comp Price
        { width: 10 }, // Comp Shipping
        { width: 10 }, // Comp Total
        { width: 15 }, // Comp Seller
        { width: 10 }, // Comp Feedback %
        { width: 10 }, // Comp Score
        { width: 10 }, // Comp Location
        { width: 10 }, // Comp Available
        { width: 10 }, // Comp Sold
        { width: 40 }, // Comp URL
        { width: 15 }, // Comp Item ID
        { width: 10 }, // Price Diff $
        { width: 10 }, // Price Diff %
        { width: 12 }, // Price Position
        { width: 12 }, // Feedback Adv
        { width: 10 }, // Search Rank
        { width: 10 }, // Fast Selling
        { width: 10 }, // Low Stock
        { width: 10 }, // Top Rated
        { width: 10 }, // Free Ship
      ];
      worksheet['!cols'] = columnWidths;
      
      // Add freeze panes and filters
      worksheet['!freeze'] = { xSplit: 2, ySplit: 1 };
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      worksheet['!autofilter'] = { ref: XLSX.utils.encode_range(range) };
      
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Side by Side Comparison');
      
      // Create summary sheet
      const summary = this.createSummary(data);
      const summarySheet = XLSX.utils.json_to_sheet(summary);
      XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary Analysis');
      
      // Create opportunities sheet
      const opportunities = this.findOpportunities(data);
      const oppSheet = XLSX.utils.json_to_sheet(opportunities);
      XLSX.utils.book_append_sheet(workbook, oppSheet, 'Action Items');
      
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
    const productGroups = {};
    let underpriced = 0;
    let overpriced = 0;
    let competitive = 0;
    
    data.forEach(item => {
      const myProduct = item['My Product'] || 'Unknown';
      
      if (!productGroups[myProduct]) {
        productGroups[myProduct] = {
          count: 0,
          avgCompPrice: 0,
          myPrice: parseFloat(item['My Price']) || 0,
          myWatchers: parseInt(item['My Watchers']) || 0,
          position: item['Price Position']
        };
      }
      
      const group = productGroups[myProduct];
      group.count++;
      
      const compPrice = parseFloat(item['Comp Price']) || 0;
      if (compPrice > 0) {
        group.avgCompPrice += compPrice;
      }
      
      if (item['Price Position'] === 'UNDERPRICED') underpriced++;
      if (item['Price Position'] === 'OVERPRICED') overpriced++;
      if (item['Price Position'] === 'COMPETITIVE') competitive++;
    });

    const summary = [
      { 'Metric': 'COMPETITIVE ANALYSIS SUMMARY', 'Value': '' },
      { 'Metric': '', 'Value': '' },
      { 'Metric': 'Total Products Analyzed', 'Value': Object.keys(productGroups).length },
      { 'Metric': 'Total Competitors Found', 'Value': data.length },
      { 'Metric': '', 'Value': '' },
      { 'Metric': 'PRICING POSITION:', 'Value': '' },
      { 'Metric': '  Underpriced Items', 'Value': underpriced + ' (opportunity to raise prices)' },
      { 'Metric': '  Competitive Items', 'Value': competitive + ' (well positioned)' },
      { 'Metric': '  Overpriced Items', 'Value': overpriced + ' (need price adjustment)' },
      { 'Metric': '', 'Value': '' },
      { 'Metric': 'PRODUCT ANALYSIS:', 'Value': '' }
    ];

    Object.entries(productGroups).forEach(([product, stats]) => {
      const avgPrice = stats.count > 0 ? (stats.avgCompPrice / stats.count).toFixed(2) : 0;
      const priceDiff = stats.myPrice > 0 ? ((avgPrice - stats.myPrice) / stats.myPrice * 100).toFixed(1) : 0;
      
      summary.push({
        'Metric': `  ${product.substring(0, 40)}...`,
        'Value': `Your: $${stats.myPrice} | Market Avg: $${avgPrice} | Diff: ${priceDiff}%`
      });
    });

    return summary;
  }

  findOpportunities(data) {
    const opportunities = [];
    
    // Find underpriced items
    const underpriced = data.filter(d => d['Price Position'] === 'UNDERPRICED');
    if (underpriced.length > 0) {
      opportunities.push({
        'Action': 'RAISE PRICES',
        'Product': underpriced[0]['My Product'],
        'Current Price': underpriced[0]['My Price'],
        'Competitor Avg': underpriced[0]['Comp Price'],
        'Suggested Action': 'Increase price by 5-10% to capture more margin',
        'Priority': 'HIGH'
      });
    }
    
    // Find overpriced items
    const overpriced = data.filter(d => d['Price Position'] === 'OVERPRICED');
    if (overpriced.length > 0) {
      opportunities.push({
        'Action': 'LOWER PRICES',
        'Product': overpriced[0]['My Product'],
        'Current Price': overpriced[0]['My Price'],
        'Competitor Avg': overpriced[0]['Comp Price'],
        'Suggested Action': 'Reduce price to match market or add value',
        'Priority': 'HIGH'
      });
    }
    
    // Find items where competitors have free shipping but you don't
    const shippingOpp = data.filter(d => 
      d['Free Shipping'] === 'YES' && 
      parseFloat(d['My Price']) > parseFloat(d['Comp Total Price'])
    );
    if (shippingOpp.length > 0) {
      opportunities.push({
        'Action': 'ADD FREE SHIPPING',
        'Product': shippingOpp[0]['My Product'],
        'Current Price': shippingOpp[0]['My Price'],
        'Competitor Advantage': 'Free Shipping',
        'Suggested Action': 'Consider offering free shipping to compete',
        'Priority': 'MEDIUM'
      });
    }
    
    return opportunities.length > 0 ? opportunities : [
      { 'Action': 'No immediate actions needed', 'Product': '', 'Current Price': '', 'Competitor Avg': '', 'Suggested Action': 'Continue monitoring', 'Priority': 'LOW' }
    ];
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
class BuyApiCompetitorAnalyzer {
  constructor() {
    this.logger = new Logger();
    this.rateLimiter = new RateLimiter();
    this.apiClient = new BuyApiClient(this.logger, this.rateLimiter);
    this.csvProcessor = new CsvProcessor(this.logger);
    this.listingsLoader = new MyListingsLoader(this.logger);
    this.analyzer = new CompetitorAnalyzer(this.apiClient, this.logger);
    this.exporter = new ExcelExporter(this.logger);
  }

  async validateEnvironment() {
    const requiredFields = [
      'EBAY_CLIENT_ID',
      'EBAY_CLIENT_SECRET'
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
      this.logger.info('eBay Buy API - Competitor Analysis Tool v2.0');
      this.logger.info('Side-by-side comparison of your listings vs competitors');
      this.logger.info('='.repeat(80));
      
      await this.validateEnvironment();
      
      const environment = process.env.EBAY_ENVIRONMENT || 'sandbox';
      const inputFile = process.env.INPUT_CSV || 'bedding_categories.csv';
      const outputFile = process.env.BUYER_OUTPUT_FILE || `buyer_listings_${Date.now()}.xlsx`;
  const outputMode = CONFIG.SCALE.OUTPUT_MODE; // 'excel' or 'csv'
      
  this.logger.info(`Environment: ${environment}`);
  this.logger.info(`Input CSV: ${inputFile}`);
  this.logger.info(`Output file: ${outputFile}`);
      this.logger.info(`Scale: concurrency=${CONFIG.SCALE.CONCURRENCY}, batch=${CONFIG.SCALE.BATCH_SIZE}, output_mode=${outputMode}`);
      
      // Load your listings first
      const { listings: myListings, map: myListingsMap } = 
        await this.listingsLoader.loadMyListings('my_listings.xlsx');
      
  // Load items from CSV
  const csvItems = await this.csvProcessor.loadItemsFromCsv(inputFile);
  // Determine how many items to process. If MAX_ITEMS env var is set, use it. Otherwise analyze all.
  const maxItems = process.env.MAX_ITEMS ? parseInt(process.env.MAX_ITEMS) : csvItems.length;
  this.logger.info(`Max items to analyze: ${maxItems === csvItems.length ? 'ALL' : maxItems}`);
      
      if (csvItems.length === 0) {
        this.logger.warn('No items found in CSV');
        return { success: true, recordCount: 0 };
      }
      
  // Prepare items to analyze (slice if maxItems provided)
  const itemsToAnalyze = maxItems && maxItems < csvItems.length ? csvItems.slice(0, maxItems) : csvItems;
  this.logger.info(`Analyzing ${itemsToAnalyze.length} items from CSV`);
      
      // Choose large-scale streaming CSV mode when requested or when many items
      if (outputMode === 'csv' || itemsToAnalyze.length > 200) {
        const csvFile = outputFile.replace(/\.xlsx?$/i, '.csv');
        // Prepare headers same as Excel exporter expects
        const sampleHeaders = [
          'My Product','My Category','My Item ID','My SKU','My Price','My Quantity','My Watchers','My Views','My Sold','My URL',
          'Comp Title','Comp Price','Comp Shipping','Comp Total Price','Comp Seller','Comp Feedback %','Comp Feedback Score','Comp Location',
          'Comp Available','Comp Sold','Comp URL','Comp Item ID','Price Diff ($)','Price Diff (%)','Price Position','Feedback Advantage',
          'Search Rank','Fast Selling','Low Stock','Top Rated','Free Shipping'
        ];

        const csvExporter = new IncrementalCsvExporter(this.logger, csvFile, sampleHeaders);
        await csvExporter.init();

        // onRowCallback receives an array of competitor objects and writes them in batches
        const onRow = async (competitorRows) => {
          await csvExporter.appendRows(competitorRows);
        };

        // Run parallel analyzer (streams results to CSV)
        await this.analyzer.analyzeAllItemsParallel(itemsToAnalyze, myListingsMap, onRow);

        await csvExporter.close();
        const stats = await fs.stat(csvFile);
        this.logger.info(`Completed competitor analysis in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
        this.logger.info(`Total API requests: ${this.logger.getStats().totalRequests}, errors: ${this.logger.getStats().totalErrors}`);
        this.logger.info(`CSV file: ${csvFile} (${this.exporter.formatFileSize(stats.size)})`);
        return { success: true, filename: csvFile, recordCount: -1 };
      } else {
        // Small-scale default behavior: collect in memory and write Excel
        const competitorData = await this.analyzer.analyzeAllItems(itemsToAnalyze, myListingsMap);
        if (competitorData.length === 0) {
          this.logger.warn('No competitor data collected. Exiting without creating an Excel file.');
          return { success: true, recordCount: 0 };
        }
        const exportResult = await this.exporter.exportToExcel(competitorData, outputFile);
        const stats = this.logger.getStats();
        this.logger.info('='.repeat(80));
        this.logger.info(`Completed competitor analysis in ${((Date.now() - startTime) / 1000).toFixed(1)}s`);
        this.logger.info(`Total API requests: ${stats.totalRequests}, errors: ${stats.totalErrors}`);
        this.logger.info(`Excel file: ${exportResult.filename} (${this.exporter.formatFileSize(exportResult.fileSize)})`);
        this.logger.info(`Total competitor records: ${exportResult.recordCount}`);
        this.logger.info('='.repeat(80));
        return { success: true, filename: exportResult.filename, recordCount: exportResult.recordCount };
      }
 
    } catch (err) {
      this.logger.error('Unhandled error during run', err);
      return { success: false, error: err.message || String(err) };
    }
  }
}

// CLI entrypoint
if (require.main === module) {
  const app = new BuyApiCompetitorAnalyzer();
  app.run()
    .then(result => {
      if (result && result.success) {
        process.exit(0);
      } else {
        process.exit(1);
      }
    })
    .catch(err => {
      console.error('Fatal error:', err);
      process.exit(1);
    });
}

// ==================== INCREMENTAL CSV EXPORTER (for large-scale runs) ====================
class IncrementalCsvExporter {
  constructor(logger, filename, headers) {
    this.logger = logger;
    this.filename = filename;
    this.headers = headers;
    this.stream = null;
    this.wroteHeader = false;
  }

  async init() {
    // create write stream (overwrite existing)
    this.stream = fss.createWriteStream(this.filename, { flags: 'w' });
    // write header row
    if (this.headers && this.headers.length > 0) {
      this.stream.write(this.headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\n');
      this.wroteHeader = true;
    }
  }

  async appendRows(rows) {
    if (!this.stream) await this.init();
    for (const row of rows) {
      const line = this.headers.map(h => {
        const v = row[h] != null ? String(row[h]) : '';
        return `"${v.replace(/"/g, '""')}"`;
      }).join(',');
      this.stream.write(line + '\n');
    }
  }

  async close() {
    if (!this.stream) return;
    return new Promise((resolve) => {
      this.stream.end(() => {
        this.logger.info(`CSV export closed: ${this.filename}`);
        resolve();
      });
    });
  }
}
