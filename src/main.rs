use std::collections::HashMap;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::{Duration, Instant};

use anyhow::{Result, anyhow};
use chrono::Utc;
use dotenvy::dotenv;
use quick_xml::de::from_str;
use reqwest::Client;
use serde_json::Value;
use tokio::sync::Mutex;

// ==================== CONFIGURATION ====================
struct Config;
impl Config {
    const TRADING_API_VERSION: &'static str = "1291";
    const SANDBOX_URL: &'static str = "https://api.sandbox.ebay.com/ws/api.dll";
    const PRODUCTION_URL: &'static str = "https://api.ebay.com/ws/api.dll";
    const REQUESTS_PER_SECOND: usize = 2;
    const DEFAULT_ENTRIES_PER_PAGE: usize = 100;
    const MAX_ENTRIES_PER_PAGE: usize = 200;
}

// ==================== LOGGER ====================
#[derive(Clone)]
struct Logger {
    start_time: Instant,
    request_count: Arc<AtomicUsize>,
    error_count: Arc<AtomicUsize>,
}

impl Logger {
    fn new() -> Self {
        Self {
            start_time: Instant::now(),
            request_count: Arc::new(AtomicUsize::new(0)),
            error_count: Arc::new(AtomicUsize::new(0)),
        }
    }

    fn log(&self, level: &str, message: &str) {
        let timestamp = Utc::now().to_rfc3339();
        let elapsed = self.start_time.elapsed().as_secs_f32();
        println!(
            "[{}] [{:.1}s] [{:>5}] {}",
            timestamp, elapsed, level, message
        );
    }

    fn info(&self, message: &str) {
        self.log("INFO", message);
    }
    fn warn(&self, message: &str) {
        self.log("WARN", message);
    }
    fn error(&self, message: &str) {
        self.error_count.fetch_add(1, Ordering::Relaxed);
        self.log("ERROR", message);
    }

    fn increment_request(&self) {
        self.request_count.fetch_add(1, Ordering::Relaxed);
    }

    fn get_stats(&self) -> (usize, usize, f32) {
        let requests = self.request_count.load(Ordering::Relaxed);
        let errors = self.error_count.load(Ordering::Relaxed);
        let elapsed = self.start_time.elapsed().as_secs_f32();
        (requests, errors, elapsed)
    }
}

// ==================== RATE LIMITER ====================
struct RateLimiter {
    times: Mutex<Vec<Instant>>,
}

impl RateLimiter {
    fn new() -> Self {
        Self {
            times: Mutex::new(Vec::new()),
        }
    }

    async fn wait_for_slot(&self) {
        loop {
            let mut times = self.times.lock().await;
            let now = Instant::now();
            times.retain(|t| now.duration_since(*t) < Duration::from_secs(1));
            if times.len() < Config::REQUESTS_PER_SECOND {
                times.push(now);
                break;
            }
            drop(times);
            tokio::time::sleep(Duration::from_millis(100)).await;
        }
    }
}

// ==================== TRADING API CLIENT ====================
struct TradingApiClient {
    logger: Logger,
    rate_limiter: Arc<RateLimiter>,
    client: Client,
    app_id: String,
    oauth_token: String,
    site_id: String,
    environment: String,
}

impl TradingApiClient {
    fn new(logger: Logger, rate_limiter: Arc<RateLimiter>) -> Self {
        Self {
            logger,
            rate_limiter,
            client: Client::new(),
            app_id: std::env::var("EBAY_APP_ID")
                .or_else(|_| std::env::var("EBAY_CLIENT_ID"))
                .unwrap_or_default(),
            oauth_token: std::env::var("EBAY_ACCESS_TOKEN").unwrap_or_default(),
            site_id: std::env::var("EBAY_SITE_ID").unwrap_or_else(|_| "0".to_string()),
            environment: std::env::var("EBAY_ENVIRONMENT")
                .unwrap_or_else(|_| "sandbox".to_string()),
        }
    }

    async fn make_request(&self, call_name: &str, request_body: &str) -> Result<Value> {
        self.rate_limiter.wait_for_slot().await;
        self.logger.increment_request();

        let xml_request = format!(
            r#"<?xml version="1.0" encoding="utf-8"?><{call}Request xmlns="urn:ebay:apis:eBLBaseComponents"><Version>{ver}</Version>{body}</{call}Request>"#,
            call = call_name,
            ver = Config::TRADING_API_VERSION,
            body = request_body
        );

        let url = if self.environment == "production" {
            Config::PRODUCTION_URL
        } else {
            Config::SANDBOX_URL
        };

        let resp = self
            .client
            .post(url)
            .header("Content-Type", "text/xml")
            .header("X-EBAY-API-CALL-NAME", call_name)
            .header("X-EBAY-API-SITEID", &self.site_id)
            .header("X-EBAY-API-APP-NAME", &self.app_id)
            .header("X-EBAY-API-VERSION", Config::TRADING_API_VERSION)
            .header(
                "X-EBAY-API-COMPATIBILITY-LEVEL",
                Config::TRADING_API_VERSION,
            )
            .header("X-EBAY-API-REQUEST-ENCODING", "XML")
            .header("X-EBAY-API-IAF-TOKEN", &self.oauth_token)
            .body(xml_request)
            .send()
            .await?
            .text()
            .await?;

        let value: Value = from_str(&resp).map_err(|e| anyhow!("Failed to parse XML: {}", e))?;

        if let Some(ack) = value.get("Ack").and_then(|v| v.as_str()) {
            if ack == "Failure" || ack == "PartialFailure" {
                return Err(anyhow!("eBay API returned {:?}", value.get("Errors")));
            }
        }

        Ok(value)
    }

    async fn sleep(&self, ms: u64) {
        tokio::time::sleep(Duration::from_millis(ms)).await;
    }
}

// ==================== DATA FETCHER ====================
struct TradingDataFetcher {
    api_client: Arc<TradingApiClient>,
    logger: Logger,
}

impl TradingDataFetcher {
    fn new(api_client: Arc<TradingApiClient>, logger: Logger) -> Self {
        Self { api_client, logger }
    }

    async fn fetch_all_seller_listings(&self) -> Result<Vec<Value>> {
        self.logger
            .info("Starting complete seller listings fetch using Trading API...");
        let active = self
            .fetch_mye_bay_category("ActiveList", "active listings")
            .await?;
        Ok(active)
    }

    async fn fetch_mye_bay_category(
        &self,
        list_type: &str,
        description: &str,
    ) -> Result<Vec<Value>> {
        self.logger.info(&format!("Fetching {}...", description));
        let mut all_items = Vec::new();
        let mut page = 1;
        let entries_per_page = Config::MAX_ENTRIES_PER_PAGE;

        loop {
            let body = format!(
                "<{}><Include>true</Include><Sort>TimeLeft</Sort><Pagination><EntriesPerPage>{}</EntriesPerPage><PageNumber>{}</PageNumber></Pagination></{}>",
                list_type, entries_per_page, page, list_type
            );
            let response = self
                .api_client
                .make_request("GetMyeBaySelling", &body)
                .await?;
            let list_container = response.get(list_type).cloned().unwrap_or(Value::Null);
            let items = Self::extract_items(&list_container);
            if items.is_empty() {
                break;
            }
            all_items.extend(items.into_iter());
            let has_more = list_container
                .get("HasMoreItems")
                .and_then(|v| v.as_str())
                .map(|s| s == "true")
                .unwrap_or(false);
            if !has_more {
                break;
            }
            page += 1;
            self.api_client.sleep(500).await;
        }
        self.logger.info(&format!(
            "{} fetched: {} items",
            description,
            all_items.len()
        ));
        Ok(all_items)
    }

    fn extract_items(list_container: &Value) -> Vec<Value> {
        list_container
            .get("ItemArray")
            .and_then(|ia| ia.get("Item"))
            .map(|items| {
                if let Some(arr) = items.as_array() {
                    arr.clone()
                } else {
                    vec![items.clone()]
                }
            })
            .unwrap_or_default()
    }
}

// ==================== DATA PROCESSOR ====================
struct TradingDataProcessor {
    logger: Logger,
}

impl TradingDataProcessor {
    fn new(logger: Logger) -> Self {
        Self { logger }
    }

    fn process_seller_listings(&self, items: &[Value]) -> Vec<HashMap<String, String>> {
        self.logger.info("Processing active listings...");
        items.iter().map(|item| self.process_item(item)).collect()
    }

    fn process_item(&self, item: &Value) -> HashMap<String, String> {
        let mut map = HashMap::new();
        map.insert("Item ID".to_string(), Self::get_str(item, &["ItemID"]));
        map.insert("SKU".to_string(), Self::get_str(item, &["SKU"]));
        map.insert("Title".to_string(), Self::get_str(item, &["Title"]));
        map.insert(
            "Category Name".to_string(),
            Self::get_str(item, &["PrimaryCategory", "CategoryName"]),
        );
        map.insert(
            "Start Price".to_string(),
            Self::get_str(item, &["StartPrice", "_"]),
        );
        map.insert(
            "Current Price".to_string(),
            Self::get_str(item, &["SellingStatus", "CurrentPrice", "_"]),
        );
        map.insert(
            "Currency".to_string(),
            Self::get_str(item, &["SellingStatus", "CurrentPrice", "currencyID"]),
        );
        map.insert("Quantity".to_string(), Self::get_str(item, &["Quantity"]));
        map.insert(
            "Quantity Sold".to_string(),
            Self::get_str(item, &["SellingStatus", "QuantitySold"]),
        );
        map.insert(
            "Listing Status".to_string(),
            Self::get_str(item, &["SellingStatus", "ListingStatus"]),
        );
        map.insert(
            "Start Time".to_string(),
            Self::get_str(item, &["ListingDetails", "StartTime"]),
        );
        map.insert(
            "End Time".to_string(),
            Self::get_str(item, &["ListingDetails", "EndTime"]),
        );
        map.insert(
            "View Item URL".to_string(),
            Self::get_str(item, &["ListingDetails", "ViewItemURL"]),
        );
        map.insert(
            "Seller ID".to_string(),
            Self::get_str(item, &["Seller", "UserID"]),
        );
        map.insert(
            "Payment Methods".to_string(),
            self.extract_payment_methods(item.get("PaymentMethods")),
        );
        map
    }

    fn get_str(value: &Value, path: &[&str]) -> String {
        let mut current = value;
        for key in path {
            match current.get(*key) {
                Some(v) => current = v,
                None => return String::new(),
            }
        }
        current.as_str().unwrap_or("").to_string()
    }

    fn extract_payment_methods(&self, pm: Option<&Value>) -> String {
        if let Some(pm) = pm {
            if let Some(s) = pm.as_str() {
                return s.to_string();
            }
            if let Some(arr) = pm.as_array() {
                return arr
                    .iter()
                    .filter_map(|v| v.as_str())
                    .collect::<Vec<_>>()
                    .join(", ");
            }
            if let Some(obj) = pm.as_object() {
                if let Some(val) = obj.get("Payment") {
                    if let Some(arr) = val.as_array() {
                        return arr
                            .iter()
                            .filter_map(|v| v.as_str())
                            .collect::<Vec<_>>()
                            .join(", ");
                    } else if let Some(s) = val.as_str() {
                        return s.to_string();
                    }
                }
            }
        }
        String::new()
    }
}

// ==================== EXCEL EXPORTER ====================
use rust_xlsxwriter::Workbook;

struct ExcelExporter {
    logger: Logger,
}

impl ExcelExporter {
    fn new(logger: Logger) -> Self {
        Self { logger }
    }

    fn export_to_excel(
        &self,
        data: &[HashMap<String, String>],
        filename: &str,
    ) -> Result<ExportResult> {
        self.logger
            .info(&format!("Creating Excel file: {}", filename));
        if data.is_empty() {
            return Err(anyhow!("No data to export"));
        }

        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        let headers: Vec<String> = data[0].keys().cloned().collect();
        for (col, header) in headers.iter().enumerate() {
            worksheet.write_string(0, col as u16, header)?;
        }

        for (row_idx, row) in data.iter().enumerate() {
            for (col, header) in headers.iter().enumerate() {
                let value = row.get(header).map(|s| s.as_str()).unwrap_or("");
                worksheet.write_string((row_idx + 1) as u32, col as u16, value)?;
            }
        }

        workbook.save(filename)?;
        let metadata = std::fs::metadata(filename)?;
        self.logger
            .info(&format!("Excel file created: {}", filename));
        Ok(ExportResult {
            filename: filename.to_string(),
            record_count: data.len(),
            file_size: metadata.len(),
        })
    }
}

struct ExportResult {
    filename: String,
    record_count: usize,
    file_size: u64,
}

// ==================== MAIN APPLICATION ====================
struct TradingApiExporter {
    logger: Logger,
    rate_limiter: Arc<RateLimiter>,
    api_client: Arc<TradingApiClient>,
    data_fetcher: TradingDataFetcher,
    data_processor: TradingDataProcessor,
    excel_exporter: ExcelExporter,
}

impl TradingApiExporter {
    fn new() -> Self {
        let logger = Logger::new();
        let rate_limiter = Arc::new(RateLimiter::new());
        let api_client = Arc::new(TradingApiClient::new(logger.clone(), rate_limiter.clone()));
        let data_fetcher = TradingDataFetcher::new(api_client.clone(), logger.clone());
        let data_processor = TradingDataProcessor::new(logger.clone());
        let excel_exporter = ExcelExporter::new(logger.clone());
        Self {
            logger,
            rate_limiter,
            api_client,
            data_fetcher,
            data_processor,
            excel_exporter,
        }
    }

    async fn validate_environment(&self) -> Result<()> {
        let required = ["EBAY_ACCESS_TOKEN", "EBAY_CLIENT_ID"];
        let missing: Vec<&str> = required
            .iter()
            .filter(|k| std::env::var(k).is_err())
            .cloned()
            .collect();
        if !missing.is_empty() {
            return Err(anyhow!(
                "Missing required environment variables: {}",
                missing.join(", ")
            ));
        }
        self.logger.info("Environment validation passed");
        Ok(())
    }

    async fn run(&self) -> Result<ExportResult> {
        let start = Instant::now();
        self.logger.info(&"=".repeat(80));
        self.logger
            .info("eBay Trading API - Active Listings Exporter (Rust)");
        self.logger.info(&"=".repeat(80));
        self.validate_environment().await?;
        let environment =
            std::env::var("EBAY_ENVIRONMENT").unwrap_or_else(|_| "sandbox".to_string());
        let output_file = std::env::var("OUTPUT_FILE").unwrap_or_else(|_| {
            format!("ebay_all_listings_{}.xlsx", chrono::Utc::now().timestamp())
        });
        self.logger.info(&format!("Environment: {}", environment));
        self.logger.info(&format!("Output file: {}", output_file));

        let listings = self.data_fetcher.fetch_all_seller_listings().await?;
        if listings.is_empty() {
            self.logger.warn("No listings found");
            return Ok(ExportResult {
                filename: output_file,
                record_count: 0,
                file_size: 0,
            });
        }

        let processed = self.data_processor.process_seller_listings(&listings);
        let export_result = self
            .excel_exporter
            .export_to_excel(&processed, &output_file)?;
        let (reqs, errors, elapsed) = self.logger.get_stats();
        self.logger.info(&"=".repeat(80));
        self.logger
            .info("ACTIVE LISTINGS EXPORT COMPLETED SUCCESSFULLY");
        self.logger
            .info(&format!("Duration: {:.1} seconds", elapsed));
        self.logger.info(&format!("Total API requests: {}", reqs));
        self.logger.info(&format!(
            "Active listings exported: {}",
            export_result.record_count
        ));
        self.logger
            .info(&format!("Output file: {}", export_result.filename));
        self.logger.info(&"=".repeat(80));
        Ok(export_result)
    }
}

// ==================== ENTRY POINT ====================
#[tokio::main]
async fn main() {
    dotenv().ok();
    let exporter = TradingApiExporter::new();
    if let Err(e) = exporter.run().await {
        eprintln!("\nFatal error: {}", e);
        std::process::exit(1);
    }
}
