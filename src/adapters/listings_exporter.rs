use reqwest::Client;
use serde_json::Value;
use std::env;

pub async fn fetch_active_listings() -> Result<Vec<Value>, Box<dyn std::error::Error>> {
    let token = env::var("EBAY_ACCESS_TOKEN")?;
    let url = "https://api.ebay.com/sell/listing/v1/listing?listingStatus=ACTIVE&limit=200";
    let client = Client::new();
    let resp = client.get(url).bearer_auth(token).send().await?;

    if !resp.status().is_success() {
        let status = resp.status();
        let text = resp.text().await.unwrap_or_default();
        return Err(format!("eBay API request failed: {} - {}", status, text).into());
    }

    let json: Value = resp.json().await?;
    let listings = json
        .get("listings")
        .and_then(|v| v.as_array())
        .cloned()
        .unwrap_or_default();
    Ok(listings)
}
