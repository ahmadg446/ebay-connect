mod adapters;

use adapters::listings_exporter::fetch_active_listings;
use serde_json::to_string_pretty;

#[tokio::main]
async fn main() {
    match fetch_active_listings().await {
        Ok(listings) => {
            println!("Total active listings: {}", listings.len());
            if let Some(first) = listings.first() {
                match to_string_pretty(first) {
                    Ok(pretty) => println!("First listing: {}", pretty),
                    Err(err) => eprintln!("Failed to format listing: {}", err),
                }
            }
        }
        Err(err) => eprintln!("Error fetching listings: {}", err),
    }
}
