# eBay Listings Exporter - Rust Version

Fetch active eBay listings using a simple Rust CLI.

## Setup

1. Ensure `EBAY_ACCESS_TOKEN` is set in your environment with a valid eBay user access token.
2. Run the exporter:

```bash
cargo run
```

The program logs the total number of active listings and prints a snippet of the first listing returned by the API.
