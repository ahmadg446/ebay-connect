# eBay Listings Exporter - Clean Version

XML parsing format
rate limited to 2req/sec

## Setup

1. Install dependencies:
```bash
npm install xlsx
```

2. Create `.env` file with your credentials:
```bash
cp .env.template .env
# Edit .env with your actual values
```

## Step 1: Get User Token

```bash
node ebay-user-token.js
```

This will output a consent URL. Open it in browser, authorize, and copy the `code` parameter from redirect URL.

Add the code to your `.env`:
```bash
EBAY_AUTH_CODE=v^1.1#i^1...
```

Run again to get access token:
```bash
node ebay-user-token.js
```

Copy the access token to your `.env`:
```bash
EBAY_ACCESS_TOKEN=v^1.1#i^1...
```

## Step 2: Export Listings

```bash
node ebay-listings-exporter.js
```

## Files

- `ebay-user-token.js` - Gets user authentication token
- `ebay-listings-exporter.js` - Exports listings to Excel
- `.env` - Your configuration (create from template)

## Notes

- User tokens expire (usually 2 hours)
- Refresh tokens can extend this
- Production requires approved eBay developer app
