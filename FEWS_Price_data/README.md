# FEWS NET Haiti Price Database

This module downloads and stores market price data for Haiti from the [FEWS NET Data Warehouse API](https://fdw.fews.net/api/).

## Overview

- **Data Source**: FEWS NET (Famine Early Warning Systems Network)
- **Coverage**: Haiti market prices from January 2005 to present
- **Database**: DuckDB (local, single-file analytical database)
- **Update Frequency**: Monthly (data collected by CNSA/FEWS NET)

## Quick Start

```bash
# Install dependencies
pip install duckdb pandas requests

# Initialize database and perform full sync
python sync_fews_db.py --init
python sync_fews_db.py --full

# Check database status
python sync_fews_db.py --stats

# Run incremental sync (for updates)
python sync_fews_db.py --sync
```

## Data Available

### Markets (11)

| Market | Department |
|--------|------------|
| Cap Haitien | Nord |
| Cayes | Sud |
| Fond-des-Negres | Nippes |
| Gonaives | Artibonite |
| Hinche | Centre |
| Jacmel | Sud-Est |
| Jeremie | Grand'Anse |
| Ouanaminthe | Nord-Est |
| Port-au-Prince, Croix-de-Bossales | Ouest |
| Port-de-Paix | Nord-Ouest |

### Products (43)

**Staples**: Rice (7 varieties), Maize Meal (3), Wheat Flour, Wheat Grain, Sorghum

**Legumes**: Beans (Black, Red, Lima, Pinto)

**Oils**: Refined Vegetable Oil (7 brands)

**Other Food**: Sugar, Salt, Spaghetti (6 brands), Milk (2 brands), Tomato Paste (3 brands)

**Fuel**: Diesel, Gasoline, Kerosene, Charcoal

## Database Schema

The database uses a star schema design with dimension tables and a central fact table.

```
┌─────────────┐     ┌─────────────────────┐     ┌─────────────┐
│   markets   │     │  price_observations │     │  products   │
├─────────────┤     ├─────────────────────┤     ├─────────────┤
│ id (PK)     │◄────│ market_id (FK)      │────►│ id (PK)     │
│ fews_id     │     │ product_id (FK)     │     │ name        │
│ fnid        │     │ unit_id (FK)        │     │ cpcv2       │
│ name        │     │ source_id (FK)      │     │ product_src │
│ admin_1     │     │ period_date         │     └─────────────┘
│ admin_2     │     │ value               │
│ latitude    │     │ exchange_rate       │     ┌─────────────┐
│ longitude   │     │ common_currency_price│────►│    units    │
└─────────────┘     └─────────────────────┘     ├─────────────┤
                                                │ id (PK)     │
                                                │ name        │
                                                │ unit_type   │
                                                └─────────────┘
```

### Tables

#### `markets` (Dimension)
Stores market locations with geographic coordinates.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key (auto-increment) |
| fews_id | INTEGER | FEWS NET market ID |
| fnid | VARCHAR | FEWS NET ID (e.g., HT0000M001) |
| name | VARCHAR | Market name |
| admin_1 | VARCHAR | Department (e.g., Nord) |
| admin_2 | VARCHAR | Commune (e.g., Cap Haitien) |
| country_code | VARCHAR | ISO country code (HT) |
| latitude | DOUBLE | Geographic latitude |
| longitude | DOUBLE | Geographic longitude |

#### `products` (Dimension)
Stores product/commodity information.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key (auto-increment) |
| name | VARCHAR | Product name (e.g., Beans (Black)) |
| cpcv2 | VARCHAR | UN CPC v2 code |
| cpcv2_description | VARCHAR | CPC description |
| product_source | VARCHAR | Local or Import |
| is_staple_food | BOOLEAN | Staple food indicator |

#### `units` (Dimension)
Stores measurement units.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key (auto-increment) |
| name | VARCHAR | Unit name (e.g., 6_lb, 175_g) |
| unit_type | VARCHAR | Unit type (e.g., Weight) |
| common_unit | VARCHAR | Standardized unit (kg) |
| conversion_factor | DOUBLE | Conversion to common unit |

#### `data_sources` (Dimension)
Stores data source organizations.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key (auto-increment) |
| fews_id | INTEGER | FEWS NET source ID |
| name | VARCHAR | Organization name |
| document_name | VARCHAR | Source document |

#### `price_observations` (Fact Table)
Central fact table storing all price observations.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key (auto-increment) |
| market_id | INTEGER | FK to markets |
| product_id | INTEGER | FK to products |
| unit_id | INTEGER | FK to units |
| source_id | INTEGER | FK to data_sources |
| period_date | DATE | Observation date (end of period) |
| start_date | DATE | Start of collection period |
| price_type | VARCHAR | Price type (Retail) |
| currency | VARCHAR | Currency code (HTG) |
| value | DOUBLE | Price in local currency |
| exchange_rate | DOUBLE | HTG to USD rate |
| common_unit_price | DOUBLE | Price per kg |
| common_currency_price | DOUBLE | Price in USD |
| collection_status | VARCHAR | Status (Published) |
| fews_dataseries_id | INTEGER | FEWS dataseries ID |
| api_modified_at | TIMESTAMP | Last modified in API |
| imported_at | TIMESTAMP | Import timestamp |

**Unique Constraint**: `(market_id, product_id, unit_id, period_date, price_type)`

#### `import_log` (Tracking)
Tracks sync operations for auditing and incremental updates.

| Column | Type | Description |
|--------|------|-------------|
| id | INTEGER | Primary key |
| import_date | TIMESTAMP | When import ran |
| records_fetched | INTEGER | Records from API |
| records_inserted | INTEGER | New records added |
| records_updated | INTEGER | Existing records updated |
| records_skipped | INTEGER | Skipped records |
| date_range_start | DATE | Query start date |
| date_range_end | DATE | Query end date |
| status | VARCHAR | success/failed/partial |
| error_message | VARCHAR | Error details if failed |

## Pre-built Views

### `v_latest_prices`
Most recent prices for all products and markets.

```sql
SELECT * FROM v_latest_prices
WHERE product LIKE 'Rice%'
ORDER BY market, product;
```

### `v_price_timeseries`
Full time series with computed price changes.

| Column | Description |
|--------|-------------|
| market | Market name |
| product | Product name |
| period_date | Observation date |
| price | Price in HTG |
| price_usd | Price in USD |
| price_1m_ago | Price 1 month ago |
| price_1y_ago | Price 1 year ago |
| pct_change_1m | % change from 1 month ago |
| pct_change_1y | % change from 1 year ago |
| moving_avg_12m | 12-month moving average |

## Example Queries

### Latest prices for a specific product
```sql
SELECT market, product, period_date, value as price_htg, common_currency_price as price_usd
FROM v_latest_prices
WHERE product = 'Beans (Black)'
ORDER BY market;
```

### Price trends over time
```sql
SELECT market, product, period_date, price, pct_change_1y, moving_avg_12m
FROM v_price_timeseries
WHERE product = 'Rice (4% Broken)'
  AND market = 'Port-au-Prince, Croix-de-Bossales'
  AND period_date >= '2020-01-01'
ORDER BY period_date;
```

### Compare prices across markets
```sql
SELECT
    period_date,
    MAX(CASE WHEN market = 'Cap Haitien' THEN value END) as cap_haitien,
    MAX(CASE WHEN market = 'Port-au-Prince, Croix-de-Bossales' THEN value END) as port_au_prince,
    MAX(CASE WHEN market = 'Gonaives' THEN value END) as gonaives
FROM price_observations po
JOIN markets m ON po.market_id = m.id
JOIN products p ON po.product_id = p.id
WHERE p.name = 'Rice (4% Broken)'
GROUP BY period_date
ORDER BY period_date DESC
LIMIT 12;
```

### Year-over-year inflation by product
```sql
SELECT
    product,
    AVG(pct_change_1y) as avg_yoy_change,
    MIN(pct_change_1y) as min_yoy_change,
    MAX(pct_change_1y) as max_yoy_change
FROM v_price_timeseries
WHERE period_date >= '2024-01-01'
  AND pct_change_1y IS NOT NULL
GROUP BY product
ORDER BY avg_yoy_change DESC;
```

## Python Usage

```python
import duckdb

# Connect to database
con = duckdb.connect('FEWS_Price_data/database/fews_haiti.duckdb')

# Query to pandas DataFrame
df = con.execute("""
    SELECT * FROM v_price_timeseries
    WHERE product = 'Beans (Black)'
""").fetchdf()

# Use DuckDB's pandas integration
import pandas as pd
con.execute("SELECT * FROM df WHERE price > 500")  # Query the DataFrame directly!

con.close()
```

## File Structure

```
FEWS_Price_data/
├── README.md                    # This file
├── fewsnet_haiti_downloader.py  # API client for direct downloads
├── sync_fews_db.py              # Database sync CLI
├── database/
│   ├── schema.sql               # Database schema definitions
│   ├── fews_database.py         # Database manager class
│   └── fews_haiti.duckdb        # DuckDB database file
└── data/                        # CSV exports (optional)
```

## API Reference

The FEWS NET API is **public** and requires no authentication.

- **Base URL**: `https://fdw.fews.net/api/`
- **Price Data**: `/marketpricefacts/?country_code=HT`
- **Markets**: `/market/?country_code=HT`
- **Formats**: JSON, CSV, XML

See [FEWS NET API Documentation](https://help.fews.net/fde/v3/fews-net-api) for details.

## Maintenance

### Incremental Updates
Run weekly or monthly to fetch new data:
```bash
python sync_fews_db.py --sync
```

### Full Refresh
To reload all historical data:
```bash
rm database/fews_haiti.duckdb
python sync_fews_db.py --init
python sync_fews_db.py --full
```

### Check Import History
```bash
python sync_fews_db.py --query "SELECT * FROM import_log ORDER BY import_date DESC LIMIT 10"
```
