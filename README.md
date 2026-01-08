# Price Matrix Outlier Detection

Detects **pricing anomalies across multi-location price matrices** using adaptive, median-based clustering.  

Designed for **internal review, audit, and QA workflows** using exported Excel data.

## Overview

This project analyzes price matrices where:

- **Each row** represents an inventory item
- **Each column** reports that item's price at a specific location

The algorithm assumes that the **most common price range** for an item represents the expected price. Values that fall outside this range are flagged as potential anomalies.

The result is an **annotated Excel file** that highlights suspicious prices while preserving the original dataset.

## How It Works

### 1. Data Cleaning
- Ignores blank, zero, and non-numeric values
- Tracks `####` cells separately for visibility
- Skips rows with insufficient numeric data for analysis

### 2. Adaptive Clustering
- Prices within a row are sorted and grouped when adjacent values fall within an adaptive tolerance: `max(tolerance, row_median * 0.15)`
- This allows clustering behavior to scale naturally with higher- or lower-priced items

### 3. Cluster Stabilization
- Clusters with medians within **15% of each other** are merged
- Reduces fragmentation caused by small numeric gaps

### 4. Sanity Filtering
- Clusters whose median is below **20% of the row median** are discarded
- Prevents extreme low values from being treated as "normal" pricing

### 5. Normal Price Selection
- The **largest cluster** is treated as the expected price range
- If multiple clusters tie in size, the cluster **closest to the row median** is selected

### 6. Outlier Detection
- Any value **not belonging to the selected cluster** is flagged as a potential outlier

## Output

- **Red text** — flagged price outliers  
- **Yellow fill** — cells containing `####`  
- **Clusters column** (optional) — displays per-row cluster groupings for transparency and debugging

## Command-Line Arguments
**--input** (required): path to the input Excel file

- First column must be the item identifier (e.g., `InventoryItemName`)
- All remaining columns are treated as numeric price values

**--output** (optional): path to the output Excel file

- Default: `processed_output.xlsx`

**--tolerance** (optional): basic numeric tolerance used for clustering

- Default: `1.00`

**--skip-prefixes** (optional): Comma-separated list of item name prefixes to exclude from analysis

- Case-insensitive
- Useful for non-priceable or operational items (e.g., fees)

## Configuration Defaults

| Setting         | Value                                   | Purpose                                                  |
| --------------- | --------------------------------------- | -------------------------------------------------------- |
| `sanity_pct`    | `0.20`                                  | Prevents low-value clusters from being treated as normal |
| `adaptive_tol`  | `max(tolerance, row_median * 0.15)`     | Scales clustering by price magnitude                     |
| `merge_pct`     | `0.15`                                  | Reduces cluster fragmentation                            |

## Key Assumptions & Limitations
- **Relative, not absolute detection**
    
    The algorithm does not know what an item *should* cost. Anomalies are identified purely from observed pricing patterns.
- **No regional normalization**
    
    Regional pricing strategies, taxes, or cost-of-living differences are not considered.
- **Human review may be required**
    
    This tool highlights suspicious patterns but does not replace pricing decisions or business judgement.
- **Tolerance-sensitive behavior**
    
    Results depend on tolerance and data distribution. Different datasets may require tuning.
- **Export-based analysis only**
    
    Operates on static Excel reports and does not integrate with upstream systems.

## Design Philosophy
This project is a **decision-support tool**, not an automated pricing engine.

Its goal is to reduce review effort by surfacing suspicious pricing patterns while preserving human oversight.
