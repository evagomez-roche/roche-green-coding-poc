# 🌍 Green IT Carbon Footprint Enrichment Pipeline

This tool automates the enrichment of device inventories and calculates the micro-allocated Software Carbon Intensity (ECO-SCI methodology) for Roche's digital sustainability reporting. 

It dynamically fetches real-time grid intensity (via Climatiq) and hardware embodied emissions (via Boavizta).

## 🚀 Basic Usage

You can run the script directly from your terminal. It supports command-line arguments to specify your custom inventory files:

    python carbon_script.py --input "your_inventory_file.xlsx" --output "enriched_output.xlsx"

*(If no arguments are provided, it will look for the default file names configured in the codebase).*

## ⚙️ Smart Autodiscovery & Air-Gapped Environments

This script features a **Smart Autodiscovery** module. It will automatically ping internal metadata endpoints to detect your cloud provider (AWS, Azure, GCP) and geolocate your IP to fetch the correct regional grid intensity. If run locally, it will scan the physical CPU cores to estimate hardware allocation.

**⚠️ For Highly Secure / Firewalled Servers:**
If your CI/CD pipeline or server blocks outbound internet traffic (air-gapped environments), the autodiscovery will safely failover to predefined defaults to prevent crashes. 

To ensure 100% precision in these secure environments, please configure the following **Environment Variables** (Manual Overrides) in your deployment pipeline. The script will prioritize these values over the autodiscovery:

* CLOUD_PROVIDER: Set to "aws", "azure", "gcp", or "local".
* CLOUD_INSTANCE_TYPE: The hardware node size (e.g., "t3.medium", "Standard_D2s_v3").
* CLOUD_REGION: The ISO country code for the electricity grid (e.g., "CH" for Switzerland, "ES" for Spain).

Example of manual override via terminal:

    export CLOUD_PROVIDER="aws"
    export CLOUD_INSTANCE_TYPE="t3.large"
    export CLOUD_REGION="CH"
    python carbon_script.py

## 🔐 Security Configuration

In compliance with enterprise security standards, **do not hardcode API keys** in the script. Please ensure the Climatiq API key is injected securely into your runtime environment:

* CLIMATIQ_API_KEY: Your Climatiq enterprise API token.