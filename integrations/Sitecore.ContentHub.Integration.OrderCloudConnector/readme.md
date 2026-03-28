# Content Hub to OrderCloud Connector

**OrderCloud Connector** is a reference implementation for synchronising product and catalogue data from [Sitecore Content Hub](https://www.sitecore.com/products/content-hub) to [Sitecore OrderCloud](https://ordercloud.io/). Built with Azure Functions and designed for extensibility, this open-source connector is triggered by Experience Edge webhooks and processes updates asynchronously via Azure Service Bus.

This solution provides a robust foundation for integrating your Content Hub product data into OrderCloud’s commerce-ready APIs.

> ⚠️ **Note**
>
> The provided code is intended as a guideline and must be tailored to suit your specific implementation requirements. Please ensure thorough end-to-end testing is conducted to validate its functionality and performance in your environment.

## Project Overview

This connector is made up of three projects:

* **`Sitecore.ContentHub.Integration.OrderCloudConnector`**
  The core Azure Function App that receives update notifications and synchronises data with OrderCloud.

* **`Sitecore.ContentHub.Integration.OrderCloudConnector.AppHost`**
  A [.NET Aspire](https://learn.microsoft.com/en-us/dotnet/aspire/get-started/aspire-overview) host for local development and orchestration.

* **`Sitecore.ContentHub.Integration.OrderCloudConnector.ServiceDefaults`**
  Shared configuration and service defaults used across the Aspire solution.

## Prerequisites

To run or deploy this connector, you’ll need:

* A Sitecore Content Hub instance with **Experience Edge** enabled
* A configured **Experience Edge webhook** for update notifications
* An **OAuth client** in Content Hub
* A Content Hub user account with **read** access to:
  * `M.PCM.Product`
  * `M.PCM.Catalog`
* A Sitecore OrderCloud account

## Configuration

Set the following environment variables in either `local.settings.json` (for local development) or within your Azure Function App configuration:

### Required Settings

| Variable                     | Description                              |
| ---------------------------- | ---------------------------------------- |
| `ContentHub:BaseUrl`         | Base URL of your Content Hub instance    |
| `ContentHub:ClientId`        | OAuth Client ID for Content Hub          |
| `ContentHub:ClientSecret`    | OAuth Client Secret                      |
| `ContentHub:Username`        | Username of the import user              |
| `ContentHub:Password`        | Password for the import user             |
| `OrderCloud:ClientId`        | OAuth Client ID for OrderCloud           |
| `OrderCloud:ClientSecret`    | OAuth Client Secret for OrderCloud       |
| `ServiceBus:Name`            | Azure Service Bus namespace              |
| `ServiceBus:UpdateQueueName` | Name of the queue to use for updates     |

### Optional Settings

| Variable                            | Description                                                               |
| ----------------------------------- | ------------------------------------------------------------------------- |
| `OrderCloud:ApiUrl`                 | Base URL of OrderCloud API (default: `https://api.ordercloud.io`)         |
| `OrderCloud:AuthUrl`                | Auth URL of OrderCloud (defaults to value of `ApiUrl`)                    |
| `Schema:CatalogEntityDefinition`    | Entity definition for catalog in Content Hub (default: `M.PCM.Catalog`)   |
| `Schema:CatalogNameProperty`        | Property used for catalog name (default: `CatalogName`)                   |
| `Schema:CatalogDescriptionProperty` | Property used for catalog description (default: `CatalogDescription`)     |
| `Schema:CatalogToProductRelation`   | Relation from catalog to product (default: `PCMCatalogToProduct`)         |
| `Schema:ProductEntityDefinition`    | Entity definition for product (default: `M.PCM.Product`)                  |
| `Schema:ProductNameProperty`        | Property used for product name (default: `ProductName`)                   |
| `Schema:ProductDescriptionProperty` | Property used for product description (default: `ProductLongDescription`) |

## Functions

The Azure Function App provides both HTTP endpoints and Service Bus listeners to handle incoming webhook messages and queue-based processing.

### HTTP Endpoints

| Function           | Purpose                                                                |
| ------------------ | ---------------------------------------------------------------------- |
| `HealthCheck`      | Basic availability check – returns `"OK"`                              |
| `AddUpdateMessage` | Accepts a JSON payload and places it on the update queue               |
| `Update`           | Entry point for Experience Edge – processes webhook update immediately |

### Service Bus Listener

| Function    | Trigger      | Purpose                                            |
| ----------- | ------------ | -------------------------------------------------- |
| `RunUpdate` | Update queue | Processes a queued message and syncs to OrderCloud |

All HTTP and queue-triggered functions expect a valid Experience Edge webhook payload related to changes in product or catalogue entities.

## Usage

1. Deploy the Azure Function App or run it locally via Aspire.
2. Set environment variables based on your Content Hub and OrderCloud setup.
3. In Content Hub:

   * Configure an **Experience Edge webhook** to call either:

     * `AddUpdateMessage` (preferred for async processing via queue), or
     * `Update` (for immediate processing)
   * Conigure Experience Edge to publish entities for the folowing definitions (with appropriate conditions):
     * `M.PCM.Product`
     * `M.PCM.Catalog`
