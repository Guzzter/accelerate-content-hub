# Content Hub to Sitecore Search Connector

**Content Hub Search Connector** is a reference implementation for synchronising data from [Sitecore Content Hub](https://www.sitecore.com/products/content-hub) to [Sitecore Search](https://www.sitecore.com/products/sitecore-search). Built with Azure Functions and designed for extensibility, this open-source solution provides a robust foundation for developing custom connectors tailored to your data model and business logic.

It supports both API-driven and queue-based ingestion and uses a flexible JSON-based configuration for defining field mappings between Content Hub and Search.

> ⚠️ **Note**
>
> The provided code is intended as a guideline and must be tailored to suit your specific implementation requirements. Please ensure thorough end-to-end testing is conducted to validate its functionality and performance in your environment.
---

## Project Overview

This solution includes three core projects:

* **`Sitecore.ContentHub.Integration.SearchConnector`**
  The Azure Function App responsible for retrieving data from Content Hub and pushing it to Sitecore Search.

* **`Sitecore.ContentHub.Integration.SearchConnector.AppHost`**
  A [.NET Aspire](https://learn.microsoft.com/en-us/dotnet/aspire/get-started/aspire-overview) host project for local development and orchestration.

* **`Sitecore.ContentHub.Integration.SearchConnector.ServiceDefaults`**
  Shared configuration for Aspire-based local development.

---

## Prerequisites

To run or deploy this connector, you’ll need:

* A Sitecore Content Hub instance
* A Sitecore Search instance with an **API Push source**
* An OAuth client in Content Hub with read access (see below)
* A user account in Content Hub for authentication
* Azure Service Bus for asynchronous message handling

### Minimum Permissions

Following the principle of least privilege, the Content Hub user should only have **read** access to the entity definitions and related entities defined in your configuration. This ensures the connector can retrieve the required data without unnecessary privileges.

---

## Configuration

The connector relies on a set of environment variables and a JSON configuration file to operate.

### Environment Variables

These variables should be set either in `local.settings.json` (for local development) or in the Azure Function App configuration.

| Variable                     | Description                                |
| ---------------------------- | ------------------------------------------ |
| `ContentHub:BaseUrl`         | Base URL of your Content Hub instance      |
| `ContentHub:ClientId`        | OAuth client ID for Content Hub            |
| `ContentHub:ClientSecret`    | OAuth client secret                        |
| `ContentHub:Username`        | Username of the import user                |
| `ContentHub:Password`        | Password for the import user               |
| `Search:BaseUrl`             | Base URL of Sitecore Search                |
| `Search:AuthToken`           | API key for the Search Push source         |
| `Search:DomainId`            | Domain ID for Sitecore Search              |
| `Search:SourceId`            | Source ID for the Search Push source       |
| `ServiceBus:UpsertQueueName` | Name of the Azure Service Bus upsert queue |
| `ServiceBus:DeleteQueueName` | Name of the Azure Service Bus delete queue |

---

### `config.json` Structure

The `config.json` file defines the mapping between Content Hub entities and Sitecore Search attributes.

#### Root

```json
{
  "CultureMaps": CultureMap[],
  "DefinitionMaps": DefinitionMap[]
}
```

---

#### `CultureMap`

Maps cultures between Content Hub and Sitecore Search. Only the cultures defined here will be processed.

```json
{
  "ContentHubCulture": "en-US",
  "SearchCulture": "en_us_"
}
```

---

#### `DefinitionMap`

Defines how entities in Content Hub map to entities in Sitecore Search.

```json
{
  "ContentHubEntityDefinition": "M.PCM.Product",
  "SearchEntity": "product",
  "FieldMaps": FieldMap[]
}
```

---

### Field Maps

The `FieldMaps` array supports a range of mapping types:

#### `PropertyFieldMap`

Maps a property from Content Hub to a Search attribute.

```json
{
  "Type": "property",
  "SearchAttributeName": "name",
  "ContentHubPropertyName": "ProductName"
}
```

---

#### `OptionListFieldMap`

Maps an option list property from Content Hub to a Search attribute.

```json
{
  "Type": "optionlist",
  "SearchAttributeName": "packaging",
  "ContentHubPropertyName": "PackagingType",
  "ContentHubOptionListName": "M.PackagingType"
}
```

---

#### `RelationFieldMap`

Follows a relation chain and maps a property on the related entity.

```json
{
  "Type": "relation",
  "SearchAttributeName": "markets",
  "ContentHubRelations": [
    {
      "Name": "PCMCatalogToProduct",
      "Role": "Child"
    },
    {
      "Name": "PCMMarketToCatalog",
      "Role": "Parent"
    }
  ],
  "ContentHubRelatedPropertyName": "MarketLabel"
}
```

---

#### `PublicLinkFieldMap`

Maps a public link to a Search attribute, optionally creating the link if it doesn’t exist.

```json
{
  "Type": "publiclink",
  "SearchAttributeName": "image_url",
  "ContentHubResourceName": "downloadOriginal",
  "CreateLinkIfNotExists": true,
  "ContentHubRelations": [
    { "Name": "PCMProductToMasterAsset", "Role": "Parent" }
  ]
}
```

---

#### `ExtractedContentFieldMap`

Maps extracted document content to a Search attribute.

```json
{
  "Type": "extractedcontent",
  "SearchAttributeName": "specifications",
  "ContentHubRelations": [
    { "Name": "PCMProductToDocumentAsset", "Role": "Parent" }
  ]
}
```

---

#### `Relation`

Defines a Content Hub relation to follow when mapping nested properties.

```json
{
  "Name": "PCMProductToMasterAsset",
  "Role": "Parent"
}
```

---

## Azure Functions

The connector exposes HTTP endpoints and Service Bus listeners for maximum integration flexibility.

### HTTP Endpoints

| Function           | Description                                                              |
| ------------------ | ------------------------------------------------------------------------ |
| `HealthCheck`      | Basic health check – returns `OK`                                        |
| `RunUpsert`        | Receives a Content Hub Action Message and upserts the entity to Search   |
| `RunDelete`        | Receives a Content Hub Action Message and deletes the entity from Search |
| `AddUpsertMessage` | Queues an upsert message on the Service Bus                              |
| `AddDeleteMessage` | Queues a delete message on the Service Bus                               |

---

### Service Bus Listeners

| Function    | Queue        | Description                                              |
| ----------- | ------------ | -------------------------------------------------------- |
| `RunUpsert` | Upsert queue | Processes queued messages to upsert entities into Search |
| `RunDelete` | Delete queue | Processes queued messages to delete entities from Search |

All endpoints expect a [Content Hub Action Message](https://doc.sitecore.com/ch/en/developers/content-hub/index.html#Actions) in the request body.

---

## Usage

1. Deploy the Azure Function App or run locally using the Aspire host.
2. Create and provide a `config.json` file defining your entity and field mappings.
3. In Sitecore Content Hub:

   * Configure **upsert** and **delete** actions:
   * 
     * Use **API Call actions** for HTTP-based integration, or
     * Use **Azure Service Bus actions** for queue-based integration.
   * Set up appropriate **triggers** to invoke these actions as part of your content lifecycle.
