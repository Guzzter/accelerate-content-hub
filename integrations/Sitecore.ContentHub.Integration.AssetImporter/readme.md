# Content Hub Asset Importer

**Content Hub Asset Importer** is an open-source integration for bulk uploading assets into [Sitecore Content Hub](https://www.sitecore.com/products/content-hub). Designed for flexibility, it uses Azure Functions to detect and import new assets from an Azure Storage container.

> ⚠️ **Note**
>
> The provided code is intended as a guideline and must be tailored to suit your specific implementation requirements. Please ensure thorough end-to-end testing is conducted to validate its functionality and performance in your environment.

## Project Overview

The solution includes three components:

- `Sitecore.ContentHub.Integration.AssetImporter`  
  The core Azure Function App that performs imports and metadata exports.

- `Sitecore.ContentHub.Integration.AssetImporter.AppHost`  
  A [.NET Aspire](https://learn.microsoft.com/en-us/dotnet/aspire/get-started/aspire-overview) host project for local development and testing.

- `Sitecore.ContentHub.Integration.AssetImporter.ServiceDefaults`  
  Shared defaults and configuration for Aspire-based local development.

## Prerequisites

- An Azure Storage Account and container (or [Azurite](https://learn.microsoft.com/en-us/azure/storage/common/storage-use-azurite) for local development)
- A Sitecore Content Hub instance
- An OAuth client registered in Content Hub
- A user account with scoped permissions (see below)

### Minimum Permissions for Import User

Following the principle of least privilege, the import user should have only the permissions necessary for uploading and managing assets they create.

| Definition | Conditions | Permissions |
|-----------|------------|-------------|
| `M.UploadConfiguration` | — | `Read` |
| `M.Asset` | `M.Final.Lifecycle.Status = Created` AND created by current user | `Update` |
| `M.Asset`, `M.File` | Created by current user | `Read`, `Create` |

## Configuration

Set the following environment variables (in Azure or `local.settings.json`):

| Variable | Description |
|----------|-------------|
| `AzureWebJobsStorage` | Azure Storage connection string (use `UseDevelopmentStorage=true` for Azurite) |
| `Azure:StorageContainerName` | Name of the storage container to monitor |
| `ContentHub:BaseUrl` | Base URL of your Content Hub instance |
| `ContentHub:ClientId` | OAuth Client ID for API access |
| `ContentHub:ClientSecret` | OAuth Client Secret |
| `ContentHub:Username` | Username of the import user |
| `ContentHub:Password` | Password for the import user |

## Functions

The Azure Function App includes four endpoints:

### `AssetImporterTimer`
Runs every 15 minutes by default. It scans the configured storage container and imports new files into Content Hub.

### `AssetImporterHttp`
An HTTP-triggered version of the timer logic, useful for manual testing or triggering imports on demand.

### `MetadataExport`
Returns a JSON array of `{ id, identifier, filename }` for assets that:
- Are in the `Created` lifecycle state
- Were created by the configured user

### `MetadataExportExcel`
Returns the same data as `MetadataExport`, formatted as a downloadable Excel (.xlsx) file. This can be used for metadata enrichment or mapping.

## Usage

1. Deploy the Azure Function App, or run it locally via the Aspire host.
2. Drop files into the configured Azure Storage container.
3. Either wait for the timer function to trigger, or manually trigger the import using the HTTP endpoint.
4. Use the metadata export functions to retrieve Content Hub identifiers for the imported assets.

These identifiers can help with post-processing steps like adding metadata or managing lifecycle states.

> ⚠️ **Notes**
> - Avoid uploading files with duplicate filenames. The system relies on filenames to map assets during metadata operations.
> - Once metadata has been applied, update the asset’s `FinalLifeCycleStatus` to remove it from future exports.
