using Microsoft.Extensions.Options;
using Sitecore.ContentHub.Integration.AssetImporter.Constants;
using Sitecore.ContentHub.Integration.AssetImporter.Models.ContentHub;
using Sitecore.ContentHub.Integration.AssetImporter.Models.Options;
using Sitecore.ContentHub.Integration.AssetImporter.Services.Abstract;
using Stylelabs.M.Base.Querying;
using Stylelabs.M.Framework.Essentials.LoadConfigurations;
using Stylelabs.M.Sdk.Search;

namespace Sitecore.ContentHub.Integration.AssetImporter.Services.Concrete
{
    class ExportWorker(IOptions<ContentHubOptions> contentHubOptions, IContentHubEntityHelper entityHelper, IContentHubSearchHelper searchAfterHelper, IExcelHelper excelHelper) : IExportWorker
    {
        public async Task<IEnumerable<UploadedFileMetadata>> GenerateMetadataExport()
        {
            return await GetUploadedFilesMetadata();
        }

        public async Task<MemoryStream> GenerateMetadataExcel()
        {
            var data = await GetUploadedFilesMetadata();
            return excelHelper.GenerateExcel(data, SchemaConstants.Asset.DefinitionName);
        }

        private async Task<IEnumerable<UploadedFileMetadata>> GetUploadedFilesMetadata()
        {
            var userId = await entityHelper.GetIdFromUsernameAsync(contentHubOptions.Value.Username);
            var createdStatusId = await entityHelper.GetIdFromIdentifierAsync(DataConstants.Identifiers.FinalLifecycleStatus_Created);

            var filter = QueryFilterBuilder.Definition.Name == SchemaConstants.Asset.DefinitionName && QueryFilterBuilder.Parent(SchemaConstants.Asset.Relations.FinalLifeCycleStatusToAsset).Id == createdStatusId && QueryFilterBuilder.CreatedBy == userId;
            var searchAfterQuery = new SearchAfterQuery(
                filter,
                [new() { Field = SchemaConstants.Asset.SystemProperties.CreatedOn, Order = QuerySortOrder.Desc }])
            {
                Take = 100,
                LoadConfiguration = new EntityLoadConfigurationBuilder().WithProperties(SchemaConstants.Asset.Properties.Filename).Build(),
            };
            var results = await searchAfterHelper.SearchAfter(searchAfterQuery);
            return results.Select(x => new UploadedFileMetadata { Id = x.Id!.Value, Identifier = x.Identifier, Filename = x.GetPropertyValue<string>(SchemaConstants.Asset.Properties.Filename) });
        }


    }
}
