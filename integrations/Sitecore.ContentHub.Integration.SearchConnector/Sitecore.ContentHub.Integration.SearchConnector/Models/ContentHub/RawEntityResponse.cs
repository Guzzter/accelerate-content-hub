using Stylelabs.M.Base.Web.Api.Models;
using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.ContentHub
{
    public class RawEntityResponse
    {
        public long Id { get; set; }

        public required string Identifier { get; set; }

        public required IEnumerable<string> Cultures { get; set; }

        public required IDictionary<string, object> Properties { get; set; }

        public required IDictionary<string, object> Relations { get; set; }

        public required Link CreatedBy { get; set; }

        public required DateTime CreatedOn { get; set; }

        public required Link ModifiedBy { get; set; }

        public required DateTime ModifiedOn { get; set; }

        [JsonPropertyName("entitydefinition")]
        public required Link EntityDefinition { get; set; }

        public required IDictionary<string, IEnumerable<Link>> Renditions { get; set; }

        public string? PublicLink { get; set; }
    }
}
