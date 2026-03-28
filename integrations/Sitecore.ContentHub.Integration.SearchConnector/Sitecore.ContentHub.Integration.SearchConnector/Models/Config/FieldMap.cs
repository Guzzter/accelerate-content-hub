using Sitecore.ContentHub.Integration.SearchConnector.Models.Config.FieldMaps;
using System.Text.Json.Serialization;

namespace Sitecore.ContentHub.Integration.SearchConnector.Models.Config
{
    [JsonPolymorphic(TypeDiscriminatorPropertyName = "Type")]
    [JsonDerivedType(typeof(PropertyFieldMap), typeDiscriminator: "property")]
    [JsonDerivedType(typeof(OptionListFieldMap), typeDiscriminator: "optionlist")]
    [JsonDerivedType(typeof(RelationFieldMap), typeDiscriminator: "relation")]
    [JsonDerivedType(typeof(PublicLinkFieldMap), typeDiscriminator: "publiclink")]
    [JsonDerivedType(typeof(ExtractedContentFieldMap), typeDiscriminator: "extractedcontent")]
    public class FieldMap
    {
        public required string SearchAttributeName { get; set; }
    }
}
