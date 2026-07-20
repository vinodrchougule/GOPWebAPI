using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace GOPWebAPI.Models.MRO_Dictionary
{
    public class MRODictionaryNounModifierTemplateModel
    {
        public MRODictionaryNounModifierTemplateModel()
        {
            NounSynonyms = new List<NounSynonym>();
            NounModifierAttributes = new List<NounModifierTemplateAttribute>();
            NounModifierAttributeEVVs = new List<NounModifierTemplateAttributeEVV>();
            NounModifierUNSPSCs = new List<NounModifierUNSPSC>();
            ImageFileNames = new List<string>();
        }

        [Required]
        [StringLength(20)]
        public string VersionNameOrNo { get; set; }
        [Required]
        [StringLength(50)]
        public string Noun { get; set; }
        [Required]
        [StringLength(100)]
        public string Modifier { get; set; }
        [StringLength(4000)]
        public string NounDefinition { get; set; }
        [StringLength(4000)]
        public string NounModifierDefinitionOrGuidelines { get; set; }
        public List<NounSynonym> NounSynonyms { get; set; }
        public List<NounModifierTemplateAttribute> NounModifierAttributes { get; set; }
        public List<NounModifierTemplateAttributeEVV> NounModifierAttributeEVVs { get; set; }
        public List<NounModifierUNSPSC> NounModifierUNSPSCs { get; set; }
        public List<string> ImageFileNames { get; set; }
        [Required]
        [StringLength(50)]
        public string UserID { get; set; }
    }

    public class NounSynonym
    {
        public string Synonym { get; set; }
        public string SynonymDefinitionOrGuidelines { get; set; }
    }

    public class NounModifierTemplateAttribute
    {
        public string Attribute { get; set; }
        public string AttributeGuidelines { get; set; }
        public string Priority { get; set; }
        public string MandatoryOrOptional { get; set; }
    }

    public class NounModifierTemplateAttributeEVV
    {
        public string Attribute { get; set; }
        public string EnumeratedValidValue { get; set; }
        public string Priority { get; set; }
    }

    public class NounModifierUNSPSC
    {
        public string UNSPSCVersion { get; set; }
        public string UNSPSCCode { get; set; }
        public string UNSPSCCategory { get; set; }
    }

    public class ImageModel
    {
        public string Name { get; set; }
        public string Data { get; set; }
        public string ImageTempFileName { get; set; }
    }

    public class MRODictionaryNounModifier
    {
        public string VersionNameOrNo { get; set; }
        public string Noun { get; set; }
        public string Modifier { get; set; }
        public string NounDefinition { get; set; }
        public string NounModifierDefinitionOrGuidelines { get; set; }
        public DateTime CreatedOn { get; set; }
        public string CreatedBy { get; set; }
        public DateTime? UpdatedOn { get; set; }
        public string UpdatedBy { get; set; }
    }
}