using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GOPWebAPI.Models.GAT_Models
{
    public class DescriptionGenerator
    {
        public string SettingName { get; set; } = string.Empty;
        public string InputFileName { get; set; } = string.Empty;
        public string SpecificModifierExcluded { get; set; } = string.Empty;
        public string UploadedAbbreviationFileName { get; set; } = string.Empty;
        public string AbbreviationFileName { get; set; } = string.Empty;
        public string DelimiterAfterNoun { get; set; } = string.Empty;
        public string DelimiterAfterModifier { get; set; } = string.Empty;
        public string DelimiterAfterAttributeName { get; set; } = string.Empty;
        public string DelimiterAfterAttributeValue { get; set; } = string.Empty;
        public string DelimiterAfterAdditionalInformation { get; set; } = string.Empty;
        public string DelimiterAfterMFRName { get; set; } = string.Empty;
        public string DelimiterAfterMFRPartNo { get; set; } = string.Empty;
        public string MultipleValuesSeparator { get; set; } = string.Empty;
        public string UploadedIdentifierFileName { get; set; } = string.Empty;
        public string IdentifierFileName { get; set; } = string.Empty;
        public string PrefixForAdditionalInformation { get; set; } = string.Empty;
        public string PrefixForMFRName { get; set; } = string.Empty;
        public string PrefixForMFRPartNo { get; set; } = string.Empty;
        public string DescriptionToGenerate { get; set; } = string.Empty;
        public string TruncationType { get; set; } = string.Empty;
        public string DelimiterForTruncation { get; set; } = string.Empty;
        public string FirstOrderOfDataInDescription { get; set; } = string.Empty;
        public string SecondOrderOfDataInDescription { get; set; } = string.Empty;
        public string ThirdOrderOfDataInDescription { get; set; } = string.Empty;
        public string FourthOrderOfDataInDescription { get; set; } = string.Empty;
        public string FifthOrderOfDataInDescription { get; set; } = string.Empty;
        public bool IsNounExcluded { get; set; } = false;
        public bool IsModifierExcluded { get; set; } = false;
        public bool IsAttributeNameExcluded { get; set; } = false;
        public bool IsAttributeValueExcluded { get; set; } = false;
        public bool IsAdditionalInformationExcluded { get; set; } = false;
        public bool IsMFRNameExcluded { get; set; } = false;
        public bool IsMFRPartNoExcluded { get; set; } = false;
        public bool IsToInterpretAdditionalInformation { get; set; } = false;
        public bool IsToInterpretAllAttributeValues {  get; set; } = false;
        public bool IsToIncludeAttributeNameFromAdditionalInformation { get; set; } = false;
        public bool IsToIncludeMaximumValues {  get; set; } = false;
        public bool IsNounToBeAbbreviated { get; set; } = false;
        public bool IsModifierToBeAbbreviated { get; set; } = false;
        public bool IsAttributeNameToBeAbbreviated { get; set; } = false;
        public bool IsAttributeValueToBeAbbreviated { get; set; } = false;
        public bool IsAdditionalInformationToBeAbbreviated { get; set; } = false;
        public bool IsMFRNameToBeAbbreviated { get; set; } = false;
        public bool IsToApplyIdentifiers { get; set; } = false;
        public bool IsToAddSpaceBeforeORAfterIdentifier { get; set; } = false;
        public bool IsToApplyIdentifierToAdditionalInformation { get; set; } = false;
        public bool IsToIncludeAttributeNameWithNULLValues { get; set; } = false;
        public bool IsToIncludeAllOtherMFRNames { get; set; } = false;
        public bool IsToIncludeAllOtherMFRPartNos { get; set; } = false;
        public bool IsToPrefixAllMFRNames { get; set; } = false;
        public bool IsToPrefixAllMFRPartNos { get; set; } = false;
        public int CharacterLimit { get; set; } = 0;
        public int DGSettingID { get; set; } = 0;
    }
}