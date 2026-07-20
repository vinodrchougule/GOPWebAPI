using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection.Emit;
using System.Web;
using GOPWebAPI.Helpers;
using GOPWebAPI.Models.GAT_Models;

namespace GOPWebAPI.DAL
{
    public class DALGATDescriptionGenerator
    {
        private readonly string _connectionString;

        public DALGATDescriptionGenerator(string connectionString)
        {
            _connectionString = connectionString;
        }

        #region Read Description Generator Setting Names
        public DataTable ReadDGSettingNames()
        {
            DataTable dtDetails = new DataTable();
            
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATDGGetSavedSettingNames", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }
            return dtDetails;
        }
        #endregion

        #region Save Description Generator Settings
        public string SaveDescriptionGeneratorSettings(DescriptionGenerator model)
        {
            string strResult = string.Empty;

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand("spGATDGSettingSetData", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        #region Add input Parameters with values
                        command.CommandTimeout = 500;
                        command.Parameters.AddWithValue("@DGSettingID", model.DGSettingID);
                        command.Parameters.AddWithValue("@SettingName", model.SettingName);
                        command.Parameters.AddWithValue("@IsNounExcluded", model.IsNounExcluded);
                        command.Parameters.AddWithValue("@IsModifierExcluded", model.IsModifierExcluded);
                        command.Parameters.AddWithValue("@IsAttributeNameExcluded", model.IsAttributeNameExcluded);
                        command.Parameters.AddWithValue("@IsAttributeValueExcluded", model.IsAttributeValueExcluded);
                        command.Parameters.AddWithValue("@IsAdditionalInformationExcluded", model.IsAdditionalInformationExcluded);
                        command.Parameters.AddWithValue("@IsMFRNameExcluded", model.IsMFRNameExcluded);
                        command.Parameters.AddWithValue("@IsMFRPartNoExcluded", model.IsMFRPartNoExcluded);
                        command.Parameters.AddWithValue("@SpecificModifierExcluded", model.SpecificModifierExcluded);
                        command.Parameters.AddWithValue("@IsToInterpretAdditionalInformation", model.IsToInterpretAdditionalInformation);
                        command.Parameters.AddWithValue("@IsToIncludeAttributeNameFromAdditionalInformation", model.IsToIncludeAttributeNameFromAdditionalInformation);
                        command.Parameters.AddWithValue("@IsToIncludeMaximumValues", model.IsToIncludeMaximumValues);
                        command.Parameters.AddWithValue("@IsToInterpretAllAttributeValues", model.IsToInterpretAllAttributeValues);
                        command.Parameters.AddWithValue("@IsNounToBeAbbreviated", model.IsNounToBeAbbreviated);
                        command.Parameters.AddWithValue("@IsModifierToBeAbbreviated", model.IsModifierToBeAbbreviated);
                        command.Parameters.AddWithValue("@IsAttributeNameToBeAbbreviated", model.IsAttributeNameToBeAbbreviated);
                        command.Parameters.AddWithValue("@IsAttributeValueToBeAbbreviated", model.IsAttributeValueToBeAbbreviated);
                        command.Parameters.AddWithValue("@IsAdditionalInformationToBeAbbreviated", model.IsAdditionalInformationToBeAbbreviated);
                        command.Parameters.AddWithValue("@IsMFRNameToBeAbbreviated", model.IsMFRNameToBeAbbreviated);
                        command.Parameters.AddWithValue("@AbbreviationFileName", model.UploadedAbbreviationFileName);
                        command.Parameters.AddWithValue("@DelimiterAfterNoun", model.DelimiterAfterNoun);
                        command.Parameters.AddWithValue("@DelimiterAfterModifier", model.DelimiterAfterModifier);
                        command.Parameters.AddWithValue("@DelimiterAfterAttributeName", model.DelimiterAfterAttributeName);
                        command.Parameters.AddWithValue("@DelimiterAfterAttributeValue", model.DelimiterAfterAttributeValue);
                        command.Parameters.AddWithValue("@DelimiterAfterAdditionalInformation", model.DelimiterAfterAdditionalInformation);
                        command.Parameters.AddWithValue("@DelimiterAfterMFRName", model.DelimiterAfterMFRName);
                        command.Parameters.AddWithValue("@DelimiterAfterMFRPartNo", model.DelimiterAfterMFRPartNo);
                        command.Parameters.AddWithValue("@MultipleValuesSeparator", model.MultipleValuesSeparator);
                        command.Parameters.AddWithValue("@IsToApplyIdentifiers", model.IsToApplyIdentifiers);
                        command.Parameters.AddWithValue("@IdentifierFileName", model.UploadedIdentifierFileName);
                        command.Parameters.AddWithValue("@IsToAddSpaceBeforeORAfterIdentifier", model.IsToAddSpaceBeforeORAfterIdentifier);
                        command.Parameters.AddWithValue("@IsToApplyIdentifierToAdditionalInformation", model.IsToApplyIdentifierToAdditionalInformation);
                        command.Parameters.AddWithValue("@PrefixForAdditionalInformation", model.PrefixForAdditionalInformation);
                        command.Parameters.AddWithValue("@PrefixForMFRName", model.PrefixForMFRName);
                        command.Parameters.AddWithValue("@PrefixForMFRPartNo", model.PrefixForMFRPartNo);
                        command.Parameters.AddWithValue("@IsToIncludeAttributeNameWithNULLValues", model.IsToIncludeAttributeNameWithNULLValues);
                        command.Parameters.AddWithValue("@IsToIncludeAllOtherMFRNames", model.IsToIncludeAllOtherMFRNames);
                        command.Parameters.AddWithValue("@IsToIncludeAllOtherMFRPartNos", model.IsToIncludeAllOtherMFRPartNos);
                        command.Parameters.AddWithValue("@IsToPrefixAllMFRNames", model.IsToPrefixAllMFRNames);
                        command.Parameters.AddWithValue("@IsToPrefixAllMFRPartNos", model.IsToPrefixAllMFRPartNos);
                        command.Parameters.AddWithValue("@DescriptionToGenerate", model.DescriptionToGenerate);
                        command.Parameters.AddWithValue("@TruncationType", model.TruncationType);
                        command.Parameters.AddWithValue("@CharacterLimit", model.CharacterLimit);
                        command.Parameters.AddWithValue("@DelimiterForTruncation", model.DelimiterForTruncation);
                        command.Parameters.AddWithValue("@FirstOrderOfDataInDescription", model.FirstOrderOfDataInDescription);
                        command.Parameters.AddWithValue("@SecondOrderOfDataInDescription", model.SecondOrderOfDataInDescription);
                        command.Parameters.AddWithValue("@ThirdOrderOfDataInDescription", model.ThirdOrderOfDataInDescription);
                        command.Parameters.AddWithValue("@FourthOrderOfDataInDescription", model.FourthOrderOfDataInDescription);
                        command.Parameters.AddWithValue("@FifthOrderOfDataInDescription", model.FifthOrderOfDataInDescription);
                        #endregion

                        // Output parameter to capture the message
                        SqlParameter outMsgParam = new SqlParameter("@Result", SqlDbType.VarChar, 500)
                        {
                            Direction = ParameterDirection.Output
                        };
                        command.Parameters.Add(outMsgParam);

                        connection.Open();
                        command.ExecuteNonQuery();
                        strResult = Convert.ToString(command.Parameters["@Result"].Value);
                        connection.Close();
                    }
                }

                return strResult;
            }
            catch (Exception ex) 
            {
                ExceptionLogging.SendExceptionToDB(ex);
                strResult = "Error: " + ex.Message;
                return strResult;
            }
        }
        #endregion

        #region Read Description Generator Saved Setting
        public DataTable ReadDescriptionGeneratorSavedSetting(string SettingName)
        {
            DataTable dtDetails = new DataTable();

            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("spGATDGSettingGetData", con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.CommandTimeout = 500;
                        cmd.Parameters.AddWithValue("@SettingName", SettingName);
                        con.Open();
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dtDetails);
                        con.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogging.SendExceptionToDB(ex);
            }

            return dtDetails;
        }
        #endregion
    }
}