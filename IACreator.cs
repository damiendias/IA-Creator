using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Web;
using EPiServer;
using EPiServer.Core;
using EPiServer.DataAbstraction;
using EPiServer.PlugIn;
using EPiServer.Scheduler;
using EPiServer.ServiceLocation;
using System.Linq;
using System.Text.RegularExpressions;
using EPiServer.Logging;
using EPiServer.Filters;
using ExcelDataReader;

namespace IACreator
{
    [ScheduledPlugIn(DisplayName = "IA Creator")]
    public class IACreatorJob : ScheduledJobBase
    {
        private bool _stopSignaled;
        public const string CSVFilePathKeyName = "CSVFilePath";
        public const string ParentPageIdKeyName = "ParentPageId";
        public const string DefaultFileName = "Content.xlsx";

        private Injected<IContentTypeRepository> _contentTypeRepository;
        private Injected<IContentRepository> _contentRepository;

        private static readonly ILogger _logger = LogManager.GetLogger();

        /// <summary>
        /// Gets the CSV file path if specified in appsettings otherwise defaults a specific file in app_data folder
        /// </summary>
        public string CSVFilePath => String.IsNullOrEmpty(ConfigurationManager.AppSettings[CSVFilePathKeyName])
                                        ? Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data", DefaultFileName)
                                        : Path.Combine(HttpRuntime.AppDomainAppPath, "App_Data", ConfigurationManager.AppSettings[CSVFilePathKeyName]);


        /// <summary>
        /// Gets the parent page to start the import. First tries to retrieve the id from appsettings and defaults to the
        /// start page if not provided.
        /// </summary>
        public ContentReference ParentContentRef => String.IsNullOrEmpty(ConfigurationManager.AppSettings[ParentPageIdKeyName])
                                                        ? PageReference.StartPage
                                                        : new ContentReference(Int32.Parse(ConfigurationManager.AppSettings[ParentPageIdKeyName]));


        public IACreatorJob()
        {
            IsStoppable = true;
        }

        /// <summary>
        /// Called when a user clicks on Stop for a manually started job, or when ASP.NET shuts down.
        /// </summary>
        public override void Stop()
        {
            _stopSignaled = true;
        }



        /// <summary>
        /// Called when a scheduled job executes
        /// </summary>
        /// <returns>A status message to be stored in the database log and visible from admin mode</returns>
        public override string Execute()
        {
            if (!File.Exists(CSVFilePath))
            {
                return "No file found to process";
            }

            int contentCount = 0;

            using (var stream = File.Open(CSVFilePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateOpenXmlReader(stream))
                {
                    var table = reader.AsDataSet().Tables[0];

                    List<DataRow> rows = (from DataRow row in table.Rows
                                          select row).ToList();

                    var currentParentIndex = 0;
                    ContentReference currentParent = null;
                    ContentReference currentLevel = null;
                    PageData pageContent = null;

                    // parse each row and instantiate a new IContent object based on the type from the spreadsheet
                    foreach (var r in rows)
                    {
                        var contentTypeName = r.Field<string>(0);
                        var contentName = r.Field<string>(1);
                        int parentIndex = Convert.ToInt32(r.Field<double>(2));
                        var fields = r.ItemArray.Skip(3).Where(f => !String.IsNullOrWhiteSpace(f as string));
                        var contentType = _contentTypeRepository.Service.Load(contentTypeName);

                        if (contentType == null)
                        {
                            throw new NullReferenceException("The content type provided could not be matched");
                        }

                        var parentPage = GetParent(parentIndex, currentParentIndex, currentParent, currentLevel);
                        currentParent = parentPage;
                        pageContent = _contentRepository.Service.GetDefault<PageData>(parentPage, contentType.ID);

                        pageContent.Name = contentName;

                        foreach (var field in fields)
                        {

                            string fieldDetails = field.ToString();
                            var propertyName = fieldDetails.Substring(0, fieldDetails.IndexOf(":"));

                            if (string.IsNullOrWhiteSpace(propertyName)) continue;

                            var startIndex = fieldDetails.IndexOf(":") + 1;
                            var length = fieldDetails.Length - startIndex;
                            var propertyValue = fieldDetails.Substring(fieldDetails.IndexOf(":") + 1, length);

                            if (string.IsNullOrWhiteSpace(propertyValue)) continue;

                            propertyValue = Regex.Replace(propertyValue, @"\r\n?|\n", "<br />");
                            try
                            {
                                pageContent.Property[propertyName].Value = propertyValue;

                            }
                            catch (Exception ex)
                            {
                                _logger.Log(Level.Error, $"Failed to store property value with property name {propertyName}. Full exception - {ex.ToString()}");
                            }
                        }
                        pageContent.Property[MetaDataProperties.PagePeerOrder].Value = contentCount;
                        pageContent.Property[MetaDataProperties.PageChildOrderRule].Value = FilterSortOrder.Index;
                        //pageContent.

                        currentLevel = _contentRepository.Service.Save(pageContent, EPiServer.DataAccess.SaveAction.SkipValidation | EPiServer.DataAccess.SaveAction.Publish);

                        currentParentIndex = parentIndex;
                        contentCount++;
                    }
                }
            }

            if (_stopSignaled)
            {
                return "Stop of job was called";
            }

            return $"{contentCount} items imported.";
        }

        /// <summary>
        /// Determines the parent content reference based on the integer sequence from the spreadsheet.
        /// </summary>
        /// <param name="parentIndex"></param>
        /// <param name="currentParentIndex"></param>
        /// <param name="currentParent"></param>
        /// <param name="currentLevel"></param>
        /// <returns>ContentReference</returns>
        private ContentReference GetParent(int parentIndex, int currentParentIndex, ContentReference currentParent, ContentReference currentLevel)
        {
            if (parentIndex == 0)
            {
                return ParentContentRef;
            }
            else if (parentIndex == currentParentIndex)
            {
                return currentParent;
            }
            else if (parentIndex > currentParentIndex)
            {
                return currentLevel;
            }
            else
            {
                var parent = _contentRepository.Service.Get<IContent>(currentParent);
                return parent.ParentLink;
            }
        }
    }
}
