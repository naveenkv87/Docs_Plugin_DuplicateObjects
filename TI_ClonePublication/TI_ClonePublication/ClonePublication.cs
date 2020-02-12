using System;
using System.ComponentModel.Composition;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Trisoft.InfoShare.Plugins.SDK;
using Trisoft.InfoShare.Plugins.SDK.BackgroundTasks;
using Trisoft.InfoShare.API25;
using System.Xml;

namespace TI_ClonePublication
{
    /// <summary>
    /// Custom plugin that will be triggered as part of the background task
    /// </summary>
    // This attribute is used to make the class discoverable by the plugin engine.
    [Export("ClonePublication", typeof(IBackgroundTaskHandler))]
    [PartCreationPolicy(CreationPolicy.NonShared)]
    public class ClonePublication : IBackgroundTaskHandler
    {
        /// <summary>
        /// The background task configuration
        /// </summary>
        private IBackgroundTaskHandlerConfiguration _configuration;

        private IEventMonitor _eventMonitor;
        private long _progressId;

        private ILogService _logService;

        //private List<long> providedLngRefs = null;

        private string _publicationId = string.Empty;
        private string _oldFolderId = string.Empty;
        private string _newFolderId = string.Empty;
        private string sourcePubDetails = string.Empty;

        private string pubName = string.Empty;
        private string pubBaselineName = string.Empty;
        private string pubVersion = string.Empty;
        private XmlNodeList pubOutPutFormat;
        private string pubLangComb = string.Empty;
        private string pubWorkingLang = string.Empty;
        private string pubResolution = string.Empty;

        private List<Models.IshObjectDetails> SourceObjects = new List<Models.IshObjectDetails>();

        public Dictionary<string, string> objectGUIDMaps = new Dictionary<string, string>();

        private string _action = string.Empty;

        /// <summary>
        /// The name of the event - "CLONE PUBLICATION".
        /// </summary>
        public string EventTypeName
        {
            get;
            set;
        }

        /// <summary>
        /// This method will be called by the background task service after the instance of the handler is created.
        /// </summary>
        /// <param name="configuration">Parameters configured in the XML Background Task Settings.</param>
        public void Initialize(IBackgroundTaskHandlerConfiguration configuration)
        {
            _configuration = configuration;
            _configuration.Parameters.TryGetValue("Action", out _action);

            _logService = _configuration.LogService;

            string batchSizeString = string.Empty;
        }

        /// <summary>
        /// This method will be called by the background task service to run the handler.
        /// </summary>
        /// <param name="context">The context of the hander execution.</param>
        public void Run(IBackgroundTaskHandlerContext context)
        {
            _logService.Debug("Started plugin {0}", this.EventTypeName);
            _eventMonitor = context.EventMonitor;
            _progressId = context.ProgressId;

            try
            {
                int maxProgress = 100;
                int currentProgress = 0;
                List<string> pubIds = new List<string>();

                _eventMonitor.AddEventDetailWithProgress(_progressId, EventLevel.Information, "Starting the clone process", "Starting the clone process using the attached eventdata", DetailStatus.Success, EventDataType.Xml, context.InputDataStream, currentProgress, maxProgress);
                _configuration.LogService.Info("{0} - Starting the process", this.EventTypeName);

                // Read out and initialize the supplied event data supplied information
                XDocument eventData = XDocument.Load(context.InputDataStream);

                //initializeInfo(eventData);

                if (eventData.Root.Element("newfolderid") != null)
                {
                    _newFolderId = eventData.Root.Element("newfolderid").Value;
                }
                if (eventData.Root.Element("oldfolderid") != null)
                {
                    _oldFolderId = eventData.Root.Element("oldfolderid").Value;
                }
                if (eventData.Root.Element("lngcardids") != null)
                {
                    _publicationId = eventData.Root.Element("lngcardids").Value;
                }

                _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "Cloning Publication", _publicationId, DetailStatus.Success);
                _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "From Folder", _oldFolderId, DetailStatus.Success);
                _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "To Folder", _newFolderId, DetailStatus.Success);

                if (_publicationId != string.Empty || _publicationId != null)
                {
                    Trisoft.InfoShare.API25.Baseline baselineObj = new Trisoft.InfoShare.API25.Baseline();
                    Trisoft.InfoShare.API25.PublicationOutput pubObj = new Trisoft.InfoShare.API25.PublicationOutput();
                    Trisoft.InfoShare.API25.DocumentObj docObj = new Trisoft.InfoShare.API25.DocumentObj();

                    pubIds.Add(_publicationId);

                    //Get Latest version of the publication
                    string _versionXMLRequestedMetaData = "<ishfields>" +
                                                    "<ishfield name='VERSION' level='version'/>" +
                                                 "</ishfields>";

                    string _sourcePubVersionDetails = pubObj.RetrieveMetadata(pubIds, Trisoft.InfoShare.API.Enumerations.API25.StatusFilter.ISHNoStatusFilter, null, _versionXMLRequestedMetaData);
                    string _latestVersionRef = GetLatestPublicationVersionRef(_sourcePubVersionDetails);
                    _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "Version Ref of latest version of Publication", _latestVersionRef, DetailStatus.Success);
                    List<long> _ishVersioRefs = new List<long>
                    {
                        long.Parse(_latestVersionRef)
                    };

                    //Get Publication Details
                    string psXMLRequestedMetaData = "<ishfields>" +
                                                        "<ishfield name='FTITLE' ishvaluetype='value' level='logical'/>" +
                                                        "<ishfield name='VERSION' level='version'/>" +
                                                        "<ishfield name='FISHOUTPUTFORMATREF' level='lng' ishvaluetype='element'/>" +
                                                        "<ishfield name='FISHPUBLNGCOMBINATION' ishvaluetype='value' level='lng'/>" +
                                                        "<ishfield name='FISHBASELINE' level='version'/>" +
                                                        "<ishfield name='FISHREQUIREDRESOLUTIONS' level='version'/>" +
                                                        "<ishfield name='FISHPUBSOURCELANGUAGES' level='version'/>" +
                                                 "</ishfields>";

                    string sourcePubDetails = pubObj.RetrieveMetadataByIshVersionRefs(_ishVersioRefs, Trisoft.InfoShare.API.Enumerations.API25.StatusFilter.ISHNoStatusFilter, null, psXMLRequestedMetaData);
                    XmlDocument _publicationDetails = new XmlDocument();
                    _publicationDetails.LoadXml(sourcePubDetails);

                    pubName = _publicationDetails.SelectNodes("//ishfield[@name='FTITLE']")[0].InnerText;
                    pubBaselineName = _publicationDetails.SelectNodes("//ishfield[@name='FISHBASELINE']")[0].InnerText;
                    pubVersion = _publicationDetails.SelectNodes("//ishfield[@name='VERSION']")[0].InnerText;
                    pubOutPutFormat = _publicationDetails.SelectNodes("//ishfield[@name='FISHOUTPUTFORMATREF']");
                    pubLangComb = _publicationDetails.SelectNodes("//ishfield[@name='FISHPUBLNGCOMBINATION']")[0].InnerText;
                    pubWorkingLang = _publicationDetails.SelectNodes("//ishfield[@name='FISHPUBSOURCELANGUAGES']")[0].InnerText;
                    pubResolution = _publicationDetails.SelectNodes("//ishfield[@name='FISHREQUIREDRESOLUTIONS']")[0].InnerText;

                    _eventMonitor.AddEventDetailWithProgress(_progressId, EventLevel.Information, "Retrieve Metadata", "Retrieve required metadata of the source publication", DetailStatus.Success, EventDataType.Xml, sourcePubDetails, currentProgress, maxProgress);

                    //Get Baseline of the publication
                    string pubBaseLineId = baselineObj.GetBaselineId(pubBaselineName);
                    string _baseLineResponse = baselineObj.GetBaseline(pubBaseLineId, null, out _baseLineResponse);
                    _eventMonitor.AddEventDetailWithProgress(_progressId, EventLevel.Information, "Retrieve Baseline", "Retrieve Baseline details", DetailStatus.Success, EventDataType.Xml, _baseLineResponse, currentProgress, maxProgress);

                    XmlDocument _sourcePubBaselineObjXML = new XmlDocument();
                    _sourcePubBaselineObjXML.LoadXml(_baseLineResponse);
                    SourceObjects = GetAllBaslineObjects(_sourcePubBaselineObjXML, docObj);

                    foreach (Models.IshObjectDetails _objRef in SourceObjects)
                    {
                        _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "Creating Duplicate of Object", _objRef.ObjSourceLogId, DetailStatus.Success);
                        //string _sourceObj = GetObjectDetails(_objRef.ObjSourceLogId, _objRef.ObjVersion, docObj);
                        //XmlDocument _sourceObjXML = new XmlDocument();
                        //_sourceObjXML.LoadXml(_sourceObj);

                        //XmlNode _objTypeIshObject = _sourceObjXML.SelectNodes("//ishobject")[0];
                        //string _ObjIshRef = _objTypeIshObject.Attributes["ishref"].Value;
                        //string _objType = _objTypeIshObject.Attributes["ishtype"].Value;
                        //_objRef.ObjType = _objType;

                        string _ObjIshRef = _objRef.ObjSourceLogId;
                        string _objType = _objRef.ObjType;

                        if (_objType != "ISHLibrary")
                        {
                            string _folderRef = CheckAndCreateFolder(_newFolderId, _objType);
                            string _objData = _objRef.ObjData;
                            string _objDataEDT = _objRef.ObjDataEdt;

                            //Get and process data
                            byte[] _updatedDataByte;

                            byte[] encodedDataAsBytes = Convert.FromBase64String(_objData);

                            if (_objType == "ISHIllustration")
                            {
                                _updatedDataByte = encodedDataAsBytes;
                            }
                            else
                            {
                                string returnValue = Encoding.Unicode.GetString(encodedDataAsBytes);
                                string _byteOrderMarkUtf8 = Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble());
                                if (returnValue.StartsWith(_byteOrderMarkUtf8))
                                {
                                    returnValue = returnValue.Remove(0, _byteOrderMarkUtf8.Length);
                                }
                                XmlDocument _sourceObjDetails = new XmlDocument();
                                _sourceObjDetails.LoadXml(returnValue);

                                XmlNode _ishNode = _sourceObjDetails.SelectSingleNode("processing-instruction('ish')");
                                _sourceObjDetails.RemoveChild(_ishNode);
                                string _updatedData = _sourceObjDetails.InnerXml;

                                _updatedData = _updatedData.Replace(_ObjIshRef, _objRef.ObjNewLogId);
                                _updatedData = updateReferenceInContent(_updatedData);

                                _updatedDataByte = Encoding.Unicode.GetBytes(_updatedData);
                            }

                            string _metadataList = GetRequiredMetadataforIshType(_objType);
                            XmlDocument _metadataXML = new XmlDocument();
                            _metadataXML.LoadXml(_metadataList);

                            XmlNodeList _listOfMadatoryMedata = _metadataXML.SelectNodes("//ishfielddefinition[@ismandatory='true' and @allowoncreate='true']");
                            string _objVersionRef = _objRef.ObjVersion;
                            string _cuurentObjMetadata = docObj.GetMetadata(_ObjIshRef, ref _objVersionRef, pubWorkingLang, string.Empty, GenerateMetadataReqXML(_listOfMadatoryMedata));

                            XmlDocument _parentObjMeta = new XmlDocument();
                            _parentObjMeta.LoadXml(_cuurentObjMetadata);

                            string _reqMeta = _parentObjMeta.SelectNodes("//ishfields")[0].InnerXml;
                            string _createReqMetadata = "<ishfields>" + _reqMeta + "<ishfield name='FSTATUS' level='lng'>Draft</ishfield></ishfields>";
                            string _defaultobjVersion = "1";
                            string _newObjLogId = _objRef.ObjNewLogId;

                            docObj.Create(long.Parse(_folderRef), GetObjectType(_objType), ref _newObjLogId, ref _defaultobjVersion, pubWorkingLang, "Low", _createReqMetadata, _objDataEDT, _updatedDataByte);
                            _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "Duplicated object with new GUID", _newObjLogId, DetailStatus.Success);
                        }
                        else
                        {
                            _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "Creating Duplicate of Object is skipped because it is Library object", _objRef.ObjSourceLogId, DetailStatus.Success);
                            _objRef.ObjNewLogId = _objRef.ObjSourceLogId;
                        }
                    }

                    Models.IshObjectDetails _MapObject = SourceObjects.Find(x => x.ObjType == "ISHMasterDoc");

                    string newPubMeta = "<ishfields>" +
                                            "<ishfield name='FTITLE' level='logical'>" + pubName + "</ishfield>" +
                                            "<ishfield name='FISHREQUIREDRESOLUTIONS' level='version'>" + pubResolution + "</ishfield>" +
                                            "<ishfield name='FISHPUBSOURCELANGUAGES' level='version'>" + pubWorkingLang + "</ishfield>" +
                                            "<ishfield name='FISHMASTERREF' level='version'>" + _MapObject.ObjNewLogId + "</ishfield>" +
                                        "</ishfields>";

                    string _folderlocation = CheckAndCreateFolder(_newFolderId, "ISHPublication");
                    string newPubGUID = "GUID-" + System.Guid.NewGuid().ToString().ToUpper();
                    string _defaultPubversion = "1";                   

                    foreach (XmlNode _sourceOutputFormats in _publicationDetails.SelectNodes("//ishobject"))
                    {
                        XmlDocument _currentObjDetails = new XmlDocument();
                        _currentObjDetails.LoadXml(_sourceOutputFormats.OuterXml);

                        if (_currentObjDetails.SelectNodes("//ishfield[@name='FISHPUBLNGCOMBINATION']")[0].InnerText == pubWorkingLang)
                        {
                            string requiredOutFormat = _currentObjDetails.SelectNodes("//ishfield[@name='FISHOUTPUTFORMATREF']")[0].InnerText;
                            pubObj.Create(long.Parse(_folderlocation), ref newPubGUID, ref _defaultPubversion, requiredOutFormat, pubLangComb, newPubMeta);
                        }
                    }

                    _eventMonitor.AddEventDetail(_progressId, EventLevel.Information, "New Publication Created", newPubGUID, DetailStatus.Success);
                }
                else
                {
                    _logService.Info("No language card ids found in the supplied eventdata '{0}'.", eventData.ToString());
                }
                // Add a message to the logging pipeline
                _logService.Info("Background handler has finished successfully for the event " + this.EventTypeName);

                // Execution is successful, so end the event
                _eventMonitor.EndEvent(_progressId, ProgressStatus.Success, 1, 1);
            }
            catch (Exception ex)
            {
                // Add a message to the logging pipeline
                _logService.ErrorException("Background handler has failed for the event " + this.EventTypeName, ex);

                // Background task service may retry the execution if configured, in that case, you cannot end the event yet
                if (!context.WillRetryOnException(ex))
                {
                    // If the retries are not configured, the service has exhausted the retry attempts, you need to end the event
                    _eventMonitor.EndEvent(_progressId, ProgressStatus.Failed, 1, 1);
                }
                throw;
            }
        }

        /// <summary>
        /// This Method will loop the result and returns the versionref of latest version of selected Publication
        /// </summary>
        private string GetLatestPublicationVersionRef(string _sourcePubVersionDoc)
        {
            XmlDocument _publicationVersionRefDetails = new XmlDocument();
            _publicationVersionRefDetails.LoadXml(_sourcePubVersionDoc);
            XmlNodeList _listOfPubObjects = _publicationVersionRefDetails.SelectNodes("//ishobject");

            string _latestPubVersionRef = _listOfPubObjects[0].Attributes["ishversionref"].Value;
            if (_listOfPubObjects.Count > 1)
            {
                foreach (XmlNode _eachNode in _listOfPubObjects)
                {
                    string _VersionofNodeinloop = _eachNode.Attributes["ishversionref"].Value;
                    if (long.Parse(_VersionofNodeinloop) > long.Parse(_latestPubVersionRef))
                    {
                        _latestPubVersionRef = _VersionofNodeinloop;
                    }
                }
            }
            return _latestPubVersionRef;
        }

        private List<Models.IshObjectDetails> GetAllBaslineObjects(XmlDocument objXML, Trisoft.InfoShare.API25.DocumentObj documentObj25)
        {
            List<Models.IshObjectDetails> objList = new List<Models.IshObjectDetails>();
            XmlNodeList _baseObjNodes = objXML.SelectNodes("//object");
            foreach (XmlNode baseObj in _baseObjNodes)
            {
                string _objRef = baseObj.Attributes["ref"].Value;
                string _objVar = baseObj.Attributes["versionnumber"].Value;

                string _sourceObj = GetObjectDetails(_objRef, _objVar, documentObj25);
                XmlDocument _sourceObjXML = new XmlDocument();
                _sourceObjXML.LoadXml(_sourceObj);

                XmlNode _objTypeIshObject = _sourceObjXML.SelectNodes("//ishobject")[0];
                string _ObjIshRef = _objTypeIshObject.Attributes["ishref"].Value;
                string _objType = _objTypeIshObject.Attributes["ishtype"].Value;

                string _objData = _sourceObjXML.SelectNodes("//ishdata")[0].InnerText;
                string _objDataEDT = _sourceObjXML.SelectNodes("//ishdata")[0].Attributes["edt"].Value;

                string newGUID = string.Empty;
                if (_objType != "ISHLibrary")
                {
                    newGUID = "GUID-" + System.Guid.NewGuid().ToString().ToUpper();
                }
                else
                {
                    newGUID = _ObjIshRef;
                }

                objectGUIDMaps.Add(_objRef, newGUID);
                Models.IshObjectDetails _objDet = new Models.IshObjectDetails
                {
                    ObjSourceLogId = _objRef,
                    ObjVersion = _objVar,
                    ObjNewLogId = newGUID,
                    ObjType = _objType,
                    ObjData = _objData,
                    ObjDataEdt = _objDataEDT
                };
                if (!objList.Contains(_objDet))
                {
                    objList.Add(_objDet);
                }
            }
            return objList;
        }

        private string GetObjectDetails(string logicalId, string objVersion, Trisoft.InfoShare.API25.DocumentObj documentObj25)
        {
            string _resObjList = string.Empty;
            try
            {
                _resObjList = documentObj25.GetObject(logicalId, ref objVersion, pubWorkingLang, pubResolution, null, null);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _resObjList;
        }

        private string CheckAndCreateFolder(string destFolder, string folderType)
        {
            string _folderId = string.Empty;
            try
            {
                string _folderSelectionQuery = "//ishfolder[@ishfoldertype='" + folderType + "']";
                Trisoft.InfoShare.API25.Folder foldersObj = new Trisoft.InfoShare.API25.Folder();
                string subFolders = foldersObj.GetSubFolders(long.Parse(destFolder));

                XmlDocument _subFolderDetails = new XmlDocument();
                _subFolderDetails.LoadXml(subFolders);

                XmlNode requiredFolderType = _subFolderDetails.SelectNodes(_folderSelectionQuery)[0];
                if (requiredFolderType != null)
                {
                    _folderId = requiredFolderType.Attributes["ishfolderref"].Value;
                }
                else
                {
                    _folderId = foldersObj.Create(long.Parse(destFolder), GetFolderName(folderType), string.Empty, null, GetFolderType(folderType)).ToString();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _folderId;
        }

        private string GetFolderName(string fldType)
        {
            string _fldName = string.Empty;
            switch (fldType)
            {
                case "ISHModule":
                    _fldName = "Topics";
                    break;
                case "ISHIllustration":
                    _fldName = "Images";
                    break;
                case "ISHLibrary":
                    _fldName = "Libraries";
                    break;
                case "ISHMasterDoc":
                    _fldName = "Maps";
                    break;
                case "ISHNone":
                    _fldName = "New Folder";
                    break;
                case "ISHPublication":
                    _fldName = "Publications";
                    break;
                case "ISHReference":
                    _fldName = "References";
                    break;
                case "ISHTemplate":
                    _fldName = "Templates";
                    break;
            }
            return _fldName;
        }

        private Trisoft.InfoShare.API.Enumerations.API25.IshFolderType GetFolderType(string fldType)
        {
            switch (fldType)
            {
                case "ISHModule":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHModule;
                case "ISHIllustration":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHIllustration;
                case "ISHLibrary":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHLibrary;
                case "ISHMasterDoc":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHMasterDoc;
                case "ISHNone":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHNone;
                case "ISHPublication":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHPublication;
                case "ISHReference":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHReference;
                case "ISHTemplate":
                    return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHTemplate;
            }
            return Trisoft.InfoShare.API.Enumerations.API25.IshFolderType.ISHNone;
        }

        private Trisoft.InfoShare.API.Enumerations.API25.ISHType GetObjectType(string fldType)
        {
            switch (fldType)
            {
                case "ISHModule":
                    return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHModule;
                case "ISHIllustration":
                    return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHIllustration;
                case "ISHLibrary":
                    return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHLibrary;
                case "ISHMasterDoc":
                    return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHMasterDoc;
                case "ISHTemplate":
                    return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHTemplate;
            }
            return Trisoft.InfoShare.API.Enumerations.API25.ISHType.ISHTemplate;
        }

        private string updateReferenceInContent(string sourceObjContent)
        {
            string updatedSourceContent = sourceObjContent;
            foreach (var eachMap in objectGUIDMaps)
            {
                updatedSourceContent = updatedSourceContent.Replace(eachMap.Key, eachMap.Value);
            }
            return updatedSourceContent;
        }

        private string GetRequiredMetadataforIshType(string requiredType)
        {
            string _allMetadata = string.Empty;
            try
            {
                Trisoft.InfoShare.API25.Settings settingsObj25 = new Trisoft.InfoShare.API25.Settings();
                List<string> _requiredType = new List<string>
                {
                    requiredType
                };
                _allMetadata = settingsObj25.RetrieveFieldSetupByIshType(_requiredType);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return _allMetadata;
        }

        private string GenerateMetadataReqXML(XmlNodeList _sourceMedataList)
        {
            string metadataXmlReq = "<ishfields>";
            foreach (XmlNode _metaObj in _sourceMedataList)
            {
                if (_metaObj.Attributes["name"].Value != "FSTATUS")
                {
                    metadataXmlReq = metadataXmlReq + "<ishfield name='" + _metaObj.Attributes["name"].Value + "' level='" + _metaObj.Attributes["level"].Value + "'/>";
                }
            }

            metadataXmlReq += "</ishfields>";
            return metadataXmlReq;
        }
        /// <summary>
        /// This method will be called by the plugin engine after all plugins have executed
        /// </summary>
        public void Dispose()
        {
            // In our example we don't need a cleanup
        }
    }
}
