/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../Model/LiveWorkItemData.ts" />
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var LinkCommand;
    (function (LinkCommand) {
        var UCIAppContextLinkCommandFactory = (function () {
            function UCIAppContextLinkCommandFactory() {
            }
            UCIAppContextLinkCommandFactory.createAppContextLinkCommand = function () {
                if (window.top.IsUSD) {
                    return new LinkCommand.USDAppContextLinkCommand();
                }
                else {
                    return new LinkCommand.UCIAppContextLinkCommand();
                }
            };
            return UCIAppContextLinkCommandFactory;
        }());
        LinkCommand.UCIAppContextLinkCommandFactory = UCIAppContextLinkCommandFactory;
    })(LinkCommand = OmniChannelAgentSDK.LinkCommand || (OmniChannelAgentSDK.LinkCommand = {}));
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../Model/LiveWorkItemData.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var LinkCommand;
    (function (LinkCommand) {
        var UCIAppContextLinkCommand = (function () {
            function UCIAppContextLinkCommand() {
                this.sessionLiveworkitemDataMap = {};
            }
            UCIAppContextLinkCommand.prototype.shouldDisplayLinkCommand = function () {
                //Verify whether one of the tabs within the session is a Customer Summary tab
                try {
                    var tabsInSession = this.getCurrentFocussedSession().tabs.getAll();
                    for (var i = 0; i < tabsInSession.getLength(); i++) {
                        var tab = tabsInSession.get(i);
                        if (this.isCustomerSummaryTab(tab)) {
                            this.fetchLiveWorkItemData();
                            return true;
                        }
                        return false;
                    }
                }
                catch (error) {
                    throw error;
                }
            };
            UCIAppContextLinkCommand.prototype.isCustomerSummaryTab = function (tab) {
                var url = tab.currentUrl;
                if (url.indexOf("pagetype=entityrecord") !== -1 && url.indexOf("etn=" + OmniChannelAgentSDK.Constants.CONVERSATION_ENTITY_LOGICAL_NAME) !== -1 && url.indexOf("formid=" + OmniChannelAgentSDK.Constants.CUSTOMER_SUMMARY_FORM_ID) !== -1) {
                    return true;
                }
                return false;
            };
            UCIAppContextLinkCommand.prototype.getCurrentFocussedSession = function () {
                try {
                    var parentWindow = window.top;
                    return parentWindow.Xrm.App.sessions.getFocusedSession();
                }
                catch (error) {
                    throw error;
                }
            };
            UCIAppContextLinkCommand.prototype.fetchLiveWorkItemData = function () {
                var _this = this;
                return new Promise(function (resolve, reject) {
                    try {
                        var currentSession_1 = _this.getCurrentFocussedSession();
                        if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(_this.sessionLiveworkitemDataMap[currentSession_1.sessionId])) {
                            resolve(_this.sessionLiveworkitemDataMap[currentSession_1.sessionId]);
                        }
                        else {
                            Microsoft.CIFramework.Internal.sendGenericMessage(new Map(), OmniChannelAgentSDK.Constants.GET_LIVEWORKITEM_DATA_EVENT, false).then(function (response) {
                                var liveWorkitemData = JSON.parse(response);
                                _this.sessionLiveworkitemDataMap[currentSession_1.sessionId] = liveWorkitemData;
                                resolve(_this.sessionLiveworkitemDataMap[currentSession_1.sessionId]);
                            }, function (error) {
                                reject(error);
                            });
                        }
                    }
                    catch (error) {
                        reject(error);
                    }
                });
            };
            UCIAppContextLinkCommand.prototype.raiseLinkingDoneEvent = function (entityName, id, liveWorkItemId) {
                var eventPayload = {
                    entityName: entityName,
                    records: [id.toString()],
                    LiveWorkItemId: liveWorkItemId
                };
                var zfpEvent = new CustomEvent(OmniChannelAgentSDK.Constants.LINKING_DONE_EVENT, { detail: eventPayload });
                window.top.dispatchEvent(zfpEvent);
            };
            return UCIAppContextLinkCommand;
        }());
        LinkCommand.UCIAppContextLinkCommand = UCIAppContextLinkCommand;
    })(LinkCommand = OmniChannelAgentSDK.LinkCommand || (OmniChannelAgentSDK.LinkCommand = {}));
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../Model/LiveWorkItemData.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var LinkCommand;
    (function (LinkCommand) {
        var USDAppContextLinkCommand = (function () {
            function USDAppContextLinkCommand() {
            }
            USDAppContextLinkCommand.prototype.shouldDisplayLinkCommand = function () {
                // Verify if window.ocContext is initialized
                try {
                    // TODO : Test and fix the USD path
                    var ocContext = window.top.ocContext;
                    if (ocContext && ocContext.config && ocContext.config.sessionParams) {
                        return true;
                    }
                    return false;
                }
                catch (error) {
                    throw error;
                }
            };
            USDAppContextLinkCommand.prototype.fetchLiveWorkItemData = function () {
                return new Promise(function (resolve, reject) {
                    try {
                        var ocContextSessionParams = window.top.ocContext.config.sessionParams;
                        var liveWorkItemData = {};
                        liveWorkItemData[OmniChannelAgentSDK.Constants.LIVEWORKITEM_ID_ATTR] = ocContextSessionParams.LiveWorkItemId;
                        liveWorkItemData[OmniChannelAgentSDK.Constants.LIVEWORKSTREAM_ID_ATTR] = ocContextSessionParams.LiveWorkStreamId;
                        resolve(liveWorkItemData);
                    }
                    catch (error) {
                        reject(error);
                    }
                });
            };
            USDAppContextLinkCommand.prototype.raiseLinkingDoneEvent = function (entityName, id, liveWorkItemId) {
                id = id.replace(/[{}]/g, "");
                var eventPayload = {
                    entityName: entityName,
                    records: [id],
                    LiveWorkItemId: liveWorkItemId
                };
                var usdEvent = "http://event?eventname=" + OmniChannelAgentSDK.Constants.LINKING_DONE_EVENT + "&PostData=";
                var stringifiedNotifyEventData = JSON.stringify(eventPayload);
                usdEvent = usdEvent + stringifiedNotifyEventData;
                this.sendToUSD(usdEvent);
            };
            USDAppContextLinkCommand.prototype.sendToUSD = function (url) {
                if (window.top.notifyUSD) {
                    window.top.notifyUSD(url);
                }
                else {
                    window.open(url);
                }
            };
            return USDAppContextLinkCommand;
        }());
        LinkCommand.USDAppContextLinkCommand = USDAppContextLinkCommand;
    })(LinkCommand = OmniChannelAgentSDK.LinkCommand || (OmniChannelAgentSDK.LinkCommand = {}));
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var Constants = (function () {
        function Constants() {
        }
        return Constants;
    }());
    Constants.LIVEWORKITEM_ID_ATTR = "liveWorkitemId";
    Constants.LIVEWORKSTREAM_ID_ATTR = "liveWorkstreamId";
    Constants.zeroGuid = "00000000-0000-0000-0000-000000000000";
    Constants.GET_LIVEWORKITEM_DATA_EVENT = "getliveworkitemdata";
    Constants.LINKING_DONE_EVENT = "OmnichannelSessionInlineSearchAndLink";
    Constants.CONTEXT_UPDATE_FAILED_STATUS = 3;
    Constants.CONVERSATION_ENTITY_LOGICAL_NAME = "msdyn_ocliveworkitem";
    Constants.CUSTOMER_SUMMARY_FORM_ID = "5fe86453-73ea-4821-b6dd-ddc06e1755a1";
    Constants.CUSTOMER_FIELDLOGICALNAME = "msdyn_customer";
    Constants.ISSUE_FIELDLOGICALNAME = "msdyn_issueid";
    Constants.OMNICHANNEL_AGENT_SDK_NAMESPACE = "Microsoft.Omnichannel";
    Constants.TELEMETRY_RESOURCE_URL = "/WebResources/msdyn_OcAriaTelemetryLogger.js";
    Constants.LINK_CONVERSATION_LIBRARY_RESOURCE_URL = "/WebResources/msdyn_LinkConversationLibrary.js";
    OmniChannelAgentSDK.Constants = Constants;
    var EntityNames = (function () {
        function EntityNames() {
        }
        return EntityNames;
    }());
    EntityNames.Incident = "incident";
    EntityNames.Account = "account";
    EntityNames.Contact = "contact";
    EntityNames.LiveWorkItem = "msdyn_ocliveworkitem";
    OmniChannelAgentSDK.EntityNames = EntityNames;
    var EntityAttributesNames = (function () {
        function EntityAttributesNames() {
        }
        return EntityAttributesNames;
    }());
    EntityAttributesNames.LWI_OCLiveWorkItemId = "msdyn_ocliveworkitemid";
    EntityAttributesNames.LWI_LiveWorkStreamId = "msdyn_liveworkstreamid";
    EntityAttributesNames.LWI_LastSessionId = "msdyn_lastsessionid";
    EntityAttributesNames.LWI_Statuscode = "statuscode";
    EntityAttributesNames.LWI_CreatedOn = "createdon";
    EntityAttributesNames.LWI_ActiveAgentId = "msdyn_activeagentid";
    OmniChannelAgentSDK.EntityAttributesNames = EntityAttributesNames;
    var EntityRelationshipNames = (function () {
        function EntityRelationshipNames() {
        }
        return EntityRelationshipNames;
    }());
    EntityRelationshipNames.Incident_Conversation = "msdyn_incident_msdyn_ocliveworkitem";
    EntityRelationshipNames.Account_Conversation = "msdyn_account_msdyn_ocliveworkitem_Customer";
    EntityRelationshipNames.Contact_Conversation = "msdyn_contact_msdyn_ocliveworkitem_Customer";
    OmniChannelAgentSDK.EntityRelationshipNames = EntityRelationshipNames;
    var LocalizationConstants = (function () {
        function LocalizationConstants() {
        }
        return LocalizationConstants;
    }());
    LocalizationConstants.resxWebResourceName = "msdyn_OmnichannelBase";
    // Localized string id
    LocalizationConstants.OC_LinkToConversationSuccessMessage = "OC_LinkToConversationSuccessMessage";
    LocalizationConstants.OC_LinkToConversationFailureMessage = "OC_LinkToConversationFailureMessage";
    LocalizationConstants.OC_LinkToConversationUpdatedMessage = "OC_PickInternalError";
    LocalizationConstants.OC_Undefined = "OC_Undefined";
    LocalizationConstants.OC_OpenConversationFailureMessage = "OC_OpenConversationFailureMessage";
    LocalizationConstants.OC_OpenConversationSuccessMessage = "OC_OpenConversationSuccessMessage";
    LocalizationConstants.OC_SendMessageToConversationFailureMessage = "OC_SendMessageToConversationFailureMessage";
    LocalizationConstants.OC_SendMessageToConversationSuccessMessage = "OC_SendMessageToConversationSuccessMessage";
    OmniChannelAgentSDK.LocalizationConstants = LocalizationConstants;
    /*
     * Constants for making CDS OData call
     */
    var ODataConstants = (function () {
        function ODataConstants() {
        }
        return ODataConstants;
    }());
    // Data parsing constants
    ODataConstants.formattedValueKey = "@OData.Community.Display.V1.FormattedValue";
    ODataConstants.lookupLogicalNameKey = "@Microsoft.Dynamics.CRM.lookuplogicalname";
    ODataConstants.lookupNavigationPropertyKey = "@Microsoft.Dynamics.CRM.associatednavigationproperty";
    ODataConstants.lookupFieldPrefix = "_";
    ODataConstants.lookupFieldSuffix = "_value";
    OmniChannelAgentSDK.ODataConstants = ODataConstants;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="../drop/AriaTelemetryLogger.d.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    /**
     * Link Command Telemetry Logger class
     */
    var TelemetryLogger = (function () {
        function TelemetryLogger() {
            this.initializeAriaTelemetryLogger();
        }
        // Gets static instance
        TelemetryLogger.Instance = function () {
            if (!TelemetryLogger.isInitialized) {
                this._instance = new TelemetryLogger();
                TelemetryLogger.isInitialized = true;
            }
            if (!TelemetryLogger.isTelemetryLoggerInitialized) {
                this._instance.initializeAriaTelemetryLogger();
            }
            return TelemetryLogger._instance;
        };
        /**
        * Initialize telemetry logger object
        */
        TelemetryLogger.prototype.initializeAriaTelemetryLogger = function () {
            if (Util.isNullOrUndefined(OmniChannelPackage.AriaTelemetryLogger)) {
                return;
            }
            try {
                this.telemetryLogger = OmniChannelPackage.AriaTelemetryLogger.Instance();
                TelemetryLogger.isTelemetryLoggerInitialized = true;
            }
            catch (e) {
            }
        };
        /**
         * Generate new request id
         */
        TelemetryLogger.prototype.getNewReqId = function () {
            return Util.generateNewGuid();
        };
        TelemetryLogger.getAWTEventProperties = function () {
            return (typeof AWTEventProperties !== typeof undefined) ?
                AWTEventProperties : window.top.AWTEventProperties;
        };
        /**
         * Send telemetry event: streamcommandevent
         * @param reqId
         * @param sessionId
         * @param liveWorkItemId
         * @param component
         * @param message
         * @param isError
         * @param errorObject
         * @param successMessage
         * @param addtionalDetails
         */
        TelemetryLogger.prototype.sendEvent = function (reqId, sessionId, liveWorkItemId, component, message, isError, errorObject, addtionalDetails) {
            if (TelemetryLogger.isInitialized && TelemetryLogger.isTelemetryLoggerInitialized) {
                try {
                    var eventProperties = new (TelemetryLogger.getAWTEventProperties())();
                    eventProperties.setName(EventNames.linkCommandEvent);
                    eventProperties.setProperty(FieldNames.reqid, reqId);
                    eventProperties.setProperty(FieldNames.sessionid, sessionId);
                    eventProperties.setProperty(FieldNames.liveworkitemid, liveWorkItemId);
                    eventProperties.setProperty(FieldNames.component, component);
                    eventProperties.setProperty(FieldNames.message, message);
                    eventProperties.setProperty(FieldNames.isError, isError);
                    var errorMessage = FieldValues.emptyString;
                    var errorStack = FieldValues.emptyString;
                    if (isError) {
                        if (!Util.isNullOrUndefined(errorObject)) {
                            errorMessage = errorObject.message;
                            errorStack = errorObject.stack;
                        }
                    }
                    eventProperties.setProperty(FieldNames.isError, isError);
                    eventProperties.setProperty(FieldNames.errorMessage, errorMessage);
                    eventProperties.setProperty(FieldNames.errorStack, errorStack);
                    eventProperties.setProperty(FieldNames.additionalDetails, addtionalDetails);
                    this.telemetryLogger.sendEvent(eventProperties);
                }
                catch (e) {
                    // Exception in parsing properties
                    var errorEvent = new (TelemetryLogger.getAWTEventProperties())();
                    errorEvent.setName(EventNames.telemetryErrorEvent);
                    errorEvent.setProperty(FieldNames.reqid, reqId);
                    this.telemetryLogger.sendEvent(errorEvent);
                }
            }
        };
        return TelemetryLogger;
    }());
    // Static instance fields
    TelemetryLogger.isInitialized = false;
    TelemetryLogger.isTelemetryLoggerInitialized = false;
    OmniChannelAgentSDK.TelemetryLogger = TelemetryLogger;
    /**
     * Events names
     */
    var EventNames = (function () {
        function EventNames() {
        }
        return EventNames;
    }());
    EventNames.telemetryErrorEvent = "telemetryerrorevent";
    EventNames.linkCommandEvent = "linkcommandevent";
    OmniChannelAgentSDK.EventNames = EventNames;
    /**
     * Component names for stream command
     */
    var Components = (function () {
        function Components() {
        }
        return Components;
    }());
    Components.linkRecordToConversation = "LinkCommandToConversation";
    Components.shouldDisplayLinkCommand = "ShouldDisplayLinkCommand";
    Components.localizedString = "LocalizedString";
    Components.unlinkRecordFromConversation = "UnlinkRecordFromConversation";
    Components.getConversationId = "GetConversationId";
    Components.initOmnichannelAgentSDK = "InitOmnichannelAgentSDK";
    Components.getConversations = "getConversations";
    Components.openConversation = "OpenConversation";
    Components.getLinkedRecords = "GetLinkedRecords";
    Components.sendMessageToConversation = "SendMessageToConversation";
    OmniChannelAgentSDK.Components = Components;
    /**
     * Field names for maintaining consistent schema
     */
    var FieldNames = (function () {
        function FieldNames() {
        }
        return FieldNames;
    }());
    FieldNames.reqid = "reqid";
    FieldNames.sessionid = "sessionid";
    FieldNames.liveworkitemid = "liveworkitemid";
    FieldNames.component = "component";
    FieldNames.message = "message";
    FieldNames.isError = "isError";
    FieldNames.errorMessage = "errorMessage";
    FieldNames.errorStack = "errorStack";
    FieldNames.additionalDetails = "details";
    OmniChannelAgentSDK.FieldNames = FieldNames;
    /**
     * Common field values
     */
    var FieldValues = (function () {
        function FieldValues() {
        }
        return FieldValues;
    }());
    FieldValues.emptyString = "";
    OmniChannelAgentSDK.FieldValues = FieldValues;
    var Util = (function () {
        function Util() {
        }
        /**
         * Utility function to check whether object is null or un-defined
         * @param object
         */
        Util.isNullOrUndefined = function (object) {
            return typeof (object) == "undefined" || object == null;
        };
        /**
        * Generate new guid
        */
        Util.generateNewGuid = function () {
            // possible hex chars for a guid
            var hexChars = "0123456789abcdef";
            var guidSize = 36;
            var guidString = "";
            for (var i = 0; i < guidSize; i++) {
                if (i === 14) {
                    // bits 12-15 set to 0010 - indicates version number of UUID RFC
                    guidString += "4";
                }
                else if (i === 8 || i === 13 || i === 18 || i === 23) {
                    // Dashes at 8, 13, 18, 23 (count begins at 0)
                    guidString += "-";
                }
                else if (i === 19) {
                    // bits 6-7 are reserved to zero and one resp.
                    var n = Math.floor(Math.random() * 0x10);
                    // tslint:disable-next-line: no-bitwise
                    guidString += hexChars.substr((n & 0x3) | 0x8, 1);
                }
                else {
                    guidString += hexChars.substr(Math.floor(Math.random() * 0x10), 1);
                }
            }
            return guidString;
        };
        return Util;
    }());
    OmniChannelAgentSDK.Util = Util;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var Utils = (function () {
        function Utils() {
        }
        /**
         * Helper function: To check if it is null or undefined
         * @param object
         */
        Utils.isNullOrUndefined = function (object) {
            return typeof (object) == "undefined" || object == null;
        };
        Utils.getSafe = function (fn, defaultVal) {
            try {
                return fn();
            }
            catch (e) {
                return defaultVal;
            }
        };
        Utils.getRelationShipNameByEntity = function (entityLogicalName) {
            switch (entityLogicalName) {
                case OmniChannelAgentSDK.EntityNames.Incident:
                    return OmniChannelAgentSDK.EntityRelationshipNames.Incident_Conversation;
                case OmniChannelAgentSDK.EntityNames.Account:
                    return OmniChannelAgentSDK.EntityRelationshipNames.Account_Conversation;
                case OmniChannelAgentSDK.EntityNames.Contact:
                    return OmniChannelAgentSDK.EntityRelationshipNames.Contact_Conversation;
            }
        };
        Utils.loadScript = function (url) {
            return new Promise(function (resolve, reject) {
                try {
                    var script_1 = document.createElement("script");
                    script_1.src = url;
                    script_1.type = 'text/javascript';
                    script_1.addEventListener('load', function () { return resolve(script_1); }, false);
                    script_1.addEventListener('error', function () { return reject(script_1); }, false);
                    window.top.document.getElementsByTagName("head")[0].appendChild(script_1);
                }
                catch (error) {
                    reject("Could not load script " + url + ", Error:" + error);
                }
            });
        };
        Utils.generateErrorObject = function (component, functionName, errorMessage, currentReqId, additionalDetails, error) {
            return {
                component: component,
                status: "failure",
                functionName: functionName,
                message: errorMessage,
                currentReqId: currentReqId,
                additionalDetails: additionalDetails,
                error: error
            };
        };
        Utils.generateSuccessPayload = function (entityLogicalName, recordId, liveWorkItemId, message, currentReqId, additionalDetails) {
            return {
                entityLogicalName: entityLogicalName,
                recordId: recordId,
                liveWorkItemId: liveWorkItemId,
                message: message,
                status: "success",
                additionalDetails: additionalDetails,
                currentReqId: currentReqId
            };
        };
        Utils.isTopWindowAccessible = function () {
            try {
                var currentWindow = window;
                currentWindow.top.location;
                return true;
            }
            catch (err) {
                return false;
            }
        };
        Utils.generateFetchXml_GET = function (entityName, attributes, orderBy, conditions) {
            return "<fetch version=\"1.0\" output-format=\"xml-platform\" mapping=\"logical\" distinct=\"false\">\n\t\t\t\t<entity name=\"" + entityName + "\">\n\t\t\t\t" + Utils.getAttributeXml(attributes) + "\n\t\t\t\t" + Utils.getOrderByXml(orderBy) + "\n\t\t\t\t" + Utils.getFilterXml(conditions) + "\n\t\t\t\t</entity>\n\t\t\t</fetch>";
        };
        Utils.getFilterXml = function (filters) {
            var filterXml = "";
            if (Array.isArray(filters)) {
                var filter_1 = "";
                filters.forEach(function (condition) {
                    if (condition.operator !== "in") {
                        filter_1 += "<condition attribute= \"" + condition.attributeName + "\" operator= \"" + condition.operator + "\" value= \"" + condition.value + "\" />";
                    }
                    else {
                        filter_1 += "<condition attribute= \"" + condition.attributeName + "\" operator= \"" + condition.operator + "\">";
                        condition.value.forEach(function (item) {
                            filter_1 += "<value>" + item + "</value>";
                        });
                        filter_1 += "</condition>";
                    }
                });
                filterXml += "<filter type= \"and\">" + filter_1 + "</filter>";
            }
            return filterXml;
        };
        Utils.getOrderByXml = function (orderBy) {
            var orderByXml = "";
            if (Array.isArray(orderBy)) {
                orderBy.forEach(function (orderAttr) {
                    orderByXml += "<order attribute= \"" + orderAttr.attributeName + "\" descending= \"" + orderAttr.descending + "\" />";
                });
            }
            return orderByXml;
        };
        Utils.getAttributeXml = function (attrs) {
            var attributeXml = "";
            if (Array.isArray(attrs)) {
                attrs.forEach(function (attributeName) {
                    attributeXml += "<attribute name= \"" + attributeName + "\" />";
                });
            }
            return attributeXml;
        };
        return Utils;
    }());
    OmniChannelAgentSDK.Utils = Utils;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation.  All rights reserved.
*/
//This module communicate with USD
var OmnichannelUtility;
(function (OmnichannelUtility) {
    'use strict';
    /**
    * This is the OmnichannelUSDCommunicator class
    */
    var OmnichannelUSDCommunicator = (function () {
        /**
        * This is OmnichannelUSDCommunicator's constructor
        * @param USDevent This is the USDevent that the raiseUSDevent function will call
        * @param functionName This is the name of the window parameter
        * @param ControlCallback This is the function which will be set to the window parameter
        */
        function OmnichannelUSDCommunicator(callbackFunctionName, ControlCallback) {
            this.controlCallback = ControlCallback;
            this.callbackFunctionName = callbackFunctionName;
        }
        /**
        * This is the RaiseUSDEvent function
        * This sets control callback function to functioname window parameter and raises
        * the OmnichannelUSDCommunicator's USDevent
        */
        OmnichannelUSDCommunicator.prototype.RaiseUSDEvent = function () {
            if (window.top[this.callbackFunctionName] == null) {
                window.top[this.callbackFunctionName] = this.controlCallback;
            }
            /* eslint-disable @typescript-eslint/no-insecure-url */
            window.open("http://event/?eventname=" + OmnichannelUSDCommunicator.USDevent + "&callback=" + this.callbackFunctionName);
            /* eslint-enable @typescript-eslint/no-insecure-url */
        };
        return OmnichannelUSDCommunicator;
    }());
    OmnichannelUSDCommunicator.USDevent = "GetOmniChannelSessionIdInfoEvent";
    OmnichannelUtility.OmnichannelUSDCommunicator = OmnichannelUSDCommunicator;
    var OmnichannelUSDBridge = (function () {
        function OmnichannelUSDBridge() {
        }
        OmnichannelUSDBridge.isUSD = function () {
            if (window.ocContext && window.ocContext.config) {
                return true;
            }
            return false;
        };
        /**
        * Helper method to get OC context.
        * @returns OCContext object.
        */
        OmnichannelUSDBridge.getOCContext = function () {
            var context = null;
            if (window.Xrm && window.Xrm.Page.context) {
                var queryStringParameters = window.Xrm.Page.context.getQueryStringParameters();
                if (queryStringParameters.ocContext != undefined) {
                    context = queryStringParameters.ocContext;
                }
                else {
                    if (window.top.Xrm && window.top.Xrm.Page.context) {
                        queryStringParameters = window.top.Xrm.Page.context.getQueryStringParameters();
                        if (queryStringParameters.ocContext != undefined) {
                            context = queryStringParameters.ocContext;
                        }
                    }
                }
            }
            //if context not present in querystring parameters get it from window.
            if (context === null || context === undefined) {
                context = window.ocContext;
            }
            if (context && context.config) {
                return context.config;
            }
            return null;
        };
        OmnichannelUSDBridge.sendToUSD = function (url) {
            if (window.top.notifyUSD) {
                window.top.notifyUSD(url);
            }
            else {
                window.open(url);
            }
        };
        OmnichannelUSDBridge.getLiveWorkItemId = function () {
            var liveWorkItemId = "";
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.LiveWorkItemId) {
                liveWorkItemId = sessionParams.LiveWorkItemId;
            }
            return liveWorkItemId;
        };
        OmnichannelUSDBridge.getSessionId = function () {
            var sessionId = "";
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.OCSessionId) {
                sessionId = sessionParams.OCSessionId;
            }
            return sessionId;
        };
        OmnichannelUSDBridge.getSessionInfo = function () {
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.SessionInfo != null && sessionParams.SessionInfo != "") {
                var sessionInfo = {};
                var parsedSessionInfo = JSON.parse(sessionParams.SessionInfo);
                sessionInfo.msdyn_sessioncreatedon = parsedSessionInfo.SessionCreatedTime;
                sessionInfo.msdyn_queueassignedon = parsedSessionInfo.SessionQueueAssignedTime;
                sessionInfo.msdyn_agentassignedon = parsedSessionInfo.SessionAgentAssignedTime;
                sessionInfo.msdyn_agentacceptedon = parsedSessionInfo.SessionAgentAcceptedTime;
                sessionInfo.msdyn_sessionclosedon = parsedSessionInfo.SessionEndTime;
                sessionInfo.msdyn_cdsqueueid = parsedSessionInfo.QueueId;
                sessionInfo.msdyn_ocsession_sessionevent = parsedSessionInfo.SessionEvents;
                sessionInfo.msdyn_ocsession_sessionparticipant = parsedSessionInfo.SessionParticipants;
                sessionInfo.msdyn_conversationcontext = parsedSessionInfo.ConvContext;
                return sessionInfo;
            }
            return null;
        };
        OmnichannelUSDBridge.getLWIInfo = function () {
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.LWIInfo != null && sessionParams.LWIInfo != "") {
                var parsedLWIInfo = JSON.parse(sessionParams.LWIInfo);
                return parsedLWIInfo;
            }
            return null;
        };
        OmnichannelUSDBridge.getContextItems = function () {
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.RuntimeContextItems != null && sessionParams.RuntimeContextItems != "") {
                return JSON.parse(sessionParams.RuntimeContextItems);
            }
        };
        OmnichannelUSDBridge.getResponseFromChannelSpecificItems = function (key) {
            var sessionParams = this.getSessionParamsFromOccontext();
            if (sessionParams && sessionParams.SessionInfo != null && sessionParams.SessionInfo != "") {
                var parsedSessionInfo = JSON.parse(sessionParams.SessionInfo);
                if (parsedSessionInfo.ChannelSpecificItems) {
                    return parsedSessionInfo.ChannelSpecificItems[key];
                }
            }
            return null;
        };
        OmnichannelUSDBridge.getSessionParamsFromOccontext = function () {
            if (window.Xrm && window.Xrm.Page.context) {
                var queryStringParameters = window.Xrm.Page.context.getQueryStringParameters();
                if (queryStringParameters && queryStringParameters.ocContext) {
                    try {
                        var ocContext = queryStringParameters.ocContext;
                        if (ocContext.config && ocContext.config.sessionParams) {
                            return ocContext.config.sessionParams;
                        }
                    }
                    catch (error) {
                    }
                }
            }
            if (window.ocContext && window.ocContext.config) {
                return window.ocContext.config.sessionParams;
            }
            return null;
        };
        OmnichannelUSDBridge.getEndPointBaseUrl = function () {
            var endpoint = "";
            if (window.ocContext && window.ocContext.config) {
                endpoint = window.ocContext.config.omniChannelBaseUrl ? window.ocContext.config.omniChannelBaseUrl : "";
            }
            return endpoint;
        };
        return OmnichannelUSDBridge;
    }());
    OmnichannelUtility.OmnichannelUSDBridge = OmnichannelUSDBridge;
})(OmnichannelUtility || (OmnichannelUtility = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
///<reference path="../TypeDefinitions/libs/XrmClientApi.d.ts" />
///<reference path="../TypeDefinitions/USDLib.d.ts" />
///<reference path="../TypeDefinitions/libs/OmnichannelUSDCommunicator.ts" />
///<reference path="../../../../references/external/TypeDefinitions/lib.es6.d.ts"/>
///<reference path="msdyn_internal_ci_library.d.ts"/>
///<reference path="../../Solution/WebResources/msdyn_LinkConversationLibrary.d.ts" />
///<reference path="./AppContext/AppContextLinkCommand.ts" />
///<reference path="./AppContext/AppContextLinkCommandFactory.ts" />
///<reference path="./AppContext/UCIAppContextLinkCommand.ts" />
///<reference path="./AppContext/USDAppContextLinkCommand.ts" />
///<reference path="./LinkTelemetryLogger.ts" />
///<reference path="./Constants.ts" />
///<reference path="./Utils.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    /**
    * Class to link the current record to the ongoing conversationvisual
    */
    var OmnichannelLinkCommand = (function () {
        function OmnichannelLinkCommand() {
            OmnichannelLinkCommand.appContextLinkCommand = OmniChannelAgentSDK.LinkCommand.UCIAppContextLinkCommandFactory.createAppContextLinkCommand();
        }
        /**
         * Returns whether the Link command should be displayed in the current context
         */
        OmnichannelLinkCommand.shouldDisplayLinkCommand = function () {
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var parentXrm = window.top.Xrm;
            var id = parentXrm.Page.data.entity.getId();
            //If form not saved, dont show command
            if (id == OmniChannelAgentSDK.FieldValues.emptyString) {
                return false;
            }
            var additionalDetails = {};
            additionalDetails["recordGuid"] = id;
            try {
                return this.appContextLinkCommand.shouldDisplayLinkCommand();
            }
            catch (error) {
                OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(currentReqId, "", "", OmniChannelAgentSDK.Components.shouldDisplayLinkCommand, "Not a valid conversation context", false, error, JSON.stringify(additionalDetails));
            }
            return false;
        };
        /**
         * Function to link the current record with the ongoing conversation
         * @param relationshipNames Comma separated list of relationship names that need to be linked
         */
        OmnichannelLinkCommand.linkRecordToConversation = function (relationshipNames) {
            var _this = this;
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var liveWorkItemId = "";
            var liveWorkStreamId = "";
            var additionalDetails = {};
            try {
                var parentXrm = window.top.Xrm;
                var globalContext = parentXrm.Utility.getGlobalContext();
                var orgId_1 = globalContext.organizationSettings.organizationId;
                var primaryAttrValue_1 = parentXrm.Page.data.getEntity().getPrimaryAttributeValue();
                var entityType_1 = parentXrm.Page.entityReference.entityType;
                var id = parentXrm.Page.data.entity.getId();
                this.fetchLiveWorkItemData().then(function (liveWorkitemData) {
                    liveWorkItemId = liveWorkitemData[OmniChannelAgentSDK.Constants.LIVEWORKITEM_ID_ATTR];
                    liveWorkStreamId = liveWorkitemData[OmniChannelAgentSDK.Constants.LIVEWORKSTREAM_ID_ATTR];
                    additionalDetails["liveWorkItemId"] = liveWorkItemId;
                    additionalDetails["liveWorkStreamId"] = liveWorkStreamId;
                    additionalDetails["recordGuid"] = id;
                    var payload = _this.generateLinkUnlinkPayload(entityType_1, id, relationshipNames, primaryAttrValue_1, orgId_1, liveWorkItemId, liveWorkStreamId);
                    _this.linkRecordToConversationInternal(payload).then(function () {
                        _this.handleContextApiSuccess(OmniChannelAgentSDK.Utils.generateSuccessPayload(entityType_1, id, liveWorkItemId, "Linking successful", currentReqId, additionalDetails));
                    }, function (error) {
                        _this.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, " linkRecordToConversation > linkRecordToConversationInternal", "Error while making the Context API call.", currentReqId, additionalDetails, error));
                    });
                }, function (error) {
                    _this.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, " linkRecordToConversation > linkRecordToConversationInternal", "Error in fetching live work item data.", currentReqId, additionalDetails, error));
                });
            }
            catch (error) {
                this.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, " linkRecordToConversation > linkRecordToConversationInternal", "Linking to conversation failed due to some exception.", currentReqId, additionalDetails, error));
            }
        };
        /**
         * Function to link the current record with the ongoing conversation
         * @param entityLogicalName logical name of the entity being linked
         * @param recordId unique identifier representing record id
         */
        OmnichannelLinkCommand.linkRecordToConversationAPI = function (entityLogicalName, recordId) {
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var liveWorkItemId = "";
            var liveWorkStreamId = "";
            var additionalDetails = {};
            var self = this;
            return new Promise(function (resolve, reject) {
                try {
                    var parentXrm = window.top.Xrm;
                    var globalContext = parentXrm.Utility.getGlobalContext();
                    var orgId_2 = globalContext.organizationSettings.organizationId;
                    self.fetchLiveWorkItemData().then(function (liveWorkitemData) {
                        liveWorkItemId = liveWorkitemData[OmniChannelAgentSDK.Constants.LIVEWORKITEM_ID_ATTR];
                        liveWorkStreamId = liveWorkitemData[OmniChannelAgentSDK.Constants.LIVEWORKSTREAM_ID_ATTR];
                        additionalDetails["liveWorkItemId"] = liveWorkItemId;
                        additionalDetails["liveWorkStreamId"] = liveWorkStreamId;
                        additionalDetails["recordGuid"] = recordId;
                        self.getEntityPrimaryAttrValue(entityLogicalName, recordId).then(function (primaryAttrValue) {
                            var payload = self.generateLinkUnlinkPayload(entityLogicalName, recordId, OmniChannelAgentSDK.Utils.getRelationShipNameByEntity(entityLogicalName), primaryAttrValue, orgId_2, liveWorkItemId, liveWorkStreamId);
                            self.linkRecordToConversationInternal(payload).then(function () {
                                self.handleContextApiSuccess(OmniChannelAgentSDK.Utils.generateSuccessPayload(entityLogicalName, recordId, liveWorkItemId, "Linking successful", currentReqId, additionalDetails), true, true, resolve);
                            }, function (error) {
                                self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkRecordToConversationAPI > linkRecordToConversation", "Link API call failed, Error while making the Context API call.", currentReqId, additionalDetails, error), false, true, reject);
                            });
                        }, function (error) {
                            self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkRecordToConversationAPI > getEntityPrimaryAttrValue", "Link API call failed.", currentReqId, additionalDetails, error), false, true, reject);
                        });
                    }, function (error) {
                        self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkRecordToConversationAPI > fetchLiveWorkItemData", "Link API call failed, Error in fetching live work item data.", currentReqId, additionalDetails, error), false, true, reject);
                    });
                }
                catch (error) {
                    self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkRecordToConversationAPI", "Linking to conversation failed due to some exception.", currentReqId, additionalDetails, error), false, true, reject);
                }
            });
        };
        /**
         * Function to unlink the current record from the ongoing conversation
         * @param entityLogicalName logical name of the entity being linked
         * @param recordId unique identifier representing record id
         */
        OmnichannelLinkCommand.unlinkFromConversation = function (entityLogicalName, recordId) {
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var liveWorkItemId = "";
            var liveWorkStreamId = "";
            var additionalDetails = {};
            var self = this;
            return new Promise(function (resolve, reject) {
                try {
                    var parentXrm = window.top.Xrm;
                    var globalContext = parentXrm.Utility.getGlobalContext();
                    var orgId_3 = globalContext.organizationSettings.organizationId;
                    self.fetchLiveWorkItemData().then(function (liveWorkItemData) {
                        liveWorkItemId = liveWorkItemData[OmniChannelAgentSDK.Constants.LIVEWORKITEM_ID_ATTR];
                        liveWorkStreamId = liveWorkItemData[OmniChannelAgentSDK.Constants.LIVEWORKSTREAM_ID_ATTR];
                        additionalDetails["liveWorkItemId"] = liveWorkItemId;
                        additionalDetails["liveWorkStreamId"] = liveWorkStreamId;
                        additionalDetails["recordGuid"] = recordId;
                        new OmniChannelPackage.LinkConversation.OCLinkToConversation().unlinkRecordFromConversation(entityLogicalName, OmniChannelAgentSDK.Utils.getRelationShipNameByEntity(entityLogicalName), recordId, orgId_3, liveWorkStreamId, liveWorkItemId).then(function (response) {
                            if (OmniChannelAgentSDK.Utils.isNullOrUndefined(response) || response == OmniChannelAgentSDK.Constants.CONTEXT_UPDATE_FAILED_STATUS) {
                                self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.unlinkRecordFromConversation, "linkRecordToConversationAPI", "Link API call failed, ReasonFailure response from Context API", currentReqId, additionalDetails), false, true, reject);
                            }
                            else {
                                self.handleContextApiSuccess(OmniChannelAgentSDK.Utils.generateSuccessPayload(entityLogicalName, recordId, liveWorkItemId, "UnLinking successful", currentReqId, additionalDetails), false, true, resolve);
                            }
                        }, function (error) {
                            self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.unlinkRecordFromConversation, "unlinkRecordFromConversation > unlinkRecordFromConversation", "UnLink API call failed, Error while making the Context API call.", currentReqId, additionalDetails, error), false, true, reject);
                        });
                    }, function (error) {
                        self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.unlinkRecordFromConversation, "unlinkRecordFromConversation > fetchLiveWorkItemData", "UnLink API call failed, Error in fetching live work item data, Error: " + error, currentReqId, additionalDetails, error), false, true, reject);
                    });
                }
                catch (error) {
                    self.handleLinkFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "unlinkRecordFromConversation", "Unlinking from conversation failed due to some exception.", currentReqId, additionalDetails, error), false, true, reject);
                }
            });
        };
        /**
         * Opens a conversion
         * @param liveWorkItemId unique identifier representing liveWorkItem id
         * @param sessionId unique identifier representing session id
         * @param liveWorkStreamId unique identifier representing liveWorkStream id
         */
        OmnichannelLinkCommand.openConversation = function (liveWorkItemId, sessionId, liveWorkStreamId) {
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var additionalDetails = {
                liveWorkStreamId: liveWorkStreamId,
                sessionId: sessionId
            };
            var self = this;
            return new Promise(function (resolve, reject) {
                try {
                    self.fetchLiveWorkItemDataById(liveWorkItemId).then(function (liveWorkitemData) {
                        if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(liveWorkitemData)) {
                            var recordSessionId = liveWorkitemData._msdyn_lastsessionid_value;
                            var recordLastWorkStreamId = liveWorkitemData._msdyn_liveworkstreamid_value;
                            if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(sessionId) && sessionId != recordSessionId) {
                                var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI", "Record not found for Session Id provided!", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId });
                                self.handleOpenConversationFailure(errorObject, false, true, reject);
                            }
                            if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(liveWorkStreamId) && liveWorkStreamId != recordLastWorkStreamId) {
                                var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI", "Record not found for LiveWorkStream Id provided!", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId });
                                self.handleOpenConversationFailure(errorObject, false, true, reject);
                            }
                            if (OmniChannelAgentSDK.Utils.isTopWindowAccessible()) {
                                var topwindow = window.top;
                                var params = {
                                    sessionId: recordSessionId,
                                    liveWorkItemId: liveWorkItemId,
                                    liveWorkStreamId: recordLastWorkStreamId
                                };
                                var evt = new CustomEvent('OpenConversationAsSession', { detail: params });
                                topwindow.dispatchEvent(evt);
                                self.handleOpenConversationSuccess(OmniChannelAgentSDK.Utils.generateSuccessPayload("", sessionId, liveWorkItemId, "OpenConversationAsSession event is raised on UCI", currentReqId, additionalDetails), false, true, resolve);
                            }
                            else {
                                var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI", "Top level window not accessible.", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId });
                                self.handleOpenConversationFailure(errorObject, false, true, reject);
                            }
                        }
                        else {
                            var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI", "Open Conversation API call failed, Failed to fetch fetching live work item for given id.", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId });
                            self.handleOpenConversationFailure(errorObject, false, true, reject);
                        }
                    }, function (error) {
                        self.handleOpenConversationFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI > fetchLiveWorkItemData", "Open Conversation API call failed, Error in fetching live work item data.", currentReqId, additionalDetails, error), false, true, reject);
                    });
                }
                catch (error) {
                    self.handleOpenConversationFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "OpenConversationAPI", "Opening conversation failed due to some exception.", currentReqId, additionalDetails, error), false, true, reject);
                }
            });
        };
        /**
         * Sends a message to conversation
         * @param message message to be sent to conversation
         * @param toSendBox determines whether the message will be sent to the sendBox
         * @param liveWorkItemId unique identifier representing liveWorkItem id
         */
        OmnichannelLinkCommand.sendMessageToConversation = function (message, toSendBox, liveWorkItemId) {
            if (toSendBox === void 0) { toSendBox = true; }
            var self = this;
            return new Promise(function (resolve, reject) {
                try {
                    var payload = {
                        text: message,
                        liveWorkItemId: liveWorkItemId,
                        toSendBox: toSendBox
                    };
                    var evt = new CustomEvent("onsendmessage", {
                        detail: payload
                    });
                    window.top.dispatchEvent(evt);
                    self.handleSendMessageToConversationSuccess(OmniChannelAgentSDK.Utils.generateSuccessPayload("", "", liveWorkItemId, "SendMessageToConversation event is raised on UCI", "", ""), false, true, resolve);
                }
                catch (error) {
                    self.handleSendMessageToConversationFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.sendMessageToConversation, "SendMessageToConversationAPI", "Sending message to conversation failed due to some exception.", "", error), false, true, reject);
                }
            });
        };
        OmnichannelLinkCommand.getLinkedRecordsInternal = function (liveWorkItemId) {
            var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
            var result = new Array();
            var self = this;
            return new Promise(function (resolve, reject) {
                try {
                    self.fetchLiveWorkItemDataById(liveWorkItemId).then(function (liveWorkitemData) {
                        if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(liveWorkitemData)) {
                            // Get the linked customer record.
                            var customerLookup_AttributeName = OmniChannelAgentSDK.ODataConstants.lookupFieldPrefix + OmniChannelAgentSDK.Constants.CUSTOMER_FIELDLOGICALNAME + OmniChannelAgentSDK.ODataConstants.lookupFieldSuffix;
                            var customerIdValue = liveWorkitemData[customerLookup_AttributeName];
                            if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(customerIdValue)) {
                                var customerLookup_AttributeLogicalName = customerLookup_AttributeName + OmniChannelAgentSDK.ODataConstants.lookupLogicalNameKey;
                                var entityLogicalName = liveWorkitemData[customerLookup_AttributeLogicalName];
                                switch (entityLogicalName) {
                                    case OmniChannelAgentSDK.EntityNames.Account:
                                        result.push({ entityName: OmniChannelAgentSDK.EntityNames.Account, recordId: customerIdValue });
                                        break;
                                    case OmniChannelAgentSDK.EntityNames.Contact:
                                        result.push({ entityName: OmniChannelAgentSDK.EntityNames.Contact, recordId: customerIdValue });
                                        break;
                                }
                            }
                            // Get the linked incident record.
                            var issueLookup_AttributeName = OmniChannelAgentSDK.ODataConstants.lookupFieldPrefix + OmniChannelAgentSDK.Constants.ISSUE_FIELDLOGICALNAME + OmniChannelAgentSDK.ODataConstants.lookupFieldSuffix;
                            var issueIdValue = liveWorkitemData[issueLookup_AttributeName];
                            if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(issueIdValue)) {
                                result.push({ entityName: OmniChannelAgentSDK.EntityNames.Incident, recordId: issueIdValue });
                            }
                            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(currentReqId, "", liveWorkItemId, OmniChannelAgentSDK.Components.getLinkedRecords, "getLinkedRecordsInternal: Linked records fetched successfully", false, null, "");
                            resolve(result);
                        }
                        else {
                            var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getLinkedRecords, "GetLinkedRecordsInternalAPI", "Failed to fetch fetching live work item data for given id.", currentReqId, {});
                            self.handleGetLinkedRecordsFailure(errorObject, true, reject);
                        }
                    }, function (error) {
                        self.handleGetLinkedRecordsFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getLinkedRecords, "GetLinkedRecordsInternalAPI > fetchLiveWorkItemDataById", "Get linked records API call failed, Error in fetching live work item data for given id.", currentReqId, {}, error), true, reject);
                    });
                }
                catch (error) {
                    self.handleGetLinkedRecordsFailure(OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getLinkedRecords, "GetLinkedRecordsInternalAPI", "Get linked records API failed due to some exception.", currentReqId, {}, error), true, reject);
                }
            });
        };
        /**
         * Gets localized string for provided string id
         * @param stringId string id from resx file
         */
        OmnichannelLinkCommand.getLocalizedString = function (stringId, requestId) {
            var telemetryLogger;
            try {
                telemetryLogger = OmniChannelAgentSDK.TelemetryLogger.Instance();
                var localizedString = Xrm.Utility.getResourceString(OmniChannelAgentSDK.LocalizationConstants.resxWebResourceName, stringId);
                if (OmniChannelAgentSDK.Utils.isNullOrUndefined(localizedString)) {
                    localizedString = OmniChannelAgentSDK.LocalizationConstants.OC_Undefined;
                }
                return localizedString;
            }
            catch (error) {
                var eventMessage = "Exception in getting localized string";
                telemetryLogger.sendEvent(requestId, OmniChannelAgentSDK.FieldValues.emptyString, OmniChannelAgentSDK.FieldValues.emptyString, OmniChannelAgentSDK.Components.localizedString, eventMessage, true, null, OmniChannelAgentSDK.FieldValues.emptyString);
            }
        };
        /**
         * Shows toast notification on Link success/failure
         */
        OmnichannelLinkCommand.showUserNotification = function (type, level, message, title, action) {
            var parentXrm = window.top.Xrm;
            var notificationPromise;
            notificationPromise = parentXrm.UI.addGlobalNotification(type, level, message, title, action);
            return notificationPromise;
        };
        OmnichannelLinkCommand.fetchLiveWorkItemData = function () {
            return this.appContextLinkCommand.fetchLiveWorkItemData();
        };
        OmnichannelLinkCommand.fetchLiveWorkItemDataById = function (liveWorkItemId, sessionId, liveWorkStreamId) {
            return new Promise(function (resolve, reject) {
                Xrm.WebApi.retrieveRecord(OmniChannelAgentSDK.EntityNames.LiveWorkItem, liveWorkItemId).then(function (response) {
                    resolve(response);
                }, function (error) {
                    reject(error);
                });
            });
        };
        OmnichannelLinkCommand.handleContextApiSuccess = function (payload, shouldNotify, isAsync, resolver) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(payload.currentReqId, "", "", OmniChannelAgentSDK.Components.linkRecordToConversation, payload.message, false, null, payload.additionalDetails);
            if (shouldNotify) {
                // Notify the client that linking is done
                this.appContextLinkCommand.raiseLinkingDoneEvent(payload.entityLogicalName, payload.recordId, payload.liveWorkItemId);
                //Toast notification
                var message = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_LinkToConversationSuccessMessage, payload.currentReqId);
                message = message.replace("{0}", OmniChannelPackage.OCDataLayer.DataHelper.getEntityDisplayName(payload.entityLogicalName).toLocaleLowerCase());
                this.showUserNotification(1 /* toast */, 1 /* success */, message, "Record linked to live work item", null);
            }
            if (isAsync)
                resolver(payload);
        };
        OmnichannelLinkCommand.handleLinkFailure = function (errorObject, shouldNotify, isAsync, rejectHandler) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(errorObject.currentReqId, "", "", OmniChannelAgentSDK.Components.linkRecordToConversation, errorObject.message, true, errorObject.error, errorObject.additionalDetails);
            if (shouldNotify) {
                var linkingFailedMessage = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_LinkToConversationFailureMessage, errorObject.currentReqId);
                this.showUserNotification(1 /* toast */, 2 /* error */, linkingFailedMessage, "Request to link record failed", null);
            }
            if (isAsync)
                rejectHandler(errorObject);
        };
        OmnichannelLinkCommand.linkRecordToConversationInternal = function (payload) {
            return new Promise(function (resolve, reject) {
                new OmniChannelPackage.LinkConversation.OCLinkToConversation().linkRecordToConversation(payload.entityLogicalName, payload.relationshipNames, payload.primaryAttrValue, payload.recordId, payload.orgId, payload.liveWorkStreamId, payload.liveWorkItemId).then(function (linkResponse) {
                    if (OmniChannelAgentSDK.Utils.isNullOrUndefined(linkResponse) && (linkResponse[0].status == OmniChannelAgentSDK.Constants.CONTEXT_UPDATE_FAILED_STATUS)) {
                        reject("Link API call failed, ReasonFailure response from Context API");
                    }
                    else {
                        resolve();
                    }
                }, function (error) {
                    reject(error);
                });
            });
        };
        OmnichannelLinkCommand.getEntityPrimaryAttrValue = function (entityLogicalName, recordId) {
            return new Promise(function (resolve, reject) {
                var parentXrm = window.top.Xrm;
                parentXrm.Utility.getEntityMetadata(entityLogicalName, []).then(function (entityMetaData) {
                    if (OmniChannelAgentSDK.Utils.isNullOrUndefined(entityMetaData))
                        reject("Couldn't fetch entity metda data for entity: " + entityLogicalName);
                    var primaryAttrName = entityMetaData.PrimaryNameAttribute;
                    parentXrm.WebApi.retrieveRecord(entityLogicalName, recordId, "?$select=" + primaryAttrName).then(function (response) {
                        (!OmniChannelAgentSDK.Utils.isNullOrUndefined(response))
                            ? resolve(response[primaryAttrName])
                            : reject("Couldn't resovle primary attribute value for entity: " + entityLogicalName);
                    }, function (error) {
                        reject(error);
                    });
                }, function (error) {
                    reject(error);
                });
            });
        };
        OmnichannelLinkCommand.generateLinkUnlinkPayload = function (entityLogicalName, recordId, relationshipNames, primaryAttrValue, orgId, liveWorkItemId, liveWorkStreamId) {
            return {
                entityLogicalName: entityLogicalName,
                relationshipNames: relationshipNames,
                primaryAttrValue: primaryAttrValue,
                recordId: recordId,
                orgId: orgId,
                liveWorkItemId: liveWorkItemId,
                liveWorkStreamId: liveWorkStreamId
            };
        };
        OmnichannelLinkCommand.handleOpenConversationSuccess = function (payload, shouldNotify, isAsync, resolver) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(payload.currentReqId, "", "", OmniChannelAgentSDK.Components.openConversation, payload.message, false, null, payload.additionalDetails);
            if (shouldNotify) {
                //Toast notification
                var message = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_OpenConversationSuccessMessage, payload.currentReqId);
                this.showUserNotification(1 /* toast */, 1 /* success */, message, "Opening Conversation", null);
            }
            if (isAsync)
                resolver(payload);
        };
        OmnichannelLinkCommand.handleOpenConversationFailure = function (errorObject, shouldNotify, isAsync, rejectHandler) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(errorObject.currentReqId, "", "", OmniChannelAgentSDK.Components.openConversation, errorObject.message, true, errorObject.error, errorObject.additionalDetails);
            if (shouldNotify) {
                var openSessionFailedMessage = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_OpenConversationFailureMessage, errorObject.currentReqId);
                this.showUserNotification(1 /* toast */, 2 /* error */, openSessionFailedMessage, "Request to open conversation failed", null);
            }
            if (isAsync)
                rejectHandler(errorObject);
        };
        OmnichannelLinkCommand.handleGetLinkedRecordsFailure = function (errorObject, isAsync, rejectHandler) {
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(errorObject.currentReqId, "", "", OmniChannelAgentSDK.Components.getLinkedRecords, errorObject.message, true, errorObject.error, errorObject.additionalDetails);
            if (isAsync)
                rejectHandler(errorObject);
        };
        OmnichannelLinkCommand.handleSendMessageToConversationSuccess = function (payload, shouldNotify, isAsync, resolver) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(payload.liveWorkItemId, "", "", OmniChannelAgentSDK.Components.sendMessageToConversation, payload.message, false, null, payload.additionalDetails);
            if (shouldNotify) {
                var sendMessageToConversationSuccessMessage = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_SendMessageToConversationSuccessMessage, payload.currentReqId);
                this.showUserNotification(1 /* toast */, 1 /* success */, sendMessageToConversationSuccessMessage, "Sending Message", null);
            }
            if (isAsync)
                resolver(payload);
        };
        OmnichannelLinkCommand.handleSendMessageToConversationFailure = function (errorObject, shouldNotify, isAsync, rejectHandler) {
            if (shouldNotify === void 0) { shouldNotify = true; }
            if (isAsync === void 0) { isAsync = false; }
            OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(errorObject.currentReqId, "", "", OmniChannelAgentSDK.Components.sendMessageToConversation, errorObject.message, true, errorObject.error, errorObject.additionalDetails);
            if (shouldNotify) {
                var sendMessageToConversationFailureMessage = OmnichannelLinkCommand.getLocalizedString(OmniChannelAgentSDK.LocalizationConstants.OC_SendMessageToConversationFailureMessage, errorObject.currentReqId);
                this.showUserNotification(1 /* toast */, 2 /* error */, sendMessageToConversationFailureMessage, "Request to send message failed", null);
            }
            if (isAsync)
                rejectHandler(errorObject);
        };
        return OmnichannelLinkCommand;
    }());
    OmnichannelLinkCommand.Instance = new OmnichannelLinkCommand();
    OmniChannelAgentSDK.OmnichannelLinkCommand = OmnichannelLinkCommand;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
/**
* @license Copyright (c) Microsoft Corporation. All rights reserved.
*/
/// <reference path="./OmnichannelLinkCommand.ts" />
/// <reference path="./Utils.ts" />
/// <reference path="./Constants.ts" />
/// <reference path="./Model/LiveWorkItemData.ts" />
/// <reference path="../TypeDefinitions/libs/AppRuntimeClientSdk.d.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var ConversationApi = (function () {
        function ConversationApi() {
        }
        ConversationApi.prototype.getConversationId = function () {
            return new Promise(function (resolve, reject) {
                var currentReqId = OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId();
                var session = Microsoft.AppRuntime.Sessions.getFocusedSession();
                session.getContext().then(function (response) {
                    var conversationId = OmniChannelAgentSDK.Utils.getSafe(function () { return response.parameters["LiveWorkItemId"]; });
                    if (conversationId) {
                        resolve(conversationId);
                    }
                    else {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getConversationId, "getConversationId API", "Couldn't fetch LiveWorkItemId from session context", currentReqId, {});
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, true, reject);
                    }
                }, function (error) {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getConversationId, "getConversationId API", "Could not fetch focused session context.", currentReqId, {}, error);
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, true, reject);
                });
            });
        };
        ConversationApi.prototype.getLinkedRecords = function () {
            return new Promise(function (resolve, reject) {
                var session = Microsoft.AppRuntime.Sessions.getFocusedSession();
                session.getContext().then(function (response) {
                    var conversationId = OmniChannelAgentSDK.Utils.getSafe(function () { return response.parameters["LiveWorkItemId"]; });
                    if (conversationId) {
                        OmniChannelAgentSDK.OmnichannelLinkCommand.getLinkedRecordsInternal(conversationId).then(function (response) {
                            resolve(response);
                        }, function (error) {
                            reject(error); //logging impliciltiy handled
                        });
                    }
                    else {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getLinkedRecords, "getLinkedRecords API", "Couldn't fetch LiveWorkItemId from session context", "", {});
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleGetLinkedRecordsFailure(errorObject, true, reject);
                    }
                }, function (error) {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.getLinkedRecords, "getLinkedRecords API", "Could not fetch focused session context.", "", {}, error);
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleGetLinkedRecordsFailure(errorObject, true, reject);
                });
            });
        };
        ConversationApi.prototype.linkToConversation = function (entityLogicalName, recordId) {
            return new Promise(function (resolve, reject) {
                if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(entityLogicalName) && !OmniChannelAgentSDK.Utils.isNullOrUndefined(recordId)) {
                    OmniChannelAgentSDK.OmnichannelLinkCommand.linkRecordToConversationAPI(entityLogicalName, recordId).then(function (response) {
                        resolve(response);
                    }, function (error) {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkToConversation API", "Error encountered in linkRecordToConversationAPI API call", "", { entityLogicalName: entityLogicalName, recordId: recordId }, error);
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, false, reject);
                        reject(error); //logging impliciltiy handled
                    });
                }
                else {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.linkRecordToConversation, "linkToConversation API", "Missing required Parameters", "", { entityLogicalName: entityLogicalName, recordId: recordId });
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, true, reject);
                }
            });
        };
        ConversationApi.prototype.unlinkFromConversation = function (entityLogicalName, recordId) {
            return new Promise(function (resolve, reject) {
                if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(entityLogicalName) && !OmniChannelAgentSDK.Utils.isNullOrUndefined(recordId)) {
                    OmniChannelAgentSDK.OmnichannelLinkCommand.unlinkFromConversation(entityLogicalName, recordId).then(function (response) {
                        resolve(response);
                    }, function (error) {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.unlinkRecordFromConversation, "unlinkFromConversation API", "Error encountered in unlinkFromConversation API call", "", { entityLogicalName: entityLogicalName, recordId: recordId }, error);
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, false, reject);
                        reject(error); //logging impliciltiy handled
                    });
                }
                else {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.unlinkRecordFromConversation, "unlinkFromConversation API", "Missing required Parameters", "", { entityLogicalName: entityLogicalName, recordId: recordId });
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleLinkFailure(errorObject, false, true, reject);
                }
            });
        };
        ConversationApi.prototype.openConversation = function (liveWorkItemId, sessionId, liveWorkStreamId) {
            return new Promise(function (resolve, reject) {
                if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(liveWorkItemId)) {
                    OmniChannelAgentSDK.OmnichannelLinkCommand.openConversation(liveWorkItemId, sessionId, liveWorkStreamId).then(function (response) {
                        resolve(response);
                    }, function (error) {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "openSessionFromConversation API", "Error encountered in openSessionFromConversation API call", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId }, error);
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleOpenConversationFailure(errorObject, false, false, reject);
                        reject(error); //logging implicitly handled
                    });
                }
                else {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.openConversation, "openSessionFromConversation API", "Missing required Parameters", "", { liveWorkItemId: liveWorkItemId, sessionId: sessionId, liveWorkStreamId: liveWorkStreamId });
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleOpenConversationFailure(errorObject, false, true, reject);
                }
            });
        };
        /**
         * API to fetch msdyn_ocliveworkitem entity records as per the filters provided
         * @param params Mandatory input parameter of type @typeof LinkCommand.Model.LiveWorkItemFilter,
         * @param correlationId Optional correlation Id for telemetry.
         * @returns In case of succese returns a JSON object with {status: "true", result: {}}, where result is a JSON object containg the data results.
         * In case of error returns a JSON object with {status: false, error: error}, where error is the exception occured.
         */
        ConversationApi.prototype.getConversations = function (params, correlationId) {
            return new Promise(function (resolve, reject) {
                if (!params.agentId) {
                    var error = new Error("Missing key \"agentId\" in input object");
                    OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(correlationId, "", "", OmniChannelAgentSDK.Components.getConversations, error.message, true, error, "");
                    reject({
                        status: false,
                        error: error
                    });
                }
                var defaultAttributes = [OmniChannelAgentSDK.EntityAttributesNames.LWI_OCLiveWorkItemId, OmniChannelAgentSDK.EntityAttributesNames.LWI_LiveWorkStreamId, OmniChannelAgentSDK.EntityAttributesNames.LWI_LastSessionId, OmniChannelAgentSDK.EntityAttributesNames.LWI_Statuscode, OmniChannelAgentSDK.EntityAttributesNames.LWI_CreatedOn];
                //agent id filter
                var filters = [{
                        attributeName: OmniChannelAgentSDK.EntityAttributesNames.LWI_ActiveAgentId,
                        operator: "eq",
                        value: params.agentId
                    }];
                //status code filter
                if (Array.isArray(params.status)) {
                    filters.push({
                        attributeName: OmniChannelAgentSDK.EntityAttributesNames.LWI_Statuscode,
                        operator: "in",
                        value: params.status
                    });
                }
                //interval filter
                if (params.createdBeforeDays) {
                    var today = new Date();
                    today.setDate(today.getDate() - params.createdBeforeDays);
                    filters.push({
                        attributeName: OmniChannelAgentSDK.EntityAttributesNames.LWI_CreatedOn,
                        operator: "on-or-before",
                        value: today.format("yyyy-MM-dd")
                    });
                }
                var attributes = defaultAttributes.slice();
                if (Array.isArray(params.attributes)) {
                    attributes = Array.from(new Set(attributes.concat(params.attributes)));
                }
                var fetchXmlQuery = OmniChannelAgentSDK.Utils.generateFetchXml_GET(OmniChannelAgentSDK.Constants.CONVERSATION_ENTITY_LOGICAL_NAME, attributes, params.orderBy, filters);
                window.top.fetch(window.top.Xrm.Page.context.getClientUrl() +
                    ("/api/data/v9.0/" + OmniChannelAgentSDK.Constants.CONVERSATION_ENTITY_LOGICAL_NAME + "s?fetchXml=") +
                    encodeURIComponent(fetchXmlQuery), {
                    credentials: "same-origin",
                    headers: {
                        Prefer: 'odata.include-annotations="*"'
                    }
                }).then(function (response) { return response.json(); }).then(function (result) {
                    if (result.error) {
                        OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(correlationId, "", "", OmniChannelAgentSDK.Components.getConversations, result.error.message, true, result.error, "");
                        reject({
                            status: false,
                            error: result.error
                        });
                    }
                    resolve({
                        status: true,
                        result: result
                    });
                }, function (error) {
                    OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(correlationId, "", "", OmniChannelAgentSDK.Components.getConversations, error.message, true, error, "");
                    reject({
                        status: false,
                        error: error
                    });
                })["catch"](function (error) {
                    OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(correlationId, "", "", OmniChannelAgentSDK.Components.getConversations, error.message, true, error, "");
                    reject({
                        status: false,
                        error: error
                    });
                });
            });
        };
        ConversationApi.prototype.sendMessageToConversation = function (message, toSendBox, liveWorkItemId) {
            if (toSendBox === void 0) { toSendBox = true; }
            return new Promise(function (resolve, reject) {
                if (!OmniChannelAgentSDK.Utils.isNullOrUndefined(message)) {
                    OmniChannelAgentSDK.OmnichannelLinkCommand.sendMessageToConversation(message, toSendBox, liveWorkItemId).then(function (response) {
                        resolve(response);
                    }, function (error) {
                        var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.sendMessageToConversation, "sendMesssageToConversation API", "Error encountered in sendMesssageToConversation API call", "", { liveWorkItemId: liveWorkItemId, messageLength: message.length() }, error);
                        OmniChannelAgentSDK.OmnichannelLinkCommand.handleOpenConversationFailure(errorObject, false, false, reject);
                        reject(error); //logging implicitly handled
                    });
                }
                else {
                    var errorObject = OmniChannelAgentSDK.Utils.generateErrorObject(OmniChannelAgentSDK.Components.sendMessageToConversation, "sendMesssageToConversation API", "Missing required Parameters", "message", { liveWorkItemId: liveWorkItemId, messageLength: message.length() });
                    OmniChannelAgentSDK.OmnichannelLinkCommand.handleOpenConversationFailure(errorObject, false, true, reject);
                }
            });
        };
        return ConversationApi;
    }());
    OmniChannelAgentSDK.ConversationApi = ConversationApi;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
///<reference path="./Utils.ts" />
///<reference path="./ConversationApi.ts" />
var OmniChannelAgentSDK;
(function (OmniChannelAgentSDK) {
    var OmnichannelAgentSdkInitializer = (function () {
        function OmnichannelAgentSdkInitializer() {
        }
        OmnichannelAgentSdkInitializer.init = function () {
            OmnichannelAgentSdkInitializer.loadDependencies().then(function () {
                OmnichannelAgentSdkInitializer.createSdkObject(new OmniChannelAgentSDK.ConversationApi());
            }, function (error) {
                OmniChannelAgentSDK.TelemetryLogger.Instance().sendEvent(OmniChannelAgentSDK.TelemetryLogger.Instance().getNewReqId(), "", "", OmniChannelAgentSDK.Components.initOmnichannelAgentSDK, "Could not initialize OmnichannelAgentSDK", true, error, "");
            });
        };
        OmnichannelAgentSdkInitializer.createSdkObject = function (clientSdk) {
            var root = window;
            var namespaceComponents = OmniChannelAgentSDK.Constants.OMNICHANNEL_AGENT_SDK_NAMESPACE.split(".");
            if (namespaceComponents) {
                var length_1 = namespaceComponents.length;
                for (var idx = 0; idx <= length_1 - 2; idx++) {
                    var component = namespaceComponents[idx].trim();
                    root[component] = root[component] || {};
                    root = root[component];
                }
                root[namespaceComponents[length_1 - 1]] = clientSdk;
            }
        };
        OmnichannelAgentSdkInitializer.loadDependencies = function () {
            return new Promise(function (resolve, reject) {
                try {
                    var parentXrm = window.top.Xrm;
                    var orgUrl = parentXrm.Utility.getGlobalContext().getClientUrl();
                    var promises = [];
                    promises.push(OmniChannelAgentSDK.Utils.loadScript(orgUrl + OmniChannelAgentSDK.Constants.TELEMETRY_RESOURCE_URL));
                    promises.push(OmniChannelAgentSDK.Utils.loadScript(orgUrl + OmniChannelAgentSDK.Constants.LINK_CONVERSATION_LIBRARY_RESOURCE_URL));
                    Promise.all(promises).then(function () {
                        resolve();
                    }, function (error) {
                        reject("Error in resolving dependency prmoises for OmnichannelAgentSDK, Error: " + error);
                    });
                }
                catch (error) {
                    reject("Error in loading dependencies for OmnichannelAgentSDK, Error: " + error);
                }
            });
        };
        return OmnichannelAgentSdkInitializer;
    }());
    OmniChannelAgentSDK.OmnichannelAgentSdkInitializer = OmnichannelAgentSdkInitializer;
})(OmniChannelAgentSDK || (OmniChannelAgentSDK = {}));
(function () { return OmniChannelAgentSDK.OmnichannelAgentSdkInitializer.init(); })();
