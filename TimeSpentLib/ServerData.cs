//
// Copyright © Microsoft Corporation.  All Rights Reserved.
// This code released under the terms of the 
// Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html)
//
// http://archive.msdn.microsoft.com/timespent

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeWebServices;  // http://blogs.msdn.com/b/exchangedev/archive/2007/12/07/generating-exchange-web-services-proxy-classes.aspx
using System.Reflection;
using System.Net;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;



namespace TimeSpentLib
{
    /// <summary>
    /// Represents the fields from a calendar item we want to return back to the client.
    /// There are over 100 fields you could return.  For this demo we just want a few.
    /// </summary>
    public class CalendarItemData
    {
        public string Subject { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string[] Categories { get; set; }
    }

    /// <summary>
    /// This class saves a list of calendar items which are fetched from Exchange.  This
    /// class can then be used to iterate through all the items.
    /// </summary>
    public class CalendarItemList
    {
        private ArrayOfRealItemsType _Items;
        private int _index = 0;

        /// <summary>
        /// Saves off an array of Exchange calendar items we can enumerate.
        /// </summary>
        /// <param name="Items">Calendar item array</param>
        public CalendarItemList(ArrayOfRealItemsType Items)
        {
            _Items = Items;
            _index = 0;
        }

        /// <summary>
        /// Returns the count of calendar items being tracked.
        /// </summary>
        /// <returns>Count of items</returns>
        public int Count()
        {
            if (_Items == null)
                return 0;
            return _Items.Items.Count();
        }

        /// <summary>
        /// Get the next item from the calendar list.
        /// </summary>
        /// <returns>The calendar item</returns>
        public CalendarItemData GetNextItem()
        {
            CalendarItemData rtn = null;

            // Do a quick index check to avoid null reference.
            if (_index < _Items.Items.Count())
            {
                // Just pull the next item and move the index forward.
                ItemType item = _Items.Items[_index++];
                rtn = new CalendarItemData();

                // Uncomment the following if you want to see the entire set of field data
                // you can get from Exchange (over 100).
                //DebugTraceCalItem(item);

                // Now convert the full calendar item to our subset data.
                ExchangeWebServices.CalendarItemType calItem = item as ExchangeWebServices.CalendarItemType;
                rtn.Subject = calItem.Subject;
                rtn.StartDate = calItem.Start;
                rtn.EndDate = calItem.End;
                rtn.Categories = calItem.Categories;
            }
            return rtn;
        }

        /// <summary>
        /// Provide a comprehensive dump of all fields of an Exchange calendar item.  There
        /// are over 100 such fields.  This routine can make it easy to figure out what
        /// additional data you might want back on the client side.
        /// </summary>
        /// <param name="item">The calendar item to dump</param>
        private static void DebugTraceCalItem(ItemType item)
        {
            IEnumerable<PropertyInfo> calPropertiesAsc = ServerData.GetSortedPropertiesForType(typeof(CalendarItemType));
            ExchangeWebServices.CalendarItemType calItem = item as ExchangeWebServices.CalendarItemType;

            // Grab every field possible and display the results.
            foreach (var prop in calPropertiesAsc)
            {
                object value = typeof(CalendarItemType).InvokeMember(prop.Name, BindingFlags.GetProperty, null, item, new Object[0]);
                string propertyName = prop.Name;
                string propertyValue = value == null ? string.Empty : value.ToString();
                System.Diagnostics.Trace.WriteLine(String.Format("\t{0}: {1}", propertyName, propertyValue));
            }
        }
    }

    /// <summary>
    /// ServerData wraps a Exchange Web Services connection to a server and allows us to get
    /// at calendar data.  This version only handles
    /// </summary>
    public class ServerData
    {
        private ExchangeServiceBinding _ExchangeBinding;

        // Assume AWST for timezone to normalize across geographies.
        private const string TIMEZONE = "W. Australia Standard Time";  // http://msdn.microsoft.com/en-us/library/bb738399.aspx

        /// <summary>
        /// Connects to the Exchange web service API.  This method requires
        /// access using local authentication (e.g. integrated Windows security).
        /// </summary>
        /// <param name="ConnectionUrl">Path to exchange web service</param>

        #region 2007

        public void ConnectToServer2007(string ConnectionUrl)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2007);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        public void ConnectToServer2007(string ConnectionUrl, string Domain, string Username, string Password)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2007, Domain, Username, Password);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        #endregion

        #region 2007 SP1

        public void ConnectToServer2007_SP1(string ConnectionUrl)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2007_SP1);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        public void ConnectToServer2007_SP1(string ConnectionUrl, string Domain, string Username, string Password)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2007_SP1, Domain, Username, Password);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        #endregion

        #region 2010

        public void ConnectToServer2010(string ConnectionUrl)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        public void ConnectToServer2010(string ConnectionUrl, string Domain, string Username, string Password)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010, Domain, Username, Password);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        #endregion

        #region 2010 SP1

        public void ConnectToServer2010_SP1(string ConnectionUrl)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010_SP1);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        public void ConnectToServer2010_SP1(string ConnectionUrl, string Domain, string Username, string Password)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010_SP1, Domain, Username, Password);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        #endregion

        #region 2010 SP2

        public void ConnectToServer2010_SP2(string ConnectionUrl)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010_SP2);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        public void ConnectToServer2010_SP2(string ConnectionUrl, string Domain, string Username, string Password)
        {
            _ExchangeBinding = GetExchangeProxy(ConnectionUrl, ExchangeVersionType.Exchange2010_SP2, Domain, Username, Password);

            // We need to scope the TimeZone with the request or else we get back GMT for all calendar items.
            // This involves retrieving the TimeZone definition data from the server and associating it
            // with our request.
            ScopeTimeZone(_ExchangeBinding);
        }

        #endregion

        /// <summary>
        /// Returns an enumerator of all the calendar items you are interested in.
        /// </summary>
        /// <returns></returns>
        public CalendarItemList GetCalendarItems(DateTime StartDate, DateTime EndDate)
        {
            CalendarItemList cdata = null;

            // Create the request for all Calendar items between the specified dates
            FindItemType findItemRequest = new FindItemType();

            // Bring back ALL properties on the items returned
            ItemResponseShapeType itemProperties = new ItemResponseShapeType();
            itemProperties.BaseShape = DefaultShapeNamesType.AllProperties;
            findItemRequest.ItemShape = itemProperties;

            // We want items from the Calendar "Folder"
            DistinguishedFolderIdType[] folderIDs = { new DistinguishedFolderIdType { Id = DistinguishedFolderIdNameType.calendar } };
            findItemRequest.ParentFolderIds = folderIDs;

            // Apply filtering to request
            BuildQueryFilter(StartDate, EndDate, findItemRequest);

            // Apply ordering to request
            BuildQueryOrdering(findItemRequest);

            // Shallow traversal
            findItemRequest.Traversal = ItemQueryTraversalType.Shallow;

            // Send request, receive response
            FindItemResponseType findItemResponse = _ExchangeBinding.FindItem(findItemRequest);

            // Pull the message from the response
            ArrayOfResponseMessagesType responseMessages = findItemResponse.ResponseMessages;
            ResponseMessageType responseMessage = responseMessages.Items[0];

            // Throw exception if the response indicates an error
            if (responseMessage.ResponseCode != ResponseCodeType.NoError)
            {
                throw new Exception(string.Format("Error: {0}", responseMessage.MessageText));
            }

            // All is good so process response details
            if (responseMessage is FindItemResponseMessageType)
            {
                FindItemResponseMessageType msgType = (responseMessage as FindItemResponseMessageType);
                FindItemParentType parentType = msgType.RootFolder;

                object obj = parentType.Item;
                if (obj is ArrayOfRealItemsType)
                {
                    ArrayOfRealItemsType items = (obj as ArrayOfRealItemsType);
                    cdata = new CalendarItemList(items);
                }
            }

            return cdata;
        }

        public bool CreateAppointment(DateTime Start, DateTime End, string Subject, string Description)
        {
            bool Result = false;

            return Result;
        }

        /// <summary>
        /// ExchangeServiceBinding is the proxy used to contact the exchange server.
        /// Here we set up the binding to the Microsoft Exchange.
        /// </summary>
        /// <returns></returns>
        private static ExchangeServiceBinding GetExchangeProxy(string ConnectionUrl, ExchangeVersionType ExchangeVersion)
        {
            ExchangeServiceBinding esb = new ExchangeServiceBinding();
            esb.RequestServerVersionValue = new RequestServerVersion();
            esb.RequestServerVersionValue.Version = ExchangeVersion;
            // Ensure the request uses default NT credentials
            esb.UseDefaultCredentials = true;
            esb.Url = ConnectionUrl;
            return esb;
        }

        private static ExchangeServiceBinding GetExchangeProxy(string ConnectionUrl, ExchangeVersionType ExchangeVersion, string Domain, string Username, string Password)
        {
            ExchangeServiceBinding esb = new ExchangeServiceBinding();
            esb.RequestServerVersionValue = new RequestServerVersion();
            esb.RequestServerVersionValue.Version = ExchangeVersion;
            // Ensure the request does not use default NT credentials
            esb.UseDefaultCredentials = false;

            esb.Credentials = new NetworkCredential(Username, Password, Domain);

            esb.Url = ConnectionUrl;
            return esb;
        }

        /// <summary>
        /// Here we scope the TimeZone to EST for our request so that all date items that come back in the response use this TimeZone. 
        /// </summary>
        /// <param name="esb">Exchange proxy</param>
        private static void ScopeTimeZone(ExchangeServiceBinding esb)
        {
            // Build the TimeZone request
            GetServerTimeZonesType gstzRequest = new GetServerTimeZonesType();
            gstzRequest.Ids = new string[] { TIMEZONE };
            gstzRequest.ReturnFullTimeZoneData = true;
            gstzRequest.ReturnFullTimeZoneDataSpecified = true;
            
            // Send the request to Exchange and receive response
            GetServerTimeZonesResponseType gstzResponse = esb.GetServerTimeZones(gstzRequest);
            
            // Pull out TimeZone definition from response
            GetServerTimeZonesResponseMessageType responseMsg = gstzResponse.ResponseMessages.Items[0] as GetServerTimeZonesResponseMessageType;
            TimeZoneDefinitionType[] timezones = responseMsg.TimeZoneDefinitions.TimeZoneDefinition;
            TimeZoneDefinitionType tzdt = timezones[0];
            
            // Associate the TimeZone definition with the Exchange proxy
            esb.TimeZoneContext = new TimeZoneContextType();
            esb.TimeZoneContext.TimeZoneDefinition = tzdt;
        }

        /// <summary>
        /// Build the StartDate &lt;= items &gt;= EndDate filter and associate it with the request object
        /// </summary>
        /// <param name="StartDate">Start range of calendar items</param>
        /// <param name="EndDate">End range of calendar items</param>
        /// <param name="findItemRequest">the request that will be sent to exchange that we want to apply filtering to</param>
        private void BuildQueryFilter(DateTime StartDate, DateTime EndDate, FindItemType findItemRequest)
        {
            // This is the filter object that we will associate with the request
            RestrictionType restriction = new RestrictionType();

            // Apply start date filtering to request
            PathToUnindexedFieldType startDateProperty = new PathToUnindexedFieldType();
            startDateProperty.FieldURI = UnindexedFieldURIType.calendarStart;
            FieldURIOrConstantType startDateValue = new FieldURIOrConstantType();
            startDateValue.Item = new ConstantValueType();
            (startDateValue.Item as ConstantValueType).Value = StartDate.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");
            IsGreaterThanType isGreaterThan = new IsGreaterThanType();
            isGreaterThan.Item = startDateProperty;
            isGreaterThan.FieldURIOrConstant = startDateValue;

            // Apply end date filtering to request
            PathToUnindexedFieldType endDateProperty = new PathToUnindexedFieldType();
            endDateProperty.FieldURI = UnindexedFieldURIType.calendarEnd;
            FieldURIOrConstantType endDateValue = new FieldURIOrConstantType();
            endDateValue.Item = new ConstantValueType();
            (endDateValue.Item as ConstantValueType).Value = EndDate.ToUniversalTime().ToString("yyyy-MM-ddThh:mm:ssZ");
            IsLessThanType isLessThan = new IsLessThanType();
            isLessThan.Item = endDateProperty;
            isLessThan.FieldURIOrConstant = endDateValue;

            // Aggregate the filtering in AND logic
            AndType inBetweenStartAndEnd = new AndType();
            inBetweenStartAndEnd.Items = new SearchExpressionType[] { isGreaterThan, isLessThan };

            // Pass the complete condition to the request
            restriction.Item = inBetweenStartAndEnd;
            findItemRequest.Restriction = restriction;
        }

        /// <summary>
        /// Ensure that the response provides us with calendar items ordered by ascending date
        /// </summary>
        /// <param name="findItemRequest">the request that will be sent to exchange that we want to apply the ordering to</param>
        private static void BuildQueryOrdering(FindItemType findItemRequest)
        {
            // Ensure items are ordered by start date
            PathToUnindexedFieldType startDateOrder = new PathToUnindexedFieldType();
            startDateOrder.FieldURI = UnindexedFieldURIType.calendarStart;
            FieldOrderType[] fieldsOrder = new FieldOrderType[] { new FieldOrderType { Item = startDateOrder, Order = SortDirectionType.Ascending } };
            findItemRequest.SortOrder = fieldsOrder;
        }

        /// <summary>
        /// Get all properties for the provided type and return them in alphabetical order
        /// </summary>
        /// <param name="type">the Type we want to enumerate the properties of</param>
        /// <returns></returns>
        public static IEnumerable<PropertyInfo> GetSortedPropertiesForType(Type type)
        {
            // Use reflection to get all the properties on the type
            type = typeof(ExchangeWebServices.CalendarItemType);
            PropertyInfo[] properties = type.GetProperties();

            // Sort the properties by name
            IEnumerable<PropertyInfo> calPropertiesAsc =
                from prop in properties
                orderby prop.Name
                select prop;

            return calPropertiesAsc;
        }

    }
}
