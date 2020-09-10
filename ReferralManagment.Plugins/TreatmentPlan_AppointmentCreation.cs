using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ReferralManagment.Plugins

{

    public class TreatmentPlan_AppointmentCreation : IPlugin

    {
        Entity _treatmentPlan, treatmentPlan;
        DateTime _authorizationStartDatetimeUTC ,_authorizationendDatetimeUTC , _planstartingdateTimeUTC = DateTime.UtcNow;
        DateTime _authorizationStartDatetimeLocal, _authorizationendDatetimeLocal , _planstartingdateTimeLocal = DateTime.Now;
        int _authorizationVisitsLeft, _userTimeZoneCode = 0;
        List<Guid> _createdAppointmentIds = new List<Guid>();
        List<string> _daysofWeek = new List<string>();

        
        public void Execute(IServiceProvider serviceProvider)

        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService _service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.PostEntityImages.Contains("PostImage")) // Post Image Check
            {
                try
                {
                    tracingService.Trace("Plugin : Referral Managment. TreatmentPlan_AppointmentCreation has started at UTC Time : " + DateTime.UtcNow.ToString());

                    _treatmentPlan = (Entity)context.PostEntityImages["PostImage"];
                    treatmentPlan = (Entity)context.InputParameters["Target"];

                    if (_treatmentPlan != null && _treatmentPlan.LogicalName == "clrcomp_treatmentplan" && _treatmentPlan.Contains("clrcomp_daysofweek") && _treatmentPlan.Contains("clrcomp_createappointment") && _treatmentPlan.Contains("clrcomp_startdateandtime") && (_treatmentPlan.GetAttributeValue<DateTime>("clrcomp_startdateandtime").Date <= DateTime.MaxValue.Date || _treatmentPlan.GetAttributeValue<DateTime>("clrcomp_startdateandtime").Date >= DateTime.MinValue.Date))
                    {
                        _userTimeZoneCode = RetrieveCurrentUsersTimeZoneCode(_service);
                         _planstartingdateTimeUTC = _treatmentPlan.GetAttributeValue<DateTime>("clrcomp_startdateandtime");
                         _planstartingdateTimeLocal= RetrieveLocalTimeFromUTCTime(_service, _planstartingdateTimeUTC, _userTimeZoneCode);
                        var _createAppointment = _treatmentPlan.GetAttributeValue<OptionSetValue>("clrcomp_createappointment").Value;
                        OptionSetValueCollection days = _treatmentPlan.GetAttributeValue<OptionSetValueCollection>("clrcomp_daysofweek");

                        for (int i = 0; i < days.Count; i++)
                        {
                            if (days[i].Value == 397780000) _daysofWeek.Add("Monday");
                            if (days[i].Value == 397780001) _daysofWeek.Add("Tuesday");
                            if (days[i].Value == 397780002) _daysofWeek.Add("Wednesday");
                            if (days[i].Value == 397780003) _daysofWeek.Add("Thursday");
                            if (days[i].Value == 397780004) _daysofWeek.Add("Friday");
                            if (days[i].Value == 397780005) _daysofWeek.Add("Saturday");

                        }               


                        if (_createAppointment == 397780000) // Checking if Create Appointment is Yes
                        {
                            tracingService.Trace("Searching Authorizations");

                            var fetch = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                          <entity name='clrcomp_authorization'>
                            <attribute name='clrcomp_authorizationid'/>
                            <attribute name='clrcomp_name'/>
                            <attribute name='createdon'/>
                            <attribute name='clrcomp_visitsleft'/>
                            <attribute name='clrcomp_visitsauthorized'/>
                            <attribute name='clrcomp_enddate'/>
                             <attribute name='clrcomp_startdate'/>
                            <order descending='false' attribute='clrcomp_name'/>
                            <filter type='and'>
                              <condition attribute='clrcomp_treatmentplan' value ='" + _treatmentPlan.Id + @"' operator='eq'/>      
                              <condition attribute='statecode' operator='eq' value='0'/>
                            </filter>
                          </entity>
                        </fetch>";

                            EntityCollection entitycollection = _service.RetrieveMultiple(new FetchExpression(fetch));

                            if (entitycollection.Entities.Count > 0)
                            {

                                tracingService.Trace("Authorization Retrieved count : " + entitycollection.Entities.Count.ToString());
                                _authorizationendDatetimeUTC = (entitycollection.Entities[0].GetAttributeValue<DateTime>("clrcomp_enddate"));
                                _authorizationendDatetimeLocal= RetrieveLocalTimeFromUTCTime(_service, _authorizationendDatetimeUTC, _userTimeZoneCode);
                                _authorizationStartDatetimeUTC = entitycollection.Entities[0].GetAttributeValue<DateTime>("clrcomp_startdate");
                                _authorizationStartDatetimeLocal = RetrieveLocalTimeFromUTCTime(_service, _authorizationStartDatetimeUTC, _userTimeZoneCode);
                                _authorizationVisitsLeft = entitycollection.Entities[0].GetAttributeValue<int>("clrcomp_visitsleft");

                                tracingService.Trace("Visits Left :" + _authorizationVisitsLeft);

                                if (_planstartingdateTimeLocal.Date <= _authorizationendDatetimeLocal.Date || _authorizationendDatetimeUTC <= DateTime.MinValue)

                                {
                                    if (_planstartingdateTimeLocal.Date <= _authorizationStartDatetimeLocal.Date) { _planstartingdateTimeLocal = _authorizationStartDatetimeLocal; }

                                    List<DateTime> _appointmentsDate = GetDatesBetween(_planstartingdateTimeLocal, _authorizationendDatetimeLocal, _daysofWeek, _authorizationVisitsLeft, tracingService, _service);

                                    foreach (DateTime date in _appointmentsDate)
                                    {
                                        Entity Appointment = new Entity("appointment");
                                        Appointment["subject"] = "Auto Created Appointment";
                                        Appointment["regardingobjectid"] = new EntityReference("clrcomp_treatmentplan", _treatmentPlan.Id);
                                        Appointment["scheduledstart"] = date;
                                        Appointment["scheduledend"] = date.AddMinutes(30);
                                        _createdAppointmentIds.Add((Guid)_service.Create(Appointment));
                                        tracingService.Trace("Appointment Created Successfully  ");
                                    }

                                    if (_createdAppointmentIds.Count > 0)
                                    {
                                        treatmentPlan.Attributes["clrcomp_createappointment"] = new OptionSetValue(397780001);
                                        tracingService.Trace("Updating Treatment Plan");
                                        _service.Update(treatmentPlan);
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new InvalidPluginExecutionException("Error in Appointment Creation Plugin . Error Message : " + ex.Message + ex.InnerException);
                }
            }
        }

        public List<DateTime> GetDatesBetween(DateTime _startdate, DateTime _endDate, List<string> selecteddays, int visits, ITracingService tracingService, IOrganizationService service)
        {
            var flag = false;
            DateTime date;
            var count = 0;

            List<DateTime> businessClosureDate = new List<DateTime>();
            List<DateTime> allDates= new List<DateTime>();
            var bfetch = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                          <entity name='msdyn_businessclosure'>
                            <attribute name='msdyn_businessclosureid'/>
                            <attribute name='msdyn_name'/>
                            <attribute name='msdyn_starttime'/>
                            <attribute name='msdyn_endtime'/>
                            <order descending='false' attribute='msdyn_name'/>
                            <filter type='and'>
                              <condition attribute='statecode' operator='eq' value='0'/>
                            </filter>
                          </entity>
                        </fetch>";

            EntityCollection entitycollection = service.RetrieveMultiple(new FetchExpression(bfetch));

            if (entitycollection.Entities.Count > 0)
            {
                tracingService.Trace("Business Closure Retrieved count : " + entitycollection.Entities.Count.ToString());
                foreach (Entity b in entitycollection.Entities)
                {
                    //Checking for Correct Busines Closure Date accrding to User Time Zone 
                    if(_userTimeZoneCode>=90) businessClosureDate.Add(b.GetAttributeValue<DateTime>("msdyn_endtime").Date);
                    if (_userTimeZoneCode < 90) businessClosureDate.Add(b.GetAttributeValue<DateTime>("msdyn_starttime").Date);
                    tracingService.Trace("Business Closure UTC Date : " + (b.GetAttributeValue<DateTime>("msdyn_endtime").Date).ToString());
                }
            }
       


            if (!_endDate.Equals(DateTime.MinValue.ToLocalTime()))
            {
                for (date = _startdate; date <= _endDate; date = date.AddDays(1))
                {
                    if (!businessClosureDate.Contains(date.Date))
                    {
                        tracingService.Trace("Entered for loop for Date check");
                        tracingService.Trace(date.ToString("dddd").ToString());
                        string daytocheck = date.ToString("dddd").ToString();

                                if (selecteddays.Contains(daytocheck)) 
                                {
                                    tracingService.Trace("Days Matched");
                                    flag = true;
                                }
                                else { flag = false; }

                                if (flag == true)
                                {
                                    tracingService.Trace("Entered for loop for Flag check");
                                    if (count < visits)
                                    {
                                        allDates.Add(date);
                                        count++;
                                    }
                                    else { return allDates; };
                                }
                                tracingService.Trace(count.ToString());                           
                        
                    }
                }
            }
            else
            {
                for (date = _startdate; date <= DateTime.MaxValue; date = date.AddDays(1))
                {
                    if (!businessClosureDate.Contains(date.Date))
                    {
                        tracingService.Trace("Entered for loop for Date check");
                        tracingService.Trace(date.ToString("dddd").ToString());
                        string daytocheck = date.ToString("dddd").ToString();

                                if (selecteddays.Contains(daytocheck))
                                {
                                    tracingService.Trace("Days Matched");
                                    flag = true;
                                }
                                else { flag = false; }

                                if (flag == true)
                                {
                                    tracingService.Trace("Entered for loop for Flag check");
                                    if (count < visits)
                                    {
                                        allDates.Add(date);
                                        count++;
                                    }
                                    else { return allDates; };
                                }
                                tracingService.Trace(count.ToString());
                    }
                }
            }
            return allDates;
        }
        //function for retrieving User Time Zone Code
        public int RetrieveCurrentUsersTimeZoneCode(IOrganizationService service)
        {
            var currentUserSettings = service.RetrieveMultiple(
            new QueryExpression("usersettings")
            {
                ColumnSet = new ColumnSet("localeid", "timezonecode"),
                Criteria = new FilterExpression
                {
                    Conditions =
               {
            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
               }
                }
            }).Entities[0].ToEntity<Entity>();
            return (int)currentUserSettings.Attributes["timezonecode"];
        }

        //function for converting UTC Time Zone to Local
        private static DateTime RetrieveLocalTimeFromUTCTime(IOrganizationService _serviceProxy, DateTime utcTime, int? _timeZoneCode)

        {

            if (!_timeZoneCode.HasValue)

                return utcTime;



            var request = new LocalTimeFromUtcTimeRequest

            {

                TimeZoneCode = _timeZoneCode.Value,

                UtcTime = utcTime.ToUniversalTime()

            };



            var response = (LocalTimeFromUtcTimeResponse)_serviceProxy.Execute(request);

            return response.LocalTime;

        }

    }

}




