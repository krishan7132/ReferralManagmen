using System.Runtime.Serialization;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Device.Location;
using System.Linq;


namespace ReferralManagment.Plugins
{
    public class TreatmentPlan_AppointmentCreation : IPlugin
    {
        DateTime authstart;
        private readonly string _unsecureString;
 // Cheking for Secure/Unecure String
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
           
            if (context.PostEntityImages.Contains("PostImage"))
            {
                try
                {
                    Entity _treatmentPlan = (Entity)context.PostEntityImages["PostImage"];

                    //Entity _treatmentPlan = (Entity)context.InputParameters["Target"];

                    if (_treatmentPlan != null && _treatmentPlan.LogicalName == "clrcomp_treatmentplan" && _treatmentPlan.Contains("clrcomp_daysofweek"))
                    {
                        tracingService.Trace("Inside Operational Loop");
                        var starting = (_treatmentPlan.GetAttributeValue<DateTime>("clrcomp_startdateandtime")).ToLocalTime();
                        bool _createAppointment = _treatmentPlan.GetAttributeValue<bool>("clrcomp_createappointment");
                        string[] daysofweek = new string[6];
                        // List<String> daysofweek = new List<string>();
                        //  OptionSetValueCollection days = (OptionSetValueCollection)_treatmentPlan["clrcomp_daysofweek"];
                        OptionSetValueCollection days = _treatmentPlan.GetAttributeValue<OptionSetValueCollection>("clrcomp_daysofweek");
                        for (int i = 0; i < days.Count; i++)
                        {
                            if (days[i].Value == 397780000) daysofweek[i] = "Monday";
                            if (days[i].Value == 397780001) daysofweek[i] = "Tuesday";
                            if (days[i].Value == 397780002) daysofweek[i] = "Wednesday";
                            if (days[i].Value == 397780003) daysofweek[i] = "Thursday";
                            if (days[i].Value == 397780004) daysofweek[i] = "Friday";
                            if (days[i].Value == 397780005) daysofweek[i] = "Saturday";
                        }

                        tracingService.Trace("Days" + days + daysofweek[0] + daysofweek[1] + daysofweek[2] + daysofweek[3] + daysofweek[4] + daysofweek[5]);

                        var ending = DateTime.Now;
                        var visits = 0;
                        if (_createAppointment)
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
                            EntityCollection entitycollection = service.RetrieveMultiple(new FetchExpression(fetch));
                            if (entitycollection.Entities.Count > 0)
                            {

                                tracingService.Trace("Authorization Retrieved count : " + entitycollection.Entities.Count.ToString());
                                ending = (entitycollection.Entities[0].GetAttributeValue<DateTime>("clrcomp_enddate")).ToLocalTime();
                                authstart = entitycollection.Entities[0].GetAttributeValue<DateTime>("clrcomp_startdate").ToLocalTime();
                                visits = entitycollection.Entities[0].GetAttributeValue<int>("clrcomp_visitsleft");
                                tracingService.Trace("Visits Left :" + visits);
                                tracingService.Trace("End Date" + ending.ToString());
                           


                            //    string[] daysofweek = { };  //Value from picklist as string (Eg Monday, Sunday etc
                            if (starting.Date <= ending.Date)
                            {
                                if (starting.Date <= authstart.Date) { starting = authstart; }
                                List<DateTime> dates = GetDatesBetween(starting, ending, daysofweek, visits, tracingService, service);
                                tracingService.Trace(dates.ToString());
                                foreach (DateTime date in dates)
                                {

                                    Entity Appointment = new Entity("appointment");
                                    Appointment["subject"] = "Auto Created Appointment";
                                    Appointment["regardingobjectid"] = new EntityReference("clrcomp_treatmentplan", _treatmentPlan.Id);
                                    Appointment["scheduledstart"] = date.ToLocalTime();
                                    Appointment["scheduledend"] = date.ToLocalTime();
                                    service.Create(Appointment);
                                    tracingService.Trace("Appointment Created with id :");
                                }
                            }

                        }
                        }
                    }
                }
                catch (Exception ex)
                {

                    throw new InvalidPluginExecutionException("Error in Appointment Creation Plugin . Error Message : " +ex.Message);
                }
            }

        }



        public List<DateTime> GetDatesBetween(DateTime startDate, DateTime endDate, String[] selecteddays, int visits, ITracingService tracingService, IOrganizationService service)

        {
            var flag = false;
            DateTime date;
            List<DateTime> allDates = new List<DateTime>();
            List<DateTime> businessClosureDate = new List<DateTime>();
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
                    businessClosureDate.Add(b.GetAttributeValue<DateTime>("msdyn_endtime").ToLocalTime().Date);

                    tracingService.Trace("Business Closure Local Date : " + (b.GetAttributeValue<DateTime>("msdyn_endtime").ToLocalTime().Date).ToString());
                    tracingService.Trace("Business Closure UTC Date : " + (b.GetAttributeValue<DateTime>("msdyn_endtime").ToUniversalTime().Date).ToString());
                    tracingService.Trace("Business Closure Date : " + (b.GetAttributeValue<DateTime>("msdyn_endtime").Date).ToString());
                }
            }
                var count = 0;
            tracingService.Trace("Start Date" + startDate.ToString());
            for (date = startDate; date <= endDate; date = date.AddDays(1))
            {
                if (!businessClosureDate.Contains(date.Date))
                {
                    tracingService.Trace("Entered for loop for Date check");
                    tracingService.Trace(date.ToString("dddd").ToString());
                    string daytocheck = date.ToString("dddd").ToString();
                    foreach (string s in selecteddays)
                    {
                        if (s != null)
                        {
                            tracingService.Trace("Entered op");
                            if (s.Contains(daytocheck)) //dateValue.ToString("dddd")   if (selecteddays.Contains(date.DayOfWeek.ToString()))
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
                                    allDates.Add(date.Date);
                                    count++;

                                }
                            }
                            tracingService.Trace(count.ToString());

                        }
                    }
                }
            }
            return allDates;
        }
    }
}

