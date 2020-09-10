
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Device.Location;
using System.Linq;


namespace ReferralManagment.Plugins
{
    public class TreatmentPlan_FacilityCreation : IPlugin
    {
        private readonly string _unsecureString;
        public TreatmentPlan_FacilityCreation(string unsecureString, string secureString)
        {
            if (String.IsNullOrWhiteSpace(unsecureString))
            {
                throw new InvalidPluginExecutionException("Unsecure and secure strings (For Distance Calculation) are required for this plugin to execute.");
            }

            _unsecureString = unsecureString;
        }// Cheking for Secure/Unecure String
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity && context.Depth == 1)
            {
                try
                {
                    #region declaring Variable
                    //var _disDdata = new Dictionary<double, Guid>();
                    List<location> _items = new List<location>();
                    double _sLatitude;
                    double _sLongitude;
                    double _dLatitude;
                    double _dLongitude;
                    int _faciitiesCreated = 0;

                    #endregion

                    Entity _treatmentPlan = (Entity)context.InputParameters["Target"];

                    if (_treatmentPlan != null && _treatmentPlan.LogicalName == "clrcomp_treatmentplan" && _treatmentPlan.Contains("clrcomp_patient"))
                    {
                        EntityReference contact = _treatmentPlan.GetAttributeValue<EntityReference>("clrcomp_patient");
                        Guid contactid = contact.Id;
                        tracingService.Trace("Patient ID : " + contactid);

                        #region Retrieve Patient Address

                        Entity _contact = service.Retrieve("contact", contactid, new ColumnSet(true));
                        if (_contact.Contains("address1_latitude") && _contact.Contains("address1_longitude"))
                        {
                            _sLatitude = _contact.GetAttributeValue<double>("address1_latitude");
                            _sLongitude = _contact.GetAttributeValue<double>("address1_longitude");
                            tracingService.Trace("Patient Address --  Latitude : " + _sLatitude + "Longitiude" + _sLongitude);
                        }
                        else { throw new InvalidPluginExecutionException("Patient Address is not present in System. Please update Latitude and Longitude Value."); }

                        #endregion

                        #region Retrieve Facilitis Address and Calculate Address

                        var fetch = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                            <entity name='account'>
                            <attribute name='address1_longitude'/>
                            <attribute name='address1_latitude'/>
                            <order descending='false' attribute='name'/>
                            <filter type='and'>
                             <condition attribute='statecode' value ='0' operator='eq'/>
                             <condition attribute='address1_latitude' operator='not-null'/>
                             <condition attribute='address1_longitude' operator='not-null'/>
                            </filter>
                            </entity>
                            </fetch>";




                        EntityCollection entitycollection = service.RetrieveMultiple(new FetchExpression(fetch));
                        if (entitycollection.Entities.Count > 0)
                        {
                            tracingService.Trace("Facilities Retrieved count : " + entitycollection.Entities.Count.ToString());
                            for (var i = 0; i < entitycollection.Entities.Count; i++)
                            {

                                _dLatitude = entitycollection.Entities[i].GetAttributeValue<double>("address1_latitude");
                                _dLongitude = entitycollection.Entities[i].GetAttributeValue<double>("address1_longitude");
                                double _temfordistance = Calculate(_sLatitude, _sLongitude, _dLatitude, _dLongitude);
                                if (_temfordistance > 0)
                                {
                                    //_disDdata.Add(_temfordistance, entitycollection.Entities[i].Id);
                                    _items.Add(new location { distance = _temfordistance, accid = entitycollection.Entities[i].Id });
                                }

                            }

                        }
                        #endregion

                        #region Delete Existing Facilities
                        var delfetch = @"<fetch distinct='false' mapping='logical' output-format='xml-platform' version='1.0'>
                            <entity name='clrcomp_nearestfaciltiies'>
                            <attribute name='clrcomp_nearestfaciltiiesid'/>
                            <order descending='false' attribute='clrcomp_name'/>
                            <filter type='and'>
                             <condition attribute='clrcomp_treatmentplan' value ='" + _treatmentPlan.Id + @"' operator='eq'/>
                            </filter>
                            </entity>
                            </fetch>";

                        EntityCollection _delentitycollection = service.RetrieveMultiple(new FetchExpression(delfetch));
                        if (entitycollection.Entities.Count > 0)
                        {
                            tracingService.Trace("Facilities Retrieved count : " + entitycollection.Entities.Count.ToString());
                            foreach (Entity delent in _delentitycollection.Entities)
                            {
                                service.Delete(delent.LogicalName, delent.Id);
                            }



                            // Array.Sort(_disDdata);
                        }
                        #endregion

                        tracingService.Trace("Nearest Facilities Distance : ");


                        for (int index = 0; index < _items.Count; index++)
                        {
                            location List = _items[index];
                            tracingService.Trace("Nearest Facilities Distance : " + List.distance);
                            if (List.distance < Convert.ToDouble(_unsecureString) && _faciitiesCreated < 3)
                            {
                                tracingService.Trace("Maximum Distance : " + Convert.ToDouble(_unsecureString));
                                Entity _account = service.Retrieve("account", List.accid, new ColumnSet(true));
                                #region Create Nearest Facilities
                                Entity _nearestFacility = new Entity("clrcomp_nearestfaciltiies");
                                _nearestFacility.Attributes["clrcomp_distanceinmiles"] = List.distance;
                                _nearestFacility.Attributes["clrcomp_name"] = _account.GetAttributeValue<string>("name");
                                _nearestFacility.Attributes["clrcomp_latitude"] = _account.GetAttributeValue<double>("address1_latitude");
                                _nearestFacility.Attributes["clrcomp_longitude"] = _account.GetAttributeValue<double>("address1_longitude");
                                _nearestFacility.Attributes["clrcomp_treatmentplan"] = new EntityReference("clrcomp_treatmentplan", _treatmentPlan.Id);
                                _nearestFacility.Attributes["clrcomp_facility"] = new EntityReference("account", List.accid);
                                Guid _nearestFacilityId = service.Create(_nearestFacility);
                                if (_nearestFacilityId != null) { _faciitiesCreated++; }
                                #endregion
                            }
                        }


                    }
                }
                catch (Exception ex)
                {

                    throw new InvalidPluginExecutionException(ex.Message);
                }
            }

        }

        public static double Calculate(double _sLatitude, double _sLongitude, double _dLatitude, double _dLongitude)
        {
            var sCoord = new GeoCoordinate(_sLatitude, _sLongitude);
            var dCoord = new GeoCoordinate(_dLatitude, _dLongitude);

            return (sCoord.GetDistanceTo(dCoord) * 0.000621371); //converting meters into Miles
        }
    }
}

