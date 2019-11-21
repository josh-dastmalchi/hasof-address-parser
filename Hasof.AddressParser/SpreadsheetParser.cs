using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDataReader;

namespace Hasof.AddressParser
{
    public class SpreadsheetParser
    {
        int nameIndex;
        int phoneIndex;
        int? street1Index;
        int? street2Index;
        int? cityIndex;
        int? stateIndex;
        int? zipIndex;
        int? locationIndex;
        int iconIndex;
        int? straightIndex;
        int? rumIndex;
        int? caskIndex;
        public List<Vendor> Parse(IExcelDataReader reader)
        {
            var vendors = new List<Vendor>();
            SetIndexesBasedOnHeaders(reader);
            //do
            //{
            while (reader.Read())
            {
                try
                {
                    string address;
                    string iconUrl;
                    bool carriesStraight = false;
                    bool carriesRum = false;
                    bool carriesCask = false;
                    var name = GetValue(reader, nameIndex);
                    var phone = PhoneFormatter.Format(GetValue(reader, phoneIndex));
                    if (locationIndex.HasValue)
                    {
                        address = GetValue(reader, locationIndex.Value);
                    }
                    else
                    {
                        var street1 = GetValue(reader, street1Index.Value);
                        string street2 = string.Empty;
                        if (street2Index.HasValue)
                        {
                            street2 = GetValue(reader, street2Index.Value);
                        }
                        var city = GetValue(reader, cityIndex.Value);
                        var state = GetValue(reader, stateIndex.Value);
                        var zip = GetValue(reader, zipIndex.Value);
                        address = $"{street1} {street2} {city}, {state} {zip}";
                    }
                    if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(address))
                    {
                        continue;
                    }

                    iconUrl = GetValue(reader, iconIndex);

                    if (straightIndex.HasValue)
                    {
                        carriesStraight = IsTruthy(GetValue(reader, straightIndex.Value));
                    }

                    if (rumIndex.HasValue)
                    {
                        carriesRum = IsTruthy(GetValue(reader, rumIndex.Value));
                    }

                    if (caskIndex.HasValue)
                    {
                        carriesCask = IsTruthy(GetValue(reader, caskIndex.Value));
                    }

                    var vendor = new Vendor
                    {
                        Name = name,
                        Address = address,
                        Phone = phone,
                        IconUrl = iconUrl,
                        CarriesBarrelRestedGin = carriesRum,
                        CarriesCaskStrengthStraightBourbon = carriesCask,
                        CarriesStraightBourbon = carriesStraight
                    };
                    vendors.Add(vendor);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    throw;
                }
            }
            //} while (reader.NextResult());

            return vendors;
        }

        private void SetIndexesBasedOnHeaders(IExcelDataReader reader)
        {
            reader.Read();
            var headers = new List<string>();
            for (var index = 0; index < reader.FieldCount; index++)
            {
                var headerText = GetValue(reader, index) ?? string.Empty;
                headers.Add(headerText.ToLower());
            }

            var streets = headers.Where(x => x.Contains("street") || x.Contains("address")).ToList();
            var cities = headers.Where(x => x.Contains("city")).ToList();
            var states = headers.Where(x => x.Contains("state")).ToList();
            var names = headers.Where(x => x.Contains("name")).ToList();
            var zips = headers.Where(x => x.Contains("zip") || x.Contains("postal")).ToList();
            var phones = headers.Where(x => x.Contains("phone")).ToList();
            var locations = headers.Where(x => x.Contains("location")).ToList();
            var icons = headers.Where(x => x.Contains("icon")).ToList();
            var carriesStraight = headers.Where(x => string.Equals("straight bourbon", x, StringComparison.OrdinalIgnoreCase)).ToList();
            var carriesRum = headers.Where(x => string.Equals("Cask Strength Straight Bourbon", x, StringComparison.OrdinalIgnoreCase)).ToList();
            var carriesCask = headers.Where(x => string.Equals("Barrel Rested Gin", x, StringComparison.OrdinalIgnoreCase)).ToList();

            if (cities.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a city: " + string.Join(" , ", cities));
            }
            if (states.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a state: " + string.Join(" , ", states));
            }
            if (names.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a name: " + string.Join(" , ", names));
            }
            if (zips.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a zip code: " + string.Join(" , ", zips));
            }
            if (phones.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a phone number: " + string.Join(" , ", phones));
            }

            if (locations.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as a location: " + string.Join(" , ", locations));
            }

            if (icons.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as an icon: " + string.Join(" , ", icons));
            }

            if (carriesStraight.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as an icon: " + string.Join(" , ", icons));
            }

            if (carriesRum.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as an icon: " + string.Join(" , ", icons));
            }

            if (carriesCask.Count > 1)
            {
                throw new ParsingFormatException("More than one column was identified as an icon: " + string.Join(" , ", icons));
            }
            var name = names.SingleOrDefault();

            if (name == null)
            {
                throw new ParsingFormatException("There must be a column containing the customer name.");
            }

            var phone = phones.SingleOrDefault();
            if (phone == null)
            {
                throw new ParsingFormatException("There must be a column containing the phone number.");
            }

            var icon = icons.SingleOrDefault();
            if (icon == null)
            {
                throw new ParsingFormatException("There must be a column containing the icon url.");
            }
            nameIndex = headers.IndexOf(name);
            phoneIndex = headers.IndexOf(phone);
            iconIndex = headers.IndexOf(icon);

            var location = locations.FirstOrDefault();
            if (location != null)
            {
                locationIndex = headers.IndexOf(location);
                return;
            }

            var street1 = streets.FirstOrDefault();
            var street2 = streets.Skip(1).FirstOrDefault();
            var city = cities.FirstOrDefault();
            var state = states.FirstOrDefault();
            var zip = zips.FirstOrDefault();


            if (street1 == null || city == null || state == null || zip == null)
            {
                throw new ParsingFormatException("There must either be a column for location, or columns for all of: street, city, state, zip");
            }
            street1Index = headers.IndexOf(street1);
            cityIndex = headers.IndexOf(city);
            stateIndex = headers.IndexOf(state);
            zipIndex = headers.IndexOf(zip);

            if (street2 != null)
            {
                street2Index = headers.IndexOf(street2);
            }

            if (carriesStraight.Any())
            {
                straightIndex = headers.IndexOf(carriesStraight[0]);
            }
            if (carriesCask.Any())
            {
                caskIndex = headers.IndexOf(carriesStraight[0]);
            }
            if (carriesRum.Any())
            {
                rumIndex = headers.IndexOf(carriesStraight[0]);
            }

        }


        private string GetValue(IExcelDataReader reader, int ordinal)
        {
            try
            {
                var obj = reader.GetValue(ordinal);
                return Convert.ToString(obj);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private static bool IsTruthy(string s)
        {
            if (s == null)
            {
                return false;
            }
            return
                string.Equals("y", s, StringComparison.OrdinalIgnoreCase) ||
                string.Equals("t", s, StringComparison.OrdinalIgnoreCase) ||
                string.Equals("true", s, StringComparison.OrdinalIgnoreCase) ||
                string.Equals("yes", s, StringComparison.OrdinalIgnoreCase);

        }
    }
}