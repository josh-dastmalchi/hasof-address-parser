using ExcelDataReader;
using GoogleMapsApi;
using GoogleMapsApi.Entities.Geocoding.Request;
using GoogleMapsApi.Entities.Geocoding.Response;
using GoogleMapsApi.Entities.PlacesDetails.Request;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Hasof.AddressParser
{
    public partial class ParserForm : Form
    {
        public ParserForm()
        {
            InitializeComponent();
            this.DragEnter += Form1_DragEnter;
            this.DragDrop += Form1_DragDrop;
        }

        private async Task ParseFile(string path)
        {
            try
            {

                Enabled = false;
                Cursor.Current = Cursors.WaitCursor;
                toolStripStatusLabel1.Text = @"Processing spreadsheet...";
                string apiKey = ConfigurationManager.AppSettings["google-maps-api-key"];
                List<Vendor> vendors;
                try
                {


                    using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                    {
                        using (var reader = ExcelReaderFactory.CreateReader(stream))
                        {
                            var parser = new SpreadsheetParser();
                            vendors = parser.Parse(reader);

                        }
                    }
                }
                catch (ParsingFormatException e)
                {
                    MessageBox.Show(e.Message, @"Parsing Error");
                    FinishedProcessing();
                    return;
                }
                catch (Exception e)
                {
                    MessageBox.Show(@"There was an unexpected problem parsing the file. " + e.Message, @"Unknown error.");
                    FinishedProcessing();
                    return;
                }
                var outputLines = new List<string>();
                outputLines.Add("var locations = [");
                var errors = new List<string>();
                for (var index = 0; index < vendors.Count; index++)
                {
                    toolStripStatusLabel1.Text = $"Processing vendor {index + 1} of {vendors.Count}";
                    var geocodingRequest = new GeocodingRequest
                    {
                        Address = vendors[index].Address,
                        ApiKey = apiKey

                    };

                    var coded = await GoogleMaps.Geocode.QueryAsync(geocodingRequest);
                    if (coded.Status != Status.OK)
                    {
                        errors.Add($"Unable to geocode {vendors[index].Name}. Google reply: " + coded.Status);
                        continue;
                    }
                    var result = coded.Results.First();
                    if (result.PartialMatch)
                    {
                        errors.Add($"Only got a partial match for {vendors[index].Name}.");
                    }

                    var placeDetailsRequest = new PlacesDetailsRequest
                    {
                        PlaceId = result.PlaceId,
                        ApiKey = apiKey
                    };

                    var placesDetailsResponse = await GoogleMaps.PlacesDetails.QueryAsync(placeDetailsRequest);
                    if (placesDetailsResponse.Status != GoogleMapsApi.Entities.PlacesDetails.Response.Status.OK)
                    {
                        errors.Add($"Unable to find the url for {vendors[index].Name} (line {index + 1})");
                        continue;
                    }
                    try
                    {
                        if (placesDetailsResponse.Result.Types.Contains("locality") || placesDetailsResponse.Result.Types.Contains("post_box"))
                        {
                            errors.Add($"Google returned an address for the vendor {vendors[index].Name} that didn\"t look like a normal address. Ignoring it.");
                        }
                        else
                        {
                            var streetNumber = result.AddressComponents.SingleOrDefault(x => x.Types.Contains("street_number"))?.ShortName;
                            var route = result.AddressComponents.SingleOrDefault(x => x.Types.Contains("route"))?.ShortName;
                            var locality = result.AddressComponents.SingleOrDefault(x => x.Types.Contains("locality"))?.ShortName;
                            var state = result.AddressComponents.SingleOrDefault(x => x.Types.Contains("administrative_area_level_1"))?.ShortName;
                            var postalCode = result.AddressComponents.SingleOrDefault(x => x.Types.Contains("postal_code"))?.ShortName;
                            var partialMatchText = result.PartialMatch ? "// Partial match double check me!" : string.Empty;
                            var address = $"{streetNumber} {route}, {locality}, {state} {postalCode}";
                            // manually formatting json, what have I done
                            outputLines.Add(
                                $"{{\"name\" : \"{vendors[index].Name}\", \"address\" : \"{address}\", \"phone\" : \"{vendors[index].Phone}\", \"googleMapsUrl\" : \"{placesDetailsResponse.Result.URL}\", \"latitude\": \"{result.Geometry.Location.Latitude}\", \"longitude\" :\"{result.Geometry.Location.Longitude}\", \"iconUrl\" : \"{vendors[index].IconUrl}\"}},{partialMatchText}");

                        }
                    }
                    catch (Exception)
                    {
                        errors.Add($"Failed to process {vendors[index].Name} (line {index + 1})");
                    }

                    Application.DoEvents();
                }


                outputLines.Add("];");

                if (errors.Any())
                {
                    MessageBox.Show(string.Join(Environment.NewLine, errors), @"Unable to process some addresses");
                }

                richTextBox1.Lines = outputLines.ToArray();
                richTextBox1.Focus();

            }
            catch (Exception e)
            {
                MessageBox.Show("Something went really wrong: " + Environment.NewLine + e, "Major Error");
            }
            finally
            {
                FinishedProcessing();
            }
        }

        private void FinishedProcessing()
        {
            Enabled = true;
            Cursor.Current = Cursors.Default;
            toolStripStatusLabel1.Text = @"Ready";
        }
        private async void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = @"Excel Files (*.xls, *.xlsx)|*.xls;*.xlsx",
                Multiselect = false
            };

            var result = openFileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                await ParseFile(openFileDialog.FileName);
            }
        }
        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        async void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Any())
            {
                await ParseFile(files.First());
            }
        }
    }
}
