using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Net.Http;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web.Script.Serialization;
using System.IO;

namespace DistanceOptimizer
{
    public partial class Form1 : Form
    {
        private List<Employee> employeeList;
        private List<OfficeLocation> officeList;
        private List<List<RouteResponse>> durationsMatrix = new List<List<RouteResponse>>();
        JavaScriptSerializer serializer = new JavaScriptSerializer();

        //private string googleApiKey = "AIzaSyDjbANNRU5l9-spaj-88NRxxWchSpw6Yxk";
        private string googleApiKey = "AIzaSyAfrthhkvKDTfGW7gNxskXifRuoPQ8C9y0";
        private string urlBase = "https://maps.googleapis.com/maps/api/distancematrix/json?units=metric&origins=";
        private string urlDestinations = "&destinations=";
        private string urlMode = "&mode=";
        private string urlKey = "&key=";
        private string urlDepartureTime = "&departure_time=";
        private string urlArrivalTime = "&arrival_time="; // can only be used when travel mode = public transit!
        // February 13 2016 @ 9:00am (UTC) - just a random monday
        private string urlTimeValue9AM = "1486976400";
        // one hour earlier...
        private string urlTimeValue8AM = "1486972800";
        private string urlTrafficModel = "&traffic_model=";
        private string urlTrafficModelValuePessimistic = "pessimistic";
        private string urlTrafficModelValueBestGuess = "best_guess";

        const string EmployeeDataFileName = "EmployeeList.json";
        const string OfficeDataFileName = "OfficeList.json";
        const string DurationMatrixFileName = "DurationMatrix.json";

        public Form1()
        {
            InitializeComponent();
            ClearInputData();
        }

        private void ClearInputData()
        {
            employeeList = new List<Employee>();
            officeList = new List<OfficeLocation>();
            //durationsMatrix = new List<List<RouteResponse>>();
            //listBox3.Items.Clear();
            dataGridView3.Rows.Clear();
            dataGridView1.Rows.Clear();
        }

        private string GetFullURL(string origin, string destination, TransitMode mode)
        {
            string trafficModel = checkBox1.Checked ? urlTrafficModelValuePessimistic : urlTrafficModelValueBestGuess;

            if (mode == TransitMode.transit)
                return string.Concat(urlBase, origin, urlDestinations, destination, urlArrivalTime, urlTimeValue9AM, urlMode, "transit", urlTrafficModel, trafficModel, urlKey, googleApiKey);
            return string.Concat(urlBase, origin, urlDestinations, destination, urlDepartureTime, urlTimeValue8AM, urlMode, mode.ToString(), urlTrafficModel, trafficModel, 
                urlKey, googleApiKey);
        }

        private RouteResponse GetResponse(JavaScriptSerializer serializer, string origin, string destination, TransitMode mode)
        {
            var httpClient = new HttpClient();
            var response = httpClient.GetAsync(GetFullURL(origin, destination, mode)).Result;

            if (response.IsSuccessStatusCode)
            {
                var responseContent = response.Content;
                return serializer.Deserialize<RouteResponse>(responseContent.ReadAsStringAsync().Result);
            }

            return null;
        }

        private void LoadSavedDataXls(string filename)
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(filename);
            Excel.Worksheet sheet = wb.ActiveSheet;

            dynamic addressValue;
            dynamic nameValue;
            dynamic transitValue;

            for (int i = 1; i <= sheet.UsedRange.Count; i++)
            {
                nameValue = sheet.UsedRange.Cells[i, 1].Value2;
                addressValue = sheet.UsedRange.Cells[i, 2].Value2;
                transitValue = sheet.UsedRange.Cells[i, 4].Value2;
                TransitMode mode = TransitMode.driving;

                if (addressValue == null || addressValue == "")
                    continue;

                Enum.TryParse<TransitMode>(transitValue, out mode);
                employeeList.Add(new Employee() { Name = nameValue, Address = addressValue, TransitMode = mode });
            }

            int j = 1;
            dynamic officeAddressValue = sheet.UsedRange.Cells[j, 5].Value2;
            while (officeAddressValue != null && officeAddressValue != "")
            {
                officeList.Add(new OfficeLocation() { Address = officeAddressValue.ToString() });
                ++j;

                officeAddressValue = sheet.UsedRange.Cells[j, 5].Value2;
            }
        }

        private void DrawEmployeeList()
        {
            dataGridView1.Rows.Clear();

            foreach (Employee employee in employeeList)
            {
                string[] newRow = new string[] { employee.Name ?? "Name", employee.Address, employee.TransitModeString };
                dataGridView1.Rows.Add(newRow);
            }
        }

        private void LoadSavedDataJson(string filename)
        {
            string jsonContent = File.ReadAllText(filename);
            if (string.Equals(Path.GetFileName(filename), OfficeDataFileName))
            {
                officeList = serializer.Deserialize<List<OfficeLocation>>(jsonContent);
            }
            else if (string.Equals(Path.GetFileName(filename), EmployeeDataFileName))
            {
                employeeList = serializer.Deserialize<List<Employee>>(jsonContent);
            }
            else if (string.Equals(Path.GetFileName(filename), DurationMatrixFileName))
            {
                durationsMatrix = serializer.Deserialize<List<List<RouteResponse>>>(jsonContent);
            }
        }

        private void LoadSavedData(string filename)
        {
            if (Path.GetExtension(filename).StartsWith(".xls"))
                LoadSavedDataXls(filename);
            else if (Path.GetExtension(filename).StartsWith(".json"))
                LoadSavedDataJson(filename);
            else
                throw new Exception("Unexpected input file format!");
        }

        private void ConnectTheDots()
        {
            listBox2.Items.Clear();
            for (int i = 0; i < officeList.Count; i++)
            {
                for (int j = 0; j < employeeList.Count; j++)
                {
                    RouteResponse responseObj = durationsMatrix[j][i];

                    if (responseObj.Status == "NOT_FOUND" || responseObj.Status == "ZERO_RESULTS")
                    {
                        if (responseObj.origin_addresses.First() == "")
                            listBox2.Items.Add("Address not found: " + employeeList[j].Address);
                        else if (responseObj.destination_addresses.First() == "")
                            listBox2.Items.Add("Office not found: " + officeList[i].Address);
                        else
                            listBox2.Items.Add("Route not found for: " + employeeList[j].Address + "-->" + officeList[i].Address);

                        continue;
                    }

                    double bestDurationInfo = responseObj.GetBestDurationInfo;

                    if (bestDurationInfo == 0)
                    {
                        listBox2.Items.Add("Problem calculating duration for route: " + employeeList[j].Address + "-->" + officeList[i].Address);
                        continue;
                    }

                    officeList[i].AverageDuration += bestDurationInfo;
                    officeList[i].MaximalDuration = Math.Max(officeList[i].MaximalDuration, bestDurationInfo);
                    officeList[i].MinimalDuration = Math.Min(officeList[i].MinimalDuration, bestDurationInfo);
                    officeList[i].NonZeroDurations++;
                }
            }

            officeList.Sort((item1, item2) => (item1.CalculateAverageDuration.CompareTo(item2.CalculateAverageDuration)));
            for (int i = 0; i < officeList.Count; i++)
            {
                //listBox3.Items.Add(officeList[i].Address + ": " + (officeList[i].CalculateAverageDuration / 60).ToString("F") + " (" + (officeList[i].MinimalDuration / 60).ToString("F") + " - " + (officeList[i].MaximalDuration / 60).ToString("F") + ")");
                string[] newRow = new string[] { officeList[i].Name ?? "Name", officeList[i].Address, (officeList[i].CalculateAverageDuration / 60).ToString("F"), " (" + (officeList[i].MinimalDuration / 60).ToString("F") + " - " + (officeList[i].MaximalDuration / 60).ToString("F") + ")" };
                dataGridView3.Rows.Add(newRow);
            }
        }

        private void GetDurationsFromGoogle()
        {
            durationsMatrix = new List<List<RouteResponse>>();

            for (int i = 0; i < employeeList.Count; i++)
                durationsMatrix.Add(new List<RouteResponse>());

            for (int i = 0; i < officeList.Count; i++)
            {
                for (int j = 0; j < employeeList.Count; j++)
                {
                    RouteResponse responseObj = GetResponse(serializer, employeeList[j].Address, officeList[i].Address, employeeList[j].TransitMode);
                    durationsMatrix[j].Add(responseObj);
                }
            }
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            //ClearInputData();
            LoadSavedData(openFileDialog1.FileName);
            DrawEmployeeList();
        }

        private void LoadDataBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private int GetSelectedRowIndex()
        {
            //var selectedRowsCollection = dataGridView1.SelectedRows.Cast<DataGridViewRow>();
            var selectedCellsCollection = dataGridView1.SelectedCells.Cast<DataGridViewCell>();
            // if (selectedRowsCollection.Any())
            //    return selectedRowsCollection.First().Index;
            // else 
            if (selectedCellsCollection.Any())
                return selectedCellsCollection.First().RowIndex;

            return -1;
        }

        private void DisplayIndividualDetails()
        {
            dataGridView2.Rows.Clear();
            int selectedRowIndex = GetSelectedRowIndex();            
            if (selectedRowIndex < 0)
            {
                MessageBox.Show("Please select an employee to display distances.");
                return;
            }
            List<RouteResponse> responses = durationsMatrix[selectedRowIndex];
            foreach (RouteResponse response in responses)
            {
                string[] newRow = new string[] { response.destination_addresses.First(), (response.GetBestDurationInfo / 60).ToString("F"), (response.Distance / 1000).ToString("F") + " km"};
                dataGridView2.Rows.Add(newRow);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetDurationsFromGoogle();
            ConnectTheDots();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView2.Rows.Clear();
            DisplayIndividualDetails();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            File.WriteAllText(DurationMatrixFileName, serializer.Serialize(durationsMatrix));
            File.WriteAllText(OfficeDataFileName, serializer.Serialize(officeList));
            File.WriteAllText(EmployeeDataFileName, serializer.Serialize(employeeList));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ConnectTheDots();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
