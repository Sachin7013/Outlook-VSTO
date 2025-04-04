using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Newtonsoft.Json;
using MSOutlook = Microsoft.Office.Interop.Outlook;


namespace AlgoOrange.Outllok.AddInn
{
    public partial class ChatControl : UserControl
    {
        public ChatControl()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
        // Add the missing event handler method for button5_Click
        private void Button5_Click(object sender, EventArgs e)
        {
            // Add your logic here
        }
        // Add the missing event handler method for button5_Click
        private void button5_Click(object sender, EventArgs e)
        {
            // Add your logic here
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }


        private void richTextBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {

        }
        private void button3_Click_1(object sender, EventArgs e)
        {
            try
            {
                MSOutlook.Application outlookApp = new MSOutlook.Application();
                MSOutlook.Explorer explorer = outlookApp.ActiveExplorer();
                MSOutlook.Selection selection = explorer.Selection;

                if (selection.Count > 0)
                {
                    MSOutlook.MailItem mailItem = selection[1] as MSOutlook.MailItem;
                    if (mailItem != null)
                    {
                        string emailContent = mailItem.Body;
                        richTextBox1.Text = emailContent;
                    }
                    else
                    {
                        richTextBox1.Text = "No email selected.";
                    }
                }
                else
                {
                    richTextBox1.Text = "No email selected.";
                }
            }
            catch (Exception ex)
            {
                richTextBox1.Text = $"Exception: {ex.Message}";
            }
        }

        private void SendXLSDataToAlgoAPI(string jsonData, string userQuery = null, string chatHistory = null)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    client.BaseAddress = new Uri("http://127.0.0.1:8000/");

                    var wrappedData = new { data = JsonConvert.SerializeObject(jsonData), userQuery = userQuery, chatHistory = chatHistory };
                    string wrappedJsonData = JsonConvert.SerializeObject(wrappedData);

                    StringContent content = new StringContent(wrappedJsonData, Encoding.UTF8, "application/json");

                    Debug.WriteLine("Sending JSON data: " + wrappedJsonData);

                    HttpResponseMessage response = client.PostAsync("/excel/query", content).Result;

                    if (response.IsSuccessStatusCode)
                    {
                        string result = response.Content.ReadAsStringAsync().Result;

                        var jsonResponse = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(result);

                        this.Invoke((MethodInvoker)delegate { richTextBox1.Text = jsonResponse.message.ToString(); });
                    }
                    else
                    {
                        this.Invoke((MethodInvoker)delegate { richTextBox1.Text = $"Error: {response.StatusCode}"; });
                    }
                }
                catch (Exception ex)
                {
                    this.Invoke((MethodInvoker)delegate { richTextBox1.Text = $"Exception: {ex.Message}"; });
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }
    }
} 
