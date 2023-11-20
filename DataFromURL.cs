using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Xml;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Xml.Linq;

namespace Assignment3_JSON_XML_CSV
{
    internal class DataFromURL
    {
        string json_URL_Data = "";
        string json_bookstore_data = "";

        public void URL_Reader()
        {
            Console.WriteLine("=============== Data from URL ==================\n");
            string URL = "https://jsonplaceholder.typicode.com/users";

            //send request to server onn the URL
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);

            //get response from server and handle data
            HttpWebResponse response = request.GetResponse() as HttpWebResponse;

            using (Stream responseStream = response.GetResponseStream())
            {
                StreamReader readData = new StreamReader(responseStream, Encoding.UTF8);
                json_URL_Data = readData.ReadToEnd();
                if (json_URL_Data != null)
                {
                    Console.Write("data retrieved from {0} successful\n", URL);
                }
                else
                {
                    Console.WriteLine("No Data ---------------");
                }
            }
        } // end read method 

        // Assignment 3 Question 1 Solution
        // Fetch the following data from the server to generate following output
        public void Display_user_data()
        {
            var jsonData = JArray.Parse(json_URL_Data);

            Console.WriteLine("\n*********************************** User Directory ************************************************************************");
            Console.WriteLine("===========================================================================================================================");
            Console.WriteLine("Name                      | Email                     |Phone                   |Address");
            Console.WriteLine("===========================================================================================================================");

            foreach (var jData in jsonData)
            {
                Console.WriteLine("{0,-25} | {1,-25} | {2,-22} | {3, -30}", jData["name"], jData["email"], jData["phone"],
                    jData["address"]["suite"] + "-" + jData["address"]["street"] + "," + jData["address"]["city"]);

            }
        } // end of Display_user_data method

        // Assignment 3 Question 2 Solution
        // Create a file Users.XLSX to store result for slide - 2
        public void createNewExcelFile(string filePath)
        {
            string myFilename = "Users.xlsx";
            string myFilePath = $"{filePath}\\{myFilename}"; // desktop directory initialized in filepath

            //check if the file already exists
            if (File.Exists(myFilePath))
            {
                File.Delete(myFilePath); // delete before creating a new one
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPkg = new ExcelPackage(new FileInfo(myFilePath)))
            {
                ExcelWorksheet myWorkSheet = excelPkg.Workbook.Worksheets.Add("Users");
                myWorkSheet.Cells["A1"].Value = "ID";
                myWorkSheet.Cells["B1"].Value = "Name";
                myWorkSheet.Cells["C1"].Value = "Email";
                myWorkSheet.Cells["D1"].Value = "Phone";
                myWorkSheet.Cells["E1"].Value = "Address";


                var json2ExcelData = JArray.Parse(json_URL_Data);

                for (int i = 0; i < json2ExcelData.Count; i++)
                {
                    var jedItem = json2ExcelData[i];
                    int curRow = 0; // current row in excel
                    curRow = i + 2; // set row 2 in excel for insertion of data while the 1st row holds the column names

                    myWorkSheet.Cells[curRow, 1].Value = jedItem["id"].ToString();
                    myWorkSheet.Cells[curRow, 2].Value = jedItem["name"].ToString();
                    myWorkSheet.Cells[curRow, 3].Value = jedItem["email"].ToString();
                    myWorkSheet.Cells[curRow, 4].Value = jedItem["phone"].ToString();

                    //Concatenate the address in the proper & recommended format
                    var Address = jedItem["address"]["suite"] + "-" +
                        jedItem["address"]["street"] + "," + jedItem["address"]["city"];
                    myWorkSheet.Cells[curRow, 5].Value = Address.ToString();
                }

                //Autofit Columns

                myWorkSheet.Cells[myWorkSheet.Dimension.Address].AutoFitColumns();
                excelPkg.Save();
                Console.WriteLine("\nFile {0} Created on Desktop", myFilename);
            }
        } // end createNewExcelFile method

        // Assignment 3 Question 3a Solution
        // Work with following files to generate the following output
        public void setupBookstore_JSON(string payload)
        {
            using (StreamReader sr = new StreamReader(payload))
            {
                json_bookstore_data = sr.ReadToEnd();

                if (json_bookstore_data != null)
                {
                    Console.WriteLine();
                }
                else
                { Console.WriteLine("No Data ---------------"); }
            }

            //parsing the json file into a jSON object setupBookstore_XML
            var new_jsonData = JObject.Parse(json_bookstore_data);
            Console.WriteLine("=========================== Read JSON File ============================");
            Console.WriteLine("Category     |Title                 |Author               |Price");
            Console.WriteLine("=======================================================================");
            for (int k = 0; k < 4; k++)
            {
                Console.WriteLine("{0,-12} | {1,-20} | {2,-19} | {3, -20}",
                    new_jsonData["bookstore"]["book"][k]["_category"].ToString(),
                    new_jsonData["bookstore"]["book"][k]["title"]["__text"].ToString(),
                    new_jsonData["bookstore"]["book"][k]["author"].ToString(),
                    new_jsonData["bookstore"]["book"][k]["price"].ToString());

            }

        } //end setupBookstore_JSON method

        // Assignment 3 Question 3b Solution
        // Work with following files to generate the following output
        public void setupBookstore_XML(string payload)
        {
            XmlDocument myBookstore = new XmlDocument();
            myBookstore.Load(payload);

            // Select all book nodes
            XmlNodeList bookNodes = myBookstore.SelectNodes("/books/book");

            // Extract and display information for each book
            Console.WriteLine("=========================== Read XML File ============================");
            Console.WriteLine("Category     |Title                 |Author               |Price");
            Console.WriteLine("=======================================================================");
            foreach (XmlNode bookNode in bookNodes)
            {
                XmlAttribute categoryAttribute = bookNode.Attributes["category"];
                string category = "";
                if (categoryAttribute != null)
                {
                    category = categoryAttribute.Value;
                }
                else
                    continue;

                //get info from other tags
                string title = bookNode.SelectSingleNode("title").InnerText;
                string author = bookNode.SelectSingleNode("author").InnerText;
                string price = bookNode.SelectSingleNode("price").InnerText;

                Console.WriteLine("{0,-12} | {1,-20} | {2,-19} | {3, -20}", category, title, author, price);

            } //end setupBookstore_XML method


        }

        // Assignment 3 Question 3c Solution
        // Work with following files to generate the following output and
        // create books.xlsx file
        public void setupBookstore_CSV(string payload) 
        {
            List<string> bookRecords = new List<string>();
            SortedList<int, string> Category = new SortedList<int, string>();
            SortedList<int, string> Title = new SortedList<int, string>();
            SortedList<int, string> Author = new SortedList<int, string>();
            SortedList<int, double> Price = new SortedList<int, double>();
            SortedList<int, string> Year = new SortedList<int, string>();

            // preambles needed to create a nex excel file
            string myFilename = "books.xlsx";
            string _myFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), myFilename);

            //check if the file already exists
            if (File.Exists(_myFilePath))
            {
                File.Delete(_myFilePath); // delete before creating a new one
            }
            //set package permission
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            int pryKey = 0; //primary key for every book record

            //extract data from CSV file
            //List<string> values = new List<string>();
            using (StreamReader sr = new StreamReader(payload))
            {
                sr.ReadLine(); //ignore header

                while (sr.Peek() != -1)
                {

                    string line = sr.ReadLine();
                    bookRecords = line.Split(',').ToList();

                    Category.Add(pryKey, bookRecords[0]);
                    Title.Add(pryKey, bookRecords[1]);
                    Author.Add(pryKey, bookRecords[2]);
                    Year.Add(pryKey, bookRecords[3]);
                    Price.Add(pryKey, double.Parse(bookRecords[4]));
                    pryKey += 1;

                } // end while
            }

            //generate table format with CSV input file
            Console.WriteLine("=========================== Read CSV File ============================");
            Console.WriteLine("Category     |Title                 |Author               |Price");
            Console.WriteLine("=======================================================================");
            for (int j = 0; j < Author.Count; j++)
            {
                Console.WriteLine("{0,-12} | {1,-20} | {2,-19} | {3, -20}", Category[j], Title[j], Author[j], Price[j]);
            }


            //create the excels spreadsheet , "book.xlsx" from the csv content
            Console.WriteLine("\nCreating {0} on desktop ...", myFilename);
            using (var _excelPkg = new ExcelPackage(new FileInfo(_myFilePath)))
            {
                ExcelWorksheet _WkSt = _excelPkg.Workbook.Worksheets.Add("Books");
                _WkSt.Cells["A1"].Value = "Category";
                _WkSt.Cells["B1"].Value = "Title";
                _WkSt.Cells["C1"].Value = "Author";
                _WkSt.Cells["D1"].Value = "Year";
                _WkSt.Cells["E1"].Value = "Price";

                for (int k = 0; k < Author.Count; k++)
                {
                    
                    int _curRow = 0; // current row in excel
                    _curRow = k + 2; // set row 2 in excel for insertion of data while the 1st row holds the column names

                    _WkSt.Cells[_curRow, 1].Value = Category[k];
                    _WkSt.Cells[_curRow, 2].Value = Title[k];
                    _WkSt.Cells[_curRow, 3].Value = Author[k];
                    _WkSt.Cells[_curRow, 4].Value = Year[k];
                    _WkSt.Cells[_curRow, 5].Value = Price[k];
                }

                //Autofit Columns

                _WkSt.Cells[_WkSt.Dimension.Address].AutoFitColumns();
                _excelPkg.Save();
                Console.WriteLine("\nFile {0} Created on Desktop", myFilename);

            }
        } //end setupBookstore_CSV method

        // Adding Complexity 
        // Converting JSON to XML

        public string Json_2_Xml()
        {
            // Convert JSON to JObject
            //jsonText points to the location of the convertd json file
            string jStrings = @"{""widget"": {
                                    ""debug"": ""on"",
                                    ""window"": {
                                        ""title"": ""Sample Konfabulator Widget"",
                                        ""name"": ""main_window"",
                                        ""width"": 500,
                                        ""height"": 500
                                    },
                                    ""image"": { 
                                        ""src"": ""Images/Sun.png"",
                                        ""name"": ""sun1"",
                                        ""hOffset"": 250,
                                        ""vOffset"": 250,
                                        ""alignment"": ""center""
                                    },
                                    ""text"": {
                                        ""data"": ""Click Here"",
                                        ""size"": 36,
                                        ""style"": ""bold"",
                                        ""name"": ""text1"",
                                        ""hOffset"": 250,
                                        ""vOffset"": 100,
                                        ""alignment"": ""center"",
                                        ""onMouseUp"": ""sun1.opacity = (sun1.opacity / 100) * 90;""
                                    }
                                }}  ";
            JObject jsonObject = JObject.Parse(jStrings);

            // Create an XDocument and add the converted JSON
            XDocument xDoc = new XDocument(new XElement("root"));
            AddJsonToXml(jsonObject, xDoc.Root);

            // Return the XML as a string
            return xDoc.ToString();
        }

        public void AddJsonToXml(JObject jsonObj, XElement xmlSource)
        {
            foreach (var prpty in jsonObj.Properties())
            {
                if (prpty.Value is JObject nestedObject)
                {
                    // build nested objects with recursion
                    XElement nestedXml = new XElement(prpty.Name);
                    xmlSource.Add(nestedXml);
                    AddJsonToXml(nestedObject, nestedXml);
                }
                else if (prpty.Value is JArray array)
                {
                    // Handle arrays
                    foreach (var item in array)
                    {
                        XElement arrayElement = new XElement(prpty.Name);
                        xmlSource.Add(arrayElement);

                        if (item is JObject arrayObject)
                        {
                            // Recursively add nested objects in the array
                            AddJsonToXml(arrayObject, arrayElement);
                        }
                        else
                        {
                            // Add array item as a simple element
                            arrayElement.Add(new XElement("item", item));
                        }
                    }
                }
                else
                {
                    // Add simple elements
                    xmlSource.Add(new XElement(prpty.Name, prpty.Value));
                }
            }
        }


    }
}
