// See https://aka.ms/new-console-template for more information
//using JSONExercise;
using Assignment3_JSON_XML_CSV;
using System.ComponentModel.Design.Serialization;
using System.Text.Json;

Console.WriteLine("Assignment 3 Solution by Habeeb");
//set path to desktop directory
string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
string rootDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent + "";
string inputData_JSON = rootDir + "\\books.json";
string inputData_XML = rootDir + "\\books.xml";
string inputData_CSV = rootDir + "\\books.csv";




DataFromURL dfu = new DataFromURL();
dfu.URL_Reader();
dfu.Display_user_data();
Console.WriteLine();
dfu.createNewExcelFile(desktopDir);
Console.WriteLine();
dfu.setupBookstore_XML(inputData_XML);
Console.WriteLine();
dfu.setupBookstore_JSON(inputData_JSON);
Console.WriteLine();
dfu.setupBookstore_CSV(inputData_CSV);
Console.WriteLine("A Utility Service for Converting JSON to XML\nConversion in process ...\n");
string newXML_string = dfu.Json_2_Xml();

Console.WriteLine(newXML_string);