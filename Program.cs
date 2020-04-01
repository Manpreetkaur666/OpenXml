using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CSV.Models;
using CSV.Models.Utilities;
using System.Xml.Serialization;
using System.Net;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;

namespace CSV
{
    class Program
    {

        static string remoteUploadFileDestination = "/StudentId FirstName LastName/info2.csv";
        private static StringValue relationshipId;
        private static uint rowindex;
        private static object sheetData;

        static void Main(string[] args)
        {

            Student myrecord = new Student { StudentId = "200429017", FirstName = "Manpreet", LastName = "Kaur" };
            //Console.WriteLine(UploadFile(localUploadFilePath, url + remoteUploadFileDestination));




            List<string> directories = FTP.GetDirectory(Constants.FTP.BaseUrl);
            List<Student> students = new List<Student>();

            foreach (var directory in directories)
            {
                Student student = new Student() { AbsoluteUrl = Constants.FTP.BaseUrl };
                student.FromDirectory(directory);

                //Console.WriteLine(student);
                string infoFilePath = student.FullPathUrl + "/" + Constants.Locations.InfoFile;

                bool fileExists = FTP.FileExists(infoFilePath);
                if (fileExists == true)
                {

                    string csvPath = $@"/Users/manpreetkaur/Desktop/data/{directory}.csv";


                    // FTP.DownloadFile(infoFilePath, csvPath);
                    byte[] bytes = FTP.DownloadFileBytes(infoFilePath);
                    string csvData = Encoding.Default.GetString(bytes);

                    string[] csvlines = csvData.Split("\r\n", StringSplitOptions.RemoveEmptyEntries);

                    if (csvlines.Length != 2)
                    {
                        Console.WriteLine("Error in CSV format");
                    }
                    else
                    {
                        student.FromCSV(csvlines[1]);
                        //Console.WriteLine("  \t Age of Student is: {0} ", student.age);
                    }

                    Console.WriteLine("Found info file:");
                }
                else
                {
                    Console.WriteLine("Could not find info file:");
                }

                Console.WriteLine("\t" + infoFilePath);

                string imageFilePath = student.FullPathUrl + "/" + Constants.Locations.ImageFile;

                bool imageFileExists = FTP.FileExists(imageFilePath);

                if (imageFileExists == true)
                {

                    Console.WriteLine("Found image file:");
                }
                else
                {
                    Console.WriteLine("Could not find image file:");
                }

                Console.WriteLine("\t" + imageFilePath);

                students.Add(student);
                Console.WriteLine(directory);

                Console.WriteLine(" \t Count of student is: {0}", students.Count);
                Console.WriteLine("  \t Age of Student is: {0} ", student.age);

            }

            Student me = students.SingleOrDefault(x => x.StudentId == myrecord.StudentId);
            Student meUsingFind = students.Find(x => x.StudentId == myrecord.StudentId);

            var avgage = students.Average(x => x.age);
            var minage = students.Min(x => x.age);
            var maxage = students.Max(x => x.age);


            Console.WriteLine("  \n\t Name Searched With Query: {0} ", meUsingFind);
            Console.WriteLine("  \t Average of Student age is: {0} ", avgage);
            Console.WriteLine("  \t Minimum of Student age is: {0} ", minage);
            Console.WriteLine("  \t Maximum of Student age is: {0} ", maxage);


            string studentsCSVPath = $"{Constants.Locations.DataFolder}//students.csv";
            //Establish a file stream to collect data from the response
            using (StreamWriter fs = new StreamWriter(studentsCSVPath))
            {
                foreach (var student in students)
                {
                    fs.WriteLine(student.ToCSV());
                }
            }

            string studentsjsonPath = $"{Constants.Locations.DataFolder}//students.json";
            //Establish a file stream to collect data from the response
            using (StreamWriter fs = new StreamWriter(studentsjsonPath))
            {
                foreach (var student in students)
                {
                    string Student = Newtonsoft.Json.JsonConvert.SerializeObject(student);
                    fs.WriteLine(Student.ToString());
                    //Console.WriteLine(jStudent);
                }
            }

            //string studentsxmlPath = $"{Constants.Locations.DataFolder}//students.xml";
            //XmlSerializer serializer = new XmlSerializer(typeof(Student));
            //using (StreamWriter fs = new StreamWriter(studentsxmlPath))
            //{
            //    serializer.Serialize(fs, students);
            //}


            string studentsxmlPath = $"{Constants.Locations.DataFolder}//students.xml";
            //Establish a file stream to collect data from the response
            using (StreamWriter fs = new StreamWriter(studentsxmlPath))
            {
                //foreach (var student in students)
                //{
                //    // XmlSerializer xs = new XmlSerializer(student.GetType());
                //    XmlSerializer xs = new XmlSerializer(typeof(Student));

                //    //xs.Serialize(fs, student);
                //    fs.WriteLine(student);

                XmlSerializer x = new XmlSerializer(students.GetType());
                x.Serialize(fs, students);
                Console.WriteLine();

                //XmlSerializer x = new XmlSerializer(myrecord.GetType());
                //x.Serialize(Console.Out, myrecord);
                //Console.ReadKey();


                //Test myTest = new Test() { value1 = "Value 1", value2 = "Value 2" };
                //XmlSerializer x = new XmlSerializer(myTest.GetType());
                //x.Serialize(Console.Out, myTest);
                //Console.ReadKey();


            }


            //create word document


            string studentsWordPath = $"{Constants.Locations.DataFolder}//students.docx";
            //string studentsImagePath = $"{Constants.Locations.ImagesFolder}";
            // Create a document by supplying the filepath

            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(studentsWordPath, WordprocessingDocumentType.Document))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                //ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
                //Paragraph para = body.AppendChild(new Paragraph(new Run
                //     (new Break() { Type = BreakValues.Page })));


                //Paragraph newPara = new Paragraph(new Run
                //     (new Break() { Type = BreakValues.Page },
                //     new Text("text on the new page")));

                Run run = para.AppendChild(new Run());
                //mainPart = wordDocument.MainDocumentPart;

                //ImagePart imagePart = mainPart.AddImagePart(studentsImagePath);
                //using (StreamWriter fs = new StreamWriter(studentsWordPath))
                //


                foreach (var student in students)
                {
                    run.AppendChild(new Text("My name is :  "));
                    //run.AppendChild(new Text(student.ToString()));
                    run.AppendChild(new Text(student.FirstName.ToString()));
                    run.AppendChild(new Text("  ,  "));

                    run.AppendChild(new Text("My Student id is: "));


                    run.AppendChild(new Text(student.StudentId.ToString()));
                    run.AppendChild(new Text("  ,  "));

                    //para = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                    //mainPart = wordDocument.MainDocumentPart;

                    //mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
                    //mainPart.Document.Save();

                    //run.AppendChild(new Paragraph(new Run
                    // (new Break() { Type = BreakValues.Page })));

                    run.AppendChild(new Break() { Type = BreakValues.Page });





                }
            }


            string studentsExcelPath = $"{Constants.Locations.DataFolder}//students.xlsx";
            //using (StreamWriter fs = new StreamWriter(studentsxmlPath))
            //{
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
                Create(studentsExcelPath, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document.
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook.
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                    AppendChild<Sheets>(new Sheets());

                // Append a new worksheet and associate it with the workbook.
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "mySheet"
                };


                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                var excelRows = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().ToList();
                //var excelcolumns = sheetData.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>().ToList();
                int rowindex = 1;
                //int columnindex = 1;
                foreach (var student in students)
                {

                    Row row = new Row();
                   // DocumentFormat.OpenXml.Spreadsheet.Columns cs = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                    row.RowIndex = (UInt32)rowindex;
                Cell cell = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.FirstName.ToString())
                        
                    
                        
                    };
                Cell cell1 = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.LastName.ToString())



                };
                Cell cell2 = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.StudentId.ToString())



                };
                Cell cell3 = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.MyRecord.ToString())



                };
                Cell cell4 = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.age.ToString())



                };
                Cell cell5 = new Cell()
                {

                    DataType = CellValues.String,
                    CellValue = new CellValue(student.DateOfBirthDT.ToString())



                };

                //Cell cell6 = new Cell()
                //{

                //    DataType = CellValues.String,
                //    CellValue = new CellValue(student.Guid.ToString())



                //};



                row.Append(cell);
                row.Append(cell1);
                row.Append(cell2);
                row.Append(cell3);
                row.Append(cell4);
                row.Append(cell5);
                sheetData.Append(row);



                    //how to write the data in cell
                    rowindex++;
                }

                sheets.Append(sheet);

                workbookpart.Workbook.Save();

                // Close the document.
                spreadsheetDocument.Close();

                

            








            
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            








            //string document = $"{Constants.Locations.DataFolder}//students.docx";
            //using (WordprocessingDocument wordprocessingDocument =
            //WordprocessingDocument.Open(document, true))
            //{
            //    MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

            //    ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            //    using (FileStream stream = new FileStream(studentsWordPath, FileMode.Open))
            //    {
            //        imagePart.FeedData(stream);
            //    }

            //    //NewMethod(wordprocessingDocument, mainPart, imagePart);

            //    AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));



            //    foreach (var student in students)
            //    {

            //        // Define the reference of the image.
            //        var element =
            //             new Drawing(
            //                 new DW.Inline(
            //                     new DW.Extent() { Cx = 990000L, Cy = 792000L },
            //                     new DW.EffectExtent()
            //                     {
            //                         LeftEdge = 0L,
            //                         TopEdge = 0L,
            //                         RightEdge = 0L,
            //                         BottomEdge = 0L
            //                     },
            //                     new DW.DocProperties()
            //                     {
            //                         Id = (UInt32Value)1U,
            //                         Name = "Picture 1"
            //                     },
            //                     new DW.NonVisualGraphicFrameDrawingProperties(
            //                         new A.GraphicFrameLocks() { NoChangeAspect = true }),
            //                     new A.Graphic(
            //                         new A.GraphicData(
            //                             new PIC.Picture(
            //                                 new PIC.NonVisualPictureProperties(
            //                                     new PIC.NonVisualDrawingProperties()
            //                                     {
            //                                         Id = (UInt32Value)0U,
            //                                         Name = "New Bitmap Image.jpg"
            //                                     },
            //                                     new PIC.NonVisualPictureDrawingProperties()),
            //                                 new PIC.BlipFill(
            //                                     new A.Blip(
            //                                         new A.BlipExtensionList(
            //                                             new A.BlipExtension()
            //                                             {
            //                                                 Uri =
            //                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
            //                                             })
            //                                     )
            //                                     {
            //                                         Embed = relationshipId,
            //                                         CompressionState =
            //                                         A.BlipCompressionValues.Print
            //                                     },
            //                                     new A.Stretch(
            //                                         new A.FillRectangle())),
            //                                 new PIC.ShapeProperties(
            //                                     new A.Transform2D(
            //                                         new A.Offset() { X = 0L, Y = 0L },
            //                                         new A.Extents() { Cx = 990000L, Cy = 792000L }),
            //                                     new A.PresetGeometry(
            //                                         new A.AdjustValueList()
            //                                     )
            //                                     { Preset = A.ShapeTypeValues.Rectangle }))
            //                         )
            //                         { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            //                 )
            //                 {
            //                     DistanceFromTop = (UInt32Value)0U,
            //                     DistanceFromBottom = (UInt32Value)0U,
            //                     DistanceFromLeft = (UInt32Value)0U,
            //                     DistanceFromRight = (UInt32Value)0U,
            //                     EditId = "50D07946"
            //                 });

            //        // Append the reference to body, the element should be in a Run.
            //        wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));





            //}

            //}












            //}

            return;

            //FileInfo newfileinfo = new FileInfo(newfilePath);
            //Image studentImage = Imaging.Base64ToImage(student.ImageData);
            //studentImage.Save(newfileinfo.FullName, ImageFormat.Jpeg);
        }

        //private static void NewMethod(WordprocessingDocument wordprocessingDocument, MainDocumentPart mainPart, ImagePart imagePart)
        //{
        //    AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
        //}
    }
}


    //    public static string UploadFile(string sourceFilePath, string destinationFileUrl, string username = Constants.FTP.UserName, string password = Constants.FTP.Password)
    //    {
    //        string output;

    //        // Get the object used to communicate with the server.
    //        FtpWebRequest request = (FtpWebRequest)WebRequest.Create(destinationFileUrl);

    //        request.Method = WebRequestMethods.Ftp.UploadFile;

    //        // This example assumes the FTP site uses anonymous logon.
    //        request.Credentials = new NetworkCredential(username, password);

    //        // Copy the contents of the file to the request stream.
    //        byte[] fileContents = GetStreamBytes(sourceFilePath);

    //        //Get the length or size of the file
    //        request.ContentLength = fileContents.Length;

    //        //Write the file to the stream on the server
    //        using (Stream requestStream = request.GetRequestStream())
    //        {
    //            requestStream.Write(fileContents, 0, fileContents.Length);
    //        }

    //        //Send the request
    //        using (FtpWebResponse response = (FtpWebResponse)request.GetResponse())
    //        {
    //            output = $"Upload File Complete, status {response.StatusDescription}";
    //        }
    //        Thread.Sleep(Constants.FTP.OperationPauseTime);

    //        return (output);
    //    }

    //    private static byte[] GetStreamBytes(string sourceFilePath)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}


    



    



