using System;
using System.Reflection.Metadata;
using CSV.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;

namespace CSV.WordDocument
{
    public class wordprocess
    {
        public wordprocess()
        {
        }

        string studentsWordPath = $"{Constants.Locations.DataFolder}//students.docx";
            //string studentsImagePath = $"{Constants.Locations.ImagesFolder}";
            // Create a document by supplying the filepath. 
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(studentsWordPath, WordprocessingDocumentType.Document))
            {
                
                    

                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();


    // Create the document structure and add some text.
    mainPart.Document = new Document();
    Body body = mainPart.Document.AppendChild(new Body());
    Paragraph para = body.AppendChild(new Paragraph());
    Run run = para.AppendChild(new Run());
                //mainPart = wordDocument.MainDocumentPart;

                //ImagePart imagePart = mainPart.AddImagePart(studentsImagePath);
                //using (StreamWriter fs = new StreamWriter(studentsWordPath))
                //{
                foreach (var student in students)
                    {
                    run.AppendChild(new Text("My name is :  "));
                    //run.AppendChild(new Text(student.ToString()));
                    run.AppendChild(new Text(student.FirstName.ToString()));
                    run.AppendChild(new Text("  ,  "));

                    run.AppendChild(new Text("My Student id is: "));
                   
                 
                    run.AppendChild(new Text(student.StudentId.ToString()));
                    run.AppendChild(new Text("  ,  "));

                    para = new Paragraph(new Run(new Break() { Type = BreakValues.Page }));
                    mainPart = wordDocument.MainDocumentPart;

                    mainPart.Document.Body.InsertAfter(para, mainPart.Document.Body.LastChild);
                    mainPart.Document.Save();






                    //using (FileStream stream = new FileStream(studentsImagePath, FileMode.Open))
                    //{
                    //    imagePart.FeedData(stream);
                    //}

                    //AddImageToBody(wordDocument, mainPart.GetIdOfPart(imagePart));
                }

                //}
             

            }
    }
}
