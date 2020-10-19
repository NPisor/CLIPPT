using System;
using System.IO;
using System.Linq;
using System.Net;
using HtmlAgilityPack;
using Syncfusion.Presentation;

namespace ConsoleApp1
{
    public class Program
    {
        public static void Main()
        {
            Console.WriteLine("How many items do you need?: ");
            string input = Console.ReadLine();
            int number;
            Int32.TryParse(input, out number);
            Console.WriteLine("What would you like to search for?: ");
            var search = Console.ReadLine();
            var document = new HtmlWeb().Load("https://www.shutterstock.com/search/" + search);
            WebClient client = new WebClient();
            var urls = document.DocumentNode.Descendants("img")
                                            .Select(e => e.GetAttributeValue("src", null))
                                            .Where(s => !String.IsNullOrEmpty(s));
            int x = 0;
            IPresentation pptxDoc = Presentation.Create();
            foreach (string item in urls)
            {
                if (x == number){
                    break;
                }
                client.DownloadFile(item, search + x + ".jpg");
                Console.WriteLine(item);
                ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                Console.WriteLine("Enter title text: ");
                var title = Console.ReadLine();
                Stream pictureStream = File.Open("./" + search + x + ".jpg", FileMode.Open);
                slide.Pictures.AddPicture(pictureStream, slide.SlideSize.Width/2, slide.SlideSize.Height/2, 250, 250);
                IShape titleTextBox = slide.AddTextBox(slide.SlideSize.Width/2, 10, 500, 500);
                IParagraph paragraph = titleTextBox.TextBody.AddParagraph();
                ITextPart textPart = paragraph.AddTextPart();
                textPart.Text = title;
                pictureStream.Dispose();
                x++;
            }            


            pptxDoc.Save("Sample.pptx");
            pptxDoc.Close();





            Console.ReadLine();
        }
    }    
}
