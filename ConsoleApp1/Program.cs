/* import necessary libraries */
using System;
using System.IO;
using System.Linq;
using System.Net;

/* Using HTMLAgilityPack to handle HTML tags and parse */
using HtmlAgilityPack;

/* Using SyncFusion for creating .pptx file programatically */
using Syncfusion.Presentation;

namespace ConsoleApp1
{
    public class Program
    {
        public static void Main()
        {
            /* Take in values for numbers of slides and images to parse from HTML */
            Console.WriteLine("How many items do you need?: ");
            string input = Console.ReadLine();
            int number;
            Int32.TryParse(input, out number);
            
            /* Take in user response for image content */
            Console.WriteLine("What would you like to search for?: ");
            var search = Console.ReadLine();
            
            /* Load Shutterstock site and create new WebClient and scrape searching for <img src> tag */
            var document = new HtmlWeb().Load("https://www.shutterstock.com/search/" + search);
            WebClient client = new WebClient();
            var urls = document.DocumentNode.Descendants("img")
                                            .Select(e => e.GetAttributeValue("src", null))
                                            .Where(s => !String.IsNullOrEmpty(s));
            int x = 0;
            
            /* Initialize new PPT Presentation using SyncFusion */
            IPresentation pptxDoc = Presentation.Create();
            
            /* Iterate over urls of images, download to local directory number of specified images */
            foreach (string item in urls)
            {
                if (x == number){
                    break;
                }
                client.DownloadFile(item, search + x + ".jpg");
                Console.WriteLine(item);
                
                /* Create new Blank slide, take in User title text */
                ISlide slide = pptxDoc.Slides.Add(SlideLayoutType.Blank);
                Console.WriteLine("Enter title text: ");
                var title = Console.ReadLine();
                
                /* Open local image files scraped from web and add to slide with user defined title of each slide */
                Stream pictureStream = File.Open("./" + search + x + ".jpg", FileMode.Open);
                slide.Pictures.AddPicture(pictureStream, slide.SlideSize.Width/2, slide.SlideSize.Height/2, 250, 250);
                IShape titleTextBox = slide.AddTextBox(slide.SlideSize.Width/2, 10, 500, 500);
                IParagraph paragraph = titleTextBox.TextBody.AddParagraph();
                ITextPart textPart = paragraph.AddTextPart();
                textPart.Text = title;
                pictureStream.Dispose();
                x++;
            }      
            
            /* Save PPT project */
            pptxDoc.Save("Sample.pptx");
            pptxDoc.Close();

            Console.ReadLine();
        }
    }    
}
