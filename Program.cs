using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace FirstApp
{
    public class Program
    {
        static void Main(string[] args)
        {
            try
            {
                try
                {
                    string command = args[0];

                    switch (command)
                    {
                        case "readDoc":
                            try
                            {
                                string level_1_option = args[1];

                                switch (level_1_option)
                                {
                                    case "-i":
                                        try
                                        {
                                            string filepath = args[2];
                                            string text;

                                            if (File.Exists(filepath))
                                            {
                                                text = File.ReadAllText(filepath);
                                                string prefixFile = ".";

                                                if (args[3].Equals("-o") && args[4] is not null)
                                                {

                                                    string newDirectory = args[4];
                                                    if (!Directory.Exists(newDirectory))
                                                    {
                                                        Directory.CreateDirectory(newDirectory);
                                                    }
                                                    prefixFile = newDirectory;
                                                }
                                                string newFile = Path.Combine(prefixFile, DateTime.Now.ToString("dd-MM-yy_HHmmss") + "_data.txt");

                                                using (FileStream fs = File.Create(newFile)){}

                                                if(Path.GetExtension(filepath).Equals(".docx", StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, false)){
                                                        
                                                        Body body =  wordDoc.MainDocumentPart!.Document.Body!;
                                                        text = body!.InnerText;
                                                        }

                                                        Console.WriteLine(text);
                                                    
                                                }
                                                //StringWriter writer = new();

                                                File.AppendAllText(newFile, text);
                                            }
                                            else
                                            {
                                                Console.WriteLine("Doesn't exist the file");
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine("Give a " + level_1_option + " option value");
                                            if (ex is not null)
                                            {
                                                Console.WriteLine(ex);
                                            }

                                        }

                                        break;

                                    default:
                                        Console.WriteLine(command + " doesn't have a option with the name " + level_1_option);
                                        break;

                                }
                            }
                            catch
                            {
                                Console.WriteLine("Give a valid option");
                            }
                            break;

                        default:
                            Console.WriteLine("The app doesn't have a method with the name " + command);
                            break;
                    }
                }
                catch
                {
                    Console.WriteLine("Give a valid command");
                }

            }
            catch (Exception e)
            {

                Console.WriteLine(e);
            }

        }
    }
}