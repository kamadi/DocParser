using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DocParser
{
    class WordDocumentHelper
    {
        private string filePath;
        private string directoryPath;
       
        private Microsoft.Office.Interop.Word.Application wordApplication;
        public WordDocumentHelper(string filePath) {
            this.filePath = filePath;
            this.directoryPath = Path.GetDirectoryName(filePath) + "\\"+DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss"); 
            wordApplication = new Microsoft.Office.Interop.Word.Application();
        }

        public void parse() {
            var document = wordApplication.Documents.Open(filePath);
            Console.WriteLine("reading started");
            Console.WriteLine(directoryPath);
            
            List<string> names = new List<string>();
            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;

            String name;

            for (var i = 1; i <= wordApplication.ActiveDocument.InlineShapes.Count; i++)
            {
                var inlineShapeId = i;
                name = String.Format("img_{0}.png", unixTimestamp + "_" + i);
                names.Add(name);
            }

            replaceImages(names);

            wordApplication.Quit();
        
        }

        public void replaceImages(List<string> names)
        {
            System.IO.Directory.CreateDirectory(directoryPath);
            System.IO.Directory.CreateDirectory(directoryPath+"\\"+"картинки");
            object miss = System.Reflection.Missing.Value;
            object path = filePath;
            object readOnly = false;
            Microsoft.Office.Interop.Word.Document docs = wordApplication.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);

            List<string> lista = new List<string>();
            List<Range> ranges = new List<Range>();
            int j = 0;
            Range range;
            string name;
            Int32 unixTimestamp = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            int i = 1;
            Console.WriteLine("images number:" + docs.InlineShapes.Count);
            foreach (Microsoft.Office.Interop.Word.InlineShape ilshp in docs.InlineShapes)
            {
                if (ilshp.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    ilshp.Select();
                    wordApplication.Selection.Copy();
                    name = String.Format("img_{0}.png", unixTimestamp + "_" + i);
                    // Check data is in the clipboard
                    if (Clipboard.GetDataObject() != null)
                    {
                        var data = Clipboard.GetDataObject();

                        // Check if the data conforms to a bitmap format
                        if (data != null && data.GetDataPresent(DataFormats.Bitmap))
                        {
                            // Fetch the image and convert it to a Bitmap
                            var image = (Image)data.GetData(DataFormats.Bitmap, true);
                            var currentBitmap = new Bitmap(image);
                            // Save the bitmap to a file
                            Console.WriteLine("save image:" + names[j]);
                            currentBitmap.Save(directoryPath + "\\картинки\\" + names[j]);
                        }
                    }

                    ilshp.Application.Selection.MoveEnd();

                    ranges.Add(ilshp.Range);
                    range = ilshp.Range;
                    ilshp.Delete();
                    range.InsertAfter("{" + names[j] + "}");

                    //break;

                    i++;
                    Microsoft.Office.Interop.Word.Paragraph prfo = docs.Paragraphs.Add(miss);
                }

                j++;

            }

            docs.Close(ref miss, ref miss, ref miss);
            
            generateJson();
        }


        public void generateJson() {
            object miss = System.Reflection.Missing.Value;
            object path = filePath;
            object readOnly = false;
            Microsoft.Office.Interop.Word.Document docs = wordApplication.Documents.Open(ref path, ref miss, ref readOnly, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss, ref miss);
            List<Question> questions = new List<Question>();
            Question question = null;
            Answer answer = null;
            string lastAction = "";
            int rightAnswer = 0;

            string temp = "";
            Console.WriteLine("!*test".Substring(1, "!*test".Length - 1));
            for (int i = 1; i < docs.Paragraphs.Count; i++)
            {
                string s = docs.Paragraphs[i].Range.Text; // get paragraph text
                //byte[] bytes = Encoding.Default.GetBytes(s);
                //s = Encoding.UTF8.GetString(bytes);

                Console.WriteLine(s);
                if (!String.IsNullOrEmpty(s))
                {
                    if (s[0] == '#')
                    {
                        if (lastAction == "answer_created")
                        {
                            questions.Add(question);
                        }
                        question = new Question();
                        question.Answers = new List<Answer>();
                        question.Type = "one";
                        lastAction = "new_question";
                        rightAnswer = 0;
                    }
                    else if (s[0] == '*')
                    {
                        if (s[1] == '!')
                        {
                            question.Content = Util.convert(s.Substring(2));
                            temp += question.Content;
                            question.Type = "one";
                            Console.WriteLine("question before:" + question.Content);
                            lastAction = "question_created";
                        }
                        else
                        {
                            answer = new Answer();
                            lastAction = "answer_created";
                            if (s[1] == '+')
                            {
                                answer.Content = Util.convert(s.Substring(2));
                                answer.Type = 1;
                                rightAnswer++;
                            }
                            else
                            {
                                answer.Content = Util.convert(s.Substring(1));
                                answer.Type = 0;
                            }
                            question.Answers.Add(answer);

                        }
                    }
                    else
                    {
                        if (lastAction == "question_created")
                        {
                            question.Content = Util.convert(String.Concat(question.Content, s));
                        }
                        else if (lastAction == "answer_created")
                        {
                            question.Answers[question.Answers.Count - 1].Content = Util.convert(String.Concat(answer.Content, s));
                        }
                    }

                    if (rightAnswer > 1)
                    {
                        question.Type = "multiple";
                    }

                    Console.WriteLine("action:" + lastAction);
                }

            }
            if (lastAction == "answer_created")
            {
                questions.Add(question);
            }
            string json = JsonConvert.SerializeObject(questions.ToArray());
            Console.WriteLine(json);
            System.IO.File.WriteAllText(directoryPath+"\\вопросы.json", json, System.Text.Encoding.GetEncoding("UTF-8"));

            docs.Close();

        }
    }
}
