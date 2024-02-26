using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;


namespace Links
{
    class Program
    {
        static List<Media> GetMediaList()
        {
            var mediaList = new List<Media>();
            var sr = new StreamReader("mediaList.txt");
            while (!sr.EndOfStream)
            {
                string _tmp = sr.ReadLine();
                var _tmpMedia = _tmp.Split('#');
                var _media = new Media();
                _media.name = _tmpMedia[0];
                _media.link = _tmpMedia[1];
                mediaList.Add(_media);
            }
            sr.Close();
            return mediaList;
        }

        static List<string> ReadFile(string fileName)
        {
            var links = new List<string>();
            var doc = DocX.Load(fileName);
            var _links = doc.Paragraphs.ToList();
            foreach(var item in _links)
            {
                links.Add(item.Text);
            }
            doc.Save();
            return links;
        }

        static void Main(string[] args)
        {
            var mediaList = GetMediaList();
            var linksNeg = ReadFile("data\\Negative.docx");
            var linksPoz = ReadFile("data\\Pozitive.docx");
            var linksNeu = ReadFile("data\\Neutral.docx");
            var notFoundNeg = new List<string>();
            var notFoundPoz = new List<string>();
            var notFoundNeu = new List<string>();
            foreach(var item in linksNeg)
            {
                bool found = false;
                foreach(var media in mediaList)
                {
                    if (item.Contains(media.link))
                    {
                        found = true;
                        media.neg++;
                        break;
                    }
                }
                if (!found && item!="")
                {
                    notFoundNeg.Add(item);
                }
            }

            foreach (var item in linksPoz)
            {
                bool found = false;
                foreach (var media in mediaList)
                {
                    if (item.Contains(media.link))
                    {
                        found = true;
                        media.poz++;
                        break;
                    }
                }
                if (!found && item != "")
                {
                    notFoundPoz.Add(item);
                }
            }

            foreach (var item in linksNeu)
            {
                bool found = false;
                foreach (var media in mediaList)
                {
                    if (item.Contains(media.link))
                    {
                        found = true;
                        media.neu++;
                        break;
                    }
                }
                if (!found && item != "")
                {
                    notFoundNeu.Add(item);
                }
            }

            for (int i = 0; i < mediaList.Count; i++)
            {
                for(int j = 1; j < mediaList.Count; j++)
                {
                    if (mediaList[j] > mediaList[j - 1])
                    {
                        var _tmp = mediaList[j];
                        mediaList[j] = mediaList[j - 1];
                        mediaList[j - 1] = _tmp;
                    }
                }
            }

            using (var document = DocX.Create("data\\Result.docx"))
            {
                // Save this document to disk.
                var table = document.InsertTable(1, 2);
                table.SetColumnWidth(0, 880.74, false);
                table.SetColumnWidth(1, 600.2, false);
                
                for (int i = 0; i < mediaList.Count; i++)
                {
                    if(mediaList[i].neg>0 || mediaList[i].poz > 0 || mediaList[i].neu > 0)
                    {
                        table.Rows[i].Cells[0].Paragraphs[0].Append(mediaList[i].name).Color(Color.Black).Font("Times New Roman").FontSize(12);
                        if(mediaList[i].neg > 0)
                        {
                            table.Rows[i].Cells[1].Paragraphs[0].Append("-" + mediaList[i].neg).Color(Color.Red).Bold().Font("Times New Roman").FontSize(12);
                        }
                        if (mediaList[i].poz > 0)
                        {
                            if(mediaList[i].neg > 0)
                            {
                                table.Rows[i].Cells[1].Paragraphs[0].Append("  ").Color(Color.Green).Bold().Font("Times New Roman").FontSize(12);
                            }
                            table.Rows[i].Cells[1].Paragraphs[0].Append("+" + mediaList[i].poz).Color(Color.Green).Bold().Font("Times New Roman").FontSize(12);
                        }
                        if (mediaList[i].neu > 0)
                        {
                            if(mediaList[i].neg > 0 || mediaList[i].poz > 0)
                            {
                                table.Rows[i].Cells[1].Paragraphs[0].Append("  ").Color(Color.Blue).Bold().Font("Times New Roman").FontSize(12);
                            }
                            table.Rows[i].Cells[1].Paragraphs[0].Append(mediaList[i].neu.ToString()).Color(Color.Blue).Bold().Font("Times New Roman").FontSize(12);
                        }
                        table.Rows[i].Cells[0].Paragraphs[0].Alignment = Alignment.center;
                        table.Rows[i].Cells[1].Paragraphs[0].Alignment = Alignment.center;
                        if(i<(mediaList.Count - 1))
                        {
                            table.InsertRow();
                        }
                    }
                }
                int rowsCount = table.RowCount;
                if (table.Rows[rowsCount - 1].Cells[0].Paragraphs[0].Text == "")
                {
                    table.RemoveRow();
                }
                if (notFoundNeg.Count>0 || notFoundPoz.Count>0 || notFoundNeu.Count > 0)
                {
                    var p0 = document.InsertParagraph();
                    var p1 = document.InsertParagraph().Font("Times New Roman").FontSize(12).Bold();
                    p1.Alignment = Alignment.center;
                    p1.Append("Следующие гиперссылки не были распознаны").FontSize(12).Bold();
                    var p2 = document.InsertParagraph();
                    if (notFoundNeg.Count > 0)
                    {
                        foreach(var item in notFoundNeg)
                        {
                            var par = document.InsertParagraph().Color(Color.Red).Font("Times New Roman").FontSize(12);
                            par.Append(item).Color(Color.Red).Font("Times New Roman").FontSize(12);
                            var p_ = document.InsertParagraph();
                        }
                    }
                    if (notFoundPoz.Count > 0)
                    {
                        foreach (var item in notFoundPoz)
                        {
                            var par = document.InsertParagraph().Color(Color.Green).Font("Times New Roman").FontSize(12);
                            par.Append(item).Color(Color.Green).Font("Times New Roman").FontSize(12);
                            var p_ = document.InsertParagraph();
                        }
                    }
                    if (notFoundNeu.Count > 0)
                    {
                        foreach (var item in notFoundNeu)
                        {
                            var par = document.InsertParagraph().Color(Color.Blue).Font("Times New Roman").FontSize(12);
                            par.Append(item).Color(Color.Blue).Font("Times New Roman").FontSize(12);
                            var p_ = document.InsertParagraph();
                        }
                    }
                }
                else
                {
                    var p0 = document.InsertParagraph();
                    var p1 = document.InsertParagraph().Font("Times New Roman").FontSize(12).Bold();
                    p1.Alignment = Alignment.center;
                    p1.Append("Все гиперссылки были распознаны").FontSize(12).Bold();
                }
                document.Save();
            }
            Console.WriteLine("Результаты записаны в файл Result.docx");
            Console.ReadKey();
        }
    }
}
