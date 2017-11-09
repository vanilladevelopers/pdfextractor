/* 
 Proof of Concept for PDF splitting based on a specific word within 
 the text (e.g., a document containing the word "session" followed 
 by a sequential integer id). The program outputs the total processing time
 and the error detection rate for the said word and the id. It also demonstrates
 the use of meta-data manipulation and cropping via the 'Stamper' class.
*/
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Diagnostics;

namespace PdfExtractor
{
  //Helper class that stores the coordinates of the word found plus its text
  public class RectAndText
  {
    public Rectangle Rect;
    public String Text;
    public RectAndText(Rectangle rect, String text)
    {
      this.Rect = rect;
      this.Text = text;
    }
  }

  public class MyLocationTextExtractionStrategy : LocationTextExtractionStrategy
  {
    //Hold each coordinate
    public List<RectAndText> myPoints = new List<RectAndText>();

    //Automatically called for each chunk of text in the PDF
    public override void RenderText(TextRenderInfo renderInfo)
    {
      base.RenderText(renderInfo);

      //Get the bounding box for the chunk of text
      var bottomLeft = renderInfo.GetDescentLine().GetStartPoint();
      var topRight = renderInfo.GetAscentLine().GetEndPoint();

      //Create a rectangle from it
      var rect = new iTextSharp.text.Rectangle(
                                              bottomLeft[Vector.I1],
                                              bottomLeft[Vector.I2],
                                              topRight[Vector.I1],
                                              topRight[Vector.I2]
                                              );

      //Add this to our main collection
      this.myPoints.Add(new RectAndText(rect, renderInfo.GetText()));
    }
  }
  
  class Program
  {
    static float gapV = 9.0f;

    static string outputDir = "";
    static string _filePath = "";

    struct Info
    {
      public string sessionId;
      public Rectangle rect;

      public void Print()
      {
        Console.Out.WriteLine("Session Id: " + sessionId + "; Rect: (" + rect.Left + "," + rect.Top + "," + rect.Right + "," + rect.Bottom + ")");
      }
    }

    static void Extract(bool recoveryEnabled)
    {
      List<int> errorSessionList = new List<int>();
      int sessionsX = 0;
      int global = 0;
      var pdfReader = new PdfReader(_filePath);
      int lastPage;
      string format = "D" + pdfReader.NumberOfPages.ToString().Length;
      int lastValidIdFromPreviousPage = -1;

      for (int p = 0; p < pdfReader.NumberOfPages; p++)
      {
        Console.Clear();
        Console.SetCursorPosition(0, 0);
        Console.WriteLine("Processing page " + (p + 1).ToString(format) + " of " + pdfReader.NumberOfPages);

        string textFromPage = "";
        lastPage = p;
        while (lastPage + 1 <= pdfReader.NumberOfPages && lastPage + 1 != pdfReader.NumberOfPages)
        {
          lastPage++;
          textFromPage = PdfTextExtractor.GetTextFromPage(pdfReader, lastPage + 1);
          if (textFromPage.Contains("Sess"))
          {
            break;
          }
        }

        string segments = PdfTextExtractor.GetTextFromPage(pdfReader, p + 1);
        List<string> sessionList = new List<string>();
        while (segments.Length > 0)
        {
          int i1 = segments.IndexOf("Sess");

          //text goes over more than 2 pages but we already got the session info
          if (i1 == -1)
          {
            break;
          }

          int i2 = segments.IndexOf("Sess", i1 + "Session".Length);
          if (i2 != -1)
          {
            string[] kArr = segments.Substring(i2, 20).Split(' ');

            if (!kArr[1].Contains("#"))
            {
              string temp = segments.Substring(i1, i2 - i1);
              sessionList.Add(temp);
              segments = segments.Substring(i2, segments.Length - i2);

            }
            else
            {
              i2 = segments.IndexOf("Sess", i2 + "Session".Length);
              if (i2 != -1)
              {
                string temp = segments.Substring(i1, i2 - i1);
                sessionList.Add(temp);
                segments = segments.Substring(i2, segments.Length - i2);
              }
              else
              {
                string temp = segments.Substring(i1, segments.Length - i1);
                sessionList.Add(temp);
                segments = "";
              }
            }
          }
          else
          {
            string temp = segments.Substring(i1, segments.Length - i1);
            sessionList.Add(temp);
            segments = "";
          }
        }

        List<Info> infoList = new List<Info>();
        int sessions = 0;
        var tes = new MyLocationTextExtractionStrategy();
        var ex = PdfTextExtractor.GetTextFromPage(pdfReader, p + 1, tes);
        for (int chunk = 0; chunk < tes.myPoints.Count; chunk++)
        {
          if (tes.myPoints[chunk].Text.Contains("Sess"))
          {
            if (chunk + 1 < tes.myPoints.Count && tes.myPoints[chunk + 1].Text.Contains("#"))
            {
              continue;
            }

            Info info;
            info.sessionId = "NA";
            info.rect = tes.myPoints[chunk].Rect;

            //get the Session number
            Regex regex = new Regex(@"^[0-9]+$");
            string[] tokens = sessionList[sessions].Split(' ');
            for (int t = 0; t < tokens.Length; t++)
            {
              if (tokens[t].Contains("Sess"))
              {
                /**** do a regex to any number of digits****/
                if (t + 1 < tokens.Length && regex.IsMatch(tokens[t + 1].Trim()))
                {
                  info.sessionId = tokens[t + 1].Trim();
                }
                break;
              }
            }
            //info.Print();
            infoList.Add(info);
            sessions++;
          }
        }

        if (infoList.Count == 0)
        {
          Console.WriteLine("Fatal: PDF resolution did not allow to recover word rendering locations");
          return;
        }

        /*** Recovering session numbers lost ***/
        if (recoveryEnabled)
        {
          int indexFound = -1;
          int s = 0;
          for (; s < infoList.Count; s++)
          {
            if (indexFound != -1)
            {
              if (!infoList[s].sessionId.Contains("NA"))
              {
                if (int.Parse(infoList[indexFound].sessionId) - int.Parse(infoList[s].sessionId) == indexFound - s)
                {
                  break;
                }
              }
            }

            if (!infoList[s].sessionId.Contains("NA"))
            {
              indexFound = s;
              if (lastValidIdFromPreviousPage != -1)
              {
                if (Math.Abs(lastValidIdFromPreviousPage - int.Parse(infoList[s].sessionId)) == s + 1)
                {
                  break;
                }
              }
            }
          }
          if (s == infoList.Count && indexFound != s - 1)
          {
            indexFound = -1;
          }

          if (indexFound != -1)
          {
            int val = int.Parse(infoList[indexFound].sessionId);
            for (int tt = 0; tt < infoList.Count; tt++)
            {
              Info info = infoList[tt];
              if (info.sessionId.Contains("NA") && tt < indexFound)
              {
                int newVal = val - Math.Abs(indexFound - tt);
                info.sessionId = (newVal).ToString();
                infoList[tt] = info;
                errorSessionList.Add(newVal);
              }
              else if (info.sessionId.Contains("NA") && tt > indexFound)
              {
                int newVal = val + Math.Abs(indexFound - tt);
                info.sessionId = (newVal).ToString();
                infoList[tt] = info;
                errorSessionList.Add(newVal);
              }
              else if (!info.sessionId.Contains("NA"))
              {
                if (lastValidIdFromPreviousPage != -1)
                {
                  int iVal = lastValidIdFromPreviousPage + tt + 1;
                  if (int.Parse(infoList[tt].sessionId) != iVal)
                  {
                    //Correcting OCR error
                    info.sessionId = iVal.ToString();
                    infoList[tt] = info;
                    errorSessionList.Add(iVal);
                  }
                }
                else
                {
                  int iVal = int.Parse(infoList[indexFound].sessionId) - indexFound + tt;
                  if (int.Parse(infoList[tt].sessionId) != iVal)
                  {
                    //Correcting OCR error
                    info.sessionId = iVal.ToString();
                    infoList[tt] = info;
                    errorSessionList.Add(iVal);
                  }
                }
              }
            }
            lastValidIdFromPreviousPage = int.Parse(infoList[infoList.Count - 1].sessionId);
          }
          else //try again in case the whole page failed
          {
            if (lastValidIdFromPreviousPage != -1)
            {
              int ss = lastValidIdFromPreviousPage + 1;
              for (int tt = 0; tt < infoList.Count; tt++)
              {
                Info info = infoList[tt];
                if (info.sessionId.Contains("NA"))
                {
                  info.sessionId = (ss + tt).ToString();
                  infoList[tt] = info;
                }
              }
              if (infoList.Count - 1 >= 0)
              {
                lastValidIdFromPreviousPage = int.Parse(infoList[infoList.Count - 1].sessionId);
              }
            }
            else
            {
              if (!int.TryParse(infoList[infoList.Count - 1].sessionId, out lastValidIdFromPreviousPage))
              {
                Console.WriteLine("Fatal: Could not recover from this error on page " + (p + 1));
                lastValidIdFromPreviousPage = -1;
              }
            }
          }

          //double check
          foreach (Info info in infoList)
          {
            if (info.sessionId.Contains("NA"))
            {
              Console.WriteLine("Could not recover session Id on page " + (p + 1));
              sessionsX++;
            }
          }
        }
        else
        {
          //just count
          foreach (Info info in infoList)
          {
            if (info.sessionId.Contains("NA"))
            {
              sessionsX++;
            }
          }
        }
        
        //Single-page template
        Document doc1 = new Document();
        PdfCopy copy1;
        PdfImportedPage importedPage1;
        using (FileStream fs1 = new FileStream(outputDir + "template1.pdf", FileMode.Create))
        {
          copy1 = new PdfCopy(doc1, fs1);
          doc1.Open();
          importedPage1 = copy1.GetImportedPage(pdfReader, p + 1);
          copy1.AddPage(importedPage1);
          doc1.Close();
        }

        //Multi-page template
        int pTemp;
        Document docN = new Document();
        PdfCopy copyN = null;
        PdfImportedPage importedPageN = null;
        using (FileStream fs = new FileStream(outputDir + "templateN.pdf", FileMode.Create))
        {
          copyN = new PdfCopy(docN, fs);
          docN.Open();
          pTemp = p;
          while (pTemp <= lastPage && pTemp != pdfReader.NumberOfPages)
          {
            importedPageN = copyN.GetImportedPage(pdfReader, pTemp + 1);
            copyN.AddPage(importedPageN);
            pTemp++;
          }
          docN.Close();
        }

        //retrieve the info from for the last page where the next session is
        Info infoEnd;
        infoEnd.sessionId = "Last";
        Rectangle r = pdfReader.GetPageSize(lastPage + 1);
        infoEnd.rect = new Rectangle(0, r.Top, 0, r.Bottom); ///swapping top

        tes = new MyLocationTextExtractionStrategy();
        ex = PdfTextExtractor.GetTextFromPage(pdfReader, lastPage + 1, tes);
        for (int chunk = 0; chunk < tes.myPoints.Count; chunk++)
        {
          if (tes.myPoints[chunk].Text.Contains("Sess"))
          {
            infoEnd.rect = tes.myPoints[chunk].Rect;
            break;
          }
        }

        //go through all info
        for (int info = 0; info < infoList.Count; info++)
        {
          string outputFilename = infoList[info].sessionId.Trim();

          if (File.Exists(outputDir + outputFilename + ".pdf"))
          {
            outputFilename += "-" + (global + 1).ToString("D4");
          }
          global++;

          outputFilename += ".pdf";

          PdfStamper stamper;
          PdfDictionary page;
          if (info + 1 < infoList.Count)
          {
            var pdfReader2 = new PdfReader(outputDir + "template1.pdf");
            using (FileStream fs = new FileStream(outputDir + outputFilename, FileMode.Create))
            {
              stamper = new PdfStamper(pdfReader2, fs);
              page = pdfReader2.GetPageN(1);
              page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, infoList[info + 1].rect.Top + gapV, importedPage1.Width, infoList[info].rect.Top + gapV }));
              stamper.MarkUsed(page);

              Dictionary<String, String> meta = pdfReader2.Info;
              meta.Add("Keywords", infoList[info].sessionId);
              stamper.MoreInfo = meta;

              stamper.Close();
              pdfReader2.Close();
            }
          }
          else
          {
            PdfReader pdfReader2;

            //last page
            if (p + 1 == pdfReader.NumberOfPages)
            {
              pdfReader2 = new PdfReader(outputDir + "template1.pdf");
              stamper = new PdfStamper(pdfReader2, new FileStream(outputDir + outputFilename, FileMode.Create));

              page = pdfReader2.GetPageN(1);
              page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, 0, importedPageN.Width, infoList[info].rect.Top + gapV }));
              stamper.MarkUsed(page);

              Dictionary<String, String> meta2 = pdfReader2.Info;
              meta2.Add("Keywords", infoList[info].sessionId);
              stamper.MoreInfo = meta2;

              stamper.Close();
              pdfReader2.Close();
              break;
            }
            else //last sesion 
            {
              pdfReader2 = new PdfReader(outputDir + "templateN.pdf");
              stamper = new PdfStamper(pdfReader2, new FileStream(outputDir + outputFilename, FileMode.Create));

              page = pdfReader2.GetPageN(1);
              page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, 0, importedPageN.Width, infoList[info].rect.Top + gapV }));
              stamper.MarkUsed(page);

              //session with more than 2 pages              
              if (lastPage - p > 1)
              {
                int newP = 1;
                int max = lastPage - p;
                while (newP <= max)
                {
                  page = pdfReader2.GetPageN(newP + 1);

                  if (newP != max)
                  {
                    page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, 0, importedPageN.Width, importedPageN.Height }));
                  }
                  else
                  {
                    page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, infoEnd.rect.Top + gapV, importedPageN.Width, importedPageN.Height }));
                  }

                  stamper.MarkUsed(page);

                  newP++;
                }
                p = lastPage - 1;
              }
              else
              {
                page = pdfReader2.GetPageN(2);
                page.Put(PdfName.CROPBOX, new PdfArray(new float[] { 0, infoEnd.rect.Top + gapV, importedPageN.Width, importedPageN.Height }));

                stamper.MarkUsed(page);
              }
            }

            Dictionary<String, String> meta = pdfReader2.Info;
            meta.Add("Keywords", infoList[info].sessionId);
            stamper.MoreInfo = meta;

            stamper.Close();
            pdfReader2.Close();
          }
        }
      }

      Console.WriteLine("Sessions extracted: " + global);
      if (recoveryEnabled)
      {
        Console.WriteLine("Session ID recovery was attempted. Percentage of undetected Session IDs after OCR reconstruction was " + (((float)sessionsX / (float)global) * 100).ToString("0.00") + "%");
      }
      else
      {
        Console.WriteLine("No Session ID recovery was attempted. Percentage of undetected Session IDs with no In/Out Digits was " + (((float)sessionsX / (float)global) * 100).ToString("0.00") + "%");
      }

      Console.WriteLine();
    }

    static void Main(string[] args)
    {
      if (args.Length == 2 || args.Length == 3)
      {
        if (File.Exists(args[0]) && System.IO.Path.GetExtension(args[0]).ToLower() == ".pdf")
        {
          if (Directory.Exists(args[1]) || args[1] == "")
          {
            Stopwatch sw = new Stopwatch();
            sw.Start();

            _filePath = args[0];
            outputDir = (args[1] == "") ? "" : args[1] + "\\";
            if (args.Length == 3)
            {
              if (args[2] == "-r")
              {
                Extract(true);
                return;
              }
              else
              {
                Console.WriteLine("Invalid option\n");
              }
            }
            else
            {
              Extract(false);
              return;
            }

            sw.Stop();

            if (sw.Elapsed.Seconds > 1)
            {
              Console.WriteLine("Elapsed time was = {0} mins {1} secs", sw.Elapsed.Minutes, sw.Elapsed.Seconds);
            }
            else
            {
              Console.WriteLine("Elapsed time was less than a second");
            }
          }
          else
          {
            Console.WriteLine("Invalid output directory");
          }
        }
        else
        {
          Console.WriteLine("Invalid input file");
        }
      }
    }
  }
}
