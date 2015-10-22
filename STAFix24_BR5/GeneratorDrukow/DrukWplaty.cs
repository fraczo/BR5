using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.IO;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using Microsoft.SharePoint.Utilities;
using System.Text.RegularExpressions;

namespace GeneratorDrukow
{
    public static class DrukWplaty
    {
        public static bool Attach_DrukWplaty(SPWeb web, SPListItem item, string nazwaPliku, string odbiorca, string numerRachunku, double kwotaDoZaplaty, string zleceniodawca, string tytulem)
        {
            //const int maxLen = 27;
            const int maxLen = 46;
            odbiorca = odbiorca.ToUpper();
            zleceniodawca = zleceniodawca.ToUpper();
            tytulem = tytulem.ToUpper();

            //unormowanie nazwy pliku
            nazwaPliku = CleanupFileName(nazwaPliku);

            if (numerRachunku.Length != 26)
            {
                return false;
            }

            if (item.Attachments.Count > 0)
            {
                foreach (string fname in item.Attachments)
                {
                    if (fname == nazwaPliku)
                    {
                        item.Attachments.DeleteNow(fname);
                        break;
                    }
                }
            }

            string pdfFilePath = @"SiteAssets/Templates/DW-Przelew.pdf";
            int x0 = 89;
            int y0 = 76;

            //string pdfFilePath = @"SiteAssets/Templates/PolecenieWplaty-Template.pdf";
            //int x0 = 95;
            //int y0 = 115;

            //string pdfFilePath = @"SiteAssets/Templates/MyTemplate.pdf";

            SPFile file = web.GetFile(pdfFilePath);

            if (file.Exists)
            {
                int bufferSize = 20480;
                byte[] byteBuffer = new byte[bufferSize];
                //byteBuffer = File.ReadAllBytes(pdfFilePath);
                byteBuffer = file.OpenBinary();

                MemoryStream coverSheetContent = new MemoryStream();

                coverSheetContent.Write(byteBuffer, 0, byteBuffer.Length);
                int t = PdfReader.TestPdfFile(coverSheetContent);
                PdfDocument document = PdfReader.Open(coverSheetContent);
                PdfPage page = document.Pages[0];

                // Create an empty XForm object with the specified width and height
                // A form is bound to its target document when it is created. The reason is that the form can 
                // share fonts and other objects with its target document.
                XForm form = new XForm(document, XUnit.FromMillimeter(138), XUnit.FromMillimeter(77));

                // Create an XGraphics object for drawing the contents of the form.
                XGraphics formGfx = XGraphics.FromForm(form);

                // Draw a large transparent rectangle to visualize the area the form occupies
                XColor back = XColors.Orange;
                back.A = 0.2;
                XSolidBrush brush = new XSolidBrush(back);
                formGfx.DrawRectangle(brush, -10000, -10000, 20000, 20000);

                // On a form you can draw...

                // ... text

                XFont font = new XFont("Verdana", 10, XFontStyle.Bold); //XFont("Verdana", 10, XFontStyle.Regular)

                string odbiorca2 = string.Empty;
                if (odbiorca.Length > maxLen)
                {
                    odbiorca2 = odbiorca.Substring(maxLen, odbiorca.Length - maxLen);
                    if (odbiorca2.Length > maxLen)
                    {
                        odbiorca2 = odbiorca2.Substring(0, maxLen);
                    }
                    odbiorca = odbiorca.Substring(0, maxLen);
                }

                formGfx.DrawString(odbiorca, font, XBrushes.Navy, 8, 8, XStringFormats.TopLeft);
                formGfx.DrawString(odbiorca2, font, XBrushes.Navy, 8, 33, XStringFormats.TopLeft);
                //formGfx.DrawString("Numer rachunku odbiorcy przekazu pocztowego", new XFont("Verdana", 10, XFontStyle.Regular), XBrushes.Navy, 8, 57, XStringFormats.TopLeft);

                int n = 0;
                formGfx.DrawString(numerRachunku.Substring(n, 2), font, XBrushes.Navy, 13, 57, XStringFormats.TopLeft); n = n + 2;

                int offset = 57;
                int targetX = 48;

                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX, 57, XStringFormats.TopLeft); n = n + 4;
                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX + 1 * offset, 57, XStringFormats.TopLeft); n = n + 4;
                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX + 2 * offset, 57, XStringFormats.TopLeft); n = n + 4;
                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX + 3 * offset, 57, XStringFormats.TopLeft); n = n + 4;
                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX + 4 * offset, 57, XStringFormats.TopLeft); n = n + 4;
                formGfx.DrawString(numerRachunku.Substring(n, 4), font, XBrushes.Navy, targetX + 5 * offset, 57, XStringFormats.TopLeft);

                formGfx.DrawString("***" + String.Format("{0:#,0.00}", kwotaDoZaplaty) + "***", font, XBrushes.Navy, 220, 81, XStringFormats.TopLeft);

                int zlote = (int)kwotaDoZaplaty;
                int grosze = (int)(100 * kwotaDoZaplaty) % 100;
                string kwota = String.Format("{0} {1}",
                    KwotaSlownie.LiczbaSlownie(zlote),
                    grosze + "/100");

                formGfx.DrawString("***" + kwota + "***", new XFont("Verdana", 10, XFontStyle.Regular), XBrushes.Navy, 8, 106, XStringFormats.TopLeft);

                string zleceniodawca2 = string.Empty;
                if (zleceniodawca.Length > maxLen)
                {
                    zleceniodawca2 = zleceniodawca.Substring(maxLen, zleceniodawca.Length - maxLen);
                    if (zleceniodawca2.Length > maxLen)
                    {
                        zleceniodawca2 = zleceniodawca2.Substring(0, maxLen);
                    }
                    zleceniodawca = zleceniodawca.Substring(0, maxLen);
                }
                formGfx.DrawString(zleceniodawca, font, XBrushes.Navy, 8, 130, XStringFormats.TopLeft);
                formGfx.DrawString(zleceniodawca2, font, XBrushes.Navy, 8, 154, XStringFormats.TopLeft);

                string tytulem2 = string.Empty;
                if (tytulem.Length > maxLen)
                {
                    tytulem2 = tytulem.Substring(maxLen, tytulem.Length - maxLen);
                    if (tytulem2.Length > maxLen)
                    {
                        tytulem2 = tytulem2.Substring(0, maxLen);
                    }
                    tytulem = tytulem.Substring(0, maxLen);
                }
                formGfx.DrawString(tytulem, font, XBrushes.Navy, 8, 178, XStringFormats.TopLeft);
                formGfx.DrawString(tytulem2, font, XBrushes.Navy, 8, 201, XStringFormats.TopLeft);




                XPen pen = XPens.LightBlue.Clone();
                pen.Width = 2.5;

                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Draw the form on the page of the document in its original size
                gfx.DrawImage(form, x0, y0);
                gfx.DrawImage(form, x0, y0 + 300);

                MemoryStream ms = new MemoryStream();
                document.Save(ms);

                item.Attachments.AddNow(nazwaPliku, ms.GetBuffer());
                //item.SystemUpdate();

                return true;
            }

            else
            {
                return false;
            }


        }

        public static bool Attach_DrukWplatyPD(SPWeb web, SPListItem item, string nazwaPliku, string odbiorca, string numerRachunku, double kwotaDoZaplaty, string zleceniodawca, string nip, string typIdentyfikatora, string okres, string symbolFormularza, string identyfikacjaZobowiazania)
        {
            const int maxLen = 27;
            odbiorca = odbiorca.ToUpper();
            zleceniodawca = zleceniodawca.ToUpper();

            //unormowanie nazwy pliku
            nazwaPliku = CleanupFileName(nazwaPliku);

            if (numerRachunku.Length != 26)
            {
                return false;
            }

            if (item.Attachments.Count > 0)
            {
                foreach (string fname in item.Attachments)
                {
                    if (fname == nazwaPliku)
                    {
                        item.Attachments.DeleteNow(fname);
                        break;
                    }
                }
            }

            string pdfFilePath = @"SiteAssets/Templates/DW-Podatek.pdf";
            int x0 = 90;
            int dx = 136;
            int y0 = 74;
            int dy = 76;
            double ofset0 = 14.3;
            int c01 = 3;
            int r01 = 7;
            int r02 = 31;
            int r03 = 55;
            int c02 = 218;
            int r04 = 81;
            int r05 = 103;
            int r06 = 127;
            int r07 = 152;
            int r08 = 176;
            int r09 = 200;


            SPFile file = web.GetFile(pdfFilePath);

            if (file.Exists)
            {
                int bufferSize = 20480;
                byte[] byteBuffer = new byte[bufferSize];
                //byteBuffer = File.ReadAllBytes(pdfFilePath);
                byteBuffer = file.OpenBinary();

                MemoryStream coverSheetContent = new MemoryStream();

                coverSheetContent.Write(byteBuffer, 0, byteBuffer.Length);
                int t = PdfReader.TestPdfFile(coverSheetContent);
                PdfDocument document = PdfReader.Open(coverSheetContent);
                PdfPage page = document.Pages[0];

                XForm form = new XForm(document, XUnit.FromMillimeter(dx), XUnit.FromMillimeter(dy));
                XGraphics formGfx = XGraphics.FromForm(form);

                // Draw a large transparent rectangle to visualize the area the form occupies
#if DEBUG
                XColor back = XColors.Orange;
#else
                XColor back = XColors.White;
#endif

                back.A = 0.2;
                XSolidBrush brush = new XSolidBrush(back);
                formGfx.DrawRectangle(brush, -10000, -10000, 20000, 20000);

                XFont font = new XFont("Verdana", 10, XFontStyle.Bold); //XFont("Verdana", 10, XFontStyle.Regular)

                string odbiorca2 = string.Empty;
                if (odbiorca.Length > maxLen)
                {
                    odbiorca2 = odbiorca.Substring(maxLen, odbiorca.Length - maxLen);
                    if (odbiorca2.Length > maxLen)
                    {
                        odbiorca2 = odbiorca2.Substring(0, maxLen);
                    }
                    odbiorca = odbiorca.Substring(0, maxLen);
                }

                PlotText(odbiorca, c01, r01, ofset0, formGfx, font);

                PlotText(odbiorca2, c01, r02, ofset0, formGfx, font);

                PlotText(numerRachunku, c01, r03, ofset0, formGfx, font);

                PlotText("X", c01 + (int)(9 * ofset0) + 2, r04 - 3, ofset0, formGfx, font);
                PlotText((string)(String.Format("{0:#0.00}", kwotaDoZaplaty) + "************").Substring(0, 12), c02, r04, ofset0, formGfx, font);

                int zlote = (int)kwotaDoZaplaty;
                int grosze = (int)(100 * kwotaDoZaplaty) % 100;
                string kwota = String.Format("{0} {1}",
                    KwotaSlownie.LiczbaSlownie(zlote),
                    grosze + "/100");

                formGfx.DrawString("*" + kwota + "*", new XFont("Verdana", 10, XFontStyle.Regular), XBrushes.Navy, c01, r05, XStringFormats.TopLeft);

                string zleceniodawca2 = string.Empty;
                if (zleceniodawca.Length > maxLen)
                {
                    zleceniodawca2 = zleceniodawca.Substring(maxLen, zleceniodawca.Length - maxLen);
                    if (zleceniodawca2.Length > maxLen)
                    {
                        zleceniodawca2 = zleceniodawca2.Substring(0, maxLen);
                    }
                    zleceniodawca = zleceniodawca.Substring(0, maxLen);
                }
                PlotText(zleceniodawca, c01, r06, ofset0, formGfx, font);
                PlotText(zleceniodawca2, c01, r07, ofset0, formGfx, font);

                PlotText(nip, c01, r08, ofset0, formGfx, font);
                PlotText(typIdentyfikatora, c01 + (int)(15 * ofset0), r08, ofset0, formGfx, font);
                PlotText(okres, c01 + (int)(19 * ofset0), r08, ofset0, formGfx, font);
                PlotText(symbolFormularza, c01, r09, ofset0, formGfx, font);
                PlotText(identyfikacjaZobowiazania, c01 + (int)(7 * ofset0), r09, ofset0, formGfx, font);


                XPen pen = XPens.LightBlue.Clone();
                pen.Width = 2.5;

                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Draw the form on the page of the document in its original size
                gfx.DrawImage(form, x0, y0);
                gfx.DrawImage(form, x0, y0 + 300);

                MemoryStream ms = new MemoryStream();
                document.Save(ms);

                item.Attachments.AddNow(nazwaPliku, ms.GetBuffer());
                //item.SystemUpdate();

                return true;
            }

            else
            {
                return false;
            }


        }

        private static void PlotText(string napis, int c01, int r01, double offset, XGraphics formGfx, XFont font)
        {
            for (int i = 0; i < napis.Length; i++)
            {
                string znak = napis.Substring(i, 1);

                PlotChar(znak, c01 + (int)(i * offset), r01, formGfx, font);
            }
        }

        private static void PlotChar(string odbiorca, int c01, int r01, XGraphics formGfx, XFont font)
        {
            formGfx.DrawString(odbiorca, font, XBrushes.Navy, c01, r01, XStringFormats.TopLeft);
        }

        private static string CleanupFileName(string nazwaPliku)
        {
            //string illegal = "\"M\"\\a/ry/ h**ad:>> a\\/:*?\"| li*tt|le|| la\"mb.?";
            string regexSearch = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
            Regex r = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            return r.Replace(nazwaPliku, "");

        }

        public static bool Attach_DrukWplatyZUS(SPWeb web, SPListItem item, string nazwaPliku, string numerRachunku, double kwotaDoZaplaty, string nadawca, string nip, string typIdentyfikatora, string drugiIdentyfikator, string identyfikatorDeklaracji)
        {
            const int maxLen = 27;

            nazwaPliku = CleanupFileName(nazwaPliku);

            if (numerRachunku.Length != 26)
            {
                return false;
            }
            else
            {
                numerRachunku = String.Format(@"{0}                  {1}", numerRachunku.Substring(0, 2), numerRachunku.Substring(20, 1));
            }

            if (item.Attachments.Count > 0)
            {
                foreach (string fname in item.Attachments)
                {
                    if (fname == nazwaPliku)
                    {
                        item.Attachments.DeleteNow(fname);
                        break;
                    }
                }
            }

            string pdfFilePath = @"SiteAssets/Templates/DW-ZUS.pdf";
            int x0 = 92;
            int dx = 135;
            int y0 = 76;
            int dy = 78;
            int formoffset = 304;

            double ofset0 = 14.15; //odstępo pomiędzy znakami
            double ofset1 = 24.1; //odstęp pomiędzy liniami

            int c01 = 3;
            int r01 = 58;
            int r02 = r01 + (int)(ofset1 * 1);
            int r03 = r01 + (int)(ofset1 * 2);
            int c02 = 215;
            int r04 = r01 + (int)(ofset1 * 3);
            int r05 = r01 + (int)(ofset1 * 4);
            int r06 = r01 + (int)(ofset1 * 5);
            int r07 = r01 + (int)(ofset1 * 6);



            SPFile file = web.GetFile(pdfFilePath);

            if (file.Exists)
            {
                int bufferSize = 20480;
                byte[] byteBuffer = new byte[bufferSize];
                byteBuffer = file.OpenBinary();

                MemoryStream coverSheetContent = new MemoryStream();

                coverSheetContent.Write(byteBuffer, 0, byteBuffer.Length);
                int t = PdfReader.TestPdfFile(coverSheetContent);
                PdfDocument document = PdfReader.Open(coverSheetContent);
                PdfPage page = document.Pages[0];

                XForm form = new XForm(document, XUnit.FromMillimeter(dx), XUnit.FromMillimeter(dy));
                XGraphics formGfx = XGraphics.FromForm(form);
#if DEBUG
                XColor back = XColors.Orange;
#else
                XColor back = XColors.White;
#endif
                back.A = 0.2;
                XSolidBrush brush = new XSolidBrush(back);
                formGfx.DrawRectangle(brush, -10000, -10000, 20000, 20000);

                XFont font = new XFont("Verdana", 10, XFontStyle.Bold);

                PlotText(numerRachunku, c01, r01, ofset0, formGfx, font);

                PlotText((string)(String.Format("{0:#0.00}", kwotaDoZaplaty) + "************").Substring(0, 12), c02, r02, ofset0, formGfx, font);

                int zlote = (int)kwotaDoZaplaty;
                int grosze = (int)(100 * kwotaDoZaplaty) % 100;
                string kwota = String.Format("{0} {1}",
                    KwotaSlownie.LiczbaSlownie(zlote),
                    grosze + "/100");

                formGfx.DrawString("*" + kwota + "*", new XFont("Verdana", 10, XFontStyle.Regular), XBrushes.Navy, c01, r03, XStringFormats.TopLeft);

                string nadawca2 = string.Empty;
                if (nadawca.Length > maxLen)
                {
                    nadawca2 = nadawca.Substring(maxLen, nadawca.Length - maxLen);
                    if (nadawca2.Length > maxLen)
                    {
                        nadawca2 = nadawca2.Substring(0, maxLen);
                    }
                    nadawca = nadawca.Substring(0, maxLen);
                }
                PlotText(nadawca, c01, r04, ofset0, formGfx, font);
                PlotText(nadawca2, c01, r05, ofset0, formGfx, font);

                PlotText(nip, c01, r06, ofset0, formGfx, font);
                PlotText(typIdentyfikatora, c01 + (int)(11 * ofset0), r06, ofset0, formGfx, font);
                PlotText(drugiIdentyfikator, c01 + (int)(13 * ofset0), r06, ofset0, formGfx, font);

                PlotText(identyfikatorDeklaracji, c01, r07, ofset0, formGfx, font);


                XPen pen = XPens.LightBlue.Clone();
                pen.Width = 2.5;

                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Draw the form on the page of the document in its original size
                gfx.DrawImage(form, x0, y0);
                gfx.DrawImage(form, x0, y0 + formoffset);

                MemoryStream ms = new MemoryStream();
                document.Save(ms);

                item.Attachments.AddNow(nazwaPliku, ms.GetBuffer());
                //item.SystemUpdate();

                return true;
            }

            else
            {
                return false;
            }

        }
    }
}
