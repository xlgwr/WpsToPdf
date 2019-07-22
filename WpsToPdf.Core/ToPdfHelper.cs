using Excel;
using PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Word;

namespace ConsoleApp1
{
    public class ToPdfHelper : IDisposable
    {
        dynamic wps;
        public ToPdfHelper(string typeName)
        {
            if (typeName == "xls" || typeName == "xlsx")
            {
                typeName = "KET.Application";
            }
            else if (typeName == "ppt" || typeName == "pptx")
                typeName = "KWPP.Application";
            else
                typeName = "KWps.Application";

            //创建wps实例，需提前安装wps
            Type type = Type.GetTypeFromProgID(typeName);

            if (type == null)
                type = Type.GetTypeFromProgID("wps.Application");

            wps = Activator.CreateInstance(type);
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveUrl"></param>
        /// <param name="wpsFilename"></param>
        /// <returns></returns>
        public string XlsWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            var pdfSavePath = Path.ChangeExtension(fileSource, "pdf");
            try
            {
                XlFixedFormatType targetType = XlFixedFormatType.xlTypePDF;
                object missing = Type.Missing;

                //xls 转pdf
                dynamic doc = wps.Application.Workbooks.Open(fileSource, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                doc.ExportAsFixedFormat(targetType, pdfSavePath, XlFixedFormatQuality.xlQualityStandard, true, false, missing, missing, missing, missing);
                //设置隐藏菜单栏和工具栏
                //wps.setViewerPreferences(PdfWriter.HideMenubar | PdfWriter.HideToolbar);
                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                //targetPath = GetEXCELtoPDF.CreatePDFs(saveUrl, targetPath);
                throw;
            }
            finally
            {
                Dispose();
            }
            return Path.ChangeExtension(wpsFilename, "pdf");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveUrl"></param>
        /// <param name="wpsFilename"></param>
        /// <returns></returns>
        public string WordWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            var pdfPath = Path.ChangeExtension(fileSource, "pdf");
            try
            {
                //用wps 打开word不显示界面
                dynamic doc = wps.Documents.Open(fileSource, Visible: false);
                //doc 转pdf
                doc.ExportAsFixedFormat(pdfPath, WdExportFormat.wdExportFormatPDF);

                //设置隐藏菜单栏和工具栏
                //wps.setViewerPreferences(PdfWriter.HideMenubar | PdfWriter.HideToolbar);


                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                //targetPath = GetEXCELtoPDF.CreatePDFs(saveUrl, targetPath);
                throw;
            }
            finally
            {
                Dispose();
            }
            return Path.ChangeExtension(wpsFilename, "pdf");
        }
        public string PPTWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            var pdfPath = Path.ChangeExtension(fileSource, "pdf");
            try
            {

                //ppt 转pdf                
                dynamic doc = wps.Presentations.Open(fileSource, MsoTriState.msoCTrue, MsoTriState.msoCTrue, MsoTriState.msoCTrue);

                object missing = Type.Missing;

                //doc.ExportAsFixedFormat(pdfPath, PpFixedFormatType.ppFixedFormatTypePDF,
                //    PpFixedFormatIntent.ppFixedFormatIntentPrint,
                //    MsoTriState.msoCTrue, PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                //    PpPrintOutputType.ppPrintOutputBuildSlides,
                //      MsoTriState.msoCTrue, null, PpPrintRangeType.ppPrintAll,"",
                //      false, false, false, false, false, missing);

                doc.SaveAs(pdfPath, PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);

                //设置隐藏菜单栏和工具栏
                //wps.setViewerPreferences(PdfWriter.HideMenubar | PdfWriter.HideToolbar);


                doc.Close();
                doc = null;
            }
            catch (Exception ex)
            {
                //targetPath = GetEXCELtoPDF.CreatePDFs(saveUrl, targetPath);
                throw;
            }
            finally
            {
                Dispose();
            }
            return Path.ChangeExtension(wpsFilename, "pdf");
        }


        public void Dispose()
        {
            if (wps != null) { wps.Quit(); wps = null; }
        }
    }
}
