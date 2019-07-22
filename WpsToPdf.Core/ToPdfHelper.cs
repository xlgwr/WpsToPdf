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
        public string typeName { get; private set; }
        public string fileSource { get; private set; }
        public ToPdfHelper(string fileSource)
        {
            var typeName = Path.GetExtension(fileSource);

            this.typeName = typeName;
            this.fileSource = fileSource;

            switch (typeName)
            {
                case ".xls":
                case ".xlsx":
                    typeName = "KET.Application";
                    break;
                case ".ppt":
                case ".pptx":
                    typeName = "KWPP.Application";
                    break;
                default:
                    typeName = "KWps.Application";
                    break;
            }
            //创建wps实例，需提前安装wps
            Type type = Type.GetTypeFromProgID(typeName);

            if (type == null)
                type = Type.GetTypeFromProgID("wps.Application");

            wps = Activator.CreateInstance(type);
        }
        public string SavePdf(string wpsFilename)
        {
            string result = "";
            switch (typeName)
            {
                case ".xls":
                case ".xlsx":
                    result = XlsWpsToPdf(fileSource, wpsFilename);
                    break;
                case ".ppt":
                case ".pptx":
                    result = PPTWpsToPdf(fileSource, wpsFilename);
                    break;
                default:
                    result = WordWpsToPdf(fileSource, wpsFilename);
                    break;
            }
            return result;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveUrl"></param>
        /// <param name="wpsFilename"></param>
        /// <returns></returns>
        private string XlsWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            string path = Path.GetDirectoryName(fileSource);
            var pdfSavePath = Path.Combine(path, string.Format("{0}_{1}.pdf", wpsFilename, Guid.NewGuid().ToString()));
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
            return pdfSavePath;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="saveUrl"></param>
        /// <param name="wpsFilename"></param>
        /// <returns></returns>
        private string WordWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            string path = Path.GetDirectoryName(fileSource);
            var pdfSavePath = Path.Combine(path, string.Format("{0}_{1}.pdf", wpsFilename, Guid.NewGuid().ToString()));

            try
            {
                //用wps 打开word不显示界面
                dynamic doc = wps.Documents.Open(fileSource, Visible: false);
                //doc 转pdf
                doc.ExportAsFixedFormat(pdfSavePath, WdExportFormat.wdExportFormatPDF);

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
            return pdfSavePath;
        }
        private string PPTWpsToPdf(string fileSource, string wpsFilename)
        {
            if (wpsFilename == null)
            {
                throw new ArgumentNullException("wpsFilename");

            }
            string path = Path.GetDirectoryName(fileSource);
            var pdfSavePath = Path.Combine(path, string.Format("{0}_{1}.pdf", wpsFilename, Guid.NewGuid().ToString()));
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

                doc.SaveAs(pdfSavePath, PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);

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
            return pdfSavePath;
        }
        public void Dispose()
        {
            if (wps != null) { wps.Quit(); wps = null; }
        }
    }
}
