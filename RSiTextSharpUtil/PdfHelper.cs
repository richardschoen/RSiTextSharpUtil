using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Drawing.Imaging;
using System.Drawing;
using System.Windows.Media.Imaging;

namespace RSiTextSharpUtil

{
    public class PDFHelper
    {
        string _lasterror = "";
        string _laststacktrace = "";

        /// <summary>
        /// Get last error if any
        /// </summary>
        /// <returns>Last error value</returns>
        public string GetLastError()
        {
            return _lasterror;
        }
        /// <summary>
        /// Get last StackTrace value if any
        /// </summary>
        /// <returns>Last stack trace value</returns>
        public string GetLastStackTrace()
        {
            return _laststacktrace;
        }

        /// <summary>
        /// Merge one or more image or document files into a single PDF
        /// Supported Formats: bmp, gif, jpg, jpeg, png, tif, tiff, pdf (including multi-page tiff and pdf files)
        /// Original code: https://www.ryadel.com/en/merge-multiple-gif-png-jpg-pdf-image-files-single-pdf-file-asp-net-c-sharp-core-itextsharp/
        /// </summary>
        public bool MergeIntoPDF(string outputPdfFile,bool replace=false, params ByteArrayInfo[] infoArray)
        {

            _lasterror = "";
            _laststacktrace = "";

            // If output file exists and not replacing, bail out. 
            if (File.Exists(outputPdfFile)){
                if (!replace)
                {
                    throw new Exception(String.Format("Output PDF file {0} already exists and replace not selected. Process cancelled.",outputPdfFile));
                }
                else
                {
                    File.Delete(outputPdfFile);
                }
            }

            // If we do have a single PDF file, write out the original file data without doing anything
            if (infoArray.Length == 1 && infoArray[0].FileExtension.Trim('.').ToLower() == "pdf")
            {
                File.WriteAllBytes(outputPdfFile,infoArray[0].Data);
                return true;
            }

            // patch to fix the "PdfReader not opened with owner password" error.
            // ref.: https://stackoverflow.com/questions/17691013/pdfreader-not-opened-with-owner-password-error-in-itext
            PdfReader.unethicalreading = true;

            using (Document doc = new Document())
            {
                doc.SetPageSize(PageSize.LETTER);

                using (var ms = new MemoryStream())
                {
                    // PdfWriter wri = PdfWriter.GetInstance(doc, ms);
                    using (PdfCopy pdf = new PdfCopy(doc, ms))
                    {
                        doc.Open();

                        foreach (ByteArrayInfo info in infoArray)
                        {
                            try
                            {
                                doc.NewPage();
                                Document imageDocument = null;
                                PdfWriter imageDocumentWriter = null;
                                switch (info.FileExtension.Trim('.').ToLower())
                                {
                                    case "bmp":
                                    case "gif":
                                    case "jpg":
                                    case "jpeg":
                                    case "png":
                                        using (imageDocument = new Document())
                                        {
                                            using (var imageMS = new MemoryStream())
                                            {
                                                using (imageDocumentWriter = PdfWriter.GetInstance(imageDocument, imageMS))
                                                {
                                                    imageDocument.Open();
                                                    if (imageDocument.NewPage())
                                                    {
                                                        var image = iTextSharp.text.Image.GetInstance(info.Data);
                                                        image.Alignment = Element.ALIGN_CENTER;
                                                        image.ScaleToFit(doc.PageSize.Width - 10, doc.PageSize.Height - 10);
                                                        if (!imageDocument.Add(image))
                                                        {
                                                            throw new Exception("Unable to add image to page!");
                                                        }
                                                        imageDocument.Close();
                                                        imageDocumentWriter.Close();
                                                        using (PdfReader imageDocumentReader = new PdfReader(imageMS.ToArray()))
                                                        {
                                                            var page = pdf.GetImportedPage(imageDocumentReader, 1);
                                                            pdf.AddPage(page);
                                                            imageDocumentReader.Close();
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                    case "tif":
                                    case "tiff":
                                        //Get the frame dimension list from the image of the file
                                        using (var imageStream = new MemoryStream(info.Data))
                                        {
                                            using (System.Drawing.Image tiffImage = System.Drawing.Image.FromStream(imageStream))
                                            {
                                                //get the globally unique identifier (GUID) 
                                                Guid objGuid = tiffImage.FrameDimensionsList[0];
                                                //create the frame dimension 
                                                FrameDimension dimension = new FrameDimension(objGuid);
                                                //Gets the total number of frames in the .tiff file 
                                                int noOfPages = tiffImage.GetFrameCount(dimension);

                                                //get the codec for tiff files
                                                ImageCodecInfo ici = null;
                                                foreach (ImageCodecInfo i in ImageCodecInfo.GetImageEncoders())
                                                    if (i.MimeType == "image/tiff")
                                                        ici = i;

                                                foreach (Guid guid in tiffImage.FrameDimensionsList)
                                                {
                                                    for (int index = 0; index < noOfPages; index++)
                                                    {
                                                        FrameDimension currentFrame = new FrameDimension(guid);
                                                        tiffImage.SelectActiveFrame(currentFrame, index);
                                                        using (MemoryStream tempImg = new MemoryStream())
                                                        {

                                                            //var encoder = new TiffBitmapEncoder();
                                                            //encoder.Compression = TiffCompressOption.Zip;
                                                            //encoder.Frames.Add(BitmapFrame.Create(image));
                                                            //encoder.Save(stream);
                                                            // Encoder parameters for CCITT G4
                                                            //EncoderParameters eps = new EncoderParameters(1);
                                                            //eps.Param[0] = new EncoderParameter(Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                                                            //tiffImage.Save(tempImg,ImageCodecInfo.GetImageEncoders( ImageFormat.Tiff,eps);
                                                            // Original code saved as TIFF before adding to PDF. PNG is smaller
                                                            //tiffImage.Save(tempImg, ImageFormat.Tiff);
                                                            
                                                            // Save image as PNG before writing to Ong file. Png format is smaller
                                                            tiffImage.Save(tempImg,ImageFormat.Png);
                                                            using (imageDocument = new Document())
                                                            {
                                                                using (var imageMS = new MemoryStream())
                                                                {
                                                                    using (imageDocumentWriter = PdfWriter.GetInstance(imageDocument, imageMS))
                                                                    {
                                                                        imageDocument.Open();
                                                                        if (imageDocument.NewPage())
                                                                        {
                                                                            var image = iTextSharp.text.Image.GetInstance(tempImg.ToArray());
                                                                            //image.CompressionLevel = Element.CCITTG4;
                                                                            // Set image DPI
                                                                            //image.SetDpi(100,100);
                                                                            image.Alignment = Element.ALIGN_CENTER;
                                                                            image.ScaleToFit(doc.PageSize.Width - 10, doc.PageSize.Height - 10);
                                                                            if (!imageDocument.Add(image))
                                                                            {
                                                                                throw new Exception("Unable to add image to page!");
                                                                            }
                                                                            imageDocument.Close();
                                                                            imageDocumentWriter.Close();
                                                                            using (PdfReader imageDocumentReader = new PdfReader(imageMS.ToArray()))
                                                                            {
                                                                                var page = pdf.GetImportedPage(imageDocumentReader, 1);
                                                                                pdf.AddPage(page);
                                                                                imageDocumentReader.Close();
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        break;
                                    case "pdf":
                                        using (var reader = new PdfReader(info.Data))
                                        {
                                            for (int i = 0; i < reader.NumberOfPages; i++)
                                            {
                                                pdf.AddPage(pdf.GetImportedPage(reader, i + 1));
                                            }
                                            pdf.FreeReader(reader);
                                            reader.Close();
                                        }
                                        break;
                                    default:
                                        // not supported image format:
                                        // skip it (or throw an exception if you prefer)
                                        throw new Exception(String.Format("Unsupported image format for file {0}",info.FileName));
                                }
                            }
                            catch (Exception e)
                            {
                                e.Data["FileName"] = info.FileName;
                                _lasterror = e.Message;
                                _laststacktrace = e.StackTrace;
                                return false;
                            }
                        }
                        if (doc.IsOpen()) doc.Close();
                        
                        //  Write output to the new consolidated PDF file
                        File.WriteAllBytes(outputPdfFile,ms.GetBuffer());
                        return true;
                    }
                }
            }
        }
    }
}

