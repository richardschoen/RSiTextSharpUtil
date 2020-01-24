namespace RSiTextSharpUtil
{
    /// <summary>
    /// POCO class to store byte[] and other useful informations regarding the data.
    /// </summary>
    public class ByteArrayInfo
    {
        public ByteArrayInfo(byte[] fileData, string fileName)
        {
            Data = fileData;
            FileName = fileName;
            FileExtension = System.IO.Path.GetExtension(FileName).ToLower();
        }

        public byte[] Data { get; set; }

        /// <summary>
        /// The File Name (es. "TestFile.pdf")
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// The File Extension, including the dot (es. ".pdf")
        /// </summary>
        public string FileExtension { get; set; }
    }
}
