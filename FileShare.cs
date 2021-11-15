using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure; // Namespace for Azure Configuration Manager
using Microsoft.Office.Interop.Word;
using Microsoft.WindowsAzure.Storage; // Namespace for Storage Client Library
using Microsoft.WindowsAzure.Storage.File; // Namespace for Azure Files

namespace Microsoft.BotBuilderSamples
{
    public static class FileShare
    {
        private static string Key = "Ft16G8GZpjyTn8pDSaOyMfuDBJOU9f60yhHZ2NoHaWj8s/QAT+bNF8aXM8JMKbcZI1vYPU1R4whEA6bzlkktfg==";
        private static string StrgeAccName = "560d0c765857dsvm";
        private static string StorageConnectionString = "DefaultEndpointsProtocol=https;AccountName=" + StrgeAccName + ";AccountKey=" + Key;
        public static string fileName = "Edison.txt";
        private static string shareName = "fileshare";
        private static CloudFileShare fileShare;
        private static CloudFileClient fileClient;
        private static CloudFileDirectory rootDir;

        public static void Connect()
        {
            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(StorageConnectionString);

            // Create a CloudFileClient object for credentialed access to Azure Files.
            fileClient = storageAccount.CreateCloudFileClient();
            // Get a reference to the file share we created previously.

            // Get a reference to the root directory for the share.
            rootDir = fileClient.GetShareReference(shareName).GetRootDirectoryReference();
        }

        public static List<string> GetFilesList()
        {
            var fileList = rootDir.ListFilesAndDirectoriesSegmentedAsync(null);
            List<string> fileNames = fileList.Result.Results.OfType<CloudFile>().Select(b => b.Name).ToList();

            return fileNames;
        }

        public static string Retrieve()
        {
            
                // Get a reference to the file we created previously.
            CloudFile file = rootDir.GetFileReference(fileName);

            switch (file.Name.Substring(file.Name.LastIndexOf(".") + 1))
            {
                case "docx": return DocfileReader(file);

                case "txt": return TextfileReader(file);

                case "pdf": return PDFfileReader(file);
            }

            return "Empty File";
        }

        private static string DocfileReader(CloudFile file)
        {
            string value = "";
            Application application = new Application();
            Document document = application.Documents.Open(file);
            int count = document.Words.Count;
            for (int i = 1; i <= count; i++)
            {
                // Write the word.
                string text = document.Words[i].Text;
                value = value + text;
                //Console.WriteLine("Word {0} = {1}", i, text);
            }
            // Close word.
            application.Quit();

            return value;
        }
        private static string TextfileReader(CloudFile file)
        {
            Stream InputStream = file.OpenReadAsync().Result;
            using (var reader = new StreamReader(InputStream, Encoding.UTF8))
            {
                string value = reader.ReadToEnd();
                // Do something with the value
                return value;
            }
        }
        private static string PDFfileReader(CloudFile file)
        {
            throw new NotImplementedException();
        }
    }
}
