using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace SfStileTabellaWord
{
    /// <summary>
    /// Un table style definisce un insieme di formattazioni a livello di tabelle, righe, 
    /// celle e paragrafi che possono essere applicati ad una tabella. L'istanza WTableStyle 
    /// rappresenta lo stile di ua tabella in un documento Word.
    /// Nota:
    ///     Essential DocIO attualmente fornisce supporto per table styles soltanto nei
    ///     formati DOCX e WordML. DocIO può conservare i table styles sia quelli built-in 
    ///     che quelli personalizzati quando si aprono e si salvano formati DOCX, WordML.
    ///     Viene inoltre preservata la visualizzazione nelle conversioni da Word a PDF, 
    ///     da Word a Image, e da Word a HTML.
    /// Il codice di esempio che segue illustra come applicare i built-in table styles alla 
    /// tabella.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            StileTabellaWordAsync();
        }

        private async void StileTabellaWordAsync()
        {
            //Creates an instance of WordDocument class
            WordDocument document = new WordDocument();

            //Apre un documento word esistente nella istanza DocIO
            //document.Open("Table.docx", FormatType.Docx);
            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
            StorageFile storageFile;
            try
            {
                storageFile = await local.GetFileAsync("Table.docx");
            }
            catch (Exception)
            {
                return;
            }
            var streamFile = await storageFile.OpenStreamForReadAsync();
            document.Open(streamFile, FormatType.Docx);

            WSection section = document.Sections[0];

            WTable table = section.Tables[0] as WTable;

            //Applies "LightShading" built-in style to table

            table.ApplyStyle(BuiltinTableStyle.LightShading);

            //Saves and closes the document instance 

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "TableStyle.docx");

            DefaultLaunch("TableStyle.docx");
        }

        async void DefaultLaunch(string stFile)
        {

            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;

            var file = await local.GetFileAsync(stFile);


            if (file != null)
            {
                // Launch the retrieved file
                var success = await Windows.System.Launcher.LaunchFileAsync(file);

                if (success)
                {
                    // File launched
                }
                else
                {
                    // File launch failed
                }
            }
            else
            {
                // Could not find file
            }
        }

        async Task<StorageFile> Save(MemoryStream stream, string filename)
        {

            stream.Position = 0;

            StorageFile stFile;

            //if (!(Windows.Foundation.Metadata.ApiInformation.IsTypePresent("Windows.Phone.UI.Input.HardwareButtons")))

            //{

            //    FileSavePicker savePicker = new FileSavePicker();

            //    savePicker.DefaultFileExtension = ".docx";

            //    savePicker.SuggestedFileName = filename;

            //    savePicker.FileTypeChoices.Add("Word Documents", new List<string>() { ".docx" });

            //    stFile = await savePicker.PickSaveFileAsync();

            //}

            //else

            //{

            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;

            stFile = await local.CreateFileAsync(filename, CreationCollisionOption.ReplaceExisting);

            //}
            if (stFile != null)

            {

                using (IRandomAccessStream zipStream = await stFile.OpenAsync(FileAccessMode.ReadWrite))

                {

                    // Write compressed data from memory to file

                    using (Stream outstream = zipStream.AsStreamForWrite())

                    {

                        byte[] buffer = stream.ToArray();

                        outstream.Write(buffer, 0, buffer.Length);

                        outstream.Flush();

                    }

                }

            }
            return stFile;
        }
    }
}
