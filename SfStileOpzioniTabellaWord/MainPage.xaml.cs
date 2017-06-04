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

namespace SfStileOpzioniTabellaWord
{
    /// <summary>
    /// Dopo che hai applicato uno stile alla tabella, puoi abilitare o disabilitare 
    /// la speciale formattazione della tabella. Esistono sei opzioni: prima colonna, 
    /// ultima colonna, banded rows, 
    /// banded columns, header row and last row.
    /// 
    /// The following code example illustrates how to enable and disable the special table 
    /// formatting options of the table styles
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            StileOpzioniTabellaWordAsync();
        }

        private async void StileOpzioniTabellaWordAsync()
        {
            //Crea una istanza della classe WordDocument
            WordDocument document = new WordDocument();
            //Aggiunge una sezione ad un documento word
            IWSection sectionFirst = document.AddSection();
            //Aggiunge una tabella ad un documento word
            IWTable tableFirst = sectionFirst.AddTable();
            //Dimensiona la tabella
            tableFirst.ResetCells(3, 2);
            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            document.Open(stream, FormatType.Docx);

            WSection section = document.Sections[0];

            WTable table = section.Tables[0] as WTable;

            //Applica alla tabella lo stile built-in "LightShading"
            table.ApplyStyle(BuiltinTableStyle.LightShading);

            //Abilita una formattazione speciale per le banded columns della tabella 
            table.ApplyStyleForBandedColumns = true;

            //Abilita una formattazione speciale per le banded rows of the table
            table.ApplyStyleForBandedRows = true;

            //Disabilita la formattazione speciale per la prima colonna della tabella
            table.ApplyStyleForFirstColumn = false;

            //Abilita una formattazione speciale per la riga di testata della tabella
            table.ApplyStyleForHeaderRow = true;

            //Abilita una formattazione speciale per l'ultima colonna della tabella
            table.ApplyStyleForLastColumn = true;

            //Disabilita la formattazione speciale per l'ultima riga della tabella
            table.ApplyStyleForLastRow = false;

            //Salva e chiudi l'istanza del documento
            //Salva il documento su memory stream
            MemoryStream memoryStream = new MemoryStream();
            await document.SaveAsync(memoryStream, FormatType.Docx);

            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(memoryStream, "TableStyle.docx");
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

