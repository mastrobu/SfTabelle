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

namespace SfMergeCellsVertTabellaWord
{
    /// <summary>
    /// Puoi unire due o più celle della tabella in una singola cella situate nella stessa riga
    /// o colonna.
    ///  
    /// Il seguente codice di esempio illustra come creare una tabella contenente 
    /// merged cells verticali.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            MergeCellsVertTabellaWord();
        }

        private async void MergeCellsVertTabellaWord()
        {
            //Crea una istanza della classe WordDocument

            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();

            section.AddParagraph().AppendText("Vertical merging of Table cells");

            IWTable table = section.AddTable();

            table.ResetCells(2, 2);

            //Aggiunge contenuto alle celle della tabella

            table[0, 0].AddParagraph().AppendText("First row, First cell");

            table[0, 1].AddParagraph().AppendText("First row, Second cell");

            table[1, 0].AddParagraph().AppendText("Second row, First cell");

            table[1, 1].AddParagraph().AppendText("Second row, Second cell");

            //Specifica che il vertical merge inizia dalla prima cella della prima riga

            table[0, 0].CellFormat.VerticalMerge = CellMerge.Start;

            //Modifica il contenuto della cella

            table[0, 0].Paragraphs[0].Text = "Vertically merged cell";

            //Specifica che il vertical merge continua sulla prima cela della seconda riga

            table[1, 0].CellFormat.VerticalMerge = CellMerge.Continue;

            ////Salva e chiudi l'istanza del documento
            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "VerticalMerge.docx");

            DefaultLaunch("VerticalMerge.docx");
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
