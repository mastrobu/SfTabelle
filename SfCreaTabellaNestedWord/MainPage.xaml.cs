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

namespace SfCreaTabellaNestedWord
{
    /// <summary>
    /// Puoi creare una nested table aggiungendo una nuova tabella dentro una
    /// cella. 
    /// Il codice che segue illustra come aggiungere una tabella ad una cella.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            CreaTabellaNestedWordAsync();
        }

        private async void CreaTabellaNestedWordAsync()
        {
            //Crea una istanza della classe WordDocument
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();

            section.AddParagraph().AppendText("Price Details");

            IWTable table = section.AddTable();

            table.ResetCells(3, 2);

            table[0, 0].AddParagraph().AppendText("Item");

            table[0, 1].AddParagraph().AppendText("Price($)");

            table[1, 0].AddParagraph().AppendText("Items with same price");

            //Aggiunge una nested table alla cella (seconda riga, prima cella).
            IWTable nestTable = table[1, 0].AddTable();

            //Crea il numero specificato di righe e colonne per la nested table
            nestTable.ResetCells(3, 1);

            //Accede all'istanza della cella della nested table (prima riga, prima cella)
            WTableCell nestedCell = nestTable.Rows[0].Cells[0];

            //Specifica la larghezza della nested cell
            nestedCell.Width = 200;

            //Aggiunge il contenuto alla nested cell
            nestedCell.AddParagraph().AppendText("Apple");

            //Accede all'istanza della cella della nested table (seconda riga, prima cella)
            nestedCell = nestTable.Rows[1].Cells[0];

            //Specifica la larghezza della nested cell
            nestedCell.Width = 200;

            //Aggiunge il contenuto alla nested cell
            nestedCell.AddParagraph().AppendText("Orange");

            //Accede all'istanza della cella della nested table (terza riga, prima cella)
            nestedCell = nestTable.Rows[2].Cells[0];

            //Specifica la larghezza della nested cell
            nestedCell.Width = 200;

            //Aggiunge il contenuto alla nested cell
            nestedCell.AddParagraph().AppendText("Mango");

            //Accede all'istanza della cella della nested table (seconda riga, seconda cella)
            nestedCell = table.Rows[1].Cells[1];

            table[1, 1].AddParagraph().AppendText("85");

            table[2, 0].AddParagraph().AppendText("Pomegranate");

            table[2, 1].AddParagraph().AppendText("70");

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "NestedTable.docx");

            DefaultLaunch("NestedTable.docx");
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
