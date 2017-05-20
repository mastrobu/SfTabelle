using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.Storage.Streams;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace SfCreaTabellaWord
{
    /// <summary>
    /// Il codice che segue illustra come creare una semplice tabella 
    /// con un numero predefinito di righe e celle.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();

            CreaTabellaWordAsync();
        }

        private async void CreaTabellaWordAsync()
        {

            WordDocument document = new WordDocument();

            //Aggiunge una sezione ad un documento word
            IWSection section = document.AddSection();

            //Aggiunge un nuovo paragrafo ad un documento word e aggiunge testo al paragrafo
            IWTextRange textRange = section.AddParagraph().AppendText("Price Details");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;
            section.AddParagraph();

            //Aggiunge una nuova tabella al documento Word
            IWTable table = section.AddTable();

            //Specifica il numero totale di righe e colonne
            table.ResetCells(3, 2);

            //Accede all'istanza della cella (prima riga, prima cella) e aggiunge il contenuto alla cella
            textRange = table[0, 0].AddParagraph().AppendText("Item");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.Bold = true;

            //Accede all'istanza della cella (prima riga, seconda cella) e aggiunge il contenuto alla cella
            textRange = table[0, 1].AddParagraph().AppendText("Price($)");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 12;
            textRange.CharacterFormat.Bold = true;

            //Accede all'istanza della cella (seconda riga, prima cella) e aggiunge il contenuto alla cella
            textRange = table[1, 0].AddParagraph().AppendText("Apple");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            //Accede all'istanza della cella (seconda riga, seconda cella) e aggiunge il contenuto alla cella
            textRange = table[1, 1].AddParagraph().AppendText("50");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            //Accede all'istanza della cella (terza riga, prima cella) e aggiunge il contenuto alla cella
            textRange = table[2, 0].AddParagraph().AppendText("Orange");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            //Accede all'istanza della cella (terza riga, seconda cella) e aggiunge il contenuto alla cella
            textRange = table[2, 1].AddParagraph().AppendText("30");
            textRange.CharacterFormat.FontName = "Arial";
            textRange.CharacterFormat.FontSize = 10;

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);
            
            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "Table.docx");

            DefaultLaunch("Table.docx");


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
