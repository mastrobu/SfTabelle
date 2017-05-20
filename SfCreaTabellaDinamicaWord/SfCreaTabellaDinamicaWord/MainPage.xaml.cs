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

namespace SfCreaTabellaDinamicaWord
{
    /// <summary>
    /// Il codice che segue illustra come creare una semplice tabella aggiungendo
    /// dinamicamente righe.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            CreaTabellaDinamicaWordAsync();
        }

        private async void CreaTabellaDinamicaWordAsync()
        {
            //Crea una istanza della classe WordDocument
            WordDocument document = new WordDocument();

            IWSection section = document.AddSection();

            section.AddParagraph().AppendText("Price Details");

            section.AddParagraph();

            //Aggiunge una nuova tabella al dovumento Word
            IWTable table = section.AddTable();

            //Aggiunge la prima riga alla tabella
            WTableRow row = table.AddRow();

            //Aggiunge la prima cella nella prima riga 
            WTableCell cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Item");

            //Aggiunge la seconda cella nella prima riga 
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Price($)");

            //Aggiunge la seconda riga alla tabella
            row = table.AddRow(true, false);

            //Aggiunge la prima cella nella seconda riga 
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Apple");

            //Aggiunge la seconda cella nella seconda riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("50");

            //Aggiunge la terza riga alla tabella
            row = table.AddRow(true, false);

            //Aggiunge la prima cella nella terza riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Orange");

            //Aggiunge la seconda cella nella terza riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("30");

            //Aggiunge la quarta riga alla tabella
            row = table.AddRow(true, false);

            //Aggiunge la prima cella nella quarta riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Banana");

            //Aggiunge la seconda cella nella quarta riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("20");

            //Aggiunge la quinta riga alla tabella
            row = table.AddRow(true, false);

            //Aggiunge la prima cella nella quinta riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("Grapes");

            //Aggiunge la seconda cella nella quinta riga
            cell = row.AddCell();

            //Specifica la larghezza della cella
            cell.Width = 200;

            cell.AddParagraph().AppendText("70");

            ////Saves and closes the document instance
            //document.Save("Table.docx", FormatType.Docx);

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

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
