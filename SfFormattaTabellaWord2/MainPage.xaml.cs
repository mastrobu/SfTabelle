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

namespace SfFormattaTabellaWord2
{
    /// <summary>
    /// Il seguente esempio di codice illustra come caricare un documento esistente 
    /// e applicare opzioni di formattazione cella quali VerticalAlignment, TextDirection, 
    /// Paddings, Borders, etc.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            FormattaTabellaWord2Async();
        }

        private async void FormattaTabellaWord2Async()
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

            //Accede all'istanza della prima riga della tabella
            WTableRow row = table.Rows[0];

            //Specifica l'altezza della riga
            row.Height = 20;

            //Specifica il tipo di atezza riga
            row.HeightType = TableRowHeightType.AtLeast;

            //Accede all'istanza della prima cella della riga
            WTableCell cell = row.Cells[0];

            //Specifica il back ground color della cella
            cell.CellFormat.BackColor = Color.FromArgb(192, 192, 192);

            //Specifica lo stesso padding della tabella come false per preservare il cell padding corrente
            cell.CellFormat.SamePaddingsAsTable = false;

            //Specifica il left, right, top e bottom padding della cella
            cell.CellFormat.Paddings.Left = 5;
            cell.CellFormat.Paddings.Right = 5;
            cell.CellFormat.Paddings.Top = 5;
            cell.CellFormat.Paddings.Bottom = 5;

            //Specifica l'allineamento verticale del contenuto del testo
            cell.CellFormat.VerticalAlignment = Syncfusion.DocIO.DLS.VerticalAlignment.Middle;

            //Accede all'istanza della seconda cella della riga
            cell = row.Cells[1];

            cell.CellFormat.BackColor = Color.FromArgb(192, 192, 192);

            cell.CellFormat.SamePaddingsAsTable = false;

            //Specifica il left, right, top e bottom padding della cella
            cell.CellFormat.Paddings.All = 5;

            cell.CellFormat.VerticalAlignment = Syncfusion.DocIO.DLS.VerticalAlignment.Middle;

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "TableCellFormatting.docx");

            DefaultLaunch("TableCellFormatting.docx");
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

