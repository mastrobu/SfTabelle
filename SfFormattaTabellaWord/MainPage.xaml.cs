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

namespace SfFormattaTabellaWord
{
    /// <summary>
    /// The following code example illustrates how to load an existing document 
    /// and apply table formatting options such as Borders, LeftIndent, Paddings, 
    /// IsAutoResize, etc.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            FormattaTabellaWordAsync();
        }

        private async void FormattaTabellaWordAsync()
        {
            //Creates an instance of WordDocument class (Empty Word Document)

            WordDocument document = new WordDocument();

            //Opens an existing Word document into DocIO instance
            //document.Open("Table.docx", FormatType.Docx);
            StorageFolder local = Windows.Storage.ApplicationData.Current.LocalFolder;
            StorageFile storageFile = await local.GetFileAsync("Table.docx");
            var streamFile = await storageFile.OpenStreamForReadAsync();
            document.Open(streamFile, FormatType.Docx);

            //Accede all'istanza della prima section del documento Word
            WSection section = document.Sections[0];

            //Accede all'istanza della prima tabella nella section
            WTable table = section.Tables[0] as WTable;

            //Specifica il titolo della tabella
            table.Title = "PriceDetails";

            //Specifica la descrizione della tabella
            table.Description = "This table shows the price details of various fruits";

            //Specifica l'indentazione della tabella
            table.IndentFromLeft = 50;

            //Specifica il colore di background della tabella
            table.TableFormat.BackColor = Color.FromArgb(192, 192, 192);

            //Specifica l'allineamento orizzontale della tabella
            table.TableFormat.HorizontalAlignment = RowAlignment.Left;

            //Specifica il padding left, right, top and bottom di tutte le celle della tabella
            table.TableFormat.Paddings.All = 10;

            //Specifica l'auto resize della tabella per ridimensionare automaticamente tutte le celle sulla base del loro contenuto
            table.TableFormat.IsAutoResized = true;

            //Specifica la dimensione riga del border top, bottom, left and right della tabella
            table.TableFormat.Borders.LineWidth = 2f;

            //Specifica la dimensione riga del border orizzontale
            table.TableFormat.Borders.Horizontal.LineWidth = 2f;

            //Specifica la dimensione riga del border verticale
            table.TableFormat.Borders.Vertical.LineWidth = 2f;

            //Specifica il top, bottom, left and right border color delle tabelle
            table.TableFormat.Borders.Color = Color.Red;

            //Specifica il border color orizzontale della tabella
            table.TableFormat.Borders.Horizontal.Color = Color.Red;

            //Specifica il border color vericale della tabella
            table.TableFormat.Borders.Vertical.Color = Color.Red;

            //Specifica il tipo di border della tabella
            table.TableFormat.Borders.BorderType = BorderStyle.Double;

            //Accede all'istanza della prima riga della tabella
            WTableRow row = table.Rows[0];

            //Specifica l'altezza della riga
            row.Height = 20;

            //Specifica il tipo di altezza della riga
            row.HeightType = TableRowHeightType.AtLeast;

            //Salva il documento su memory stream
            MemoryStream stream = new MemoryStream();
            await document.SaveAsync(stream, FormatType.Docx);

            //Libera le risorse impegnate dall'istanza WordDocument
            document.Close();

            //Salva lo stream come file di documento Word nella macchina locale
            StorageFile stFile = await Save(stream, "TableFormatting.docx");

            DefaultLaunch("TableFormatting.docx");
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
