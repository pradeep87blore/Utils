using System;
using System.IO;
using System.Threading;
using System.Windows;


// Source for Word file merge logic: https://stackoverflow.com/questions/35982526/merging-word-documents-in-folder-using-c-sharp 

namespace WordMerger
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] fileList = null;
        string outputFilePath = "";
        bool insertPageBreaks = true; // Insert breaks between docs
        bool bSuccessfulOperation = true;

        int pageCount = 0, wordCount = 0, lineCount = 0;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Initialized(object sender, EventArgs e)
        {
            DisableControlsTillSuccess(false);
        }

        public void Merge() // Shall be called on a separate thread and shall make use of the member variables
        {
            object missing = System.Type.Missing;
            object pageBreak = Microsoft.Office.Interop.Word.WdBreakType.wdSectionBreakNextPage;
            object outputFile = outputFilePath;

            // Create a new Word application
            Microsoft.Office.Interop.Word._Application wordApplication = new Microsoft.Office.Interop.Word.Application();

            bSuccessfulOperation = true;
            try
            {
                // Create a new file
                Microsoft.Office.Interop.Word.Document wordDocument = wordApplication.Documents.Add(
                                              ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                // Make a Word selection object.
                Microsoft.Office.Interop.Word.Selection selection = wordApplication.Selection;

                //Count the number of documents to insert;
                int documentCount = fileList.Length;

                //A counter that signals that we shoudn't insert a page break at the end of document.
                int breakStop = 0;

                int iCounter = 0;
                int iFileCount = fileList.Length;

                // Iterate through the entire list of files
                foreach (string file in fileList)
                {
                    Console.WriteLine("Merging file: " + file);
                    breakStop++;
                    // Insert the files to our template
                    selection.InsertFile(
                                                file
                                            , ref missing
                                            , ref missing
                                            , ref missing
                                            , ref missing);

                    //Do we want page breaks added after each documents?
                    if (insertPageBreaks && breakStop != documentCount)
                    {
                        selection.InsertBreak(ref pageBreak);
                    }
                    iCounter++;
                    UpdateProgressBar(((float)iCounter / iFileCount) * 100);
                }

                // Save the document to it's output file.
                wordDocument.SaveAs(
                                ref outputFile
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing
                            , ref missing);

                // Get some stats about the merged file
                lineCount = wordDocument.Sentences.Count;
                wordCount = wordDocument.Words.Count;

                Microsoft.Office.Interop.Word.WdStatistic stat = Microsoft.Office.Interop.Word.WdStatistic.wdStatisticPages;
                pageCount = wordDocument.ComputeStatistics(stat, ref missing);
                wordDocument.Save();

                // Clean up!
                wordDocument = null;
            }
            catch (Exception ex)
            {
                //I didn't include a default error handler so i'm just throwing the error
                Console.WriteLine(ex);
                bSuccessfulOperation = false;
            }
            finally
            {
                // Finally, Close our Word application
                wordApplication.Quit(ref missing, ref missing, ref missing);
            }


            DisplayResult(); // Done, show the results
        }
    

    void UpdateProgressBar(float value)
        {
            Dispatcher.Invoke(() =>
            {
                progressbar_Progress.Value = value;
            });
            
        }

    void DisableControlsTillSuccess(bool bDisable)
        {
            textBlock_outputFileDetails.IsEnabled = bDisable;
            button_openMergedFile.IsEnabled = bDisable;
        }

    private void Button_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                DisableControlsTillSuccess(false); 

                textBlock_opStatus.Text = "";

                fileList = Directory.GetFiles(dialog.SelectedPath,  "*.docx");

                foreach(var file in fileList)
                {
                    listBox_fileList.Items.Add(file);
                }

                outputFilePath = (dialog.SelectedPath + "\\MergedFile.docx");

            }
        }

        private void Button1_Click(object sender, RoutedEventArgs e)
        {
            string[] documentsToMerge = fileList;

            textBlock_opStatus.Text = "Converting...";

            Thread worker = new Thread(new ThreadStart(Merge));
            worker.Start();
        }

        private void Button_openMergedFile_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(outputFilePath); // Opens the merged file
        }

        private void DisplayResult()
        {
            Dispatcher.Invoke(() =>
            {
                if (bSuccessfulOperation)
                {
                    System.Windows.MessageBox.Show("Completed merging the file successfully");
                    textBlock_opStatus.Text = "Success";
                    DisableControlsTillSuccess(true);

                    FillStatistics();
                }
                else
                {
                    System.Windows.MessageBox.Show("Failed the merge operation");
                    textBlock_opStatus.Text = "Failed to merge all files";
                }
            });
            
        }

        private void FillStatistics()
        {
            textBlock_outputFileDetails.Text =
                string.Format("Page count: {0}, Line Count: {1}, Word Count: {2}", pageCount, lineCount, wordCount);
        }
    }
}

