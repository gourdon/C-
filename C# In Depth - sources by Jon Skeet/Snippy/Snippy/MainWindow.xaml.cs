using System;
using System.CodeDom.Compiler;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Reflection;

namespace Snippy
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string filename = null;
        MenuItem currentlySelectedOptions = null;

        string Filename
        {
            get { return filename; }
            set
            {
                filename = value;
                Title = filename==null ? "Snippy" : Path.GetFileName(filename)+" - Snippy";
            }
        }

        SnippetOptions CurrentOptions
        {
            get { return (SnippetOptions)currentlySelectedOptions.Tag; }
        }

        public MainWindow()
        {
            InitializeComponent();
            LoadSnippetOptions();
            Console.SetOut(new ParagraphWriter(output));
        }

        private void LoadSnippetOptions()
        {
            foreach (SnippetOptions options in SnippetOptions.LoadAvailableOptions())
            {
                MenuItem item = new MenuItem();
                item.Header = options.Name;
                item.IsCheckable = true;
                item.Tag = options;
                item.Checked += HandleOptionsSelection;
                item.IsChecked = false;
                OptionsMenu.Items.Add(item);
            }

            ((MenuItem)OptionsMenu.Items[0]).IsChecked = true;
        }

        void HandleOptionsSelection(object sender, RoutedEventArgs args)
        {
            MenuItem item = (MenuItem)args.Source;
            if (currentlySelectedOptions != null &&
                item != currentlySelectedOptions)
            {
                currentlySelectedOptions.IsChecked = false;
            }
            currentlySelectedOptions = item;
        }

        void New(object sender, RoutedEventArgs e)
        {
            Filename=null;
            codeBox.Clear();
            output.Inlines.Clear();
        }

        void Save(object sender, RoutedEventArgs e)
        {
            if (filename==null)
            {
                SaveAs(sender, e);
                return;
            }
            Snippet snippet = new Snippet(codeBox.Text, CurrentOptions);
            snippet.Save(Filename);
        }

        void SaveAs(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.DefaultExt = "scs";
            dialog.Filter="Snippets (*.scs)|*.scs|All files (*.*)|*.*";
            dialog.AddExtension = true;
            if (dialog.ShowDialog() ?? false)
            {
                Filename = dialog.FileName;
                Save(sender, e);
            }            
        }

        void Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.DefaultExt = "scs";
            dialog.Filter="Snippets (*.scs)|*.scs|All files (*.*)|*.*";
            dialog.Multiselect = false;
            if (dialog.ShowDialog() ?? false)
            {
                codeBox.Text = File.ReadAllText(dialog.FileName);
                Filename = dialog.FileName;
            }
        }

        void Export(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.DefaultExt = "cs";
            dialog.Filter="C# Files (*.cs)|*.cs|All files (*.*)|*.*";
            if (dialog.ShowDialog() ?? false)
            {
                Snippet snippet = new Snippet(codeBox.Text, CurrentOptions);
                snippet.Export(dialog.FileName);
            }
        }

        /// <summary>
        /// Performs rudimentary conversion from "normal" code. It's pretty ugly, but it suffices.
        /// </summary>
        void Import(object sender, RoutedEventArgs e)        
        {
            string[] lines = codeBox.Text.Split(new[] { "\r\n" }, StringSplitOptions.None);
            // First remove leading and trailing blank lines
            var linesList = lines.SkipWhile(x => x.Trim() == "")
                                 .Select(x => x.TrimEnd())
                                 .ToList();
            while (linesList.Last() == "")
            {
                linesList.RemoveAt(linesList.Count - 1);
            }
            // Now left-justify
            int prefixLength = linesList.Where(x => x.Trim() != "").Min(line => line.TakeWhile(c => c == ' ').Count());
            linesList = linesList.Select(line => line.Trim() == "" ? "" : line.Substring(prefixLength)).ToList();
            // Now remove [STAThread]
            linesList = linesList.Where(line => line.Trim() != "[STAThread]").ToList();
            // Now find Main() and change it accordingly
            int startOfMain = linesList.IndexOf("static void Main()");
            if (startOfMain != -1 && startOfMain < linesList.Count - 2)
            {
                linesList.RemoveAt(startOfMain);
                linesList[startOfMain] = "...";
                if (linesList.Last() == "}")
                {
                    linesList.RemoveAt(linesList.Count - 1);
                }
                for (int i = startOfMain + 1; i < linesList.Count; i++)
                {
                    if (linesList[i].StartsWith("    "))
                    {
                        linesList[i] = linesList[i].Substring(4);
                    }
                }
            }
            codeBox.Text = string.Join("\r\n", linesList);
        }

        void CompileAndRun(object sender, RoutedEventArgs e)
        {
            Snippet snippet = new Snippet(codeBox.Text, CurrentOptions);
            CompilerResults results = snippet.Compile();

            output.Inlines.Clear();

            foreach (CompilerError error in results.Errors)
            {
                Run run = new Run("Line " + error.Line + ": " + error.ErrorText);
                run.Foreground = error.IsWarning ? Brushes.MediumBlue : Brushes.Red;
                output.Inlines.Add(run);
                output.Inlines.Add(new LineBreak());
            }

            if (results.Errors.HasErrors)
            {
                Run run = new Run("(Build failed)");
                run.Foreground = Brushes.Red;
                output.Inlines.Add(run);
                return;
            }
            else
            {
                Run run = new Run("(Build successful)");
                run.Foreground = Brushes.MediumBlue;
                output.Inlines.Add(run);
                output.Inlines.Add(new LineBreak());
                Type snippetType = results.CompiledAssembly.GetType("Snippet");
                try
                {
                    snippetType.GetMethod("Main").Invoke(null, new object[] { new string[] { } });
                }
                catch (TargetInvocationException ex)
                {
                    Run error = new Run(string.Format("Exception: {0}", ex.InnerException));
                    error.Foreground = Brushes.Red;
                    output.Inlines.Add(error);
                }
            }
        }
    }
}
