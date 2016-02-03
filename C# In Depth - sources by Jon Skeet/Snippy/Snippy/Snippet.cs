using System.CodeDom.Compiler;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Microsoft.CSharp;

namespace Snippy
{
    /// <summary>
    /// Represents a snippet of code - typically the text within the text box
    /// of the UI
    /// </summary>
    public sealed class Snippet
    {
        const string ClassPrologue="\r\npublic class Snippet\r\n{\r\n";
        const string MainPrologue="    [STAThread]\r\n    public static void Main(string[] args)\r\n    {\r\n";
        const string ClassEpilogue = "    }\r\n}\r\n";
        const string UsingFormat = "using {0};\r\n";
        const string LinePragmaFormat = "#line {0}\r\n";
        const string LineFormat = "        {0}\r\n";

        string snippetText;
        SnippetOptions options;

        /// <summary>
        /// Generates a snippet from the given text, with the given options.
        /// </summary>
        public Snippet(string snippetText, SnippetOptions options)
        {
            this.snippetText = snippetText;
            this.options = options;
        }

        string GenerateCode(bool includeSnippetLineNumbers)
        {
            string nonMain = "";
            StringBuilder builder = new StringBuilder();
            foreach (string ns in options.Namespaces)
            {
                builder.AppendFormat(UsingFormat, ns);
            }

            string namespaces = builder.ToString();
            builder = new StringBuilder();
            StringReader reader = new StringReader(snippetText);
            string line;
            int lineNumber = 1;
            while ((line = reader.ReadLine()) != null)
            {
                if (includeSnippetLineNumbers && lineNumber == 1)
                {
                    builder.AppendFormat(LinePragmaFormat, 1);
                }
                // If we see "..." for the first time
                if (line.Trim() == "..." && nonMain == "")
                {
                    nonMain = builder.ToString();
                    builder = new StringBuilder();
                    if (includeSnippetLineNumbers)
                    {
                        lineNumber++;
                        builder.AppendFormat(LinePragmaFormat, lineNumber);
                    }
                }
                else
                {
                    builder.AppendFormat(LineFormat, line);
                    lineNumber++;
                }
            }

            return namespaces + ClassPrologue + nonMain + MainPrologue + builder.ToString() + ClassEpilogue;
        }

        public CompilerResults Compile()
        {
            string executable = Application.ExecutablePath;
            string baseDir = Path.GetDirectoryName(executable);

            CodeDomProvider compiler = new CSharpCodeProvider(options.CompilerOptions);
            CompilerParameters parameters = new CompilerParameters();
            parameters.WarningLevel = 4;
            parameters.GenerateExecutable = false;
            parameters.GenerateInMemory = true;
            foreach (string assembly in options.Assemblies)
            {
                string suffix = assembly.EndsWith(".exe") ? "" : ".dll";
                string prefix = "";
                if (!assembly.StartsWith("System") && !assembly.StartsWith("Microsoft") && assembly != "WindowsBase")
                {
                    prefix = baseDir + "\\";
                }
                // This is a really grotty hack, but appears to work
                if (assembly == "WindowsBase")
                {
                    string refDir = @"C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETFramework\v4.0\";
                    if (!Directory.Exists(refDir))
                    {
                        refDir = refDir.Replace(" (x86)", "");
                    }
                    prefix = refDir;
                }
                parameters.ReferencedAssemblies.Add(prefix + assembly + suffix);
            }

            return compiler.CompileAssemblyFromSource(parameters, GenerateCode(true));
        }

        /// <summary>
        /// Exports the snippet by generating code (without extra line numbers)
        /// </summary>
        public void Export(string location)
        {
            File.WriteAllText(location, GenerateCode(false));
        }

        /// <summary>
        /// Saves the snippet in its raw form
        /// </summary>
        public void Save(string location)
        {
            File.WriteAllText(location, snippetText);
        }
    }
}
