using System.IO;
using System.Text;
using System.Windows.Controls;
using System.Threading;
using System.Windows.Documents;
using System.Windows.Threading;
using System;

namespace Snippy
{
    class ParagraphWriter : TextWriter
    {
        /// <summary>
        /// Equivalent of Action in .NET 3.5
        /// </summary>
        delegate void VoidAction();

        Paragraph paragraph;

        internal ParagraphWriter(Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }

        public override void Write(char c)
        {
            Write(c.ToString());
        }

        public override void Write(char[] buffer, int index, int count)
        {
            Write(new string(buffer, index, count));
        }

        public override void Write(string text)
        {
            VoidAction action = () => paragraph.Inlines.Add(new Run(text));
            paragraph.Dispatcher.BeginInvoke(DispatcherPriority.Normal, action);
        }
    }
}
