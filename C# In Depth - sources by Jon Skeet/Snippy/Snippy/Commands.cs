using System.Windows.Input;

namespace Snippy
{
    public static class Commands
    {
        public static readonly RoutedCommand CompileAndRun = new RoutedCommand("Compile", typeof(MainWindow));
        public static readonly RoutedCommand Export = new RoutedCommand("Export", typeof(MainWindow));
        public static readonly RoutedCommand Import = new RoutedCommand("Import", typeof(MainWindow));
    }
}
