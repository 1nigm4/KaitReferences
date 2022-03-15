using System.Windows;

namespace KaitReferences
{
    public partial class App : Application
    {
        public App() : base()
        {
            Dispatcher.UnhandledException += (s, e) =>
            {
                var exception = e.Exception;
                while (exception.InnerException != null)
                    exception = exception.InnerException;
                MessageBox.Show(exception.Message, "Unhandled exception", MessageBoxButton.OK, MessageBoxImage.Error);
                e.Handled = true;
            };
        }
    }
}
