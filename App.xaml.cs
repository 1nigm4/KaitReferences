﻿using KaitReference.Services;
using System.Windows;

namespace KaitReference
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

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            Excel.App.Quit();
        }
    }
}
