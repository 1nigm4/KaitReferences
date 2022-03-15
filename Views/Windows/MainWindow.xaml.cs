using KaitReferences.Models;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace KaitReferences.Views.Windows
{
    public partial class MainWindow : Window
    {
        public static CheckBox WordVisible;
        private CollectionViewSource personsCollection;

        public MainWindow()
        {
            InitializeComponent();
            WordVisible = WVisible;
        }

        private void PersonsCollection_OnFilter(object sender, FilterEventArgs e)
        {
            if (!(e.Item is Person person)) return;
            if (person.Education.Area.Contains(SearchArea.Text, StringComparison.OrdinalIgnoreCase) &&
                person.LastName.Contains(SearchLastName.Text, StringComparison.OrdinalIgnoreCase) &&
                person.Name.Contains(SearchName.Text, StringComparison.OrdinalIgnoreCase) &&
                person.Reference.Status.Contains(SearchStatus.Text, StringComparison.OrdinalIgnoreCase)) return;
            e.Accepted = false;
        }

        private void ReferenceSearch_OnChanged(dynamic sender, TextChangedEventArgs e)
        {
            personsCollection ??= (CollectionViewSource) sender.FindResource("PersonsCollection");
            personsCollection.View.Refresh();
        }
    }
}
