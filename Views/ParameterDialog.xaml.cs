// AccessExporter/Views/ParameterDialog.xaml.cs
using System.Windows;
using AccessExporter.Models;

namespace AccessExporter.Views;

public partial class ParameterDialog : Window
{
    public ParameterDialog(string title, IReadOnlyList<ParameterInput> inputs)
    {
        InitializeComponent();
        Title = title;
        ParameterList.ItemsSource = inputs;
    }

    // The Cancel button has IsCancel="True", so WPF closes the dialog with DialogResult = false by itself
    private void Ok_Click(object sender, RoutedEventArgs e) => DialogResult = true;
}