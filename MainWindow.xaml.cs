// AccessExporter/MainWindow.xaml.cs
using System.Windows;
using AccessExporter.ViewModels;

namespace AccessExporter;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext = new MainViewModel();
    }
}