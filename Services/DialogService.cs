// AccessExporter/Services/DialogService.cs
using System.Windows;
using AccessExporter.Models;
using AccessExporter.Views;

namespace AccessExporter.Services;

public interface IDialogService
{
    // Returns the typed texts in parameter order, or null if the user canceled
    IReadOnlyList<string>? AskParameters(
        string title,
        IReadOnlyList<QueryParameter> parameters,
        IReadOnlyList<string>? lastInputs);
}

public class DialogService : IDialogService
{
    public IReadOnlyList<string>? AskParameters(
        string title,
        IReadOnlyList<QueryParameter> parameters,
        IReadOnlyList<string>? lastInputs)
    {
        var inputs = parameters
            .Select((p, i) => new ParameterInput(
                p,
                lastInputs is not null && i < lastInputs.Count ? lastInputs[i] : string.Empty))
            .ToList();

        var dialog = new ParameterDialog(title, inputs)
        {
            Owner = Application.Current.MainWindow
        };

        return dialog.ShowDialog() == true
            ? inputs.Select(i => i.Text).ToList()
            : null;
    }
}