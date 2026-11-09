// AccessExporter/Models/ParameterInput.cs
namespace AccessExporter.Models;

// One row in the parameter dialog: label + the text the user types
public class ParameterInput
{
    public ParameterInput(QueryParameter definition, string text)
    {
        Definition = definition;
        Text = text;
    }

    public QueryParameter Definition { get; }

    public string Label => $"{Definition.Name} ({Definition.Type})";

    public string Text { get; set; }
}