// AccessExporter/Models/TableItem.cs
using System.Data.OleDb;
using CommunityToolkit.Mvvm.ComponentModel;

namespace AccessExporter.Models;

public enum TableKind
{
    Table,
    LinkTable,
    ViewOrOdbc,
    Procedure
}

// Definition of one parameter of a saved query
public record QueryParameter(string Name, int Position, OleDbType Type);

public record TableInfo(string Name, TableKind Kind, IReadOnlyList<QueryParameter> Parameters);

public partial class TableItem : ObservableObject
{
    public TableItem(string name, TableKind kind, IReadOnlyList<QueryParameter> parameters)
    {
        Name = name;
        Kind = kind;
        Parameters = parameters;
    }

    public string Name { get; }

    public TableKind Kind { get; }

    public IReadOnlyList<QueryParameter> Parameters { get; }

    // Text typed in the parameter dialog last time, used to pre-fill it
    public IReadOnlyList<string>? LastInputs { get; set; }

    [ObservableProperty]
    private bool _isSelected;
}