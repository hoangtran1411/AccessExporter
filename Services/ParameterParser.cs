// AccessExporter/Services/ParameterParser.cs
using System.Data.OleDb;
using System.Globalization;
using AccessExporter.Models;

namespace AccessExporter.Services;

// Converts the text typed in the dialog into the real type the query expects
public static class ParameterParser
{
    public static ParameterValue Parse(QueryParameter parameter, string text)
    {
        var culture = CultureInfo.InvariantCulture;

        try
        {
            object value = parameter.Type switch
            {
                OleDbType.UnsignedTinyInt => byte.Parse(text, culture),
                OleDbType.SmallInt => short.Parse(text, culture),
                OleDbType.Integer => int.Parse(text, culture),
                OleDbType.BigInt => long.Parse(text, culture),
                OleDbType.Single => float.Parse(text, culture),
                OleDbType.Double => double.Parse(text, culture),
                OleDbType.Currency or OleDbType.Decimal or OleDbType.Numeric => decimal.Parse(text, culture),
                OleDbType.Date or OleDbType.DBDate or OleDbType.DBTimeStamp => DateTime.Parse(text, culture),
                OleDbType.Boolean => ParseBool(text),
                _ => text
            };

            return new ParameterValue(parameter.Name, parameter.Type, value);
        }
        catch (Exception ex) when (ex is FormatException or OverflowException)
        {
            throw new FormatException(
                $"Invalid value for {parameter.Name} ({parameter.Type}): '{text}'");
        }
    }

    private static bool ParseBool(string text) => text.Trim().ToLowerInvariant() switch
    {
        "1" or "-1" or "true" or "yes" => true,
        "0" or "false" or "no" => false,
        _ => throw new FormatException()
    };
}