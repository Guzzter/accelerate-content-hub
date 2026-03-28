namespace Sitecore.CH.TranslationGenerator.Services.Abstract
{
    public interface IConsoleHelper
    {
        char GetChar(string label);

        bool GetExpectedChar(string label, char expectedChar);

        string GetInput(string label, Func<string, bool>? validateValueFunc = null);

        string GetInput(string label, string defaultValue, Func<string, bool>? validateValueFunc = null);

        T? GetInput<T>(string label, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null);

        T? GetInput<T>(string label, T defaultValue, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null, Func<T, string>? defaultValueOutputFunc = null);

        T? GetJsonInput<T>(string label, Func<T, bool>? validateValueFunc = null);

        T? SelectItem<T>(IEnumerable<T> items, string? title = null, string? defaultItemLabel = null) where T : class;

        T? SelectItem<T>(IDictionary<T, string> items, string? title = null, string? defaultItemLabel = null);

        void WriteList(string title, IEnumerable<string> data, bool displayEmptyData = false);

        void WriteList(string title, IEnumerable<object> data, bool displayEmptyData = false);

        void WriteTable(IEnumerable<string> headers, IEnumerable<IEnumerable<string?>> data);

        void ResetStyle();

        void Write(string text, bool newLine = true, bool pad = false);

        void WriteTempLine(string line);

        void OverwriteLine(string line);
    }
}
