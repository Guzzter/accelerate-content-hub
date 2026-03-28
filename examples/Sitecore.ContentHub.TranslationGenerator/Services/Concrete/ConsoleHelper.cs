using Sitecore.CH.TranslationGenerator.Services.Abstract;
using System.Text.Json;

namespace Sitecore.CH.TranslationGenerator.Services.Concrete
{
    class ConsoleHelper : IConsoleHelper
    {
        public string GetInput(string label, Func<string, bool>? validateValueFunc = null)
        {
            return GetInput(label, x => x, validateValueFunc) ?? string.Empty;
        }

        public string GetInput(string label, string defaultValue, Func<string, bool>? validateValueFunc = null)
        {
            return GetInput($"{label} ({defaultValue})", x => ParseDefaultValue(x, defaultValue), validateValueFunc) ?? string.Empty;
        }

        public T? GetInput<T>(string label, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null)
        {
            T? value;
            while (!TryGetInput(label, out value, parseValueFunc, validateValueFunc))
                WriteError("Input was invalid.");
            return value;
        }

        public T? GetInput<T>(string label, T defaultValue, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null, Func<T, string>? defaultValueOutputFunc = null)
        {
            T? value;
            var defaultValueString = defaultValueOutputFunc == null ? defaultValue?.ToString() : defaultValueOutputFunc(defaultValue);
            while (!TryGetInput($"{label} ({defaultValueString})", defaultValue, out value, parseValueFunc, validateValueFunc))
                WriteError("Input was invalid.");
            return value;
        }

        public T? GetJsonInput<T>(string label, Func<T, bool>? validateValueFunc = null)
        {
            T? value;
            while (!TryGetJsonInput(label, out value, validateValueFunc))
                WriteError("Input was invalid.");
            return value;
        }

        public T? SelectItem<T>(IEnumerable<T> items, string? title = null, string? defaultItemLabel = null) where T : class
        {
            return SelectItem(items.ToDictionary(x => x, x => x.ToString() ?? string.Empty), title, defaultItemLabel);
        }

        public T? SelectItem<T>(IDictionary<T, string> items, string? title = null, string? defaultItemLabel = null)
        {
            if (title != null)
                Write(title);

            var dictionary = new Dictionary<int, string>();
            foreach (var item in items)
                dictionary.Add(dictionary.Count + 1, item.Value);
            if (!string.IsNullOrEmpty(defaultItemLabel))
                dictionary.Add(0, defaultItemLabel);
            foreach (var item in dictionary)
                Write($"{item.Key} - {item.Value}");

            var selectedKey = GetInput("Selection", int.Parse, x => dictionary.ContainsKey(x));

            return selectedKey == 0 ? default : items.ElementAt(selectedKey - 1)!.Key;
        }

        public void ResetStyle()
        {
            Console.ResetColor();
        }

        public void Write(string text, bool newLine = true, bool pad = false)
        {
            if (pad)
                text = PadLine(text);
            if (newLine)
                Console.WriteLine(text);
            else
                Console.Write(text);
        }

        public void WriteTempLine(string line)
        {
            Write(line, false, true);
            Console.CursorLeft = 0;
        }

        public void OverwriteLine(string line)
        {
            Console.CursorLeft = 0;
            Write(line, false, true);
        }

        public void WriteTable(IEnumerable<string> headers, IEnumerable<IEnumerable<string?>> data)
        {
            const char ColumnSeperatorChar = '|';
            const char BreakRowSeperatorChar = '-';

            Write("");

            var columnWidths = new List<int>();
            foreach (var (header, index) in headers.Select((x, i) => (x, i)))
            {
                var columnData = data.Select(x => x.ElementAt(index));
                var columnWidth = Math.Max(header.Trim().Length, columnData.Any() ? columnData.Max(x => x?.Trim().Length ?? 0) : 0) + 1;

                columnWidths.Add(columnWidth);

                Write($"{ColumnSeperatorChar} {header.Trim().PadRight(columnWidth)}", false);
            }
            Write($"{ColumnSeperatorChar}");

            Write($"{ColumnSeperatorChar}{string.Join(ColumnSeperatorChar, columnWidths.Select(w => "".PadRight(w + 1, BreakRowSeperatorChar)))}{ColumnSeperatorChar}");

            foreach (var rowData in data)
            {
                foreach (var (cellData, colIndex) in rowData.Select((x, i) => (x, i)))
                {
                    Write($"{ColumnSeperatorChar} {cellData?.Trim().PadRight(columnWidths[colIndex])}", false);
                }
                Write($"{ColumnSeperatorChar}");
            }

            Write("");
        }

        public void WriteList(string title, IEnumerable<object> data, bool displayEmptyData = false)
        {
            WriteList(title, data.Select(x => x.ToString() ?? string.Empty), displayEmptyData);
        }

        public void WriteList(string title, IEnumerable<string> data, bool displayEmptyData = false)
        {
            if (displayEmptyData || data.Any())
                WriteTable(new string[] { title }, data.Select(x => new string[] { x }));
        }

        private bool TryGetJsonInput<T>(string label, out T? value, Func<T, bool>? validateValueFunc = null)
        {
            return TryGetInput(label, out value, i => JsonSerializer.Deserialize<T>(i), validateValueFunc);
        }

        private bool TryGetInput<T>(string label, out T? value, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null)
        {
            var input = GetInput(label, Console.ReadLine);
            return TryParseInput(input, out value, parseValueFunc, validateValueFunc);
        }

        private bool TryGetInput<T>(string label, T defaultValue, out T? value, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null)
        {
            var input = GetInput(label, Console.ReadLine);
            if (string.IsNullOrEmpty(input))
            {
                value = defaultValue;
                return true;
            }
            return TryParseInput(input, out value, parseValueFunc, validateValueFunc);
        }

        private bool TryParseInput<T>(string? input, out T? value, Func<string, T> parseValueFunc, Func<T, bool>? validateValueFunc = null)
        {
            try
            {
                value = parseValueFunc(input ?? string.Empty);
                return validateValueFunc == null || validateValueFunc(value);
            }
            catch
            {
                value = default;
                return false;
            }
        }

        private T GetInput<T>(string label, Func<T> inputAction)
        {
            Write($"{label}: ", false);
            Console.BackgroundColor = ConsoleColor.White;
            Console.ForegroundColor = ConsoleColor.Black;
            var input = inputAction();
            ResetStyle();
            Write(string.Empty, true);
            return input;
        }

        private void WriteError(string text)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Write(text);
            ResetStyle();
        }

        private string PadLine(string line)
        {
            return line.PadRight(Console.BufferWidth, ' ');
        }

        public char GetChar(string label)
        {
            return GetInput(label, Console.ReadKey).KeyChar;
        }

        public bool GetExpectedChar(string label, char expectedChar)
        {
            return GetChar(label) == expectedChar;
        }

        public static string ParseDefaultValue(string value, string defaultValue)
        {
            return string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(defaultValue) ? defaultValue : value;
        }
    }
}
