using System.Text;
using StudyHelper;

// Використання: StudyHelper <файл_з_конспектом> [кількість_запитань]
Console.OutputEncoding = Encoding.UTF8;

if (args.Length < 1)
{
    Console.Error.WriteLine("Використання: StudyHelper <файл_з_конспектом> [кількість_запитань]");
    return 1;
}

string? apiKey = Environment.GetEnvironmentVariable("GEMINI_API_KEY");
if (string.IsNullOrWhiteSpace(apiKey))
{
    Console.Error.WriteLine("Не задано змінну середовища GEMINI_API_KEY.");
    return 1;
}

if (!File.Exists(args[0]))
{
    Console.Error.WriteLine($"Файл не знайдено: {args[0]}");
    return 1;
}

string text = File.ReadAllText(args[0]);
if (string.IsNullOrWhiteSpace(text))
{
    Console.Error.WriteLine("Файл порожній.");
    return 1;
}

int questionCount = args.Length > 1 && int.TryParse(args[1], out int n) && n > 0 ? n : 5;
string model = Environment.GetEnvironmentVariable("GEMINI_MODEL") ?? "gemini-2.5-flash";

try
{
    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
    var client = new GeminiClient(http, apiKey, model);

    Console.WriteLine($"Модель: {model}. Аналізую конспект...\n");
    StudyResult result = await client.AnalyzeAsync(text, questionCount);

    Console.WriteLine("=== ПЕРЕКАЗ ===");
    Console.WriteLine(result.Summary);
    Console.WriteLine();
    Console.WriteLine("=== ЗАПИТАННЯ ДЛЯ САМОПЕРЕВІРКИ ===");
    for (int i = 0; i < result.Questions.Count; i++)
    {
        Console.WriteLine($"{i + 1}. {result.Questions[i].Question}");
        Console.WriteLine($"   Відповідь: {result.Questions[i].Answer}");
    }

    return 0;
}
catch (GeminiException ex)
{
    Console.Error.WriteLine(ex.Message);
    return 2;
}
catch (TaskCanceledException)
{
    Console.Error.WriteLine("Час очікування відповіді Gemini вичерпано.");
    return 2;
}
catch (HttpRequestException ex)
{
    Console.Error.WriteLine($"Мережева помилка: {ex.Message}");
    return 2;
}
