using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace StudyHelper;

public class GeminiException(string message) : Exception(message);

// Мінімальний клієнт Gemini API (REST, метод generateContent)
public class GeminiClient(HttpClient http, string apiKey, string model)
{
    private const string BaseUrl = "https://generativelanguage.googleapis.com/v1beta/models";

    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNameCaseInsensitive = true };

    public async Task<StudyResult> AnalyzeAsync(string text, int questionCount)
    {
        string prompt =
            "Ти помічник для навчання. Нижче конспект. Склади стислий переказ (3-5 речень) " +
            $"і рівно {questionCount} запитань для самоперевірки з короткими відповідями. " +
            "Спирайся лише на текст конспекту, нічого не вигадуй. Пиши мовою конспекту.\n\n" +
            "Конспект:\n" + text;

        var body = new JsonObject
        {
            ["contents"] = new JsonArray(new JsonObject
            {
                ["parts"] = new JsonArray(new JsonObject { ["text"] = prompt })
            }),
            ["generationConfig"] = new JsonObject
            {
                ["responseMimeType"] = "application/json",
                ["responseSchema"] = BuildSchema()
            }
        };

        using var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseUrl}/{model}:generateContent")
        {
            Content = new StringContent(body.ToJsonString(), Encoding.UTF8, "application/json")
        };
        request.Headers.Add("x-goog-api-key", apiKey);

        using var response = await http.SendAsync(request);
        string json = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
        {
            throw new GeminiException(DescribeError(response.StatusCode, json));
        }

        string? payload = ExtractText(json);
        if (string.IsNullOrWhiteSpace(payload))
        {
            throw new GeminiException("Gemini повернув порожню відповідь (можливо, запит заблоковано фільтром).");
        }

        var result = JsonSerializer.Deserialize<StudyResult>(payload, JsonOptions);
        return result ?? throw new GeminiException("Не вдалося розібрати відповідь Gemini.");
    }

    private static JsonObject BuildSchema() => new()
    {
        ["type"] = "OBJECT",
        ["properties"] = new JsonObject
        {
            ["summary"] = new JsonObject { ["type"] = "STRING" },
            ["questions"] = new JsonObject
            {
                ["type"] = "ARRAY",
                ["items"] = new JsonObject
                {
                    ["type"] = "OBJECT",
                    ["properties"] = new JsonObject
                    {
                        ["question"] = new JsonObject { ["type"] = "STRING" },
                        ["answer"] = new JsonObject { ["type"] = "STRING" }
                    },
                    ["required"] = new JsonArray("question", "answer")
                }
            }
        },
        ["required"] = new JsonArray("summary", "questions")
    };

    private static string? ExtractText(string json)
    {
        var root = JsonNode.Parse(json);
        return root?["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.GetValue<string>();
    }

    private static string DescribeError(HttpStatusCode status, string json)
    {
        string detail = JsonNode.Parse(json)?["error"]?["message"]?.GetValue<string>() ?? json;
        bool keyProblem = status is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden
            || detail.Contains("API key", StringComparison.OrdinalIgnoreCase);

        return status switch
        {
            HttpStatusCode.TooManyRequests => "Перевищено ліміт запитів Gemini (429). Зачекай хвилину і спробуй знову.",
            HttpStatusCode.PaymentRequired => "Закінчились кредити проєкту в AI Studio (402). Поповни баланс або візьми ключ з іншого проєкту.",
            _ when keyProblem => $"Ключ API не прийнято ({(int)status}). Перевір змінну GEMINI_API_KEY. Деталі: {detail}",
            _ => $"Помилка Gemini API ({(int)status}): {detail}"
        };
    }
}
