namespace StudyHelper;

// Структура відповіді, яку просимо в Gemini (structured output)
public record QuizItem(string Question, string Answer);

public record StudyResult(string Summary, List<QuizItem> Questions);
