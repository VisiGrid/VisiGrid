// AI client for Ask AI feature
//
// Sends questions to AI providers and parses structured responses.
// Currently supports OpenAI only.

use serde::{Deserialize, Serialize};
use visigrid_config::ai::ResolvedAIConfig;
use visigrid_config::settings::AIProvider;

use super::context::AIContext;

/// Response from Ask AI
#[derive(Debug, Clone)]
pub struct AskResponse {
    /// Natural language explanation
    pub explanation: String,

    /// Proposed formula (if AI generated one)
    pub formula: Option<String>,

    /// Warnings about the response
    pub warnings: Vec<String>,

    /// Raw response text (for debugging)
    pub raw_response: Option<String>,
}

/// Error from Ask AI
#[derive(Debug, Clone)]
pub enum AskError {
    /// Provider not configured
    NotConfigured(String),
    /// Provider not implemented
    NotImplemented(String),
    /// API key missing
    MissingKey,
    /// Network error
    NetworkError(String),
    /// API error response
    ApiError { status: u16, message: String },
    /// Failed to parse response
    ParseError(String),
    /// Provider returned unexpected format
    InvalidResponse(String),
}

impl std::fmt::Display for AskError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AskError::NotConfigured(msg) => write!(f, "AI not configured: {}", msg),
            AskError::NotImplemented(msg) => write!(f, "Provider not implemented: {}", msg),
            AskError::MissingKey => write!(f, "API key not configured"),
            AskError::NetworkError(msg) => write!(f, "Network error: {}", msg),
            AskError::ApiError { status, message } => write!(f, "API error ({}): {}", status, message),
            AskError::ParseError(msg) => write!(f, "Failed to parse response: {}", msg),
            AskError::InvalidResponse(msg) => write!(f, "Invalid response: {}", msg),
        }
    }
}

impl std::error::Error for AskError {}

// ============================================================================
// OpenAI API types
// ============================================================================

#[derive(Serialize)]
struct OpenAIRequest {
    model: String,
    messages: Vec<OpenAIMessage>,
    temperature: f32,
    max_tokens: u32,
    response_format: Option<OpenAIResponseFormat>,
}

#[derive(Serialize)]
struct OpenAIMessage {
    role: String,
    content: String,
}

#[derive(Serialize)]
struct OpenAIResponseFormat {
    #[serde(rename = "type")]
    format_type: String,
}

#[derive(Deserialize)]
struct OpenAIResponse {
    choices: Vec<OpenAIChoice>,
    #[allow(dead_code)]
    usage: Option<OpenAIUsage>,
}

#[derive(Deserialize)]
struct OpenAIChoice {
    message: OpenAIResponseMessage,
    #[allow(dead_code)]
    finish_reason: Option<String>,
}

#[derive(Deserialize)]
struct OpenAIResponseMessage {
    content: String,
}

#[derive(Deserialize)]
struct OpenAIUsage {
    #[allow(dead_code)]
    prompt_tokens: u32,
    #[allow(dead_code)]
    completion_tokens: u32,
    #[allow(dead_code)]
    total_tokens: u32,
}

#[derive(Deserialize)]
struct OpenAIError {
    error: OpenAIErrorDetail,
}

#[derive(Deserialize)]
struct OpenAIErrorDetail {
    message: String,
    #[allow(dead_code)]
    r#type: Option<String>,
}

/// Expected JSON structure from AI
#[derive(Deserialize)]
struct AIJsonResponse {
    explanation: String,
    formula: Option<String>,
}

// ============================================================================
// Main API
// ============================================================================

/// Ask AI a question about the given context
///
/// Returns a structured response with explanation and optional formula.
/// This is a blocking call - use in a background task.
pub fn ask_ai(
    config: &ResolvedAIConfig,
    question: &str,
    context: &AIContext,
) -> Result<AskResponse, AskError> {
    // Check provider is configured and implemented
    match config.provider {
        AIProvider::None => {
            return Err(AskError::NotConfigured("AI is disabled".to_string()));
        }
        AIProvider::OpenAI => {
            // Continue with OpenAI implementation
        }
        AIProvider::Local | AIProvider::Anthropic | AIProvider::Gemini | AIProvider::Grok => {
            return Err(AskError::NotImplemented(format!(
                "{} provider not yet implemented",
                config.provider.name()
            )));
        }
    }

    // Check API key
    let api_key = config.api_key.as_ref().ok_or(AskError::MissingKey)?;

    // Build prompt
    let system_prompt = build_system_prompt();
    let user_prompt = build_user_prompt(question, context);

    // Call OpenAI API
    call_openai(api_key, &config.model, &system_prompt, &user_prompt)
}

fn build_system_prompt() -> String {
    r#"You are a spreadsheet assistant. Your role is to help users understand and analyze their data.

CRITICAL INSTRUCTIONS:
1. Return ONLY valid JSON with exactly these keys: "explanation" and "formula"
2. "explanation" must be a string explaining your analysis or answer
3. "formula" must be either a single spreadsheet formula string starting with "=" or null
4. Do NOT include any text before or after the JSON
5. Do NOT use markdown code blocks

FORMULA RULES:
- Use standard spreadsheet functions (SUM, AVERAGE, SUMIF, COUNTIF, VLOOKUP, INDEX, MATCH, etc.)
- Reference cells using A1 notation
- If no formula is appropriate, set formula to null
- Only propose ONE formula, not multiple alternatives

RESPONSE FORMAT:
{"explanation": "your explanation here", "formula": "=FORMULA(A1:B10)" or null}

Example good response:
{"explanation": "The top 5 values can be found using LARGE function with INDEX/MATCH to get corresponding names.", "formula": "=INDEX(A:A,MATCH(LARGE(B:B,1),B:B,0))"}

Example when no formula applies:
{"explanation": "The data shows a 15% increase in sales compared to last quarter.", "formula": null}"#.to_string()
}

fn build_user_prompt(question: &str, context: &AIContext) -> String {
    let mut prompt = String::new();

    prompt.push_str("CONTEXT:\n");
    prompt.push_str(&context.to_prompt_text());
    prompt.push('\n');

    prompt.push_str("QUESTION:\n");
    prompt.push_str(question);
    prompt.push('\n');

    prompt.push_str("\nRemember: Return ONLY valid JSON with \"explanation\" and \"formula\" keys.");

    prompt
}

fn call_openai(
    api_key: &str,
    model: &str,
    system_prompt: &str,
    user_prompt: &str,
) -> Result<AskResponse, AskError> {
    let client = reqwest::blocking::Client::builder()
        .timeout(std::time::Duration::from_secs(60))
        .build()
        .map_err(|e| AskError::NetworkError(e.to_string()))?;

    let request = OpenAIRequest {
        model: model.to_string(),
        messages: vec![
            OpenAIMessage {
                role: "system".to_string(),
                content: system_prompt.to_string(),
            },
            OpenAIMessage {
                role: "user".to_string(),
                content: user_prompt.to_string(),
            },
        ],
        temperature: 0.3, // Lower temperature for more consistent output
        max_tokens: 1024,
        response_format: Some(OpenAIResponseFormat {
            format_type: "json_object".to_string(),
        }),
    };

    let response = client
        .post("https://api.openai.com/v1/chat/completions")
        .header("Authorization", format!("Bearer {}", api_key))
        .header("Content-Type", "application/json")
        .json(&request)
        .send()
        .map_err(|e| AskError::NetworkError(e.to_string()))?;

    let status = response.status();

    if !status.is_success() {
        let error_text = response.text().unwrap_or_default();
        if let Ok(error) = serde_json::from_str::<OpenAIError>(&error_text) {
            return Err(AskError::ApiError {
                status: status.as_u16(),
                message: error.error.message,
            });
        }
        return Err(AskError::ApiError {
            status: status.as_u16(),
            message: error_text,
        });
    }

    let response_body: OpenAIResponse = response
        .json()
        .map_err(|e| AskError::ParseError(e.to_string()))?;

    let content = response_body
        .choices
        .first()
        .map(|c| c.message.content.clone())
        .ok_or_else(|| AskError::InvalidResponse("No choices in response".to_string()))?;

    parse_ai_response(&content)
}

fn parse_ai_response(content: &str) -> Result<AskResponse, AskError> {
    let mut warnings = Vec::new();

    // Try to parse as JSON
    let parsed: AIJsonResponse = match serde_json::from_str(content) {
        Ok(p) => p,
        Err(e) => {
            // Try to extract JSON from the response if it's wrapped in markdown
            if let Some(json_start) = content.find('{') {
                if let Some(json_end) = content.rfind('}') {
                    let json_str = &content[json_start..=json_end];
                    match serde_json::from_str(json_str) {
                        Ok(p) => {
                            warnings.push("Response contained extra text around JSON".to_string());
                            p
                        }
                        Err(_) => {
                            return Err(AskError::ParseError(format!(
                                "Failed to parse JSON: {}. Raw: {}",
                                e, content
                            )));
                        }
                    }
                } else {
                    return Err(AskError::ParseError(format!(
                        "Failed to parse JSON: {}. Raw: {}",
                        e, content
                    )));
                }
            } else {
                return Err(AskError::ParseError(format!(
                    "Response is not JSON: {}. Raw: {}",
                    e, content
                )));
            }
        }
    };

    // Canonicalize and validate formula if present
    let formula = parsed.formula.map(|f| {
        let trimmed = f.trim().to_string();
        // Auto-fix: prepend = if missing but formula looks valid
        if !trimmed.starts_with('=') && !trimmed.is_empty() {
            warnings.push("Added leading '=' to formula".to_string());
            format!("={}", trimmed)
        } else {
            trimmed
        }
    });

    Ok(AskResponse {
        explanation: parsed.explanation.trim().to_string(),
        formula,
        warnings,
        raw_response: Some(content.to_string()),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_valid_json() {
        let json = r#"{"explanation": "The sum is 100.", "formula": "=SUM(A1:A10)"}"#;
        let response = parse_ai_response(json).unwrap();
        assert_eq!(response.explanation, "The sum is 100.");
        assert_eq!(response.formula, Some("=SUM(A1:A10)".to_string()));
        assert!(response.warnings.is_empty());
    }

    #[test]
    fn test_parse_null_formula() {
        let json = r#"{"explanation": "No formula needed.", "formula": null}"#;
        let response = parse_ai_response(json).unwrap();
        assert_eq!(response.explanation, "No formula needed.");
        assert!(response.formula.is_none());
    }

    #[test]
    fn test_parse_json_with_markdown() {
        let json = "Here's the answer:\n```json\n{\"explanation\": \"Test\", \"formula\": null}\n```";
        let response = parse_ai_response(json).unwrap();
        assert_eq!(response.explanation, "Test");
        assert!(!response.warnings.is_empty()); // Should warn about extra text
    }
}
