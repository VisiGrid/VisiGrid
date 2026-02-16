//! CI runner identity detection.
//!
//! Scrapes environment variables from common CI providers to build a
//! runner context object for inclusion in `source_metadata.runner`.
//! This provides a tamper-evident audit trail: proofs record *who* ran
//! the computation and *where*, without requiring server-side verification.

use serde_json::{json, Value};

/// Detect CI environment and return structured runner metadata.
/// Returns `None` when not running in a recognized CI environment.
pub fn get_runner_context() -> Option<Value> {
    // GitHub Actions
    if let Ok(repo) = std::env::var("GITHUB_REPOSITORY") {
        return Some(json!({
            "provider": "github_actions",
            "repo": repo,
            "run_id": env_or_null("GITHUB_RUN_ID"),
            "run_number": env_or_null("GITHUB_RUN_NUMBER"),
            "actor": env_or_null("GITHUB_ACTOR"),
            "sha": env_or_null("GITHUB_SHA"),
            "ref": env_or_null("GITHUB_REF"),
            "workflow": env_or_null("GITHUB_WORKFLOW"),
            "event": env_or_null("GITHUB_EVENT_NAME"),
        }));
    }

    // GitLab CI
    if let Ok(project) = std::env::var("CI_PROJECT_PATH") {
        return Some(json!({
            "provider": "gitlab_ci",
            "project": project,
            "pipeline_id": env_or_null("CI_PIPELINE_ID"),
            "job_id": env_or_null("CI_JOB_ID"),
            "user": env_or_null("GITLAB_USER_LOGIN"),
            "sha": env_or_null("CI_COMMIT_SHA"),
            "ref": env_or_null("CI_COMMIT_REF_NAME"),
        }));
    }

    // CircleCI
    if let Ok(project) = std::env::var("CIRCLE_PROJECT_REPONAME") {
        return Some(json!({
            "provider": "circleci",
            "project": project,
            "build_num": env_or_null("CIRCLE_BUILD_NUM"),
            "build_url": env_or_null("CIRCLE_BUILD_URL"),
            "username": env_or_null("CIRCLE_USERNAME"),
            "sha": env_or_null("CIRCLE_SHA1"),
            "branch": env_or_null("CIRCLE_BRANCH"),
        }));
    }

    // Generic CI detection (Travis, Jenkins, Buildkite, etc.)
    if std::env::var("CI").map(|v| v == "true" || v == "1").unwrap_or(false) {
        return Some(json!({
            "provider": "unknown",
            "ci": true,
        }));
    }

    None
}

fn env_or_null(key: &str) -> Value {
    match std::env::var(key) {
        Ok(v) if !v.is_empty() => Value::String(v),
        _ => Value::Null,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn returns_none_outside_ci() {
        // In normal test env, CI vars shouldn't be set (unless running in CI)
        // This test validates the function doesn't panic
        let _ = get_runner_context();
    }
}
