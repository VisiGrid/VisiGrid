# CI Integration

Verify dataset integrity on every pipeline run. Two paths, same proof.

## GitHub Actions

Use the official action — zero install, runs in any GitHub workflow.

```yaml
- uses: VisiGrid/verify-action@v1
  with:
    file: output/payments.csv
    repo: acme/payments
    api_key: ${{ secrets.VISIHUB_API_KEY }}
    wait: true
```

The action writes a Job Summary with check status, diff details, and a
link to the cryptographic proof. See the
[verify-action README](https://github.com/VisiGrid/verify-action) for
all options including `source_type`, `source_identity`, and
`fail_on_check_failure`.

## Generic CI (GitLab, Jenkins, Buildkite, cron, local)

Install `vgrid` and call `vgrid publish`. Works anywhere you can run a
static binary.

### Install

```bash
# Download the latest release
curl -fsSL https://github.com/VisiGrid/VisiGrid/releases/latest/download/vgrid-linux-x86_64.tar.gz \
  | tar xz -C /usr/local/bin vgrid
```

### Authenticate

```bash
# CI: pass the token directly (no TTY needed)
vgrid login --token "$VISIHUB_API_KEY"

# Interactive: prompts for token on stdin
vgrid login
```

The token is stored in `~/.config/visigrid/auth.json`. In CI runners
with ephemeral filesystems, run `vgrid login` at the start of each job.

### Publish

```bash
vgrid publish output/payments.csv \
  --repo acme/payments \
  --source-type dbt \
  --source-identity models/payments \
  --wait \
  --fail-on-check-failure
```

When stdout is piped (the default in CI), output is JSON:

```json
{
  "run_id": "42",
  "version": 3,
  "status": "verified",
  "check_status": "pass",
  "proof_url": "https://api.visihub.app/api/repos/acme/payments/runs/42/proof"
}
```

See [cli-output-schema.md](cli-output-schema.md) for the full field
reference and stability guarantees.

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Check passed (or no check configured) |
| 41 | Integrity check failed |
| 42 | Network error |
| 43 | Server validation error |
| 44 | Timeout |

---

## GitLab CI

```yaml
verify-dataset:
  stage: test
  image: ubuntu:latest
  before_script:
    - curl -fsSL https://github.com/VisiGrid/VisiGrid/releases/latest/download/vgrid-linux-x86_64.tar.gz
        | tar xz -C /usr/local/bin vgrid
    - vgrid login --token "$VISIHUB_API_KEY"
  script:
    - dbt run --select payments
    - vgrid publish target/payments.csv
        --repo acme/payments
        --source-type dbt
        --source-identity models/payments
        --wait
        --fail-on-check-failure
  artifacts:
    when: always
    reports:
      dotenv: visihub.env
```

## Jenkins (Declarative Pipeline)

```groovy
pipeline {
    agent any
    environment {
        VISIHUB_API_KEY = credentials('visihub-api-key')
    }
    stages {
        stage('Setup') {
            steps {
                sh '''
                    curl -fsSL https://github.com/VisiGrid/VisiGrid/releases/latest/download/vgrid-linux-x86_64.tar.gz \
                      | tar xz -C /usr/local/bin vgrid
                    vgrid login --token "$VISIHUB_API_KEY"
                '''
            }
        }
        stage('Build') {
            steps {
                sh 'dbt run --select payments'
            }
        }
        stage('Verify') {
            steps {
                sh '''
                    vgrid publish target/payments.csv \
                      --repo acme/payments \
                      --source-type dbt \
                      --source-identity models/payments \
                      --wait \
                      --fail-on-check-failure
                '''
            }
        }
    }
}
```

## Cron / Local

```bash
#!/usr/bin/env bash
set -euo pipefail

# Generate the dataset
python scripts/export_payments.py > /tmp/payments.csv

# Publish and verify
result=$(vgrid publish /tmp/payments.csv \
  --repo acme/payments \
  --source-type cron \
  --wait \
  --output json)

# Check the result
status=$(echo "$result" | jq -r '.check_status')
proof=$(echo "$result" | jq -r '.proof_url')

if [ "$status" = "fail" ]; then
  echo "FAILED — proof: $proof"
  # send alert
  exit 1
fi

echo "Verified — proof: $proof"
```
