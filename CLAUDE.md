# CLAUDE.md — OmniStay Core

## Project Overview

**Repository:** Ait426/omnistay

OmniStay is a B2B SaaS integration platform that extracts reservation and revenue data from multiple external PMS (Property Management Systems), normalizes it to a global standard format, and asynchronously writes it to Excel ledgers and other outputs.

**Core philosophy:** Modularity, zero single-points-of-failure (SPOF), scalability, data integrity.

**Critical warning:** The `hotel-automation` folder (legacy ~30k-line codebase) is used for live operations. **NEVER scan, read, or modify it.** All work happens exclusively inside `omnistay-core`.

## System Architecture

The system is strictly isolated into 3 layers with a **one-way data pipeline**:

```
Adapters  ──▶  Core DB  ──▶  Workers
(collect)      (normalize)    (output)
```

### Directory Structure

```
omnistay-core/
├── adapters/       # Data collection — PMS integrations & web scraping
│                   # Fetches data and passes it to Core
│                   # ⛔ No business logic allowed here
│
├── core/           # Central control — the system's heart
│   ├── models/     # SQLAlchemy ORM models (DB schema)
│   ├── normalize/  # Global data normalization logic
│   └── config_loader.py  # Reads from /config/config.json
│
├── workers/        # One-way output — consumes Core DB data
│                   # Writes to Excel, sends notifications, etc.
│                   # ⛔ Must run independently; no knowledge of other workers or adapters
│
└── config/
    └── config.json # All environment/runtime settings (NEVER hardcode values)
```

### Data Flow Rules

- **Adapters → Core DB → Workers** (one direction only)
- Adapters CANNOT call Workers directly
- Workers CANNOT call or depend on Adapters
- All cross-layer communication goes through the Core DB

## Tech Stack

| Component          | Technology                  | Notes                                        |
| ------------------ | --------------------------- | -------------------------------------------- |
| Language           | Python 3.14+                |                                              |
| Database           | SQLite via SQLAlchemy ORM   | No raw SQL — design for future PostgreSQL migration |
| Excel Automation   | xlwings (COM control)       | Preserves formulas; no openpyxl for writes   |

## Strict Rules for AI Assistants

### Rule 1: No Hardcoding

All configurable values MUST be loaded from `/config/config.json` via `config_loader.py`:

- Excel cell coordinates (e.g., `B4`, `K25`)
- Timezone identifiers (e.g., `Asia/Seoul`)
- Currency codes (e.g., `KRW`)
- Any value that could change between deployments

**Never** embed these directly in Python source files.

### Rule 2: Global Data Standards

**Time:** All timestamps stored in the DB (`check_in_time`, `check_out_time`, etc.) MUST be in **UTC**. Convert to local timezone only at the presentation/output layer (display, Excel writes).

**Currency:** All monetary amounts (`total_amount`, etc.) MUST use Python's `Decimal` type. Every amount column MUST be paired with an ISO 4217 currency code column (e.g., `currency="KRW"`).

**Reservation status:** Only these 4 enum values are allowed system-wide:

| Enum Value          | Meaning         |
| ------------------- | --------------- |
| `EXPECTED_CHECKIN`  | Expected arrival |
| `CHECKED_IN`        | Guest checked in |
| `CHECKED_OUT`       | Guest checked out |
| `CANCELED`          | Reservation canceled |

External PMS status codes (e.g., Yanolja's `IH`, `SO`) MUST be mapped to one of these 4 values inside the adapter layer before reaching Core.

### Rule 3: Fail-Fast & Explicit Logging

- **Never** swallow errors with bare `try-except: pass`
- If Excel is locked (access denied) or data is empty, log the error clearly and halt the process immediately
- Design for operator awareness — silent failures are forbidden

### Rule 4: One-Way Data Pipeline

Enforce the strict flow: `Adapters → Core DB → Workers`

- Adapters must not import from or call Workers
- Workers must not import from or call Adapters
- Core is the only shared dependency

## Development Workflow

### Branch Strategy

- Development branches follow the pattern `claude/<description>-<session-id>`
- Always push with `git push -u origin <branch-name>`
- Never force-push without explicit permission

### Commit Conventions

- Write clear, descriptive commit messages
- Use imperative mood in commit subjects (e.g., "Add feature" not "Added feature")
- Keep subject lines under 72 characters

## Build & Test Commands

_No build or test commands configured yet. Update this section when a build system is added._

## General AI Assistant Guidelines

1. **Read before writing** — Always read existing files before modifying them
2. **Minimal changes** — Only make changes that are directly requested or clearly necessary
3. **No over-engineering** — Keep solutions simple; avoid abstractions for hypothetical future needs
4. **Security first** — Never introduce injection, XSS, or other OWASP Top 10 vulnerabilities
5. **Never touch `hotel-automation`** — That folder is off-limits, period
6. **Update this file** — When adding infrastructure (build tools, test frameworks, CI/CD), update the relevant sections here
