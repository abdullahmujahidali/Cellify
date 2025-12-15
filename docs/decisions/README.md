# Architecture Decision Records

This directory contains Architecture Decision Records (ADRs) documenting the key technical decisions made during the development of Cellify.

## What is an ADR?

An ADR is a document that captures an important architectural decision along with its context and consequences.

## Index

| ADR | Title | Status |
|-----|-------|--------|
| [001](./001-technology-stack.md) | Technology Stack | Accepted |
| [002](./002-minimal-dependencies.md) | Minimal Dependencies Philosophy | Accepted |
| [003](./003-data-structures.md) | Core Data Structures | Accepted |
| [004](./004-accessibility.md) | Accessibility Architecture | Accepted |
| [005](./005-csv-format.md) | CSV Import/Export Implementation | Accepted |
| [006](./006-xlsx-format.md) | XLSX Import/Export Implementation | Accepted |

## Template

When adding a new ADR, use this template:

```markdown
# ADR-XXX: Title

**Date:** YYYY-MM-DD
**Status:** Proposed | Accepted | Deprecated | Superseded

## Context

What is the issue that we're seeing that is motivating this decision?

## Decision

What is the change that we're proposing and/or doing?

## Consequences

What becomes easier or more difficult to do because of this change?
```

## Contributing

When making significant architectural decisions, please document them here so future contributors understand the reasoning behind our choices.
