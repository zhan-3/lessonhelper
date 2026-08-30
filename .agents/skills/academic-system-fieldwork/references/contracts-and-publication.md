# Contracts and publication

## Diagnostic evidence versus facts

Discovery output is non-publishable. A plausible endpoint, visible row, manual navigation, or successful parser run may propose a contract but must not replace an academic snapshot.

Prefer separate result types for:

```text
diagnostic observation
verified read
incomplete read
published complete snapshot
```

This boundary belongs in types and persistence APIs, not only prose or UI copy.

## Request contract

Version a fixed contract with:

- method and redacted path pattern;
- required header names;
- query/form field names;
- dynamic-field acquisition procedure;
- authentication redirect/error markers;
- response schema fingerprint;
- business identity field;
- term/category/page provenance;
- pagination termination rule;
- semantic acceptance predicates.

Request parity means reproducing shape through the authenticated browser context. Never replay captured cookies, authorization values, CSRF tokens, or stale dynamic fields.

## Three acceptance gates

### Transport

- request reached the intended protected endpoint;
- response is not a login redirect/page;
- status and content type match the contract.

### Parse

- schema keys/types match;
- pagination metadata parses;
- stable business identities are present.

### Domain

- requested term/category matches returned provenance;
- populated UI witness is not accepted as inexplicably empty;
- row counts and pagination are plausible;
- no class/reference dataset silently substitutes for personal facts.

HTTP 200 and `complete=true` are insufficient without domain acceptance.

## Identity

Use server-issued business identities. Display name, course code, row number, DOM position, and concatenated fallback keys may support presentation or diagnostics but cannot authorize execution.

## Completeness barrier

Build a manifest before collection:

```yaml
expected:
  - term: current
    category: example
    page: 1
completed: []
failed: []
```

Expand expected pages only from verified pagination metadata. Publish atomically when every expected segment is completed and none failed. Otherwise:

- retain the last complete snapshot;
- store the attempt and sanitized diagnostics separately;
- report missing segments as unknown.

## Contract drift

A fixed contract that fails should fail closed. Discovery may gather sanitized evidence for a new version, but it must not silently become the production read path or publish its observations.

## Suggested evals

- HTTP 200 with login HTML;
- schema-valid all-zero result contradicted by a UI witness;
- missing middle page;
- one failed category among several;
- wrong term provenance;
- absent business identity;
- diagnostic discovery succeeds while fixed read fails;
- failed refresh preserves the previous complete snapshot.
