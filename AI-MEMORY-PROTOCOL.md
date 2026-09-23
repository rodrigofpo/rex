# UNIVERSAL CONTEXT ENGINEERING PROTOCOL (AI-MEMORY)

This document contains standard operating procedures for interaction between development agents and the host's long-term memory system. Agents must read and follow these constraints to ensure context synchronization across tool switches without modifying or generating local database files.

---

## 💻 1. Environment & Architecture Overview

*   **Operating System:** Fedora Linux (Developer Environment)
*   **Workflow Mode:** Individual Developer (Single-user workspace)
*   **Database Isolation Strategy:** Strictly central data management. **No database files (`.db`, `.sqlite`) or temporary folders (`.ai-memory/`) should be initialized inside the project repository.** All runtime metrics, embeddings, and transaction logs reside strictly within the user data path: `~/.local/share/ai-memory/`.
*   **Repository Footprint:** The project codebase must only receive clean, version-controlled Markdown (`.md`) artifacts representing distilled context and long-term architectural rules.

---

## 🛠️ 2. Tool-Agnostic Context Strategy

Any active agent (IDEs, CLI utilities, or autonomous pipelines) must interact with this context via standard file inputs or protocol layers rather than modifying the codebase state directly:
*   **Targeted Context:** Agents must target precise code locations or specific documentation paths rather than analyzing the entire repository tree blindly.
*   **Context Ingestion:** The active runtime must ingest localized markdown notes directly from the designated documentation directory to inject historical system constraints into its system prompt before running programmatic tasks.

---

## 🔄 3. Step-by-Step Context Lifecycle for Agents

To keep a lightweight footprint, agents must execute context capture, processing, and localized markdown exporting in a strict sequence:

### Step 1: Capture Context (Central)
Whenever a significant development milestone, refactoring step, or architectural decision is concluded, record the interaction state to the central host service.
```bash
ai-memory capture --project-path . --message "Describe the implemented feature and what structural or business constraints were solved."
```

### Step 2: Distill and Export Markdown Only (Project Repository)
To update the project context without exposing database binaries, execute the consolidation pipeline (Hermes loop) pointing directly to a specific documentation directory inside the project repository.
```bash
ai-memory process --output-dir ./docs/ai-notes
```
*   **Agent Constraint:** Validate that the output folder contains exclusively `.md` files structured using open metadata standards. The agent must reject actions that generate structural `.toml` configuration profiles or `.db` storage objects inside the repository documentation folder.

### Step 3: Consume Context
When launching a new session or a new tool over this codebase, explicitly bind the localized notes folder to the new agent execution environment to restore project parity.
