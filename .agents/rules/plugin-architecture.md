---
trigger: always_on
---

# Plugin Design Principles: Modularity and Portability

When creating or maintaining plugins, skills, and rules in this workspace, always adhere to the principles of **high modularity** and **strict portability**.
 
---

## 1. Modularity and Focused Skills

1. **Single Responsibility for Skills:**
   - Avoid creating large "monolithic" skills that try to handle all possible workflows in a single `SKILL.md`.
   - Decompose plugin functionality into focused, discrete skills based on user intents (e.g., separate skills for export/conversion, editing/mutation, and compilation/generation).
   - Keep scripts dedicated to specific tasks rather than packing everything into one mega-script.

2. **Leverage Progressive Disclosure:**
   - Antigravity loads skill instructions on demand. By keeping skills granular, the agent only injects instructions relevant to the active user prompt into its context window, preserving token budget and preventing instruction bleed.

3. **Shared Rules at Plugin Level:**
   - Place domain-wide constraints, style rules, and safety guidelines under `plugins/<plugin-name>/rules/` so they apply across all skills within the plugin without duplicating instructions.

---

## 2. Strict Portability and Relative Paths

Plugins are designed to be copied, shared, and transferred between different environments:
- From workspace level (`.agents/plugins/<plugin-name>/` or `_agents/plugins/<plugin-name>/`)
- To global user configuration (`~/.gemini/config/plugins/<plugin-name>/`)
- Across different machines, user directories, and operating systems (Windows / Linux / macOS).

### Requirements:

1. **NO Absolute Paths:**
   - Never hardcode absolute filesystem paths (such as `C:\Users\...`, `/home/...`, `/tmp/...`) anywhere in:
     - `plugin.json`
     - `SKILL.md` instructions and examples
     - Markdown documentation and rule files
     - Source code and scripts

2. **Relative Markdown & Documentation Links:**
   - All links between files inside a plugin must be relative to the containing document (e.g., `[Reference](./references/guide.md)` or `[Template](../../templates/default.docx)`).

3. **Dynamic Path Resolution in Scripts:**
   - In Python scripts, resolve paths to sister files, templates, or local assets dynamically relative to the script location:
     ```python
     from pathlib import Path
     BASE_DIR = Path(__file__).resolve().parent
     TEMPLATE_DIR = BASE_DIR.parent / "templates"
     ```
   - Never rely on the caller's current working directory (CWD) for internal plugin assets.

4. **Self-Contained Execution in `SKILL.md`:**
   - Instructions in `SKILL.md` must instruct the agent to run scripts using relative paths from the skill root or plugin root:
     ```bash
     # Good: relative to skill folder
     uv run scripts/<script_name>.py <arguments>
     python scripts/<script_name>.py <arguments>
     ```
   - Do not tie execution commands to hardcoded repository directory names like `.agents/...`.
