---
trigger: model_decision
description: python code for skills generation
---

# Python Virtual Environments for Skills and Plugins

When designing or creating skills and plugins that contain Python scripts in their `scripts/` directory, always enforce dependency isolation so that scripts never install or run packages in the global Python environment.

## Requirements for Skills and Plugins

1. **Never use global pip:**
   - Prohibit commands like `pip install <package>` directly into the system or global environment.
   - All Python scripts must execute within an isolated virtual environment.

2. **Choose an Isolation Strategy:**
   - **Option A (Recommended if `uv` is available): `uv run` with inline metadata (PEP 723)**
     - Declare dependencies directly in the script header:
       ```python
       # /// script
       # dependencies = [
       #   "requests",
       #   "pydantic",
       # ]
       # ///
       ```
     - Instruct the agent in `SKILL.md` to execute the script via:
       ```bash
       uv run scripts/<script_name>.py [args]
       ```
   - **Option B: Dedicated `.venv`**
     - Maintain a `requirements.txt` or `pyproject.toml` inside the skill directory.
     - Include setup instructions in `SKILL.md` to check for and create a local `.venv` if it does not exist:
       - Windows: `<path-to-skill>/.venv/Scripts/python scripts/<script_name>.py` (or `./.venv/Scripts/python scripts/<script_name>.py` when executed from the skill directory)
       - Linux/macOS: `<path-to-skill>/.venv/bin/python scripts/<script_name>.py` (or `./.venv/bin/python scripts/<script_name>.py` when executed from the skill directory)

3. **Documentation in `SKILL.md`:**
   - Always include an explicit **Environment Setup** section explaining how the virtual environment is initialized and which command the agent must use to run the scripts.
