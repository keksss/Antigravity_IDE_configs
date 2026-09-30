---
trigger: model_decision
description: composing, drafting, generating, or replying to emails and messages
---

# Author Email Voice and Composition Style

When generating, drafting, or replying to emails and messages on behalf of the user, adhere strictly to the author's personal communication style and typography rules derived from their actual email history.

---

## 1. Punctuation and Typography Rules (MANDATORY)

1. **Strictly Short Hyphens Only (`-`):**
   - **NEVER** use em dashes (`—`) or en dashes (`–`).
   - Use **only** standard short hyphens (`-`) for both hyphens and dashes (e.g., `слово - пояснение`, `T4MTEMP-876 - [BUILD] Task Name`).
   
2. **Straight Quotes:**
   - Use standard straight ASCII quotes (`"`) rather than curly quotes or guillemets (`« »`).

3. **Restrained Punctuation:**
   - Avoid excessive exclamation marks (`!`). Use them sparingly (primarily in greetings like `Добрый день!` or brief acknowledgments).
   - End normal sentences with standard periods.

4. **Grammar and Orthography Preservation:**
   - Do **NOT** artificially sanitize, over-correct, or lecture on grammar and orthography.
   - Retain the user's authentic engineering voice, colloquial IT phrasing, and natural speech cadence. Do not convert practical dev notes into textbook academic or bureaucratic Russian/English.

---

## 2. Structure and Formatting

1. **Concise, Punchy Paragraphs:**
   - Keep paragraphs short: typically 1 to 2 sentences per paragraph.
   - Separate paragraphs with a blank line.
   - Average sentence length is ~9–10 words. Avoid complex compound clauses and run-on sentences.

2. **Zero Corporate Fluff:**
   - Go straight to the point immediately.
   - First sentence: context, status update, or fact.
   - Subsequent sentences: the specific question, block, or requested action.

3. **Links and References:**
   - Place URLs on their own line with a concise lead-in:
     - `Ссылка на спеку: https://...`
     - `https://jira.example.com/browse/... - [Component] Description`

4. **Questions:**
   - Formulate questions directly and place them in their own paragraph:
     - `Вопрос, мы это по итогу делаем?`
     - `Can we create the new namespace for this purpose?`

---

## 3. Greetings and Signoffs

### Russian Messages:
- **Greetings:**
  - `Добрый день,`
  - `<Имя>, добрый день!`
  - `Привет,`
- **Signoffs:**
  - In quick internal replies: signoff is often omitted entirely.
  - When closing:
    - `Best regards,`
    - `Best regards,\nKonstantin`
    - `Хорошо, спасибо!`
  - *Avoid overly formal "С уважением, Константин" unless explicitly requested for high-stakes external correspondence.*

### English Messages:
- **Greetings:**
  - `Hello <Name>,`
  - `Hi <Name>,`
  - `Hello everyone,`
  - `Hi All,`
- **Signoffs:**
  - `Regards,\nKonstantin`
  - `Best regards,`
  - `Best regards,\nKonstantin`

---

## 4. Technical Terminology
- Use natural engineering terms freely without hyper-formal translation:
  - RU: *спека*, *тимс*, *аппка*, *дефект*, *билить время*, *селективный экран*, *GUI отчет*.
  - EN: clear, functional developer English focused on problem-solving, FM/APIs, and system namespaces.
