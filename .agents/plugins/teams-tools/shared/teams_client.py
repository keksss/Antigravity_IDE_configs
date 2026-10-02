# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "websockets>=12.0",
# ]
# ///

"""
teams_client.py - Local connector to Microsoft Teams via WebView2 CDP WebSocket (port 9222).
Provides direct, local, safe access to Teams conversations, messages, attachments, and messaging.
"""

import asyncio
import base64
import http.client
import json
import os
from urllib.parse import urlsplit
from pathlib import Path
from typing import Any, Dict, List, Optional

class TeamsConnectionError(Exception):
    """Raised when Teams is not running or the CDP port 9222 is closed."""
    pass

def get_active_target() -> Dict[str, Any]:
    """
    Connects to CDP endpoint on port 9222 and selects the active Microsoft Teams page.
    """
    try:
        connection = http.client.HTTPConnection("127.0.0.1", 9222, timeout=3.0)
        try:
            connection.request("GET", "/json", headers={"User-Agent": "TeamsTools/1.0"})
            response = connection.getresponse()
            if response.status != 200:
                raise TeamsConnectionError(f"CDP endpoint returned HTTP {response.status}.")
            targets = json.loads(response.read().decode("utf-8"))
        finally:
            connection.close()
    except Exception as exc:
        raise TeamsConnectionError(
            "Microsoft Teams is not running or port 9222 is not available.\n"
            "Please ensure Microsoft Teams is open, and WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS contains '--remote-debugging-port=9222'."
        ) from exc

    candidates = [
        t for t in targets
        if t.get("type") == "page"
        and "teams.microsoft.com" in t.get("url", "")
        and "deepLink=default" not in t.get("url", "")
        and "about:blank" not in t.get("url", "")
    ]

    if not candidates:
        # Fallback to any teams page
        candidates = [t for t in targets if t.get("type") == "page" and "teams.microsoft.com" in t.get("url", "")]

    if not candidates:
        raise TeamsConnectionError("No active Microsoft Teams window found on port 9222.")

    # Sort to pick the primary UI window (longest title / contains 'Microsoft Teams')
    candidates.sort(key=lambda x: (
        "| Microsoft Teams" in x.get("title", ""),
        len(x.get("title", ""))
    ), reverse=True)

    return candidates[0]

async def evaluate_js(expression: str, await_promise: bool = True, timeout: float = 60.0) -> Any:
    """
    Executes a JavaScript expression inside the active Microsoft Teams WebView2 context.
    """
    import websockets

    target = get_active_target()
    ws_url = target.get("webSocketDebuggerUrl")
    if not ws_url:
        raise TeamsConnectionError("Target does not provide a webSocketDebuggerUrl.")

    async with websockets.connect(ws_url, max_size=50 * 1024 * 1024) as ws:
        msg_id = 1
        payload = {
            "id": msg_id,
            "method": "Runtime.evaluate",
            "params": {
                "expression": expression,
                "awaitPromise": await_promise,
                "returnByValue": True
            }
        }
        await ws.send(json.dumps(payload))
        
        raw_resp = await asyncio.wait_for(ws.recv(), timeout=timeout)
        data = json.loads(raw_resp)

        if "error" in data:
            raise RuntimeError(f"CDP Error: {data['error']}")

        result_obj = data.get("result", {})
        if result_obj.get("exceptionDetails"):
            exc = result_obj["exceptionDetails"]
            raise RuntimeError(f"JavaScript Exception: {exc.get('text')} - {exc.get('exception', {}).get('description')}")

        return result_obj.get("result", {}).get("value")

async def list_chats(limit: int = 30, query: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Retrieves the list of recent chats (1:1, groups, meetings) from Teams IndexedDB.
    """
    query_str = json.dumps(query.lower() if query else "")
    js_code = f"""
    (async () => {{
        const dbs = await indexedDB.databases();
        const convDb = dbs.find(d => d.name && d.name.startsWith('Teams:conversation-manager:') && d.name.includes('b527dc1b-39af-4083-9b77-1060c0f4dbf9'))
                     || dbs.find(d => d.name && d.name.startsWith('Teams:conversation-manager:'));
        if (!convDb) return {{ error: 'Conversation database not found in IndexedDB.' }};

        return new Promise((resolve) => {{
            const req = indexedDB.open(convDb.name);
            req.onsuccess = (e) => {{
                const db = e.target.result;
                if (!db.objectStoreNames.contains('conversations')) {{
                    resolve({{ error: 'Store conversations not found.' }});
                    return;
                }}
                const tx = db.transaction('conversations', 'readonly');
                const store = tx.objectStore('conversations');
                const getAllReq = store.getAll();
                getAllReq.onsuccess = (ev) => {{
                    const list = ev.target.result || [];
                    const q = {query_str};

                    const chats = [];
                    for (const c of list) {{
                        // Filter out non-chat thread types if needed
                        let title = c.chatTitle?.shortTitle || c.chatTitle || c.threadProperties?.topic || '';
                        if (typeof title === 'object') {{
                            title = title.shortTitle || title.longTitle || '';
                        }}
                        
                        let participants = [];
                        if (c.chatTitle?.avatarUsersInfo) {{
                            participants = c.chatTitle.avatarUsersInfo.map(u => ({{
                                name: u.displayName,
                                email: u.email,
                                mri: u.mri
                            }}));
                        }}

                        if (!title && participants.length > 0) {{
                            title = participants.map(p => p.name).join(', ');
                        }}

                        const lm = c.lastMessage || c.lastNonTMMessage || {{}};
                        let preview = lm.content || '';
                        preview = preview.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/\\r|\\n/g, ' ').trim();
                        if (preview.length > 120) preview = preview.substring(0, 120) + '...';

                        const timestamp = c.lastMessageTimeUtc || lm.originalarrivaltime || c.clientUpdateTime || 0;
                        const sender = lm.imdisplayname || lm.from || '';

                        const item = {{
                            id: c.id,
                            title: title || 'Без названия',
                            type: c.type || 'Chat',
                            last_sender: sender,
                            last_message: preview,
                            last_timestamp: timestamp,
                            participants: participants
                        }};

                        if (q) {{
                            const searchCorpus = (item.title + ' ' + item.last_message + ' ' + item.last_sender + ' ' + JSON.stringify(participants)).toLowerCase();
                            if (!searchCorpus.includes(q)) continue;
                        }}

                        chats.push(item);
                    }}

                    // Sort newest first
                    chats.sort((a, b) => {{
                        const tA = Number(a.last_timestamp) || 0;
                        const tB = Number(b.last_timestamp) || 0;
                        return tB - tA;
                    }});

                    resolve({{ chats: chats.slice(0, {limit}) }});
                }};
                getAllReq.onerror = (ev) => resolve({{ error: ev.target.error?.name }});
            }};
            req.onerror = (e) => resolve({{ error: e.target.error?.name }});
        }});
    }})()
    """
    res = await evaluate_js(js_code)
    if isinstance(res, dict) and "error" in res:
        raise RuntimeError(f"Failed to list chats: {res['error']}")
    return res.get("chats", []) if isinstance(res, dict) else []

async def get_chat_history(chat_id: str, limit: int = 50) -> List[Dict[str, Any]]:
    """
    Reads message history and attachments for a specific conversation from Teams IndexedDB.
    """
    escaped_chat_id = json.dumps(chat_id)
    js_code = f"""
    (async () => {{
        const targetId = {escaped_chat_id};
        const dbs = await indexedDB.databases();
        const rcDb = dbs.find(d => d.name && d.name.startsWith('Teams:replychain-manager:'));
        if (!rcDb) return {{ error: 'Replychain database not found.' }};

        return new Promise((resolve) => {{
            const req = indexedDB.open(rcDb.name);
            req.onsuccess = (e) => {{
                const db = e.target.result;
                const storeName = db.objectStoreNames.contains('replychains') ? 'replychains' : db.objectStoreNames[0];
                const tx = db.transaction(storeName, 'readonly');
                const store = tx.objectStore(storeName);
                const getReq = store.getAll();

                getReq.onsuccess = (ev) => {{
                    const items = ev.target.result || [];
                    const messages = [];

                    for (const it of items) {{
                        if (it.conversationId !== targetId) continue;
                        const msgList = it.messages || Object.values(it.messageMap || {{}});
                        for (const m of msgList) {{
                            const rawContent = m.content || '';
                            let text = rawContent.replace(/<img[^>]+itemtype=["'][^"']*(?:Emoji|Emoticon)[^"']*["'][^>]*alt=["']?([^"'>]*)["']?[^>]*>/gi, '$1 ')
                                                 .replace(/<[^>]*>/g, '')
                                                 .replace(/&nbsp;/g, ' ')
                                                 .replace(/\\u00a0/g, ' ')
                                                 .trim();
                            let files = [];

                            if (rawContent.includes('<img')) {{
                                const imgRegex = /<img\\s+[^>]*>/gi;
                                let mTag;
                                while ((mTag = imgRegex.exec(rawContent)) !== null) {{
                                    const tag = mTag[0];
                                    if (tag.includes('schema.skype.com/Emoji') || 
                                        tag.includes('schema.skype.com/Emoticon') ||
                                        tag.includes('statics.teams.cdn.office.net') ||
                                        tag.includes('/emoticons/')) {{
                                        continue;
                                    }}
                                    const srcMatch = tag.match(/src=["']([^"']+)["']/i);
                                    const altMatch = tag.match(/alt=["']([^"']*)["']/i);
                                    if (srcMatch) {{
                                        files.push({{
                                            name: (altMatch ? altMatch[1] : '') || 'embedded_image.png',
                                            type: 'image',
                                            size: 0,
                                            url: srcMatch[1]
                                        }});
                                    }}
                                }}
                            }}

                            try {{
                                if (m.properties?.files && typeof m.properties.files === 'string') {{
                                    const parsedFiles = JSON.parse(m.properties.files);
                                    for (const f of parsedFiles) {{
                                        files.push({{
                                            name: f.fileName || f.title || 'unnamed',
                                            type: f.fileType || f.type || '',
                                            size: f.fileSize || 0,
                                            url: f.botFileProperties?.url || f.objectUrl || f.fileInfo?.fileUrl || ''
                                        }});
                                    }}
                                }}
                            }} catch(err) {{}}

                            const timestamp = m.originalArrivalTime || m.clientArrivalTime || m.composetime || m.timestamp || 0;
                            const sender = m.imDisplayName || m.fromDisplayNameInToken || m.imdisplayname || m.from || m.creator || 'Unknown';

                            messages.push({{
                                id: m.id,
                                sender: sender,
                                text: text,
                                raw_html: rawContent,
                                timestamp: timestamp,
                                attachments: files
                            }});
                        }}
                    }}

                    // Sort chronologically ascending
                    messages.sort((a, b) => {{
                        const tA = Number(a.timestamp) || 0;
                        const tB = Number(b.timestamp) || 0;
                        return tA - tB;
                    }});

                    resolve({{
                        chat_id: targetId,
                        total_messages: messages.length,
                        messages: messages.slice(-{limit})
                    }});
                }};
                getReq.onerror = (ev) => resolve({{ error: ev.target.error?.name }});
            }};
            req.onerror = (e) => resolve({{ error: e.target.error?.name }});
        }});
    }})()
    """
    res = await evaluate_js(js_code)
    if isinstance(res, dict) and "error" in res:
        raise RuntimeError(f"Failed to read chat history: {res['error']}")
    return res.get("messages", []) if isinstance(res, dict) else []

async def download_attachment_to_file(download_url: str, output_path: str) -> str:
    """
    Downloads a chat attachment using active Teams authenticated session and saves locally.
    """
    max_bytes = 20 * 1024 * 1024
    parsed = urlsplit(download_url)
    hostname = (parsed.hostname or "").lower()
    trusted_domains = ("teams.microsoft.com", "sharepoint.com", "office.com", "office.net")
    if parsed.scheme != "https" or not any(
        hostname == domain or hostname.endswith("." + domain) for domain in trusted_domains
    ):
        raise ValueError("Attachment URL must use HTTPS on a trusted Microsoft or SharePoint host.")
    out = Path(output_path).resolve()
    if out.exists():
        raise FileExistsError(f"Refusing to overwrite existing attachment: {out}")
    escaped_url = json.dumps(download_url)
    js_code = f"""
    (async () => {{
        const url = {escaped_url};
        try {{
            const resp = await fetch(url, {{ credentials: 'include', redirect: 'manual' }});
            if (resp.type === 'opaqueredirect' || (resp.status >= 300 && resp.status < 400)) {{
                return {{ error: 'Attachment redirect is not allowed.' }};
            }}
            const finalUrl = new URL(resp.url);
            const allowed = ['teams.microsoft.com', 'sharepoint.com', 'office.com', 'office.net'];
            if (finalUrl.protocol !== 'https:' || !allowed.some(d => finalUrl.hostname === d || finalUrl.hostname.endsWith('.' + d))) {{
                return {{ error: 'Download redirected to an untrusted host.' }};
            }}
            if (!resp.ok) return {{ error: 'HTTP ' + resp.status + ' ' + resp.statusText }};
            const maxBytes = {max_bytes};
            const declaredSize = Number(resp.headers.get('content-length'));
            if (declaredSize > maxBytes) return {{ error: 'Attachment exceeds size limit.' }};
            const contentType = resp.headers.get('content-type') || '';
            if (contentType.toLowerCase().startsWith('text/html')) return {{ error: 'Attachment returned an HTML page.' }};
            const readerStream = resp.body?.getReader();
            if (!readerStream) return {{ error: 'Attachment response has no readable body.' }};
            const chunks = [];
            let total = 0;
            while (true) {{
                const {{ done, value }} = await readerStream.read();
                if (done) break;
                total += value.byteLength;
                if (total > maxBytes) {{ await readerStream.cancel(); return {{ error: 'Attachment exceeds size limit.' }}; }}
                chunks.push(value);
            }}
            const blob = new Blob(chunks);
            return new Promise((resolve) => {{
                const reader = new FileReader();
                reader.onloadend = () => {{
                    const base64data = reader.result.split(',')[1];
                    resolve({{ success: true, base64: base64data, size: blob.size }});
                }};
                reader.onerror = () => resolve({{ error: 'FileReader error' }});
                reader.readAsDataURL(blob);
            }});
        }} catch (err) {{
            return {{ error: String(err) }};
        }}
    }})()
    """
    res = await evaluate_js(js_code, timeout=60.0)
    if not isinstance(res, dict) or not res.get("success"):
        raise RuntimeError(f"Download failed: {res.get('error') if isinstance(res, dict) else 'Unknown error'}")

    encoded = res.get("base64")
    size = res.get("size")
    if not isinstance(encoded, str) or not isinstance(size, int) or not 0 <= size <= max_bytes or len(encoded) > 4 * ((max_bytes + 2) // 3):
        raise RuntimeError("Download response has invalid size or encoding.")
    try:
        raw_bytes = base64.b64decode(encoded, validate=True)
    except (ValueError, base64.binascii.Error) as exc:
        raise RuntimeError("Download response has invalid base64.") from exc
    if len(raw_bytes) != size:
        raise RuntimeError("Download size does not match response.")
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("xb") as output:
        output.write(raw_bytes)
    return str(out)

async def send_message_to_active_chat(text: str, expected_chat_id: Optional[str] = None) -> Dict[str, Any]:
    """
    Types text into the active Teams message composer using CKEditor instance and triggers the Send button.
    """
    escaped_text = json.dumps(text)
    escaped_chat_id = json.dumps(expected_chat_id)
    js_code = f"""
    (() => {{
        const expectedId = {escaped_chat_id};
        if (expectedId && decodeURIComponent(window.location.hash).split(/[/?#]/).includes(expectedId) === false) {{
            return {{ success: false, error: 'Active chat does not match requested chat ID.' }};
        }}
        const editor = document.querySelector('[data-tid="ckeditor"]');
        if (!editor) return {{ success: false, error: 'Message editor not found in active window.' }};

        const rawText = {escaped_text};
        // Escape HTML entities for the paragraph
        const safeHtml = rawText
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/\\n/g, '<br>');

        if (editor.ckeditorInstance && typeof editor.ckeditorInstance.setData === 'function') {{
            editor.ckeditorInstance.setData('<p>' + safeHtml + '</p>');
        }} else {{
            editor.focus();
            document.execCommand('insertText', false, rawText);
            editor.dispatchEvent(new Event('input', {{ bubbles: true }}));
        }}

        // Locate Send button
        const buttons = Array.from(document.querySelectorAll('button'));
        const sendBtn = buttons.find(b => {{
            const tid = (b.getAttribute('data-tid') || '').toLowerCase();
            const label = (b.getAttribute('aria-label') || '').toLowerCase();
            return tid === 'sendmessagecommands-send' || label.startsWith('send') || label.startsWith('отправить');
        }});

        if (!sendBtn) {{
            return {{ success: false, error: 'Send button not found.' }};
        }}

        if (sendBtn.disabled) {{
            return {{ success: false, error: 'Send button is disabled.' }};
        }}

        sendBtn.click();
        return {{ success: true, textLength: rawText.length }};
    }})()
    """
    res = await evaluate_js(js_code, await_promise=False)
    if not isinstance(res, dict) or res.get("success") is not True:
        raise RuntimeError(f"Failed to send message: {res.get('error') if isinstance(res, dict) else 'Unexpected CDP response'}")
    return res
