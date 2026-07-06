"""The brain: the LLM agent loop (run_task), conversational chat (Ask Maia), and the live
on-screen guide (guide_live / guide)."""
import os
import re
import json
import time

from openai import OpenAI

from axon import config
from axon.config import MODEL, MAX_STEPS, TOOL_TEXT_LIMIT, GUIDE_MODEL, GUIDE_REFINE
from axon.prompts import (
    SYSTEM_PROMPT, CHAT_SYSTEM_PROMPT, BROWSE_STRATEGY,
    GUIDE_SYSTEM_PROMPT, GUIDE_LIVE_SYSTEM_PROMPT, GUIDE_REFINE_PROMPT,
)
from axon.tools import TOOLS, DISPATCH
from axon import approval
from axon import memory
from axon.outlook import my_tone
from axon.context import active_context_note
from axon.llm import text_llm
from axon.util import scrub_identity


def _tone_block():
    """Inject the learned writing style so emails Axon drafts sound like the user."""
    guide = my_tone()
    if not guide:
        return ""
    return ("\n\nWHEN DRAFTING OR REPLYING TO EMAILS, write in the user's own voice. Their style:\n"
            + guide + "\n")
from axon.approval import _needs_approval, _describe_action
from axon.util import _result
from axon.mcp import _extract, MCPClient, _close_app
from axon.browse import _ensure_debug_chrome
from axon.screen import (
    _grab_screen_b64, _grab_grid_b64, _grab_dot_monitor_img, _dot_monitor_crop,
    _screen_signature, _screens_differ, _foreground_window, _virtual_screen,
    _img_b64, _grid_b64, _take_screenshot,
)

def chat(question, on_status=None, image_path=None, history=None):
    """'Ask Maia' mode: a direct conversational answer from the LLM — no tools, no desktop actions.
    Supports an attached screenshot (vision) and prior turns (history) for follow-up context.
    history is a list of {"role": "user"|"assistant", "content": "..."} from earlier in the chat."""
    if not os.getenv("OPENAI_API_KEY"):
        return "Axon isn't set up with an API key yet."
    if on_status:
        on_status("Maia is thinking...")
    has_image = bool(image_path and os.path.isfile(image_path))
    # Text-only chat routes to Mistral (cheaper); a screenshot needs vision -> OpenAI.
    client, model = (OpenAI(), MODEL) if has_image else text_llm()
    content = question
    if image_path and os.path.isfile(image_path):
        try:
            import base64
            import mimetypes
            with open(image_path, "rb") as _f:
                _b64 = base64.b64encode(_f.read()).decode("ascii")
            _mime = mimetypes.guess_type(image_path)[0] or "image/png"
            content = [
                {"type": "text", "text": (question or "Describe and answer about this screenshot.").strip()},
                {"type": "image_url", "image_url": {"url": f"data:{_mime};base64,{_b64}"}},
            ]
        except Exception:
            content = question
    _ctx = active_context_note(question)   # name the open file(s) if they said "this file"
    if _ctx:
        if isinstance(content, list):
            content[0]["text"] = content[0].get("text", "") + _ctx
        else:
            content = (content or "") + _ctx
    messages = [{"role": "system", "content": CHAT_SYSTEM_PROMPT + memory.context_block() + _tone_block()}]
    if history:
        messages.extend(history)            # prior turns, so follow-ups keep context
    messages.append({"role": "user", "content": content})
    try:
        resp = client.chat.completions.create(model=model, messages=messages, temperature=0.4)
        return scrub_identity((resp.choices[0].message.content or "").strip()) or "(no answer)"
    except Exception as e:
        return scrub_identity(f"Error: {e}")


def run_task(question, on_status=None, should_cancel=None, on_approval=None, image_path=None, on_plan=None):
    """Run a natural-language task. Calls on_status(str) with progress; returns final summary.
    If on_approval is given, it is called as on_approval(description)->bool before any action that
    sends/changes data, and the action runs only if it returns True (approval mode).
    If image_path is given (a screenshot the user attached in the composer), it is shown to the
    vision model alongside the question so the user can ask about what's on their screen."""
    def status(msg):
        if on_status:
            on_status(msg)

    def cancelled():
        return bool(should_cancel and should_cancel())

    def stop_result():
        status("[stopped] Stopped by user.")
        return "Stopped."

    if not os.getenv("OPENAI_API_KEY"):
        msg = "Axon isn't set up with an API key yet."
        status(msg)
        return msg

    client = OpenAI()
    mcp = MCPClient()
    approval._APPROVAL_CB = on_approval  # tools that draft-then-send use this
    # If the user attached a screenshot, send it to the vision model with the question so they can
    # ask about what's on their screen (and the agent can still act on it if asked).
    user_content = question
    if image_path and os.path.isfile(image_path):
        try:
            import base64
            import mimetypes
            with open(image_path, "rb") as _f:
                _b64 = base64.b64encode(_f.read()).decode("ascii")
            _mime = mimetypes.guess_type(image_path)[0] or "image/png"
            user_content = [
                {"type": "text", "text": (question or "").strip()
                    + "\n\n[The user attached a screenshot, saved at: " + image_path + " . If they are "
                      "asking ABOUT it, answer directly from the image. If they ask to email/send/save/"
                      "attach it, use this exact file path (e.g. pass it in send_email's attachments) — "
                      "do NOT take a new screenshot. Otherwise use it as context for the task.]"},
                {"type": "image_url", "image_url": {"url": f"data:{_mime};base64,{_b64}"}},
            ]
        except Exception:
            user_content = question
    _ctx = active_context_note(question)   # name the open file(s) if they said "this file"
    if _ctx:
        if isinstance(user_content, list):
            user_content[0]["text"] = user_content[0].get("text", "") + _ctx
        else:
            user_content = (user_content or "") + _ctx
    messages = [
        {"role": "system", "content": SYSTEM_PROMPT + memory.context_block() + _tone_block()},
        {"role": "user", "content": user_content},
    ]
    status("Axon intelligence is thinking...")
    try:
        for _ in range(MAX_STEPS):
            if cancelled():
                return stop_result()
            resp = client.chat.completions.create(
                model=MODEL, messages=messages, tools=TOOLS, tool_choice="auto", temperature=0
            )
            if cancelled():
                return stop_result()
            msg = resp.choices[0].message

            assistant_msg = {"role": "assistant", "content": msg.content or ""}
            if msg.tool_calls:
                assistant_msg["tool_calls"] = [
                    {"id": tc.id, "type": "function",
                     "function": {"name": tc.function.name, "arguments": tc.function.arguments}}
                    for tc in msg.tool_calls
                ]
            messages.append(assistant_msg)

            if not msg.tool_calls:
                final = scrub_identity(msg.content or "Done.")
                status("[done] " + final)
                return final

            pending_images = []
            for tc in msg.tool_calls:
                if cancelled():
                    return stop_result()
                name = tc.function.name
                try:
                    args = json.loads(tc.function.arguments or "{}")
                except Exception:
                    args = {}
                short = ", ".join(f"{k}={v}" for k, v in args.items() if k != "app")
                status(f"-> {name}({short})" if short else f"-> {name}")

                # Approval gate: pause for the user's OK before sending/changing data.
                declined = False
                if on_approval and _needs_approval(name, args):
                    if not on_approval(_describe_action(name, args)):
                        declined = True
                if cancelled():
                    return stop_result()

                if declined:
                    desc = _describe_action(name, args)
                    status(f"[skipped] {desc}")
                    result = _result(
                        f"The user DECLINED this action, so it was NOT performed: {desc}. "
                        f"Do not retry it; continue with anything else they asked, or stop and report.", False)
                elif name == "update_todos":
                    todos = args.get("todos") or []
                    if on_plan:
                        on_plan(todos)   # render as a live checklist panel in the UI
                    ndone = sum(1 for t in todos if t.get("done"))
                    result = _result("Checklist updated: %d of %d steps done." % (ndone, len(todos)))
                elif name == "close_app":
                    result = _close_app(mcp, args)  # needs the live MCP client
                elif name in DISPATCH:
                    result = DISPATCH[name](args)
                else:
                    result = mcp.call(name, args)  # fall back to open-computer-use desktop control

                if cancelled():
                    return stop_result()
                text, image = _extract(result)
                messages.append({"role": "tool", "tool_call_id": tc.id, "content": text[:TOOL_TEXT_LIMIT]})
                if image:
                    pending_images.append((name, image))

            for name, image in pending_images:
                if cancelled():
                    return stop_result()
                messages.append({
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"Screenshot after {name}:"},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image}"}},
                    ],
                })

        status("Reached step limit without finishing.")
        return "Reached step limit without finishing."
    finally:
        approval._APPROVAL_CB = None
        mcp.close()


_ELEM_RE = re.compile(
    r"^\s*(\d+)\s+(.*?)\s+(?:Secondary Actions:.*?)?Frame:\s*\{x:\s*(-?\d+),\s*y:\s*(-?\d+),"
    r"\s*width:\s*(\d+),\s*height:\s*(\d+)\}", re.I)


def _parse_elements(text):
    """Parse get_app_state's tree into {index: (label, x, y, w, h)} (absolute screen px)."""
    elems = {}
    for ln in text.splitlines():
        m = _ELEM_RE.match(ln)
        if not m:
            continue
        idx = int(m.group(1))
        label = re.sub(r"\s+", " ", m.group(2)).strip()[:60]
        elems[idx] = (label, int(m.group(3)), int(m.group(4)), int(m.group(5)), int(m.group(6)))
    return elems


def _guide_decide(question, history, b64):
    client = OpenAI()
    done_txt = "; ".join(history) if history else "(none yet)"
    # The DECIDE pass must pick the RIGHT element (needs strong reasoning), so use the main model.
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": GUIDE_LIVE_SYSTEM_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": f"GOAL: {question or 'Help me with what is on my screen.'}\n"
                                         f"Steps already guided: {done_txt}"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    )
    return json.loads(resp.choices[0].message.content or "{}")


def _guide_refine(b64, target):
    client = OpenAI()
    resp = client.chat.completions.create(
        model=GUIDE_MODEL,
        temperature=0,
        response_format={"type": "json_object"},
        messages=[
            {"role": "system", "content": GUIDE_REFINE_PROMPT},
            {"role": "user", "content": [
                {"type": "text", "text": f"Find and tightly box this element: {target}"},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            ]},
        ],
    )
    return json.loads(resp.choices[0].message.content or "{}")


def _refine_marker(img, data):
    """Coarse box -> zoom into that region with a fresh grid -> precise box -> monitor-fraction marker."""
    box1 = data.get("box")
    if not GUIDE_REFINE or not (isinstance(box1, (list, tuple)) and len(box1) == 4):
        return _marker_from(data)
    try:
        x1, x2 = sorted((float(box1[0]), float(box1[2])))
        y1, y2 = sorted((float(box1[1]), float(box1[3])))
        cx, cy = (x1 + x2) / 2, (y1 + y2) / 2
        hw = max((x2 - x1), 130)  # zoom region half-size in 0-1000 units (>=13% of monitor)
        hh = max((y2 - y1), 130)
        rl, rr = max(0.0, cx - hw), min(1000.0, cx + hw)
        rt, rb = max(0.0, cy - hh), min(1000.0, cy + hh)
        W, H = img.size
        crop = img.crop((int(rl / 1000 * W), int(rt / 1000 * H), int(rr / 1000 * W), int(rb / 1000 * H)))
        # Standard grid on the MAGNIFIED crop already resolves to ~0.5% of the full monitor per line
        # (the crop is ~25% of the screen), so it's precise AND readable.
        ref = _guide_refine(_grid_b64(crop), str(data.get("label") or "the target element"))
        b2 = ref.get("box")
        if isinstance(b2, (list, tuple)) and len(b2) == 4:
            bx1, bx2 = sorted((float(b2[0]), float(b2[2])))
            by1, by2 = sorted((float(b2[1]), float(b2[3])))
            mx1 = rl + bx1 / 1000 * (rr - rl)
            mx2 = rl + bx2 / 1000 * (rr - rl)
            my1 = rt + by1 / 1000 * (rb - rt)
            my2 = rt + by2 / 1000 * (rb - rt)
            pad = 3  # tight margin; let small buttons get small brackets
            return {"type": "box",
                    "fx": (mx1 - pad) / 1000.0, "fy": (my1 - pad) / 1000.0,
                    "fw": max(mx2 - mx1 + 2 * pad, 12) / 1000.0,
                    "fh": max(my2 - my1 + 2 * pad, 9) / 1000.0,  # allow short height for small buttons
                    "label": str(data.get("label") or "")}
    except Exception:
        pass
    return _marker_from(data)


def _marker_from(data):
    """Build a 'Click here' callout marker from the model's approximate grid point [x, y]."""
    pt = data.get("point")
    if isinstance(pt, (list, tuple)) and len(pt) == 2:
        try:
            return {"type": "clickhere",
                    "fx": min(max(float(pt[0]) / 1000.0, 0.0), 1.0),
                    "fy": min(max(float(pt[1]) / 1000.0, 0.0), 1.0),
                    "label": "Click here"}
        except Exception:
            pass
    return None


def guide_live(question, on_step=None, should_cancel=None, max_steps=25):
    """Live coaching with CLEAR TEXT instructions (no on-screen marker): describe the next thing to
    click using landmarks, wait for the user to act, then guide the next step.

    Calls on_step({"instruction", "marker": None, "done": bool}) for each step.
    """
    def cancelled():
        return bool(should_cancel and should_cancel())

    def emit(instruction, marker, done):
        if on_step:
            on_step({"instruction": instruction, "marker": marker, "done": done})

    if not os.getenv("OPENAI_API_KEY"):
        emit("Axon isn't set up with an API key yet.", None, True)
        return

    history = []
    for _ in range(max_steps):
        if cancelled():
            return
        try:
            img = _grab_dot_monitor_img()
        except Exception as e:
            emit(f"Could not capture the screen: {e}", None, True)
            return
        try:
            data = _guide_decide(question, history, _img_b64(img))
        except Exception as e:
            emit(scrub_identity(f"Guidance failed: {e}"), None, True)
            return
        if cancelled():
            return
        instruction = str(data.get("instruction") or "").strip()
        done = bool(data.get("done"))
        # Text-only guidance — no on-screen marker (the clear description is the guide).
        emit(instruction or "(thinking...)", None, done)
        if done:
            return
        if instruction:
            history.append(instruction)
        # Wait for the user to act, then for the screen to SETTLE, before re-guiding. This avoids
        # re-evaluating (and visibly moving the mark) on transient changes like hover tooltips,
        # chart animations, or a blinking cursor on dynamic pages.
        try:
            base = _screen_signature()
        except Exception:
            base = None
        waited = 0.0
        changed = False
        THRESH = 10  # ignore small/transient changes; only react to a real navigation/click
        while waited < 90 and not cancelled():
            time.sleep(0.4)
            waited += 0.4
            try:
                cur = _screen_signature()
            except Exception:
                break
            if not base:
                break
            if not changed:
                if _screens_differ(base, cur, thresh=THRESH):
                    changed = True  # the user did something; now wait for it to settle
                    base = cur
            else:
                if not _screens_differ(base, cur, thresh=THRESH):
                    break  # settled into the new state -> guide the next step
                base = cur
    if not cancelled():
        emit("That's as far as I can guide step-by-step — tell me if you're still stuck.", None, True)


def guide(question, on_status=None, should_cancel=None):
    """Coaching mode: screenshot the screen and return {steps_text, pointer:{x,y,label}|None}."""
    def status(m):
        if on_status:
            on_status(m)

    def cancelled():
        return bool(should_cancel and should_cancel())

    if not os.getenv("OPENAI_API_KEY"):
        return {"steps_text": "Axon isn't set up with an API key yet.", "pointer": None}
    status("Axon intelligence is looking at your screen...")
    if cancelled():
        return {"steps_text": "Stopped.", "pointer": None}
    try:
        b64 = _grab_screen_b64()
    except Exception as e:
        return {"steps_text": f"Could not capture the screen: {e}", "pointer": None}
    if cancelled():
        return {"steps_text": "Stopped.", "pointer": None}
    try:
        client = OpenAI()
        resp = client.chat.completions.create(
            model=MODEL,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": GUIDE_SYSTEM_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": question or "Help me with what is on my screen."},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ]},
            ],
        )
        if cancelled():
            return {"steps_text": "Stopped.", "pointer": None}
        data = json.loads(resp.choices[0].message.content or "{}")
    except Exception as e:
        return {"steps_text": scrub_identity(f"Guidance failed: {e}"), "pointer": None}
    steps = data.get("steps") or []
    steps_text = "\n".join(str(s) for s in steps) if isinstance(steps, list) else str(steps)
    pointer = data.get("pointer") if isinstance(data.get("pointer"), dict) else None
    status(steps_text or "(no steps returned)")
    return {"steps_text": steps_text, "pointer": pointer}
