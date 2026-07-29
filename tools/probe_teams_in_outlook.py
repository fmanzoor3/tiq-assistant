"""Read-only probe: can we reach Teams chat messages via Outlook COM?

Some Microsoft 365 tenants mirror Teams 1:1 / group chats into a hidden Outlook
mailbox folder (commonly "Conversation History" or "Team Chat"). If Enerjisa's
tenant does this, TIQ Assistant could read the day's Teams messages for FREE
using the same local Outlook COM access it already uses for the calendar -- no
Graph API, no IT app registration, no scraping.

This script ONLY reads. It:
  - connects to local Outlook (read-only, same as the calendar reader),
  - walks the mailbox folder tree looking for likely chat folders,
  - reports whether any were found, how many items they hold, and a small,
    redacted sample (sender + first ~80 chars) from today so you can judge
    whether the content is actually useful.

It writes nothing, sends nothing over the network, and modifies no items.

RUN THIS ON YOUR WORK LAPTOP (where your Enerjisa Outlook account is), e.g.:

    cd "C:\\path\\to\\tiq-assistant"
    python tools\\probe_teams_in_outlook.py

Then paste the output back so we can decide the retrieval approach.
"""

from __future__ import annotations

from datetime import datetime, date, timedelta


# Outlook item class names we care about. Teams-mirrored chat items usually show
# up as IPM.Note or IPM.SkypeTeams.Message depending on tenant/version.
CHAT_FOLDER_HINTS = ("conversation history", "team chat", "teams", "chat")


def _connect():
    import win32com.client
    import pythoncom

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    return namespace


def _walk_folders(folder, depth=0, max_depth=4):
    """Yield (folder, depth) for folder and its subfolders, depth-limited."""
    yield folder, depth
    if depth >= max_depth:
        return
    try:
        for sub in folder.Folders:
            yield from _walk_folders(sub, depth + 1, max_depth)
    except Exception:
        return


def _safe(attr_getter, default=""):
    try:
        return attr_getter() or default
    except Exception:
        return default


def main() -> int:
    print("=" * 70)
    print(" TIQ Assistant — Teams-in-Outlook probe (READ ONLY)")
    print("=" * 70)

    try:
        ns = _connect()
    except Exception as e:
        print(f"\n[!] Could not connect to Outlook: {e}")
        print("    Make sure Outlook desktop is installed and configured.")
        return 1

    # Enumerate all top-level stores (mailboxes) and their folder trees.
    candidate_folders = []
    all_folder_names = []

    try:
        for store_idx in range(1, ns.Folders.Count + 1):
            root = ns.Folders.Item(store_idx)
            for folder, depth in _walk_folders(root):
                name = _safe(lambda: folder.Name)
                all_folder_names.append("  " * depth + name)
                if any(hint in name.lower() for hint in CHAT_FOLDER_HINTS):
                    candidate_folders.append(folder)
    except Exception as e:
        print(f"\n[!] Error while walking folders: {e}")

    print(f"\nScanned mailbox folder tree. "
          f"Found {len(candidate_folders)} candidate chat folder(s).")

    if not candidate_folders:
        print("\n[RESULT] No Teams/chat-like folders found in Outlook.")
        print("         => The free Outlook path is NOT available on this tenant.")
        print("         We'll need Microsoft Graph (IT) or manual paste instead.")
        print("\n(First 40 folders seen, for reference:)")
        for line in all_folder_names[:40]:
            print("   " + line)
        return 0

    today = date.today()
    start = datetime.combine(today, datetime.min.time())
    end = datetime.combine(today + timedelta(days=1), datetime.min.time())

    found_any_messages = False

    for folder in candidate_folders:
        fname = _safe(lambda: folder.Name)
        try:
            items = folder.Items
            total = _safe(lambda: items.Count, 0)
        except Exception as e:
            print(f"\n- Folder '{fname}': could not read items ({e})")
            continue

        print(f"\n- Folder: '{fname}'  (total items: {total})")

        # Try to sample today's items.
        sample_count = 0
        try:
            items.Sort("[ReceivedTime]", True)
            it = items.GetFirst()
            guard = 0
            while it is not None and guard < 500:
                guard += 1
                received = None
                try:
                    r = it.ReceivedTime
                    received = datetime(r.year, r.month, r.day, r.hour, r.minute)
                except Exception:
                    pass

                if received and start <= received < end:
                    sender = _safe(lambda: it.SenderName, "(unknown)")
                    body = _safe(lambda: it.Body, "")
                    snippet = body.strip().replace("\r", " ").replace("\n", " ")[:80]
                    msgclass = _safe(lambda: it.MessageClass, "?")
                    print(f"    • [{received:%H:%M}] {sender}: {snippet!r}  "
                          f"(class={msgclass})")
                    sample_count += 1
                    found_any_messages = True
                    if sample_count >= 5:
                        break
                elif received and received < start:
                    # Sorted newest-first; once we pass today, stop.
                    break
                it = items.GetNext()
        except Exception as e:
            print(f"    (could not sample items: {e})")

        if sample_count == 0:
            print("    (no items dated today — folder may still hold older chats)")

    print("\n" + "=" * 70)
    if found_any_messages:
        print("[RESULT] Chat-like messages ARE readable via Outlook COM.")
        print("         => We can likely auto-pull the day's Teams messages for")
        print("            free, using the plumbing the app already has.")
        print("         Paste this output back so we can confirm the format.")
    else:
        print("[RESULT] Found chat-like folders but no readable messages for today.")
        print("         Send this output back — we may adjust the folder/date logic,")
        print("         or fall back to Graph API / manual paste.")
    print("=" * 70)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
