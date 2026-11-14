#!/usr/bin/env python3
"""The full SlideShortcuts page library, for a person with a SlideShortcuts key.

The skill ships with a sample of library pages (assets/pages.json). The full
catalogue comes from slideshortcuts.com, only for a valid key. The key never
passes through the chat: the person types it on slideshortcuts.com in their
own browser, and this script receives a token for the conversation.

    python3 library.py start            # prints a link and a code for the person to open
    python3 library.py wait             # waits until they have signed in there (up to 10 minutes)
    python3 library.py key              # Claude Code: a token from SLIDESHORTCUTS_EMAIL and SLIDESHORTCUTS_KEY
    python3 library.py pages --out pages-full.json   # the full catalogue, for beautify_deck.py --pages
    python3 library.py status           # whether a token is held, and until when

What is kept on this computer: the token and its expiry (12 hours), and while
a sign-in is open, its code and poll secret, in ~/.slideshortcuts/ (or
SLIDESHORTCUTS_HOME). Never the key or an email. Decks are never sent.

Every refusal prints one plain sentence and exits with status 2. Standard
library only.
"""

import argparse
import json
import os
import sys
import time
import urllib.error
import urllib.request

API = os.environ.get("SLIDESHORTCUTS_API", "https://slideshortcuts.com/.netlify/functions/beautify")
HOME = os.environ.get("SLIDESHORTCUTS_HOME", os.path.join(os.path.expanduser("~"), ".slideshortcuts"))
TOKEN_FILE = os.path.join(HOME, "beautify-token.json")
SIGNIN_FILE = os.path.join(HOME, "beautify-signin.json")

SAY = {
    "not_open": "The full library is not open yet. The skill's sample pages work as they are.",
    "license_required": "That email and key were not accepted. Check them on slideshortcuts.com/account and sign in again.",
    "license_unverified": "The key could not be checked just now. Try again in a few minutes.",
    "sign_in_required": "Sign in first: run `library.py start`, open the link, then `library.py wait`.",
    "daily_limit": "This key has fetched the library as often as a day allows. It resets at midnight UTC.",
    "slow_down": "Too many tries in a minute. Wait a minute and try again.",
    "code_expired": "That sign-in code has expired. Run `library.py start` again.",
    "code_used": "That sign-in code was already used. Run `library.py start` again.",
    "bad_code": "That sign-in code is not valid. Run `library.py start` again.",
    "unreachable": "slideshortcuts.com could not be reached. Check the connection and try again.",
}


class Refused(Exception):
    def __init__(self, code):
        Exception.__init__(self, code)
        self.code = code


def call(action, body=None, token=None, timeout=30):
    data = json.dumps(dict(body or {}, action=action)).encode("utf-8")
    req = urllib.request.Request(API, data=data, method="POST", headers={"Content-Type": "application/json",
                                                                        "User-Agent": "slide-beautify-skill"})
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            out = json.loads(resp.read().decode("utf-8") or "{}")
    except urllib.error.HTTPError as err:
        try:
            out = json.loads(err.read().decode("utf-8") or "{}")
        except ValueError:
            out = {}
        raise Refused(out.get("error") or "unreachable")
    except (urllib.error.URLError, OSError, ValueError):
        raise Refused("unreachable")
    if not out.get("ok"):
        raise Refused(out.get("error") or "unreachable")
    return out


def _save(path, value):
    os.makedirs(HOME, exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(value, fh)
    os.chmod(path, 0o600)


def _load(path):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def _keep_token(out):
    _save(TOKEN_FILE, {"token": out["token"], "expires": time.time() + float(out.get("expires_in") or 0)})


def held_token():
    t = _load(TOKEN_FILE)
    return t["token"] if t and t.get("token") and t.get("expires", 0) > time.time() + 60 else None


def start():
    out = call("start")
    _save(SIGNIN_FILE, {"code": out["code"], "poll": out["poll"], "interval": out.get("interval") or 5,
                        "until": time.time() + float(out.get("expires_in") or 600)})
    print("Open %s in your browser and sign in with the email and key of your SlideShortcuts account." % out["url"])
    print("Your code is %s. Come back here when it says you are signed in." % out["code"])


def wait(limit=600):
    s = _load(SIGNIN_FILE)
    if not s:
        raise Refused("sign_in_required")
    end = min(s["until"], time.time() + limit)
    while time.time() < end:
        out = call("poll", {"code": s["code"], "poll": s["poll"]})
        if out.get("state") == "approved" and out.get("token"):
            _keep_token(out)
            os.remove(SIGNIN_FILE)
            print("Signed in. The full library is open to this conversation for 12 hours.")
            return
        time.sleep(max(1, float(s.get("interval") or 5)))
    raise Refused("code_expired")


def key():
    email, k = os.environ.get("SLIDESHORTCUTS_EMAIL", ""), os.environ.get("SLIDESHORTCUTS_KEY", "")
    if not email or not k:
        print("Set SLIDESHORTCUTS_EMAIL and SLIDESHORTCUTS_KEY in your own shell first (never paste the key into the chat).")
        sys.exit(2)
    _keep_token(call("token", {"email": email, "key": k}))
    print("Signed in. The full library is open for 12 hours.")


def pages(out_path):
    token = held_token()
    if not token:
        raise Refused("sign_in_required")
    out = call("pages", token=token, timeout=90)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump({"version": out.get("version"), "canvas": out.get("canvas"), "pages": out.get("pages") or []}, fh)
    print("%d pages written to %s (%s fetches left today)." % (len(out.get("pages") or []), out_path, out.get("remaining")))


def status():
    t = _load(TOKEN_FILE)
    if t and t.get("expires", 0) > time.time() + 60:
        print("Signed in until %s." % time.strftime("%H:%M on %d %b", time.localtime(t["expires"])))
    else:
        print("Not signed in. The skill's sample pages work without signing in.")


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("start")
    w = sub.add_parser("wait")
    w.add_argument("--limit", type=float, default=600)
    sub.add_parser("key")
    p = sub.add_parser("pages")
    p.add_argument("--out", default="pages-full.json")
    sub.add_parser("status")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "start":
            start()
        elif a.cmd == "wait":
            wait(a.limit)
        elif a.cmd == "key":
            key()
        elif a.cmd == "pages":
            pages(a.out)
        else:
            status()
    except Refused as err:
        print(SAY.get(err.code, "The library refused the request (%s)." % err.code))
        sys.exit(2)


if __name__ == "__main__":
    main()
