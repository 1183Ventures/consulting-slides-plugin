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

It talks only to https://slideshortcuts.com, and no setting, argument or
environment variable changes that: a deck with hidden instructions cannot
send the key anywhere else. Redirects are refused. `key` refuses on a computer
set to send web traffic through a proxy or to trust its own certificates
(the code sign-in still works there: the key goes through the person's own
browser, never through this script). What the server says back is printed
only when it passes a strict check (a link only to the sign-in page, a count
only as a whole number, an error only from the known list); anything else,
and any crash, prints one fixed line.

What is kept on this computer: the token and its expiry (12 hours), and while
a sign-in is open, its code and poll secret, in ~/.slideshortcuts/ (a folder
only this user can open; files created private, links never followed). Never
the key or an email. Decks are never sent.

Every refusal prints one plain sentence and exits with status 2. Standard
library only.
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request

SITE = "https://slideshortcuts.com/"
API = SITE + ".netlify/functions/beautify"
ACTIVATE = SITE + "activate/"
HOME = os.path.join(os.path.expanduser("~"), ".slideshortcuts")
TOKEN_NAME = "beautify-token.json"
SIGNIN_NAME = "beautify-signin.json"

# In-process only: the tests swap this for a fake server. No environment
# variable, argument or file can set it.
_TRANSPORT = None

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
    "redirected": "slideshortcuts.com answered with a redirect, which this script does not follow. Try again later.",
    "bad_answer": "slideshortcuts.com sent an answer this script does not understand. Try again later.",
}
FINAL = {"code_expired", "code_used", "bad_code", "license_required", "not_open"}
UNKNOWN = "The library refused the request. Try again later, or use the skill's sample pages."
CRASH = "Something went wrong reaching the full library. The skill's sample pages work as they are."
# Settings that send this script's traffic through another machine or make it
# trust other certificates: `key` refuses while any is present
PROXY_SETTINGS = ("HTTPS_PROXY", "https_proxy", "HTTP_PROXY", "http_proxy", "ALL_PROXY", "all_proxy",
                  "SSL_CERT_FILE", "SSL_CERT_DIR", "REQUESTS_CA_BUNDLE", "CURL_CA_BUNDLE")
ERROR_NAME = re.compile(r"^[a-z_]{1,32}$")
CODE = re.compile(r"^[A-Z0-9]{4}-[A-Z0-9]{4}$")


class Refused(Exception):
    def __init__(self, code):
        Exception.__init__(self, code)
        self.code = code if isinstance(code, str) and ERROR_NAME.match(code) else "bad_answer"


class _NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise Refused("redirected")


def _opener():
    return urllib.request.build_opener(_NoRedirect)


def _http(url, data, headers, timeout):
    req = urllib.request.Request(url, data=data, method="POST", headers=headers)
    try:
        with _opener().open(req, timeout=timeout) as resp:
            return resp.status, resp.read()
    except urllib.error.HTTPError as err:
        if 300 <= err.code < 400:
            raise Refused("redirected")
        return err.code, err.read()


def call(action, body=None, token=None, timeout=30):
    if not API.startswith(SITE) or not SITE.startswith("https://"):
        raise Refused("unreachable")
    data = json.dumps(dict(body or {}, action=action)).encode("utf-8")
    headers = {"Content-Type": "application/json", "User-Agent": "slide-beautify-skill"}
    if token:
        headers["Authorization"] = "Bearer " + token
    try:
        status, raw = (_TRANSPORT or _http)(API, data, headers, timeout)
        out = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw or "{}")
    except Refused:
        raise
    except (urllib.error.URLError, OSError):
        raise Refused("unreachable")
    except ValueError:
        raise Refused("bad_answer")
    if not isinstance(out, dict):
        raise Refused("bad_answer")
    if not out.get("ok") or status >= 400:
        raise Refused(out.get("error") or "bad_answer")
    return out


def _home():
    if os.path.islink(HOME):
        raise Refused("bad_answer")
    os.makedirs(HOME, mode=0o700, exist_ok=True)
    os.chmod(HOME, 0o700)
    return HOME


def _save(name, value):
    """Created private from the start, never through a link, then moved into place."""
    home = _home()
    path = os.path.join(home, name)
    tmp = os.path.join(home, ".%s.%d.tmp" % (name, os.getpid()))
    flags = os.O_WRONLY | os.O_CREAT | os.O_EXCL | getattr(os, "O_NOFOLLOW", 0)
    fd = os.open(tmp, flags, 0o600)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            json.dump(value, fh)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def _load(name):
    path = os.path.join(HOME, name)
    try:
        fd = os.open(path, os.O_RDONLY | getattr(os, "O_NOFOLLOW", 0))
        with os.fdopen(fd, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return None


def _drop(name):
    path = os.path.join(HOME, name)
    if os.path.lexists(path):
        os.remove(path)


def _keep_token(out):
    token = out.get("token")
    if not isinstance(token, str) or not re.match(r"^[A-Za-z0-9_-]{20,200}$", token):
        raise Refused("bad_answer")
    life = out.get("expires_in")
    life = life if isinstance(life, (int, float)) and 0 < life <= 86400 else 43200
    _save(TOKEN_NAME, {"token": token, "expires": time.time() + life})


def held_token():
    t = _load(TOKEN_NAME)
    return t["token"] if t and t.get("token") and t.get("expires", 0) > time.time() + 60 else None


def start():
    out = call("start")
    code, poll = out.get("code"), out.get("poll")
    if not (isinstance(code, str) and CODE.match(code) and isinstance(poll, str) and 20 <= len(poll) <= 200):
        raise Refused("bad_answer")
    life = out.get("expires_in")
    life = life if isinstance(life, (int, float)) and not isinstance(life, bool) and 0 < life <= 900 else 600
    _save(SIGNIN_NAME, {"code": code, "poll": poll, "interval": 5, "until": time.time() + life})
    # the page shown is always our own sign-in page, never a link the server sends
    print("Open %s in your browser and sign in with the email and key of your SlideShortcuts account." % ACTIVATE)
    print("Type this code there when the page asks for it: %s. Come back here when it says you are signed in." % code)


def wait(limit=600):
    s = _load(SIGNIN_NAME)
    if not s:
        raise Refused("sign_in_required")
    end = min(s["until"], time.time() + limit)
    try:
        while time.time() < end:
            out = call("poll", {"code": s["code"], "poll": s["poll"]})
            if out.get("state") == "approved" and out.get("token"):
                _keep_token(out)
                _drop(SIGNIN_NAME)
                print("Signed in. The full library is open to this conversation for 12 hours.")
                return
            time.sleep(max(1, float(s.get("interval") or 5)))
    except Refused as err:
        if err.code in FINAL:
            _drop(SIGNIN_NAME)
        raise
    _drop(SIGNIN_NAME)
    raise Refused("code_expired")


def _proxied():
    """Whether this computer would send the request through a proxy, or trust certificates of its own."""
    if any(os.environ.get(n) for n in PROXY_SETTINGS):
        return True
    try:
        return bool({k: v for k, v in urllib.request.getproxies().items() if k in ("https", "http", "all") and v})
    except Exception:
        return True


def key():
    if _proxied():
        print("This computer sends web traffic through a proxy or trusts its own certificates, so `key` will not send "
              "your key from here. Use the code sign-in instead: run `library.py start` and open the link; your key "
              "goes through your own browser, never through this script.")
        sys.exit(2)
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
    pages_list = out.get("pages")
    if not isinstance(pages_list, list):
        raise Refused("bad_answer")
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump({"version": out.get("version"), "canvas": out.get("canvas"), "pages": pages_list}, fh)
    left = out.get("remaining")
    left = ("%d fetches left today" % left) if isinstance(left, int) and not isinstance(left, bool) and 0 <= left < 10000 else "fetches left today unknown"
    print("%d pages written to %s (%s)." % (len(pages_list), out_path, left))


def status():
    t = _load(TOKEN_NAME)
    if t and isinstance(t.get("expires"), (int, float)) and t["expires"] > time.time() + 60:
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
        print(SAY.get(err.code, UNKNOWN))
        sys.exit(2)
    except SystemExit:
        raise
    except BaseException:
        # one fixed line: nothing the server sent reaches the chat
        print(CRASH)
        sys.exit(2)


if __name__ == "__main__":
    main()
