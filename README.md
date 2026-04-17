# Lessons Learned: AppleScript + Microsoft Graph API + SharePoint

A collection of non-obvious problems and their solutions encountered while
building an Apple Mail to SharePoint filing system using Python, MSAL, and
the Microsoft Graph API. Hopefully saves someone a few hours.

---

## AppleScript

### `do shell script` returns CR-separated lines, not LF

**Problem:** Output from `do shell script` uses carriage return (`\r`) as the
line separator, not linefeed (`\n`). Using `set AppleScript's text item
delimiters to linefeed` treats the entire output as a single line.

**Solution:** Use `paragraphs of someString` instead of splitting manually.
It handles CR, LF, and CRLF natively.

```applescript
set raw to do shell script "python3 ~/myscript.py"
set theLines to paragraphs of raw
repeat with aLine in theLines
    -- process each line
end repeat
```

---

### `return` is a keyword, not a value

**Problem:** `set AppleScript's text item delimiters to return` causes a
compile error at the next line because AppleScript parses `return` as a
statement keyword, not a character value.

**Solution:** Either use `ASCII character 13`, or better yet, use `paragraphs
of` (see above) which avoids the need to handle line endings manually.

---

### `repeat with x in list` gives a reference, not a value

**Problem:** In a `repeat with x in myList` loop, `x` is an object reference,
not the actual value. Comparing `x is "someString"` always returns false.

**Solution:** Use `contents of x` or `x as string` to get the actual value,
or restructure to avoid the comparison entirely.

```applescript
repeat with x in myList
    set xVal to contents of x
    if xVal is "someString" then
        -- this works
    end if
end repeat
```

---

### Negative text indices not supported

**Problem:** `text -1 thru -1 of someString` raises a runtime error.
AppleScript does not support negative indices for text ranges.

**Solution:** Use `character (length of someString) of someString` or
restructure to avoid negative indexing.

---

### Unicode box-drawing characters break `osacompile`

**Problem:** Characters like `─`, `│`, `╔`, `╗` in comments or strings cause
`osacompile` to fail with a cryptic parse error.

**Solution:** Remove all Unicode box-drawing characters from the `.applescript`
file. Use plain ASCII for any decorative borders or separators.

---

### stdout warnings from Python corrupt `do shell script` output

**Problem:** If a Python script called via `do shell script` emits warnings to
stdout (e.g. urllib3 SSL warnings), those warnings prepend to the output and
break any string matching (e.g. `starts with "OK"` fails).

**Solutions:**
1. In Python, add `warnings.filterwarnings("ignore")` at the top of the script.
2. In AppleScript, use `contains "OK:"` rather than `starts with "OK"` to be
   more tolerant of leading noise.
3. Redirect stderr: `do shell script "python3 script.py 2>/dev/null"`

---

## Microsoft Graph API / SharePoint

### `Sites.ReadWrite.All` is not enough to create lists

**Problem:** Creating a new SharePoint list via
`POST /sites/{id}/lists` returns `403 accessDenied` even when the
authenticated user is the site owner and the token includes
`Sites.ReadWrite.All`.

**Root cause:** `Sites.ReadWrite.All` covers reading and writing *items* in
existing lists. Creating new lists requires `Sites.Manage.All`.

**Solution:** Add `Sites.Manage.All` to your MSAL scopes and re-authenticate.

```python
SCOPES = [
    "https://graph.microsoft.com/Sites.ReadWrite.All",
    "https://graph.microsoft.com/Sites.Manage.All",  # required for list creation
    "https://graph.microsoft.com/Files.ReadWrite.All",
]
```

Note: `Sites.Manage.All` is a delegated permission that requires admin consent
in the target tenant.

---

### SharePoint column internal names are set at creation and cannot contain spaces

**Problem:** When writing to a SharePoint list via Graph API, field names in
the `fields` object must match the column's *internal name*, not its display
name. If you create a column via the SharePoint UI with a display name like
"Project Number", SharePoint generates an internal name of
`Project_x0020_Number` — which will not match `"ProjectNumber"` in your code.

**Solution:** When creating columns programmatically via Graph API, set the
`name` property (internal name) explicitly without spaces, and use
`displayName` separately for the human-readable label:

```python
{
    "name": "ProjectNumber",        # internal name — no spaces, used in API calls
    "displayName": "Project Number", # display name — shown in SharePoint UI
    "text": {}
}
```

If creating columns manually via the SharePoint UI, type the column name
without spaces initially. You can rename the display name afterwards without
affecting the internal name.

---

### Date fields require full ISO 8601 format

**Problem:** Writing a date value like `"2026-04-16"` to a SharePoint date
column via Graph API returns `400 invalidRequest`.

**Solution:** Use full ISO 8601 datetime format with a time component:

```python
start_date = "2026-04-16T00:00:00Z"
```

Or if you have a plain date string:
```python
if "T" not in date_str:
    date_str = date_str + "T00:00:00Z"
```

---

### List names in Graph API URLs are case and space sensitive

**Problem:** `GET /sites/{id}/lists/ProjectRegister` returns 404 if the list
was created with the display name "Project Register" (with a space).

**Solution:** Use the exact display name including spaces. The Graph API
accepts list names with spaces directly in the URL — you do not need to
URL-encode them for the path, though `%20` also works:

```
/sites/{id}/lists/Project Register/items
/sites/{id}/lists/Project%20Register/items   # also valid
```

---

### MSAL silent token acquisition must be filtered by tenant (realm)

**Problem:** When managing multiple Microsoft 365 tenants with a shared token
cache, `app.acquire_token_silent()` may return a token for the wrong tenant if
the cache contains accounts from multiple tenants.

**Solution:** Filter cached accounts by `realm` before calling
`acquire_token_silent`:

```python
all_accounts = app.get_accounts()
accounts = [
    a for a in all_accounts
    if a.get("realm", "").lower() == tenant_id.lower()
]
if accounts:
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
```

---

### New Microsoft 365 Group sites: write permissions lag behind read permissions

**Problem:** After creating a Microsoft 365 Group via Graph API, the
SharePoint site provisions quickly and `GET /sites/{id}/lists` returns `200`
within ~30 seconds. However, attempting to create lists (`POST /sites/{id}/lists`)
continues to return `403` for several minutes afterwards.

**Root cause:** Read access to the site becomes available before the creator's
write/owner permissions fully propagate through the SharePoint permission system.

**Solution:** Implement retry logic specifically for list creation, rather than
relying on a fixed wait period:

```python
for attempt in range(12):  # retry for up to 4 minutes
    r = requests.post(url, headers=headers, json=payload)
    if r.status_code in (200, 201):
        break
    if r.status_code == 403:
        time.sleep(20)
        continue
    break  # other errors — don't retry
```

---

### `"Field 'X' is not recognized"` on list item creation

**Problem:** `POST .../lists/{name}/items` returns `400 invalidRequest` with
message `Field 'LibraryName' is not recognized` (or similar).

**Root cause:** The field name in your request payload doesn't match any
column's internal name in the list. Common causes:
- The column was never created in this list
- The column was created with a different internal name (see spaces note above)
- The column name has a typo

**Solution:** Check the actual column internal names by calling:
```
GET /sites/{id}/lists/{name}/columns
```
and comparing the `name` property of each column against your payload keys.
Remove any fields from your payload that don't exist in the list.
