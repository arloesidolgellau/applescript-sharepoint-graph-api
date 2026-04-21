# Microsoft Graph Beta Pages API — Lessons Learned

Discovered while building automated SharePoint home page provisioning
for group-connected team sites via Python + MSAL + Microsoft Graph.

---

## 1. Valid `horizontalSectionLayoutType` values

The Graph beta `/sites/{id}/pages` API uses a strict enum for section layouts.
The documentation does not list these clearly. Valid values:

| Layout type | Columns | Required column widths |
|---|---|---|
| `fullWidth` | 1 | `[0]` |
| `oneColumn` | 1 | `[12]` |
| `oneThirdLeftColumn` | 2 | `[4, 8]` |
| `oneThirdRightColumn` | 2 | `[8, 4]` |

**Values that do NOT exist and return 400:**
- `twoColumns` — not in the enum
- `multiColumn` — not in the enum

The column `width` values are strictly validated against the layout type.
Sending the wrong width (e.g. `12` for `fullWidth`, or `6` for a two-column layout)
returns:

```
400: mismatch of column width, should be 0 but found 12
```

---

## 2. PATCH vs POST on `/sites/{id}/pages`

When **creating** a new page (`POST`), you can include:
- `name`
- `pageLayout`
- `promotionKind`
- `title`
- `canvasLayout`

When **updating** an existing page (`PATCH` to `/pages/{id}/microsoft.graph.sitePage`),
only these properties are accepted:
- `@odata.type`
- `title`
- `canvasLayout`

Sending `name`, `pageLayout`, or `promotionKind` in a PATCH returns:

```
400: Property 'name' cannot be used in this request
400: Property 'pageLayout' cannot be used in this request
```

**Pattern that works:**

```python
if existing_page_id:
    patch_body = {
        "@odata.type":  "#microsoft.graph.sitePage",
        "title":        project_name,
        "canvasLayout": canvas_layout,
    }
    requests.patch(
        f"https://graph.microsoft.com/beta/sites/{site_id}"
        f"/pages/{existing_page_id}/microsoft.graph.sitePage",
        headers=headers, json=patch_body
    )
else:
    requests.post(
        f"https://graph.microsoft.com/beta/sites/{site_id}/pages",
        headers=headers, json=full_page_body
    )
```

After saving, publish with:
```
POST /beta/sites/{id}/pages/{page-id}/microsoft.graph.sitePage/publish
```

---

## 3. SharePoint SiteScriptUtility REST API

The endpoint:
```
POST /_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility/CreateSiteScript
```

Returns `404 ResourceNotFoundException` on some Microsoft 365 plans
(observed on what appears to be Business Basic).

This is **not** a permissions issue — adding `AllSites.FullControl` delegated
permission to the Azure AD app and granting admin consent does not resolve it.
All three URL variants fail identically:
- `https://{tenant}-admin.sharepoint.com/_api/...`
- `https://{tenant}.sharepoint.com/_api/...`
- `https://{tenant}.sharepoint.com/sites/{any-site}/_api/...`

**Workaround:** Site Designs are not available on this plan tier.
Use the Graph beta pages API directly to provision home page layout instead.

---

## 4. Setting a site logo

The `siteiconmanager` SharePoint REST endpoint:
```
POST {site-url}/_api/siteiconmanager/setsitelogo
```
Requires a **SharePoint-scoped OAuth token** (`https://{tenant}.sharepoint.com/.default`),
not a Microsoft Graph token. Sending a Graph token returns `401`.

**Simpler alternative that works with a Graph token:**
```python
# Look up the M365 group from the site URL slug
slug = site_url.rstrip("/").split("/")[-1]
r = requests.get(
    f"https://graph.microsoft.com/v1.0/groups"
    f"?$filter=mailNickname eq '{slug}'&$select=id",
    headers={"Authorization": f"Bearer {graph_token}"},
)
group_id = r.json()["value"][0]["id"]

# Set group photo — this becomes the site icon
with open(logo_path, "rb") as f:
    requests.put(
        f"https://graph.microsoft.com/v1.0/groups/{group_id}/photo/$value",
        headers={
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "image/png",
        },
        data=f.read(),
    )
```

Note: the group photo endpoint can return `404` briefly after group creation.
Retry with a short delay if needed.

---

## 5. SharePoint token vs Graph token

These are different OAuth resources and require separate MSAL token requests:

```python
# Microsoft Graph
graph_scopes = ["https://graph.microsoft.com/Sites.ReadWrite.All", ...]

# SharePoint REST API
sp_scopes = ["https://{tenant}.sharepoint.com/.default"]

# SharePoint Admin site (different audience from root site)
sp_admin_scopes = ["https://{tenant}-admin.sharepoint.com/.default"]
```

A token acquired for one resource will be rejected (401 or 404) by another.
MSAL caches tokens per resource — request and cache them separately.

---

*Discovered April 2026 — Python 3.9, MSAL 1.x, Microsoft Graph beta*
