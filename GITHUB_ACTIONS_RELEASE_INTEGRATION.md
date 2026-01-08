# RocketTools Release API

## Base URL
- Production: `https://rocket-tools.ru/api/releases/`
- Development: `http://localhost:9997/api/releases/`

## Endpoints

### Upload New Release

**POST** `/api/releases/upload/`

**Authentication:** Bearer token

**Headers:**
```
Authorization: Bearer YOUR_API_TOKEN
Content-Type: multipart/form-data
```

**Request fields:**
| Field | Type | Required | Description |
|-------|------|----------|-------------|
| `version` | string | Yes | Semantic version (e.g., "1.0.0") |
| `exe_file` | file | Yes | Plugin installer .exe file |
| `release_notes` | string | No | Release notes (markdown) |
| `git_tag` | string | No | Git tag name |
| `git_commit` | string | No | Git commit SHA |
| `github_release_url` | string | No | GitHub release page URL |
| `min_revit_version` | string | No | Min Revit version (default: "2021") |
| `max_revit_version` | string | No | Max Revit version (default: "2025") |

**Example:**
```bash
curl -X POST "https://rocket-tools.ru/api/releases/upload/" \
  -H "Authorization: Bearer YOUR_API_TOKEN" \
  -F "version=1.0.0" \
  -F "release_notes=Bug fixes and improvements" \
  -F "exe_file=@./RocketTools-Setup-1.0.0.exe" \
  -F "git_tag=v1.0.0" \
  -F "git_commit=abc123def456" \
  -F "github_release_url=https://github.com/user/repo/releases/tag/v1.0.0"
```

**Response (201):**
```json
{
  "message": "Release 1.0.0 uploaded successfully",
  "release": {
    "id": 1,
    "version": "1.0.0",
    "download_url": "/media/plugin_releases/2024/01/RocketTools-Setup-1.0.0.exe",
    "file_size": 12345678,
    "file_hash": "sha256..."
  }
}
```

---

### Get Latest Release

**GET** `/api/releases/latest/`

**Authentication:** None (public)

**Response (200):**
```json
{
  "id": 1,
  "version": "1.0.0",
  "download_url": "...",
  "release_notes": "...",
  "published_at": "2024-01-15T12:00:00Z"
}
```

---

### Check for Updates

**GET** `/api/releases/check-update/?current_version=0.9.0`

**Authentication:** None (public)

**Response (200):**
```json
{
  "update_available": true,
  "latest_version": "1.0.0",
  "download_url": "...",
  "release_notes": "...",
  "published_at": "2024-01-15T12:00:00Z"
}
```

---

### List All Releases

**GET** `/api/releases/`

**Authentication:** None (public)

**Query params:** `limit` (optional, default 10)

---

### Get Release by Version

**GET** `/api/releases/{version}/`

**Authentication:** None (public)

---

### Download Release (with tracking)

**GET** `/api/releases/{version}/download/`

**Authentication:** None (public)

Increments download counter and redirects to file.

---

## API Token

1. Django Admin → "API токены релизов"
2. Create token with descriptive name
3. Copy token (shown only once)
4. Use in `Authorization: Bearer <token>` header
