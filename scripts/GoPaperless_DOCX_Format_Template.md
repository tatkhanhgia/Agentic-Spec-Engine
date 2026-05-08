# GoPaperless Workflow RESTful API Specification -- Format Template & Rules

> File: `GoPaperless Workflow RESTful API Specification  V2- 20250905 - Copy.docx`
> Generated: 2026-05-04

---

## 1. Issues Found

| # | Issue | Severity | Details |
|---|-------|----------|---------|
| 1 | **Inconsistent `Normal` style formatting** | High | Mostly uses `Verdana 11 pt`, but some characters have undefined font/size. Bold and non-bold formatting are mixed inside the same style. |
| 2 | **Inconsistent `List Paragraph` style formatting** | High | Uses `Verdana`, but size varies between default and `12 pt`. Bold and non-bold formatting are mixed. |
| 3 | **Heading 1-4 lack explicit font/size/bold definitions** | High | Heading styles previously relied on Word defaults (`null` font/size/bold). This can render differently on other machines. |
| 4 | **Inconsistent table header color** | Medium | Some table headers used `0066CC`, while others used `0070C0`. All table headers must use one standard blue. |
| 5 | **Inconsistent `Normal` paragraph alignment** | Low | Some `Normal` paragraphs were centered while the rest were left-aligned. |

---

## 2. Standard Formatting Template

### 2.1. Page Setup
- **Paper size**: A4 (595.4 pt x 842.0 pt)
- **Margins**:
  - Left (L): 35.3 pt (~1.25 cm)
  - Right (R): 35.3 pt (~1.25 cm)
  - Top (T): 92.2 pt (~3.25 cm)
  - Bottom (B): 85.0 pt (~3.0 cm)

### 2.2. Default Font
- **Font**: Verdana
- **Size**: 11 pt
- **Color**: Black (Auto / RGB 000000)

### 2.3. Standard Style List

| Style | Font | Size | Bold | Color | Alignment | Space Before | Space After | Left Indent | Hanging Indent | Usage |
|-------|------|------|------|-------|-----------|--------------|-------------|-------------|----------------|-------|
| **Normal** | Verdana | 11 pt | No | Auto | Left (default) | 0 pt | 3 pt | 18 pt | (none) | Body text under headings |
| **List Paragraph** | Verdana | 11 pt | No | Auto | Left | 0 pt | 3 pt | 18 pt | (none) | Bullet/numbered lists under headings |
| **Heading 1** | Verdana | 16 pt | Yes | 0050A8 | Left | 18 pt | 8 pt | 0 pt | 18 pt | Major chapters such as INTRODUCTION and API Specification |
| **Heading 2** | Verdana | 14 pt | Yes | 0050A8 | Left | 14 pt | 6 pt | 0 pt | 18 pt | Major sections such as Target and Abbreviation |
| **Heading 3** | Verdana | 12 pt | Yes | 0050A8 | Left | 10 pt | 4 pt | 0 pt | 18 pt | API endpoints such as authentication methods and module APIs |
| **Heading 4** | Verdana | 11 pt | Yes | 0050A8 | Left | 8 pt | 3 pt | 0 pt | 18 pt | Sub-features or sub-steps under an endpoint/module |
| **Authentication endpoint hierarchy** | `Client Credentials Authentication`, `OTP Login Authentication`, and `Single Sign On Authentication` under `Authenticate - Log in` must be Heading 3, not Heading 4. |
| **Title** | Verdana | 18 pt | Yes | Auto | Left | 0 pt | 6 pt | 0 pt | (none) | Main document title such as History |
| **MID-PageOverview-TitleBig** | Verdana | 28 pt | Yes | Auto | Left | 0 pt | 0 pt | 0 pt | (none) | Cover page service name |
| **MID-PageOverView-TitleSmall** | Verdana | 20 pt | No | Auto | Left | 0 pt | 0 pt | 0 pt | (none) | Cover page document type |
| **MID-PageOverview-Small** | Verdana | 11 pt | Yes | Auto | Right | 0 pt | 0 pt | 0 pt | (none) | Cover page version and date |

**Golden rule:** Every style must explicitly define font, size, and bold in the Word style definition. Do not rely on `null`, inherited, or default Word formatting.

### 2.4. Tables

| Property | Standard Value |
|----------|----------------|
| **Header background color** | `#0070C0` for every table |
| **Header font** | Verdana 12 pt, bold, white (`FFFFFF`) |
| **Body font** | Verdana 11 pt, non-bold, black |
| **Borders** | Solid, 0.5 pt, light gray (`D9D9D9`) or black |
| **Header alignment** | Center + middle |
| **Attribute table start header alignment** | Left + middle for merged start header rows such as `Header Attributes`, `Path Attributes`, `Request Attributes`, and `Response Attributes` |
| **Attribute table start header font color** | White (`FFFFFF`) |
| **Attribute table start header width** | `5000` |
| **Attribute table start header width type** | `pct` |
| **Attribute table start header grid span** | `5` |
| **Attribute table start header row height** | `420` dxa |
| **Attribute table start header row height rule** | `atLeast` |
| **Table alignment** | `center` |
| **Table indent** | `0 dxa` |
| **Body alignment** | Left for text, center for numbers/status codes |
| **Row height** | Minimum 20 pt, automatic or fixed |

### 2.5. Table Border Rule Lookup

| Section Heading | Table Borders | Cell Borders | Border Size | Border Color | Apply Cell Borders | Usage |
|-----------------|---------------|--------------|-------------|--------------|--------------------|-------|
| **Abbreviation** | top, left, bottom, right, insideH, insideV | top, left, bottom, right | 2 | 000000 | Yes | Ensure all rows and columns have visible separators |

### 2.6. Spacing Between Content Sections

| Case | Standard Value |
|------|----------------|
| **Between content under one Heading 2 and the next Heading 2** | Do not insert a manual blank line. The next Heading 2 must create spacing using `Space Before = 14 pt`. |
| **Between API endpoints (Heading 3 to next Heading 3)** | Use exactly one blank `Normal` paragraph between the last content of the previous API and the next Heading 3, with Verdana 11 pt, non-bold, `Space Before = 0 pt`, `Space After = 0 pt`, `Left Indent = 0.25"`. |
| **Between content under one Heading 4 and the next Heading 4** | Do not insert a manual blank line. The next Heading 4 must create spacing using `Space Before = 8 pt`. |
| **First body paragraph/list after a heading** | Must align with the heading text (not the number) at `Left Indent = 0.25"` and use `Normal` or `List Paragraph`, Verdana 11 pt. |
| **Body content under headings** | `Normal` or `List Paragraph`, Verdana 11 pt, `Space Before = 0 pt`, `Space After = 3 pt`, `Left Indent = 0.25"`, line spacing Multiple 1.15. |
| **Blank paragraph if one exists between sections** | Must be `Normal`, Verdana 11 pt, non-bold, spacing before/after 0 pt. Remove it if it is only used to create visual spacing. |

### 2.7. API Body Section Rules

| Element | Text Match | Paragraph Type | Font | Size | Bold | Space Before | Space After | Left Indent | Blank Count | Order / Columns | Standard Rule |
|---------|------------|----------------|------|------|------|--------------|-------------|-------------|-------------|-----------------|---------------|
| **Before Sample Request** | between API description and `Sample Request` | Normal blank | Verdana | 11 pt | No | 0 pt | 0 pt | 0.25" | 1 | Before `Sample Request` | Use exactly one blank paragraph. |
| **Sample Request label** | `Sample Request` | Normal | Verdana | 12 pt | Yes | 0 pt | 3 pt | 0.25" | 0 | Before request frame | Standalone label paragraph. |
| **Sample Request frame** | after `Sample Request` | Keep existing frame | Keep existing | Keep existing | Keep existing | 0 pt | 3 pt | Keep existing | 0 | Before blank and `Sample Response` | Preserve existing frame/border style. |
| **Before Sample Response** | between request frame and `Sample Response` | Normal blank | Verdana | 11 pt | No | 0 pt | 0 pt | 0.25" | 1 | Before `Sample Response` | Use exactly one blank paragraph. |
| **Sample Response label** | `Sample Response` | Normal | Verdana | 12 pt | Yes | 0 pt | 3 pt | 0.25" | 0 | Before response frame(s) | Standalone label paragraph. |
| **Sample Response frame(s)** | after `Sample Response` | Keep existing frame | Keep existing | Keep existing | Keep existing | 0 pt | 3 pt | Keep existing | 0 | Success frame then error frame when both exist | Preserve existing frame/border style. |
| **Before Attributes description** | between last response frame and `Attributes description` | Normal blank | Verdana | 11 pt | No | 0 pt | 0 pt | 0.25" | 1 | Before `Attributes description` | Use exactly one blank paragraph. |
| **Attributes description label** | `Attributes description` | Normal | Verdana | 12 pt | Yes | 0 pt | 3 pt | 0.25" | 0 | Before attribute tables | Standalone label paragraph. |
| **Attributes description tables** | after `Attributes description` | Table | Template table style | 12 pt header, 11 pt body | Header bold only | Keep existing | Keep existing | Keep existing | 0 | `Header Attributes`, `Path Attributes`, `Request Attributes`, `Response Attributes`; columns: `No`, `Name`, `Type`, `Presence`, `Description` | Include only applicable tables in this order. |
| **Inline code frames** | after `Sample Request` / `Sample Response` labels | Inline shape (text box) | Keep existing | Keep existing | Keep existing | 0 pt | 3 pt | Centered | 0 | Request/Success/Error frames | Must be center-aligned and use uniform width matching the table width. |

### 2.8. Request & Response Frame Content Format

Rules for text **inside** Request, Success Response, and Error Response frames (inline code blocks).

| Property | Standard Value |
|----------|----------------|
| **Base URL pattern** | `https://{domain}/{contextPath}/XXX` |
| **Current domain** | `prd-gopaperless.mobile-id.vn` |
| **Current contextPath** | `/workflow/api` |
| **URL formatting** | Bold the URL/path on the first line of the frame. Only the HTTP method badge uses background shading + white text. |


#### 2.8.1. HTTP Method Badge

| Method | Text Highlight (Shading) | Text Color | Font | Size | Bold |
|--------|--------------------------|------------|------|------|------|
| **GET** | `047857` (green-700) | `FFFFFF` (white) | Verdana | 10 pt | Yes |
| **POST** | `C2410C` (orange-700) | `FFFFFF` (white) | Verdana | 10 pt | Yes |
| **PUT** | `1D4ED8` (blue-700) | `FFFFFF` (white) | Verdana | 10 pt | Yes |
| **DELETE** | `B91C1C` (red-700) | `FFFFFF` (white) | Verdana | 10 pt | Yes |
| **PATCH** | `6D28D9` (purple-700) | `FFFFFF` (white) | Verdana | 10 pt | Yes |

- Apply **text highlight / shading** (background color), not font color, to the method name.
- Use white text on dark solid badge backgrounds for maximum contrast and readability.
- Method badge must be the first token on the first line of the Request frame, followed by a space and the endpoint path.

#### 2.8.2. Header Key–Value Format

| Element | Format | Example |
|---------|--------|---------|
| **Header key** | Title-Case (each word capitalized), **bold**, Verdana 10 pt, black | **Content-Type** |
| **Colon separator** | `:` followed by one space, non-bold | `: ` |
| **Header value** | lowercase / exact literal, non-bold, Verdana 10 pt, black | `application/json` |

- Header key naming must follow the canonical form defined in the API specification (e.g., `Content-Type`, `Authorization`, `x-language-name`). Do not arbitrarily switch to lowercase or uppercase.
- Each header occupies its own line. Blank line after the last header before the body.

#### 2.8.3. Template Frame (Common Headers)

The following headers are considered the **base template** for nearly every Request frame:

```
GET /api/v1/example
x-language-name: en
Content-Type: application/json

```

| Header | Required | Notes |
|--------|----------|-------|
| `x-language-name` | Always | Language code (`en`, `vi`, etc.) |
| `Content-Type` | Always | Usually `application/json` |
| `Authorization` | **Conditional** | `Bearer {accessToken}` — add only when the API requires authentication. Login/Register and other public endpoints omit this header. |

> **Design rationale:** `Authorization` is intentionally **excluded** from the fixed Template Frame because not all APIs require it (e.g., authentication endpoints themselves, public health-check APIs). Adding it unconditionally leads to incorrect samples. Document it as a conditional header and include it only when the endpoint spec mandates authentication.

#### 2.8.4. JSON Body Format

| Property | Standard Value |
|----------|----------------|
| **Font** | Verdana |
| **Size** | 10 pt |
| **Color** | Black (`000000`) |
| **Indentation** | 2 spaces per nesting level (do not use Tab character) |
| **Key names** | Double-quoted, non-bold, exact case as API spec |
| **String values** | Double-quoted |
| **Number / boolean / null** | Unquoted, lowercase (`true`, `false`, `null`) |
| **Line breaks** | One line per key at the same level; nested objects start on a new indented line |
| **Trailing commas** | Forbidden (valid JSON) |

**Example:**

```json
{
  "username": "john.doe",
  "password": "securePass123",
  "rememberMe": true
}
```

#### 2.8.5. Frame Text Global Rules

| Property | Value |
|----------|-------|
| **Font** | Verdana |
| **Base size** | 10 pt (1 pt smaller than body text to distinguish code frame from prose) |
| **Line spacing** | Single |
| **Background** | Keep existing frame background (usually light gray or white) |
| **Alignment** | Left |

#### 2.8.6. Variable Placeholder Color

Variable placeholders in URL paths, headers, and JSON bodies must be colored red for visibility.

| Element | Format | Example |
|---------|--------|---------|
| **Variable placeholder** | Curly braces `{}`, Verdana 10 pt, non-bold, **red** (`FF0000`) | `{domain}`, `{contextPath}`, `{accessToken}` |

### 2.9. Mandatory Checklist

When creating or editing the DOCX API specification, enforce the following:

- [ ] **Consistent font**: All text uses Verdana. Do not use Arial, Calibri, Times New Roman, or other fonts.
- [ ] **Consistent size**: Body text uses 11 pt. Header/table-name labels and table headers use 12 pt.
- [ ] **Intentional bold usage**: Only headings, header/table-name labels, and table headers are bold unless there is a specific document reason.
- [ ] **Explicit heading styles**: Heading styles must define font, size, bold, and color directly in the Word style definition.
- [ ] **Consistent table color**: Every table header uses `#0070C0`. Do not use `#0066CC` or other blue variants.
- [ ] **Correct page margins**: L=R=35.3 pt, T=92.2 pt, B=85.0 pt.
- [ ] **No random paragraph alignment**: `Normal` must be left-aligned. Only cover-page styles should be centered.
- [ ] **Paragraph spacing**: Do not use repeated Enter/manual blank lines to create spacing. Use Space Before/After in paragraph styles.

---

## 3. How to Apply in Microsoft Word

### Step 1: Modify default styles
1. Open **Home** > **Styles** > expanded styles panel.
2. Find the target style (Normal, Heading 1, Heading 2, etc.).
3. Right-click the style > **Modify...**
4. In the Modify Style dialog:
   - **Font**: `Verdana`
   - **Size**: Match the table above
   - **Bold/Italic**: Match the table above
   - **Font color**: Match the table above
   - **Format** > **Paragraph...**: Set Space Before/After, alignment, and indent
5. If available, select **Add to template** or **New documents based on this template**.
6. Click **OK**.

### Step 2: Standardize table headers
1. Select the first row of each table.
2. Open **Table Design** > **Shading** > **More Colors** and set `0070C0`.
3. Set header text to Verdana 12 pt, bold, white (`FFFFFF`).
4. Use Format Painter or table selection to apply consistently.

### Step 3: QA before saving
- [ ] Press `Ctrl+A` and check that the ribbon shows Verdana for the selected text where applicable.
- [ ] Use **Review** > **Inspect Document** or **File** > **Info** > **Check for Issues** > **Inspect Document** to remove stray formatting when needed.

---

## 4. Suggested Improvements for the Current File

1. **Normalize `Normal`**: Apply Verdana 11 pt to the whole `Normal` style and remove accidental bold body text.
2. **Normalize `List Paragraph`**: Apply Verdana 11 pt consistently. If bold list items are intentional, use a separate style rather than direct formatting.
3. **Define Heading 1-4 explicitly**: Apply Verdana, size, bold, and blue color (`0050A8`) directly to the styles.
4. **Standardize table headers**: Convert all table header backgrounds to `0070C0` and make header text Verdana 11 pt, bold, white.
5. **Remove redundant or duplicate styles**: Review unused or overlapping styles and remove them when safe.
