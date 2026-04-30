---
name: access-vba-reserved-words
description: >-
  Detect and avoid Microsoft Access / VBA reserved words in identifiers.
  Use when naming variables, fields/controls, modules, procedures, queries, or SQL columns.
  Guide learners to safe alternatives and consistent conventions.
license: CC-BY-4.0
---

# Access/VBA Reserved Words – Naming Safety Skill

## Purpose
Help developers avoid case-insensitive collisions with:
- **VBA keywords & operators** (`Dim`, `If`, `Select`, `Function`, `ByVal`, `And`, etc.).
- **Built-in VBA function names** (`InStr`, `Left`, `Right`, `Mid`, `Len`, `Date`, `Time`, `Now`, etc.).
- **Access/DAO object model names** (`Form`, `Forms`, `Report`, `Section`, `Field`, `Fields`, `Recordset`, etc.).
- **ACE/Jet SQL keywords** (`SELECT`, `FROM`, `WHERE`, `JOIN`, `GROUP BY`, `ORDER BY`, `TABLE`, `INDEX`, `UNION`, etc.).
- **Special characters/symbols** (spaces, `'`, `"`, `.`, `!`, `?`, `*`, `+`, `-`, `=`, `<`, `>`, `#`, `%`, `$`, `&`, `@`, `\`, `/`, `^`, `~`, `{}`, `[]`, `()`).
> Access and the database engine treat names **case-insensitively**; reusing these terms as identifiers often leads to compile or runtime errors.

## Procedure
1. **Scan identifiers** in the current context (variables, procedure names, control names, field names, query columns, and embedded SQL).
2. **Flag exact case-insensitive matches** to reserved words and report each occurrence with the location.
3. **Suggest safe replacements** using descriptive names:
   - `InStr` → `posInStr` or `searchPos`
   - `Date` → `SaleDate` / `CreatedDate`
   - `Name` → `PersonName` / `ItemName`
   - `Text` → `BodyText` / `MessageText`
4. **Apply conventions**:
   - Use CamelCase, no spaces or special characters.
   - Prefer Leszynski/Reddick-style type prefixes for variables (e.g., `strName`, `lngCount`, `dtmStart`) and object prefixes (`frmMain`, `qrySales`).
5. **If renaming Access objects** already in use, prefer a refactor rather than bracketing; `[Date]` can work but renaming avoids subtle bugs.
6. **(Optional) Run a repository scan** when requested to list offenders and propose bulk renames.

## Common Offenders (teach by example)
- Identifiers: `date`, `time`, `now`, `value`, `name`, `text`, `year`
- VBA functions used as names: `instr`, `left`, `right`, `mid`, `len`
- Access/DAO object names: `form`, `forms`, `report`, `section`, `field`, `fields`
- SQL words: `select`, `from`, `where`, `join`, `order`, `group`, `union`, `table`, `index`

## Naming Recommendations
- ✅ Prefer descriptive, specific names: `SaleDate`, `IsEligible`, `TotalAmount`.
- ✅ Use CamelCase; consider type/object prefixes (`str`, `lng`, `dtm`, `frm`, `qry`) for clarity.
- ❌ Avoid reserved words and special characters in any identifier.
- ❌ Avoid generic names like `Data`, `Info`, `Temp` when they collide with engine terms.

## Example Prompts (to trigger this skill)
- “Scan this module and flag any **Access/VBA reserved-word** variable names.”
- “Suggest safe replacements for fields named `Date` and `Text` across forms.”
- “Check my SQL queries for column names that collide with ACE/Jet reserved words.”

## References
- Microsoft: Avoid using reserved words and symbols in Access
- Microsoft: SQL reserved words (ACE/Jet engine)
- VBA language keywords & reserved identifiers
- Allen Browne: Problem names & reserved words in Access