---
description: "Universal AL coding standards applied across all BC AL projects. Project-specific overrides are in al-coding-standards.instructions.md in this same folder."
applyTo: "**/*.al"
---

# AL Coding Standards - Universal

These standards apply to all Business Central AL extension projects. Project-specific settings (ID range, prefix, BC version) are defined in `al-coding-standards.instructions.md` in this same folder.

---

## File Naming

Follow the `<ObjectName>.<ObjectType>.al` pattern. The file name must match the AL object name inside it.

| Object Type | Suffix |
|---|---|
| Table | `.Table.al` |
| Table Extension | `.TableExt.al` |
| Page | `.Page.al` |
| Page Extension | `.PageExt.al` |
| Codeunit | `.Codeunit.al` |
| Report | `.Report.al` |
| Report Extension | `.ReportExt.al` |
| Enum | `.Enum.al` |
| Enum Extension | `.EnumExt.al` |
| Query | `.Query.al` |
| XMLport | `.XMLport.al` |
| Interface | `.Interface.al` |
| Permission Set | `.PermissionSet.al` |
| Profile | `.Profile.al` |

**Correct:**
```
CustomerExt.TableExt.al
SalesOrderManagement.Codeunit.al
CustomerCardExt.PageExt.al
```
**Wrong:** `Tab-Ext50100.al`, `Cod50100.al`, `CustomerExt-50100.al`

---

## Object Naming

- Use **PascalCase** for all object names
- Object names must be descriptive and reflect business purpose
- Extensions: include the base object name - e.g., `Customer Card Ext`, `Sales Header Ext`
- Codeunits for management logic: use `Mgt` or `Management` suffix - e.g., `Customer Comments Mgt`
- Event subscriber codeunits: use `EventSubscriber` or `EventHandler` suffix
- All object names and field names must be wrapped in double quotes when referenced in code

---

## Variable Naming

- **Local variables**: camelCase starting lowercase (`salesHeader`, `totalAmount`)
- **Global variables**: PascalCase starting uppercase (`CustomerBuffer`, `TotalSalesAmount`)
- **Temporary record variables**: prefix with `Temp` (`TempCustomer`, `TempItemLedgerEntry`)
- **Boolean variables**: use `Is`, `Has`, or `Can` prefix (`IsPosted`, `HasPermission`, `CanPost`)
- **Record variables**: use the table name without a suffix (`Customer`, not `CustomerRecord`)
- **Codeunit variables**: use `Mgt` or `Management` suffix (`CustomerMgt`, `SalesMgt`)
- **No Hungarian notation**: do not use type prefixes like `str`, `int`, `bol`
- **No generic names**: avoid `Rec`, `TempRec`, `Buffer` without context

Use standard BC abbreviations where appropriate: `No.`, `Qty`, `Amt`, `LCY`, `FCY`.

---

## NoImplicitWith

Always enable `NoImplicitWith` in `app.json` features. All record field references must be explicitly qualified - never rely on implicit `with` scoping.

**Wrong:**
```al
Customer.Find();
Name := 'Test';
```
**Correct:**
```al
Customer.Find();
Customer.Name := 'Test';
```

Use `Rec.` prefix for the implicit current record in table and page triggers.

---

## Object Structure

Follow this ordering inside AL objects:

1. `Properties`
2. `fields` / `keys` (tables)
3. `layout` / `actions` / `views` (pages)
4. `dataset` / `requestpage` / `rendering` (reports)
5. `trigger` blocks (OnInsert, OnModify, OnDelete, OnRename for tables; OnInit, OnOpenPage, OnAfterGetRecord for pages)
6. `procedure` blocks - public procedures before local/internal
7. `var` sections immediately before the code block that uses them

---

## Error Handling

**Prefer `ErrorInfo` over plain `Error()` for all user-facing validation errors (BC23+).**

Use `ErrorInfo` when:
- You know how to fix the error (add a Fix-it action)
- The user needs to navigate somewhere to resolve it (add a Show-it action)
- The error is a field validation (set `RecordId` to highlight the field)

```al
// Correct - actionable error with context
procedure ValidateQuantity(Quantity: Decimal; MaxQuantity: Decimal)
var
    QuantityError: ErrorInfo;
begin
    if Quantity > MaxQuantity then begin
        QuantityError.Title := 'Quantity exceeds maximum';
        QuantityError.Message :=
            StrSubstNo('The quantity %1 exceeds the allowed maximum of %2.', Quantity, MaxQuantity);
        QuantityError.DetailedMessage('The maximum is set in the setup page.');
        QuantityError.AddAction(
            StrSubstNo('Set value to %1', MaxQuantity),
            Codeunit::"Quantity Fix Handler",
            'SetToMaximum'
        );
        Error(QuantityError);
    end;
end;

// Wrong - no context, no guidance
procedure BadValidation(Quantity: Decimal)
begin
    if Quantity > 100 then
        Error('Quantity is invalid');
end;
```

**Fix-it actions:** Label pattern: "Set value to [value]" or "Set [field] to [value]". Only add if the fix is a single, permission-safe action.

**Show-it actions:** Label pattern: "Show [PageName]". Use `AddNavigationAction()` and set `PageNo`.

**Always set `RecordId`** on validation errors so BC highlights the affected field.

---

## Code Style

- **Indentation**: 4 spaces (no tabs)
- **Keywords**: lowercase (`begin`, `end`, `if`, `then`, `else`, `procedure`, `var`, `trigger`)
- **Object and type names**: PascalCase, quoted with double quotes when referencing
- **`begin`/`end`**: `begin` on the same line as `if`/`for`/`while`/`trigger`/`procedure`; `end` on its own line
- **Line length**: aim for under 120 characters
- Use `StrSubstNo()` for parameterised messages - never concatenate strings in error/message calls
- Prefer `IsNullGuid()` over comparing to `'{00000000-...}'`

---

## AL Patterns - Always / Never

### Always
- Use `FindSet()` + `repeat...until` for multi-record loops; use `FindFirst()` for single lookups
- Call `Modify(true)` / `Insert(true)` to trigger field validation unless intentionally bypassing
- Use `Get()` and check the return value before accessing record fields
- Use `LockTable()` before updating records in concurrent scenarios
- Add `TableRelation` and `ValidateTableRelation` on foreign key fields
- Use events (publishers/subscribers) to extend logic without modifying base objects
- Keep procedures short and single-purpose

### Never
- Never use `COMMIT` inside a codeunit unless it is a top-level entry point (post routine, job queue)
- Never use `CurrReport.QUIT` in reports - use `CurrReport.Skip()` instead
- Never suppress errors silently unless inside a `try` function
- Never hardcode company names, user IDs, or environment-specific values
- Never leave empty `trigger` blocks - remove them if unused
- Never use object IDs as identifiers in code - always use quoted names
- Never use `with` statements

---

## AppSource / Per-Tenant Quality

- Prefix all custom object names, field names, and enum values with the project prefix (see workspace instructions) to avoid collisions
- Do not modify base application objects directly - use extensions, events, and page customizations
- Avoid using `internal` access modifier on procedures that need to be extensible by other extensions
- Use `ObsoleteState` and `ObsoleteReason` rather than deleting fields or objects
- Test with `CLEAN` profile to catch missing permissions
