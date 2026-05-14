# Referential Integrity Rules

Google Sheets is the database for this system, and Google Apps Script enforces application-level foreign-key rules before child records are written.

## Parent-child relationships

```text
Projects(Project ID)
  -> Activities(Project ID, Activity ID)
      -> Costs(Project ID, Activity ID, Cost ID)
          -> DailyCosts(Project ID, Activity ID, Cost ID, Date)
```

## Enforced write rules

1. Activities must reference an existing project before they can be created or updated.
2. Costs must reference an existing project and activity before they can be created or updated.
3. Daily costs must reference an existing project, activity, and cost before they can be created, updated, or deleted.
4. Generated daily-cost rows created from a cost mutation are validated with the same parent checks before they are written.

## Delete behavior

Project and activity deletes use cascade-style cleanup so related child rows do not remain orphaned in the spreadsheet database.
