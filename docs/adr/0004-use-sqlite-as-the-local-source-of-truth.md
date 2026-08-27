# Use SQLite as the local source of truth

The local application will use SQLite as the sole writable source of truth for student profile, confirmed notices, academic snapshots, plans, and task history. Existing JSON files are imported once and retained as read-only migration evidence; avoiding ongoing dual writes gives snapshot replacement and failure recovery one transactional boundary while keeping all academic data local.
