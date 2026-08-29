# Fix academic reads to fixed, identity-bearing requests

The workbench reads academic data through fixed, versioned HTTP contracts that carry the authenticated browser session. The browser owns authentication only; readers perform no menu discovery and no control clicking on the normal path. Each contract is a code-level object with a version, entry method and path, data method and path, required parameters, and response markers (semester selector, result table, login page).

Grade records use `GET /cjcx/queryQmcj` to verify authentication and read the semester list, then paginate with `POST /cjcx/queryQmcj`. The personal timetable uses `POST /kbcx/queryGrkb`, and the selection query uses `POST /xsxk/queryXsxkList`. Selection execution keeps the page-provided teaching-class identity, one explicit confirmation, a pre-submit re-read, and exactly one `POST /xsxk/saveXsxk` with no automatic retry.

A fixed contract change is a failure, not a trigger for automatic discovery. When a contract changes, the existing snapshot is preserved and only sanitized diagnostic evidence is emitted. Menu traversal, interface discovery, and manual navigation observation are development-only diagnostics: the backend disables their task entry unless `ACADEMIC_WORKBENCH_DEV_DIAGNOSTICS=1`, and the frontend hides the control unless the capability is reported.

All iframe-name and menu-click read paths have been removed from both the workbench runtime and the standalone `course-progress` CLI; both now collect grade records through the fixed `GET/POST /cjcx/queryQmcj` contract carrying the authenticated browser session.
