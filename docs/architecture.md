
# Architecture

The Next.js page calls the same-origin `/api/*` path. A Next.js rewrite forwards those requests to FastAPI at `BACKEND_INTERNAL_URL` (`http://backend:8003` in Docker; `http://localhost:8003` for direct local development). The backend listens on container port 8003. Compose can publish that port on a different host port without changing the internal connection. The frontend is built as a standalone production server, and the backend runs Uvicorn.

FastAPI saves uploaded workbooks and lists valid saved workbooks from `UPLOAD_DIR`. The Docker mount maps the project's `workbooks/` directory to `/excel_files`; direct local runs default to the project `workbooks/` directory. Uploaded and agent-created files therefore use the same directory. The frontend refreshes its workbook list from `/files` or the chat response. Delete calls `DELETE /files/{filename}` and removes the selected file from the UI after the API confirms success.

For chat, FastAPI selects the requested workbook and calls the LangChain agent. The agent uses OpenRouter for reasoning and discovers the pinned `excel-mcp-server==0.1.8` tools over stdio through `MultiServerMCPClient`. Tool paths must be absolute and directly under `UPLOAD_DIR`. Successful MCP writes are tracked; after execution, FastAPI reads the saved workbook to return created files and a bounded changed-range preview. That preview is rendered inline. A preview failure does not replay an Excel write.

Chart calls are routed through a hidden `_ChartData` worksheet because the MCP chart tool creates a chart on its source worksheet. For grouped totals, the backend prepares a two-column summary from the full raw table there. A post-call check inspects the saved series and aggregate against the raw source, moves the chart to the chart-only `Charts` sheet, replaces a matching old chart, and positions it clear of other charts. Existing cells on an older `Charts` sheet are preserved on a hidden `_ChartArchive` sheet during migration. The MCP package itself remains unchanged.

Conversation state is in memory inside the backend process. It is lost when the backend restarts. Saved workbook files persist in `workbooks/`; keep that directory backed up for deployments that need durable data. `.xls` files can be uploaded and previewed, while the Excel MCP editing and chart flow targets `.xlsx`.
