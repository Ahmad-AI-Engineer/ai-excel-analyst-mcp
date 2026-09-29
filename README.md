
# AI Excel Analyst

A Next.js chat interface backed by FastAPI, OpenRouter, and the pinned Excel MCP server. Users can upload, preview, download, and delete workbooks, then ask the agent to inspect or edit them. Saved write ranges appear as short table previews in chat. The visible `Charts` sheet contains charts only; chart source summaries are kept on a hidden `_ChartData` sheet.

## Run with Docker Compose

1. Copy `.env.example` to `.env` and set `OPENROUTER_API_KEY` to a valid key. Keep `.env` private.
2. From this directory run `docker compose up --build -d`.
3. Open `http://localhost:3000`. The FastAPI health endpoint and API docs are at `http://localhost:8003/` and `http://localhost:8003/docs`.

`FRONTEND_PORT` and `BACKEND_PORT` in `.env` optionally change the host ports. The backend listens on port 8003 inside Docker, and Next.js routes `/api/*` to `http://backend:8003/*` on the Compose network. When using nondefault host ports, open the corresponding frontend port. The browser does not need the backend's internal Docker address. For a custom deployment topology, set `BACKEND_INTERNAL_URL` before building the frontend image.

Uploads and agent-created workbooks persist in `workbooks/` beside this README through the Compose bind mount. Existing files in that directory are shown on startup. The backend uses `UPLOAD_DIR=/excel_files` inside Docker; when run directly, it defaults to this project's `workbooks/` directory. You can override `UPLOAD_DIR` with an absolute directory for a direct backend run. Keep the directory writable and backed up if its workbooks matter.

## Run without Docker

Install the packages in `backend/requirements.txt`, install `uv` so `uvx` is on `PATH`, and set `OPENROUTER_API_KEY`. In `backend/`, run `uvicorn main:app --host 0.0.0.0 --port 8003`. In `frontend/`, run `npm install --legacy-peer-deps` and `npm run dev`. Next.js then routes `/api/*` to `http://localhost:8003/*`. Set `BACKEND_INTERNAL_URL` before starting or building Next.js if the backend has another reachable address.

The upload and preview API accepts `.xlsx` and `.xls`. The agent's Excel MCP editing and chart workflow is intended for `.xlsx` workbooks. The sample workbooks in `samples/` are examples; the application reads active files from `workbooks/`.

## Checks before deployment

- Confirm `.env` contains a working `OPENROUTER_API_KEY` and is not committed.
- Run `docker compose config --quiet` and `docker compose up --build -d` from this directory, then check `docker compose ps` for a healthy backend.
- Open the frontend, upload a test `.xlsx`, preview it, ask a simple read question, perform a small edit, and download the result. Verify the workbook appears in `workbooks/` and survives `docker compose down` followed by `docker compose up -d`.
- Test the delete button with a disposable workbook; deletion removes the saved file from `workbooks/`.
- Run `python -m unittest discover -s tests -v` from `backend/` in an environment with its Python dependencies. MCP integration tests also require the pinned `uvx` server.

See [docs/architecture.md](docs/architecture.md) for the request and workbook flows.
