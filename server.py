from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from app.config import ASSETS_DIR, ensure_directories
from app.auth import seed_preset_users
from app.cleanup import run_age_based_cleanup
from app.routes import auth, pages, settings, api, batch, downloads
from app.routes import customers

app = FastAPI(
    title="PhillipCapital Risk Management Credit Worksheet Processor",
    description="FastAPI PDF-to-Excel processor for Risk Management credit worksheet automation.",
    version="7.0.0",
)

app.mount("/assets", StaticFiles(directory=str(ASSETS_DIR)), name="assets")


@app.on_event("startup")
def on_startup():
    ensure_directories()
    seed_preset_users()
    try:
        run_age_based_cleanup()
    except Exception:
        pass


app.include_router(auth.router)
app.include_router(pages.router)
app.include_router(settings.router)
app.include_router(api.router)
app.include_router(batch.router)
app.include_router(downloads.router)
app.include_router(customers.router)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("server:app", host="0.0.0.0", port=8000, reload=True)
