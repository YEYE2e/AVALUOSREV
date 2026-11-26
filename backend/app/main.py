from fastapi import FastAPI
from .pilot import router as pilot_router
from .web_review import router as review_router

app = FastAPI(title="Revisión de avalúos", version="0.1.0")
app.include_router(pilot_router)
app.include_router(review_router)

@app.get("/api/health")
def health():
    return {"status": "ok"}
