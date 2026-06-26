from fastapi import FastAPI

app = FastAPI(title="Bitrix24 Business Process Document Generator")


@app.get("/health")
async def health():
    return {"status": "ok"}
