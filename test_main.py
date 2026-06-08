from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def root():
    return {"message": "Root is working"}

@app.post("/test-upload")
def test_upload():
    return {"message": "POST endpoint is alive!"}
