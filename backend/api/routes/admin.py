import os
import secrets
import datetime
from fastapi import APIRouter, Depends, Form, Request, HTTPException, status
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from sqlalchemy.orm import Session

from models.license import LicenseKey, get_db

router = APIRouter()
security = HTTPBasic()

# Templates setup
_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
templates = Jinja2Templates(directory=os.path.join(_ROOT, "templates"))

# Simple Basic Auth verification
ADMIN_USERNAME = os.getenv("ADMIN_USERNAME", "admin")
ADMIN_PASSWORD = os.getenv("ADMIN_PASSWORD", "officeai2026")

def verify_admin(credentials: HTTPBasicCredentials = Depends(security)):
    correct_username = secrets.compare_digest(credentials.username, ADMIN_USERNAME)
    correct_password = secrets.compare_digest(credentials.password, ADMIN_PASSWORD)
    if not (correct_username and correct_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username

@router.get("/", response_class=HTMLResponse)
async def admin_dashboard(request: Request, db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    licenses = db.query(LicenseKey).order_by(LicenseKey.created_at.desc()).all()
    return templates.TemplateResponse("admin.html", {"request": request, "licenses": licenses})

@router.post("/generate")
async def generate_license(client_name: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    import uuid
    # Generate a unique key
    key_str = f"RV-{uuid.uuid4().hex[:12].upper()}"
    expires = datetime.datetime.utcnow() + datetime.timedelta(days=30)
    
    new_license = LicenseKey(
        key=key_str,
        client_name=client_name,
        expires_at=expires
    )
    db.add(new_license)
    db.commit()
    
    # Redirect back to dashboard
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@router.post("/revoke")
async def revoke_license(key: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    license_obj = db.query(LicenseKey).filter(LicenseKey.key == key).first()
    if license_obj:
        license_obj.is_active = False
        db.commit()
    
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)
