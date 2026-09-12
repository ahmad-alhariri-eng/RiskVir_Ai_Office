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
    return templates.TemplateResponse(
        request=request,
        name="admin.html",
        context={"licenses": licenses}
    )

@router.post("/generate")
async def generate_license(
    client_name: str = Form(...),
    duration_type: str = Form("30"),
    custom_date: str = Form(None),
    db: Session = Depends(get_db), 
    admin: str = Depends(verify_admin)
):
    import uuid
    import datetime
    
    key_str = f"RV-{uuid.uuid4().hex[:12].upper()}"
    now = datetime.datetime.utcnow()
    
    if duration_type == "30":
        expires = now + datetime.timedelta(days=30)
    elif duration_type == "90":
        expires = now + datetime.timedelta(days=90)
    elif duration_type == "365":
        expires = now + datetime.timedelta(days=365)
    elif duration_type == "custom" and custom_date:
        try:
            expires = datetime.datetime.strptime(custom_date, "%Y-%m-%d")
        except ValueError:
            expires = now + datetime.timedelta(days=30)
    else:
        expires = now + datetime.timedelta(days=30)
        
    new_license = LicenseKey(
        key=key_str,
        client_name=client_name,
        expires_at=expires
    )
    db.add(new_license)
    db.commit()
    
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@router.post("/revoke")
async def revoke_license(key: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    license_obj = db.query(LicenseKey).filter(LicenseKey.key == key).first()
    if license_obj:
        license_obj.is_active = False
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@router.post("/activate")
async def activate_license(key: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    license_obj = db.query(LicenseKey).filter(LicenseKey.key == key).first()
    if license_obj:
        license_obj.is_active = True
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@router.post("/update_duration")
async def update_duration(key: str = Form(...), new_date: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    import datetime
    license_obj = db.query(LicenseKey).filter(LicenseKey.key == key).first()
    if license_obj:
        try:
            license_obj.expires_at = datetime.datetime.strptime(new_date, "%Y-%m-%d")
            db.commit()
        except ValueError:
            pass
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)

@router.post("/delete")
async def delete_license(key: str = Form(...), db: Session = Depends(get_db), admin: str = Depends(verify_admin)):
    license_obj = db.query(LicenseKey).filter(LicenseKey.key == key).first()
    if license_obj:
        db.delete(license_obj)
        db.commit()
    return RedirectResponse(url="/admin", status_code=status.HTTP_303_SEE_OTHER)
