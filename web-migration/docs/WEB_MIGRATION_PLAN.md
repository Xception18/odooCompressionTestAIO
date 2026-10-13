# Web Migration Plan: Desktop → Web Application

## Executive Summary

Migration strategy untuk mengubah aplikasi desktop (`customtkinter`) menjadi web application modern dengan arsitektur client-server. Target: akses multi-user, deployment cloud-ready, maintain semua fitur existing.

**Current State:** Desktop app (Windows-only, single-user, local PostgreSQL)  
**Target State:** Web app (cross-platform, multi-user, cloud-deployable)  
**Estimated Effort:** 20-27 hours (zero-cost stack) | 120-160 hours (production stack)  
**Risk Level:** Low (zero-cost) | Medium (production stack)

> **💡 Quick Start:** Jump to [Phase 6: Zero-Cost Deployment](#phase-6-zero-cost-deployment-30-minutes) for $0/month setup in 30 minutes.

### Deployment Options

| Option | Cost/Month | Setup Time | Best For |
|--------|-----------|------------|----------|
| **Zero-Cost Stack** (Recommended) | $0 | 30-45 min | Personal, demo, <100 users/month |
| **Budget Cloud** | $5-20 | 2-3 hours | Small business, <500 users/month |
| **Production Cloud** | $66-157 | 3-4 hours | Enterprise, high-traffic, SLA required |

**This plan covers all three options.** See [Phase 6: Zero-Cost Deployment](#phase-6-zero-cost-deployment-bonus) for $0/month setup.

---

## Phase 1: Architecture Planning (Week 1, 16-24 hours)

### 1.1 Technology Stack Selection

**Backend Options:**
1. **FastAPI** (Recommended)
   - Async support (critical untuk web automation)
   - Built-in OpenAPI/Swagger docs
   - Type hints native (existing codebase sudah partial typing)
   - WebSocket support untuk real-time logs
   - **Effort:** Low (Python familiarity)

2. **Django + Celery** (Alternative)
   - Full-featured (auth, admin panel, ORM)
   - Better untuk complex business logic
   - Celery untuk background tasks
   - **Effort:** Medium (learning curve steeper)

**Frontend Options:**
1. **React + Vite** (Recommended)
   - Component reuse (5 tabs → 5 route components)
   - Rich ecosystem (AG-Grid untuk tabel, React Query untuk state)
   - **Effort:** Medium (JavaScript/TypeScript learning)

2. **Vue.js** (Alternative)
   - Simpler learning curve
   - Good for forms-heavy apps
   - **Effort:** Low-Medium

3. **HTMX + Alpine.js** (Minimal JS)
   - Server-side rendering
   - Minimal frontend complexity
   - **Effort:** Low (mostly HTML/CSS)

**Recommended Stack:**
```
Frontend: React + TypeScript + Vite + TailwindCSS
Backend: FastAPI + SQLAlchemy (reuse db_pool.py) + Celery (background tasks)
Database: PostgreSQL (existing, no migration needed)
Auth: JWT tokens + OAuth2
Deployment: Docker + Nginx + Gunicorn/Uvicorn
```

### 1.2 Architecture Diagram

```
┌─────────────────────────────────────────────────────────────┐
│                      Web Browser (Client)                    │
│  React SPA: 5 routes (File Config, DB Config, Process, etc) │
└──────────────────────┬──────────────────────────────────────┘
                       │ HTTP/WebSocket
                       ↓
┌─────────────────────────────────────────────────────────────┐
│                    Nginx (Reverse Proxy)                     │
│              SSL Termination + Static Files                  │
└──────────────────────┬──────────────────────────────────────┘
                       │
        ┌──────────────┴───────────────┐
        ↓                              ↓
┌──────────────────┐         ┌─────────────────────┐
│  FastAPI Backend │←────────│  Celery Workers     │
│  (API + WebSocket)│         │  (Excel, Playwright)│
└────────┬─────────┘         └──────────┬──────────┘
         │                              │
         └──────────────┬───────────────┘
                        ↓
              ┌─────────────────────┐
              │  PostgreSQL + Redis │
              │  (Data + Task Queue)│
              └─────────────────────┘
```

### 1.3 Module Mapping

| Desktop Module | Web Equivalent | Migration Strategy |
|----------------|----------------|-------------------|
| `modules/ui/main_window.py` | React components | **Rewrite** (UI logic only) |
| `modules/excel_handler.py` | FastAPI service | **Reuse** (100% backend) |
| `modules/db_pool.py` | FastAPI dependency | **Reuse** (minimal adapt) |
| `modules/input_rencana_benda_uji.py` | Celery task | **Adapt** (async wrapper) |
| `modules/playwright_helpers.py` | Celery task | **Reuse** (containerize browser) |
| `modules/config.py` | FastAPI settings | **Adapt** (Pydantic BaseSettings) |
| `modules/logging_config.py` | Structured logging | **Enhance** (add request IDs) |

---

## Phase 2: Backend Development (Week 2, 40-48 hours)

### 2.1 FastAPI Project Structure

```
backend/
├── app/
│   ├── __init__.py
│   ├── main.py                    # FastAPI app entry
│   ├── config.py                  # Settings (from modules/config.py)
│   ├── dependencies.py            # DB pool, auth, etc.
│   ├── models/                    # SQLAlchemy models
│   │   ├── __init__.py
│   │   ├── user.py
│   │   └── benda_uji.py
│   ├── schemas/                   # Pydantic request/response schemas
│   │   ├── __init__.py
│   │   ├── excel.py
│   │   └── auth.py
│   ├── api/                       # API routes
│   │   ├── __init__.py
│   │   ├── auth.py
│   │   ├── excel.py               # Excel processing endpoints
│   │   ├── database.py            # DB operations
│   │   └── automation.py          # Playwright automation
│   ├── services/                  # Business logic (reuse modules/)
│   │   ├── __init__.py
│   │   ├── excel_service.py       # Wrapper for modules/excel_handler.py
│   │   ├── db_service.py          # Wrapper for modules/db_controller.py
│   │   └── automation_service.py  # Wrapper for modules/input_rencana_benda_uji.py
│   ├── tasks/                     # Celery tasks
│   │   ├── __init__.py
│   │   ├── excel_tasks.py
│   │   └── automation_tasks.py
│   ├── websockets/                # Real-time communication
│   │   ├── __init__.py
│   │   └── logs.py                # Stream logs to frontend
│   └── utils/
│       ├── __init__.py
│       └── auth.py                # JWT utilities
├── modules/                       # Symlink/copy from desktop app
│   ├── excel_handler.py           # REUSE AS-IS
│   ├── db_pool.py                 # REUSE AS-IS
│   ├── playwright_helpers.py      # REUSE AS-IS
│   └── ...
├── tests/
│   ├── test_api/
│   └── test_services/
├── alembic/                       # DB migrations
├── requirements.txt
└── Dockerfile
```

### 2.2 Key API Endpoints

```python
# app/api/excel.py
POST   /api/excel/upload              # Upload Excel file
GET    /api/excel/preview/{file_id}  # Preview Excel data
POST   /api/excel/process             # Start processing (async)
GET    /api/excel/status/{task_id}    # Check processing status

# app/api/database.py
GET    /api/database/connection/test  # Test DB connection
POST   /api/database/upload            # Upload processed data
GET    /api/database/records           # Fetch records (pagination)

# app/api/automation.py
POST   /api/automation/start           # Start Playwright automation
POST   /api/automation/stop            # Stop automation
GET    /api/automation/status          # Check automation status
WS     /api/automation/logs            # WebSocket log stream

# app/api/auth.py
POST   /api/auth/login                 # Login (JWT token)
POST   /api/auth/refresh               # Refresh token
GET    /api/auth/me                    # Get current user
```

### 2.3 Authentication & Authorization

```python
# app/utils/auth.py
from fastapi import Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer
from jose import JWTError, jwt

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="api/auth/login")

async def get_current_user(token: str = Depends(oauth2_scheme)):
    # Verify JWT token
    # Return user from DB
    pass

# Role-based access control
def require_role(role: str):
    def decorator(current_user = Depends(get_current_user)):
        if current_user.role != role:
            raise HTTPException(status_code=403)
        return current_user
    return decorator
```

### 2.4 Background Task Management

```python
# app/tasks/excel_tasks.py
from celery import Celery
from modules.excel_handler import ExcelBebanProcessor

celery_app = Celery('tasks', broker='redis://localhost:6379/0')

@celery_app.task(bind=True)
def process_excel_task(self, file_path: str, user_id: int):
    """Process Excel file in background"""
    try:
        # Update task state
        self.update_state(state='PROCESSING', meta={'progress': 0})
        
        # Reuse existing logic
        processor = ExcelBebanProcessor(file_path)
        result = processor.process()
        
        self.update_state(state='PROCESSING', meta={'progress': 50})
        
        # Upload to DB
        db_service.upload_data(result)
        
        return {'status': 'SUCCESS', 'rows': len(result)}
    except Exception as e:
        self.update_state(state='FAILURE', meta={'error': str(e)})
        raise
```

### 2.5 WebSocket Real-time Logs

```python
# app/websockets/logs.py
from fastapi import WebSocket
import asyncio

class LogStreamer:
    def __init__(self):
        self.connections: list[WebSocket] = []
    
    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.connections.append(websocket)
    
    async def broadcast(self, message: dict):
        for connection in self.connections:
            await connection.send_json(message)

log_streamer = LogStreamer()

@router.websocket("/ws/logs")
async def websocket_logs(websocket: WebSocket):
    await log_streamer.connect(websocket)
    try:
        while True:
            await asyncio.sleep(1)  # Keep connection alive
    except:
        log_streamer.connections.remove(websocket)
```

---

## Phase 3: Frontend Development (Week 3, 40-48 hours)

### 3.1 React Project Structure

```
frontend/
├── public/
├── src/
│   ├── main.tsx                   # Entry point
│   ├── App.tsx                    # Root component + routing
│   ├── components/                # Reusable components
│   │   ├── Layout.tsx             # App layout with navbar
│   │   ├── FileUploader.tsx
│   │   ├── DataTable.tsx          # AG-Grid wrapper
│   │   └── LogViewer.tsx          # Real-time log component
│   ├── pages/                     # Route pages (match desktop tabs)
│   │   ├── FileConfiguration.tsx  # Tab 1
│   │   ├── DatabaseConfiguration.tsx  # Tab 2
│   │   ├── ProcessControl.tsx     # Tab 3
│   │   ├── AutomationInput.tsx    # Tab 4
│   │   └── ProcessLogs.tsx        # Tab 5
│   ├── hooks/                     # Custom React hooks
│   │   ├── useWebSocket.ts        # WebSocket hook for logs
│   │   ├── useApi.ts              # API fetch wrapper
│   │   └── useAuth.ts             # Authentication hook
│   ├── services/                  # API client
│   │   ├── api.ts                 # Axios instance + interceptors
│   │   ├── excelApi.ts
│   │   ├── databaseApi.ts
│   │   └── automationApi.ts
│   ├── store/                     # State management (Zustand/Redux)
│   │   ├── authStore.ts
│   │   └── processStore.ts
│   ├── types/                     # TypeScript types
│   │   ├── excel.ts
│   │   └── automation.ts
│   └── utils/
│       └── formatters.ts
├── package.json
├── vite.config.ts
└── Dockerfile
```

### 3.2 Key Components

**FileUploader Component:**
```tsx
// src/components/FileUploader.tsx
import { useState } from 'react';
import { uploadExcelFile } from '@/services/excelApi';

export const FileUploader = () => {
  const [file, setFile] = useState<File | null>(null);
  const [uploading, setUploading] = useState(false);

  const handleUpload = async () => {
    if (!file) return;
    
    setUploading(true);
    try {
      const formData = new FormData();
      formData.append('file', file);
      const response = await uploadExcelFile(formData);
      console.log('Upload success:', response);
    } catch (error) {
      console.error('Upload failed:', error);
    } finally {
      setUploading(false);
    }
  };

  return (
    <div className="file-uploader">
      <input type="file" onChange={(e) => setFile(e.target.files?.[0] || null)} />
      <button onClick={handleUpload} disabled={uploading}>
        {uploading ? 'Uploading...' : 'Upload'}
      </button>
    </div>
  );
};
```

**WebSocket Log Viewer:**
```tsx
// src/components/LogViewer.tsx
import { useEffect, useState } from 'react';
import { useWebSocket } from '@/hooks/useWebSocket';

export const LogViewer = () => {
  const [logs, setLogs] = useState<string[]>([]);
  const { lastMessage } = useWebSocket('/ws/logs');

  useEffect(() => {
    if (lastMessage) {
      setLogs(prev => [...prev, lastMessage.data]);
    }
  }, [lastMessage]);

  return (
    <div className="log-viewer">
      {logs.map((log, idx) => (
        <div key={idx} className="log-entry">{log}</div>
      ))}
    </div>
  );
};
```

### 3.3 State Management

```typescript
// src/store/processStore.ts (Zustand)
import create from 'zustand';

interface ProcessState {
  isProcessing: boolean;
  progress: number;
  taskId: string | null;
  startProcess: (fileId: string) => Promise<void>;
  stopProcess: () => void;
}

export const useProcessStore = create<ProcessState>((set) => ({
  isProcessing: false,
  progress: 0,
  taskId: null,
  
  startProcess: async (fileId) => {
    const response = await fetch('/api/excel/process', {
      method: 'POST',
      body: JSON.stringify({ file_id: fileId })
    });
    const { task_id } = await response.json();
    set({ isProcessing: true, taskId: task_id });
  },
  
  stopProcess: () => {
    set({ isProcessing: false, progress: 0, taskId: null });
  }
}));
```

---

## Phase 4: Integration & Testing (Week 4, 24-32 hours)

### 4.1 Docker Compose Setup

```yaml
# docker-compose.yml
version: '3.8'

services:
  postgres:
    image: postgres:15
    environment:
      POSTGRES_DB: alatujidb
      POSTGRES_USER: postgres
      POSTGRES_PASSWORD: ${DB_PASSWORD}
    volumes:
      - postgres_data:/var/lib/postgresql/data
    ports:
      - "5432:5432"

  redis:
    image: redis:7-alpine
    ports:
      - "6379:6379"

  backend:
    build: ./backend
    command: uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
    volumes:
      - ./backend:/app
      - ./modules:/app/modules  # Share existing modules
    ports:
      - "8000:8000"
    environment:
      - DATABASE_URL=postgresql://postgres:${DB_PASSWORD}@postgres:5432/alatujidb
      - REDIS_URL=redis://redis:6379/0
      - SESSION_TOKEN=${SESSION_TOKEN}
    depends_on:
      - postgres
      - redis

  celery_worker:
    build: ./backend
    command: celery -A app.tasks.celery_app worker --loglevel=info
    volumes:
      - ./backend:/app
      - ./modules:/app/modules
    environment:
      - DATABASE_URL=postgresql://postgres:${DB_PASSWORD}@postgres:5432/alatujidb
      - REDIS_URL=redis://redis:6379/0
      - SESSION_TOKEN=${SESSION_TOKEN}
    depends_on:
      - postgres
      - redis
      - backend

  playwright:
    build:
      context: ./backend
      dockerfile: Dockerfile.playwright
    command: celery -A app.tasks.celery_app worker -Q playwright --loglevel=info
    volumes:
      - ./backend:/app
      - ./modules:/app/modules
    environment:
      - DATABASE_URL=postgresql://postgres:${DB_PASSWORD}@postgres:5432/alatujidb
      - REDIS_URL=redis://redis:6379/0
      - SESSION_TOKEN=${SESSION_TOKEN}
    depends_on:
      - redis

  frontend:
    build: ./frontend
    ports:
      - "3000:3000"
    volumes:
      - ./frontend:/app
      - /app/node_modules
    environment:
      - VITE_API_URL=http://localhost:8000
    depends_on:
      - backend

  nginx:
    image: nginx:alpine
    ports:
      - "80:80"
      - "443:443"
    volumes:
      - ./nginx/nginx.conf:/etc/nginx/nginx.conf
      - ./nginx/ssl:/etc/nginx/ssl
    depends_on:
      - backend
      - frontend

volumes:
  postgres_data:
```

### 4.2 Testing Strategy

**Backend Tests:**
```python
# backend/tests/test_api/test_excel.py
import pytest
from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)

def test_upload_excel():
    with open("test_data.xlsx", "rb") as f:
        response = client.post(
            "/api/excel/upload",
            files={"file": ("test.xlsx", f, "application/vnd.ms-excel")}
        )
    assert response.status_code == 200
    assert "file_id" in response.json()

def test_process_excel():
    # Upload first
    file_id = "test_file_123"
    
    # Start processing
    response = client.post("/api/excel/process", json={"file_id": file_id})
    assert response.status_code == 202
    assert "task_id" in response.json()
```

**Frontend Tests (Vitest + React Testing Library):**
```typescript
// frontend/src/components/__tests__/FileUploader.test.tsx
import { render, screen, fireEvent } from '@testing-library/react';
import { FileUploader } from '../FileUploader';

test('renders file input and upload button', () => {
  render(<FileUploader />);
  expect(screen.getByText('Upload')).toBeInTheDocument();
});

test('disables button while uploading', async () => {
  render(<FileUploader />);
  const button = screen.getByText('Upload');
  fireEvent.click(button);
  expect(button).toBeDisabled();
});
```

**E2E Tests (Playwright):**
```typescript
// e2e/excel-processing.spec.ts
import { test, expect } from '@playwright/test';

test('complete excel processing flow', async ({ page }) => {
  await page.goto('http://localhost:3000');
  
  // Login
  await page.fill('[name="email"]', 'test@example.com');
  await page.fill('[name="password"]', 'password');
  await page.click('button[type="submit"]');
  
  // Upload Excel
  await page.click('text=File Configuration');
  await page.setInputFiles('input[type="file"]', 'test_data.xlsx');
  await page.click('text=Upload');
  
  // Wait for success
  await expect(page.locator('text=Upload successful')).toBeVisible();
  
  // Start processing
  await page.click('text=Process Control');
  await page.click('text=Start Full Process');
  
  // Check logs
  await expect(page.locator('.log-viewer')).toContainText('Processing started');
});
```

---

## Phase 5: Deployment & Migration (Week 5, 16-24 hours)

### 5.1 Production Deployment Checklist

**Infrastructure:**
- [ ] VPS/Cloud setup (DigitalOcean, AWS, GCP)
- [ ] Domain + SSL certificate (Let's Encrypt)
- [ ] PostgreSQL managed instance (or self-hosted with backups)
- [ ] Redis instance (or use managed service)

**Environment Configuration:**
```bash
# .env.production
DATABASE_URL=postgresql://user:pass@prod-db.example.com:5432/alatujidb
REDIS_URL=redis://prod-redis.example.com:6379/0
SESSION_TOKEN=your_production_token
SECRET_KEY=your_jwt_secret_key_here
ALLOWED_ORIGINS=https://yourdomain.com
```

**Nginx Configuration:**
```nginx
# nginx/nginx.conf
upstream backend {
    server backend:8000;
}

upstream frontend {
    server frontend:3000;
}

server {
    listen 80;
    server_name yourdomain.com;
    return 301 https://$server_name$request_uri;
}

server {
    listen 443 ssl http2;
    server_name yourdomain.com;

    ssl_certificate /etc/nginx/ssl/fullchain.pem;
    ssl_certificate_key /etc/nginx/ssl/privkey.pem;

    # Frontend
    location / {
        proxy_pass http://frontend;
        proxy_set_header Host $host;
    }

    # Backend API
    location /api/ {
        proxy_pass http://backend;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }

    # WebSocket
    location /ws/ {
        proxy_pass http://backend;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
    }
}
```

### 5.2 Data Migration Strategy

**Option A: Gradual Migration (Recommended)**
1. Deploy web app alongside desktop app
2. Users can choose which to use
3. Both apps share same PostgreSQL database
4. Monitor usage, deprecate desktop after 2-3 months

**Option B: Hard Cutover**
1. Schedule maintenance window
2. Backup entire database
3. Deploy web app
4. Notify users of new URL
5. Provide training/documentation

**Database Migration (if needed):**
```python
# alembic/versions/001_add_user_table.py
from alembic import op
import sqlalchemy as sa

def upgrade():
    op.create_table(
        'users',
        sa.Column('id', sa.Integer, primary_key=True),
        sa.Column('email', sa.String(255), unique=True, nullable=False),
        sa.Column('hashed_password', sa.String(255), nullable=False),
        sa.Column('role', sa.String(50), default='user'),
        sa.Column('created_at', sa.DateTime, server_default=sa.func.now())
    )

def downgrade():
    op.drop_table('users')
```

### 5.3 Monitoring & Logging

```python
# backend/app/main.py
from fastapi import FastAPI
from prometheus_fastapi_instrumentator import Instrumentator

app = FastAPI()

# Prometheus metrics
Instrumentator().instrument(app).expose(app)

# Structured logging
import structlog
logger = structlog.get_logger()

@app.middleware("http")
async def log_requests(request, call_next):
    logger.info("request_started", path=request.url.path, method=request.method)
    response = await call_next(request)
    logger.info("request_finished", status_code=response.status_code)
    return response
```

---

## Migration Risks & Mitigations

| Risk | Impact | Probability | Mitigation |
|------|--------|-------------|------------|
| Playwright container fails | High | Medium | Test in Docker early; use xvfb for headless |
| Excel processing timeout | Medium | Medium | Implement progress tracking + resume capability |
| Multi-user concurrency bugs | High | Low | Load testing with locust/k6 before launch |
| User resistance to web UI | Medium | Medium | Gradual migration + training sessions |
| Token expiry handling | Low | High | Implement refresh token flow |
| Database connection pool exhaustion | High | Low | Monitor pool metrics, adjust DB_POOL_SIZE |

---

## Cost Estimation

### Development Costs (if outsourced)
**Zero-Cost Stack:**
- Backend wrapper: 8 hours × $50-100/hr = $400-800
- Frontend: 15 hours × $50-100/hr = $750-1,500
- Deployment: 2 hours × $50-100/hr = $100-200
- **Total:** $1,250-2,500

**Production Stack:**
- Backend: 40 hours × $50-100/hr = $2,000-4,000
- Frontend: 40 hours × $50-100/hr = $2,000-4,000
- DevOps/Deployment: 16 hours × $75-125/hr = $1,200-2,000
- Testing/QA: 24 hours × $40-80/hr = $960-1,920
- **Total:** $6,160-11,920

### Infrastructure Costs (monthly)

| Service | Zero-Cost | Budget | Production |
|---------|-----------|--------|------------|
| **Frontend** | Cloudflare Pages (FREE) | Cloudflare Pages (FREE) | Cloudflare Pages (FREE) |
| **Backend** | Render Free (FREE) | Render Starter ($7) | VPS 4vCPU 8GB ($40-80) |
| **Database** | Neon Free (FREE) | Neon Pro ($19) | Managed PostgreSQL ($15-50) |
| **Redis** | Upstash Free (FREE) | Upstash Pro ($10) | Managed Redis ($10-25) |
| **Monitoring** | cron-job.org (FREE) | Built-in (FREE) | Prometheus + Grafana ($0) |
| **SSL + Domain** | Cloudflare (FREE) | Cloudflare (FREE) | Let's Encrypt ($0-2) |
| **TOTAL/month** | **$0** | **$36** | **$66-157** |

**Constraints by tier:**
- **Zero-Cost:** 50-100 users/mo, 500 automation runs/mo, 0.5GB DB, 30s cold starts
- **Budget:** 200-500 users/mo, 2K automation runs/mo, 3GB DB, no cold starts
- **Production:** 1,000+ users/mo, unlimited runs, unlimited DB, 99.9% SLA

---

## Success Criteria

**Phase 1 (Planning):** ✅ Architecture approved, tech stack decided  
**Phase 2 (Backend):** ✅ All API endpoints working, 80%+ test coverage  
**Phase 3 (Frontend):** ✅ All 5 pages functional, responsive design  
**Phase 4 (Integration):** ✅ E2E tests pass, Docker Compose runs locally  
**Phase 5 (Deployment):** ✅ Production live, <2s page load, zero downtime

**Post-Launch Metrics:**
- Page load time: <2 seconds (vs desktop instant)
- API response time: <500ms (95th percentile)
- Concurrent users: 10+ without degradation
- Uptime: 99.5%+
- User satisfaction: 4/5+ rating

---

## Phase 6: Zero-Cost Deployment (30 Minutes)

### 6.1 The Free Stack Architecture

**Complete $0/month solution using free tier stacking:**

```
┌─────────────────────────────────────────────────────┐
│  Cloudflare Pages (Frontend) - FREE FOREVER         │
│  - Unlimited bandwidth, 500 builds/mo               │
└──────────────────────┬──────────────────────────────┘
                       │ HTTPS
                       ↓
┌─────────────────────────────────────────────────────┐
│  Render.com (Backend) - FREE 750h/mo                │
│  - 512MB RAM, auto-sleep after 15min               │
│  - Keep warm with cron-job.org ping                │
└──────────────────────┬──────────────────────────────┘
                       │
        ┌──────────────┴───────────────┐
        ↓                              ↓
┌────────────────────┐      ┌──────────────────────┐
│  Neon PostgreSQL   │      │  Upstash Redis       │
│  FREE 0.5GB        │      │  FREE 10K cmds/day   │
└────────────────────┘      └──────────────────────┘
```

**Usage Limits:**
- Users: 50-100/month
- Automation runs: 500/month
- Excel files: 200/month (5MB each)
- DB storage: 0.5GB (~500K records)
- Cold start: 30s after 15min idle

### 6.2 Quick Deployment (Step-by-Step)

**Prerequisites:**
- GitHub account
- Accounts on: Cloudflare, Render.com, Neon, Upstash, cron-job.org

**Step 1: Database Setup (3 minutes)**
```bash
# 1. Create Neon account: https://neon.tech
# 2. Create project (free tier auto-selected)
# 3. Copy connection string
DATABASE_URL=postgresql://user:***@ep-xxx.neon.tech/dbname?sslmode=require

# 4. Run existing migrations
psql $DATABASE_URL < database_schema.sql
```

**Step 2: Redis Setup (2 minutes)**
```bash
# 1. Create Upstash account: https://upstash.com
# 2. Create Redis database (free tier)
# 3. Copy REST URL
REDIS_URL=rediss://default:***@global-xxx.upstash.io:6379
```

**Step 3: Backend Deployment (10 minutes)**
```bash
# 1. Create render.yaml in project root
cat > render.yaml << 'EOF'
services:
  - type: web
    name: odoo-backend
    env: python
    buildCommand: "pip install -r requirements.txt && playwright install chromium"
    startCommand: "uvicorn app.main:app --host 0.0.0.0 --port $PORT"
    envVars:
      - key: DATABASE_URL
        sync: false
      - key: REDIS_URL
        sync: false
      - key: SESSION_TOKEN
        sync: false
      - key: PYTHONUNBUFFERED
        value: "1"
EOF

# 2. Push to GitHub
git add render.yaml
git commit -m "feat: add Render deployment config"
git push

# 3. Visit https://dashboard.render.com
# 4. New Web Service → Connect GitHub repo
# 5. Render auto-detects render.yaml
# 6. Add environment variables in dashboard
# 7. Deploy (takes ~5 min)
```

**Step 4: Frontend Deployment (5 minutes)**
```bash
# 1. Build React app
cd frontend
npm run build

# 2. Deploy to Cloudflare Pages
npx wrangler pages deploy dist --project-name=odoo-app

# 3. Update API URL in frontend
# frontend/src/config.ts
export const API_URL = 'https://odoo-backend.onrender.com';

# 4. Rebuild and redeploy
npm run build
npx wrangler pages deploy dist --project-name=odoo-app
```

**Step 5: Keep-Alive Setup (5 minutes)**
```bash
# 1. Visit https://cron-job.org → Sign up
# 2. Create cron job:
#    - Title: "Keep Render Warm"
#    - URL: https://odoo-backend.onrender.com/health
#    - Schedule: */14 * * * * (every 14 minutes)
#    - Enabled: Yes
# 3. Save
```

**Step 6: Test Everything (5 minutes)**
```bash
# Test backend health
curl https://odoo-backend.onrender.com/health

# Test frontend
open https://odoo-app.pages.dev

# Upload test Excel file
# Check logs in Render dashboard
```

### 6.3 Backend Adaptations for Free Tier

**Simplified task queue (no Celery):**
```python
# backend/app/tasks/simple_queue.py
import asyncio
from sqlalchemy import select
from app.models import Task
from datetime import datetime

class SimpleTaskQueue:
    """PostgreSQL-backed task queue (no Redis dependency)"""
    
    def __init__(self, db):
        self.db = db
        self.running = False
    
    async def enqueue(self, task_type: str, payload: dict):
        task = Task(
            type=task_type,
            payload=payload,
            status='pending',
            created_at=datetime.utcnow()
        )
        self.db.add(task)
        await self.db.commit()
        return task.id
    
    async def worker_loop(self):
        self.running = True
        while self.running:
            result = await self.db.execute(
                select(Task)
                .where(Task.status == 'pending')
                .order_by(Task.created_at)
                .limit(1)
            )
            task = result.scalar_one_or_none()
            
            if task:
                try:
                    task.status = 'processing'
                    await self.db.commit()
                    
                    if task.type == 'excel_process':
                        from modules.excel_handler import ExcelBebanProcessor
                        processor = ExcelBebanProcessor(task.payload['file_path'])
                        result = processor.process()
                        task.result = result
                    
                    elif task.type == 'playwright_automate':
                        from modules.input_rencana_benda_uji import run_with_custom_path_and_stop
                        result = run_with_custom_path_and_stop(
                            task.payload['excel_path'],
                            stop_event=None
                        )
                        task.result = result
                    
                    task.status = 'completed'
                except Exception as e:
                    task.status = 'failed'
                    task.error = str(e)
                
                await self.db.commit()
            
            await asyncio.sleep(5)

# backend/app/main.py
from fastapi import FastAPI
from app.tasks.simple_queue import SimpleTaskQueue

app = FastAPI()
task_queue = None

@app.on_event("startup")
async def startup():
    global task_queue
    from app.dependencies import get_db
    db = next(get_db())
    task_queue = SimpleTaskQueue(db)
    asyncio.create_task(task_queue.worker_loop())

@app.post("/api/excel/upload")
async def upload_excel(file: UploadFile):
    file_path = f"/tmp/{file.filename}"
    with open(file_path, "wb") as f:
        f.write(await file.read())
    
    task_id = await task_queue.enqueue('excel_process', {'file_path': file_path})
    return {"task_id": task_id, "status": "queued"}
```

**Lightweight Playwright config (512MB RAM):**
```python
# modules/playwright_helpers.py (update setup_driver)
def setup_driver(session_token: str = None) -> BrowserWrapper:
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=True,
        args=[
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--disable-software-rasterizer',
            '--single-process',  # Critical for low RAM
            '--disable-extensions',
            '--disable-background-networking',
            '--disable-default-apps',
            '--no-first-run',
            '--mute-audio',
        ]
    )
    # ... rest of setup
```

**Health endpoint for cron ping:**
```python
# backend/app/main.py
@app.get("/health")
async def health_check():
    """Keep Render instance warm (pinged by cron-job.org)"""
    return {
        "status": "ok",
        "timestamp": datetime.utcnow().isoformat(),
        "memory_mb": psutil.virtual_memory().used / 1024 / 1024
    }
```

### 6.4 Monitoring Free Tier Usage

**Add usage stats endpoint:**
```python
# backend/app/api/stats.py
import psutil
from datetime import datetime

@app.get("/api/stats")
async def get_usage_stats():
    """Monitor resource usage against free tier limits"""
    return {
        "memory": {
            "used_mb": psutil.virtual_memory().used / 1024 / 1024,
            "percent": psutil.virtual_memory().percent,
            "limit_mb": 512,
            "warning": psutil.virtual_memory().percent > 80
        },
        "database": {
            "size_estimate_gb": await get_db_size(),
            "limit_gb": 0.5,
            "compute_hours_used": await get_neon_hours()  # Check via Neon API
        },
        "uptime_hours": (datetime.now() - app.start_time).total_seconds() / 3600
    }
```

**Check stats weekly:**
```bash
curl https://odoo-backend.onrender.com/api/stats | jq
```

### 6.5 When to Upgrade

**Free tier sufficient if:**
- ✅ <100 users/month
- ✅ <500 automation runs/month
- ✅ <0.5GB database
- ✅ Can tolerate 30s cold starts
- ✅ Excel files <10MB

**Upgrade to Budget tier ($36/mo) when:**
- ❌ >200 users/month
- ❌ Database >0.5GB
- ❌ Need instant response (no cold starts)
- ❌ >1,000 automation runs/month

**Upgrade path:**
1. Render Starter: $7/mo (1GB RAM, no sleep)
2. Neon Pro: $19/mo (3GB DB, unlimited compute)
3. Upstash Pro: $10/mo (100K commands/day)

**Total: $36/mo** (still 45% cheaper than traditional cloud)

### 6.6 Troubleshooting Free Tier

| Issue | Cause | Fix |
|-------|-------|-----|
| **30s cold start** | Render auto-sleep | Setup cron-job.org ping every 14min |
| **Memory crash** | >512MB usage | Add `--single-process` to Playwright, process smaller batches |
| **DB suspended** | >100 compute hours/mo | Optimize queries, add caching, upgrade to Neon Pro |
| **Redis limit hit** | >10K cmds/day | Switch to PostgreSQL-backed queue (shown above) |
| **Excel timeout** | Processing >30s | Chunk large files, add progress tracking |

**Complete troubleshooting guide:** See ZERO_COST_HOSTING.md

---

## Timeline Summary

### Zero-Cost Stack (Recommended for Start)
```
Day 1: Setup & Deploy (4-6 hours)
├── Hours 1-2: Create accounts (Cloudflare, Render, Neon, Upstash)
├── Hour 3: Adapt backend for free tier (SimpleTaskQueue, lightweight Playwright)
├── Hour 4: Deploy backend to Render
├── Hour 5: Deploy frontend to Cloudflare Pages
└── Hour 6: Test end-to-end, setup cron ping

Day 2-5: Frontend Development (16-20 hours)
├── Build React components (FileUploader, LogViewer, DataTable)
├── API integration with Render backend
├── State management (Zustand)
└── Responsive design + testing

Total: 20-27 hours → Working web app at $0/month
```

### Production Stack (Enterprise)
```
Week 1: Architecture & Planning
├── Days 1-2: Tech stack research + decision
├── Days 3-4: API design + component planning
└── Day 5: Docker setup + proof of concept

Week 2: Backend Development
├── Days 6-7: FastAPI setup + auth + endpoints
├── Days 8-9: Celery tasks + WebSocket
└── Day 10: Backend testing

Week 3: Frontend Development
├── Days 11-12: React setup + routing + components
├── Days 13-14: API integration + state management
└── Day 15: Frontend testing

Week 4: Integration & Testing
├── Days 16-17: Docker Compose + E2E tests
├── Days 18-19: Performance testing + bug fixes
└── Day 20: Documentation + user guides

Week 5: Deployment
├── Days 21-22: Production setup + CI/CD
├── Day 23: Data migration + smoke tests
├── Day 24: Go-live + monitoring
└── Day 25: Handoff + training

Total: 120-160 hours → Enterprise-grade at $66-157/month
```

---

## Next Steps (Choose Your Path)

### Path A: Zero-Cost Quick Start (RECOMMENDED)
**Time: 30 minutes to working prototype**

1. **Create accounts** (10 min)
   - Cloudflare, Render, Neon, Upstash, cron-job.org
   
2. **Deploy backend** (10 min)
   - Follow Phase 6, Step 3
   - Uses existing `modules/` code (90% reuse)
   
3. **Deploy frontend** (5 min)
   - Build React app
   - Deploy to Cloudflare Pages
   
4. **Test** (5 min)
   - Upload Excel file
   - Check automation runs

**Result:** Working web app at $0/month, 50-100 users capacity

### Path B: Production Deployment
**Time: 3-4 hours setup + 120-160 hours development**

1. **Provision infrastructure** (2 hours)
   - VPS/cloud setup (DigitalOcean/AWS)
   - Managed PostgreSQL + Redis
   
2. **Follow Phases 1-5** (3-4 weeks)
   - Full Docker Compose stack
   - Celery for background tasks
   - Complete testing suite
   
3. **Go live** (1 day)
   - CI/CD pipeline
   - Monitoring + alerts
   - Production deployment

**Result:** Enterprise-grade web app at $66-157/month, unlimited scale

### Path C: Hybrid Approach
**Time: 1 week**

1. **Start with zero-cost** (Day 1)
   - Deploy to free tier
   - Validate with users
   
2. **Gather metrics** (Days 2-5)
   - Monitor usage
   - Collect user feedback
   
3. **Scale if needed** (Days 6-7)
   - Upgrade to paid tiers selectively
   - Keep frontend on Cloudflare (free)
   - Upgrade backend only if needed

**Result:** Optimal cost vs. performance

---

## Decision Matrix

| Factor | Zero-Cost | Budget | Production |
|--------|-----------|--------|------------|
| **Setup Time** | 30 min | 2-3 hours | 3-4 hours |
| **Dev Time** | 20-27 hours | 30-40 hours | 120-160 hours |
| **Monthly Cost** | $0 | $36 | $66-157 |
| **Users** | <100 | <500 | 1,000+ |
| **Cold Starts** | Yes (30s) | No | No |
| **Support** | Community | Email | Priority SLA |
| **Best For** | POC, demo | Small business | Enterprise |

**Recommendation:** Start with Path A (zero-cost), upgrade to Path C if traction, move to Path B only if business-critical.

---

## References & Resources

**FastAPI:**
- Docs: https://fastapi.tiangolo.com/
- Full-stack template: https://github.com/tiangolo/full-stack-fastapi-postgresql

**React:**
- Docs: https://react.dev/
- Vite: https://vitejs.dev/

**Playwright in Docker:**
- Official images: https://playwright.dev/docs/docker
- Headless config: https://playwright.dev/docs/ci

**Deployment:**
- DigitalOcean App Platform: https://www.digitalocean.com/products/app-platform
- Railway: https://railway.app/ (simpler, pricier)
- Render: https://render.com/ (free tier available)
