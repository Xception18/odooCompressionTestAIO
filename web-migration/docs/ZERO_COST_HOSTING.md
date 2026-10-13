# Zero-Cost Web Hosting Strategy

## TL;DR: ✅ Possible with Aggressive Free Tier Stacking

**Target Cost:** $0/month (100% free tier exploitation)  
**Trade-offs:** Usage limits, potential downtime, manual management  
**Suitable for:** Prototype, demo, low-traffic (<100 users/month), personal use  
**NOT suitable for:** Production, business-critical, high-volume automation

---

## The Free Stack (Updated October 2026)

```
┌─────────────────────────────────────────────────────┐
│  Cloudflare Pages (Frontend) - FREE FOREVER         │
│  - Unlimited bandwidth                              │
│  - 500 builds/month                                 │
│  - Custom domain support                            │
└──────────────────────┬──────────────────────────────┘
                       │ HTTPS
                       ↓
┌─────────────────────────────────────────────────────┐
│  Render.com (Backend) - FREE TIER                   │
│  - 750 hours/month (enough for 24/7 uptime)        │
│  - Auto-sleep after 15 min inactivity              │
│  - 512MB RAM, shared CPU                           │
│  - Python/Node.js support                          │
└──────────────────────┬──────────────────────────────┘
                       │
        ┌──────────────┴───────────────┐
        ↓                              ↓
┌────────────────────┐      ┌──────────────────────┐
│  Neon PostgreSQL   │      │  Upstash Redis       │
│  FREE TIER         │      │  FREE TIER           │
│  - 0.5GB storage   │      │  - 10,000 cmds/day   │
│  - 100 hours/mo    │      │  - 256MB storage     │
│  - Auto-suspend    │      │  - Max 100 conn      │
└────────────────────┘      └──────────────────────┘
```

**Total Monthly Cost:** $0.00

---

## Component Breakdown

### 1. Frontend: Cloudflare Pages (FREE ∞)

**Limits:**
- ✅ Unlimited bandwidth
- ✅ 500 builds/month (1 deploy every ~1.5 hours)
- ✅ Unlimited static requests
- ✅ Custom domain + SSL free

**Setup:**
```bash
# 1. Build React app
cd frontend
npm run build

# 2. Deploy to Cloudflare Pages
npx wrangler pages deploy dist --project-name=odoo-app

# Alternative: Connect GitHub repo (auto-deploy on push)
# Visit: https://dash.cloudflare.com/pages
```

**No limits hit unless:** Rebuilding >500 times/month (unlikely)

---

### 2. Backend: Render.com Free Tier

**Why Render over Railway/Heroku:**
- Railway: $5 minimum spend (not free)
- Heroku: Retired free tier in 2022
- Render: Still has generous free tier (as of Oct 2026)

**Limits:**
- ✅ 750 hours/month = 31.25 days (enough for 24/7)
- ⚠️ Auto-sleeps after 15 min inactivity (cold start: ~30s)
- ⚠️ 512MB RAM (tight but workable)
- ⚠️ Shared CPU (slow Excel processing)

**Setup:**
```bash
# 1. Create render.yaml
# render.yaml
services:
  - type: web
    name: odoo-backend
    env: python
    buildCommand: "pip install -r requirements.txt"
    startCommand: "uvicorn app.main:app --host 0.0.0.0 --port $PORT"
    envVars:
      - key: DATABASE_URL
        fromDatabase:
          name: neon-db
          property: connectionString
      - key: REDIS_URL
        sync: false
      - key: SESSION_TOKEN
        sync: false

# 2. Deploy
# Connect GitHub repo at https://dashboard.render.com
# Auto-deploys on git push
```

**Workaround for cold starts:**
- Use **cron-job.org** (free) to ping `/health` every 14 min
- Keeps server warm during work hours

**Playwright in 512MB RAM:**
```python
# Use playwright-python with minimal browser
from playwright.sync_api import sync_playwright

def setup_driver_lightweight():
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=True,
        args=[
            '--no-sandbox',
            '--disable-dev-shm-usage',
            '--disable-gpu',
            '--disable-software-rasterizer',
            '--disable-extensions',
            '--disable-background-networking',
            '--disable-default-apps',
            '--disable-sync',
            '--metrics-recording-only',
            '--no-first-run',
            '--mute-audio',
            '--hide-scrollbars',
            '--single-process',  # Critical for 512MB RAM
        ]
    )
    return BrowserWrapper(pw, browser, browser.new_page())
```

---

### 3. Database: Neon PostgreSQL Free Tier

**Limits:**
- ✅ 0.5GB storage (~500,000 rows of benda uji data)
- ⚠️ 100 compute hours/month (~3.3 hours/day)
- ⚠️ Auto-suspends after 5 min inactivity
- ✅ 1 database, unlimited tables

**Setup:**
```bash
# 1. Create account: https://neon.tech
# 2. Create project (free tier auto-selected)
# 3. Copy connection string

# 4. Update .env
DATABASE_URL=postgresql://user:pass@ep-xxx.neon.tech/dbname?sslmode=require

# 5. Run migrations (reuse existing schema)
psql $DATABASE_URL < database_schema.sql
```

**Staying within 100 hours/month:**
- Auto-suspend helps (inactive = no billing)
- Average usage: 30 automation runs/day × 2 min each = 60 min/day = 30 hours/month ✅

**If you exceed 100 hours:**
- Upgrade to paid tier: $19/month (still cheaper than VPS)
- OR optimize queries (use connection pooling, reduce DB calls)

---

### 4. Redis: Upstash (FREE)

**Why Redis needed:**
- Task queue for background jobs (Excel processing)
- Session storage
- Rate limiting

**Limits:**
- ✅ 10,000 commands/day (~7 commands/min sustained)
- ✅ 256MB max data size
- ✅ Max 100 concurrent connections
- ⚠️ Resets daily (not persistent beyond 24h)

**Setup:**
```bash
# 1. Create account: https://upstash.com
# 2. Create Redis database (free tier)
# 3. Copy connection string

# 4. Update .env
REDIS_URL=rediss://default:xxx@global-xxx.upstash.io:6379
```

**Staying within 10,000 commands/day:**
- 30 automation runs/day × 200 Redis ops each = 6,000 commands/day ✅
- If exceeded: Implement simple in-memory queue (lose persistence)

---

### 5. Background Tasks: No Celery (Use Simple Queue)

**Problem:** Celery needs persistent Redis/RabbitMQ (exceeds free limits)

**Solution:** Lightweight task queue in SQLite or PostgreSQL

```python
# app/tasks/simple_queue.py
import asyncio
from sqlalchemy import select
from app.models import Task

class SimpleTaskQueue:
    def __init__(self, db):
        self.db = db
        self.running = False
    
    async def enqueue(self, task_type: str, payload: dict):
        """Add task to queue (stored in PostgreSQL)"""
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
        """Process tasks from queue (runs in background thread)"""
        self.running = True
        while self.running:
            # Get next pending task
            result = await self.db.execute(
                select(Task).where(Task.status == 'pending').limit(1)
            )
            task = result.scalar_one_or_none()
            
            if task:
                try:
                    # Process task
                    if task.type == 'excel_process':
                        await process_excel(task.payload)
                    elif task.type == 'playwright_automate':
                        await run_playwright(task.payload)
                    
                    task.status = 'completed'
                except Exception as e:
                    task.status = 'failed'
                    task.error = str(e)
                
                await self.db.commit()
            
            await asyncio.sleep(5)  # Check every 5 seconds

# Start worker on app startup
@app.on_event("startup")
async def start_worker():
    asyncio.create_task(task_queue.worker_loop())
```

**No external service needed** ✅

---

### 6. File Storage: Cloudflare R2 (FREE)

**Limits:**
- ✅ 10GB storage/month
- ✅ 1 million Class A operations (PUT, LIST)
- ✅ 10 million Class B operations (GET, HEAD)
- ✅ No egress fees

**Setup:**
```bash
# 1. Enable R2 in Cloudflare dashboard
# 2. Create bucket: odoo-uploads

# 3. Install R2 client
pip install boto3

# 4. Upload Excel files
import boto3

s3 = boto3.client(
    's3',
    endpoint_url='https://xxx.r2.cloudflarestorage.com',
    aws_access_key_id=os.getenv('R2_ACCESS_KEY'),
    aws_secret_access_key=os.getenv('R2_SECRET_KEY')
)

s3.upload_file('local_file.xlsx', 'odoo-uploads', 'uploads/file.xlsx')
```

**Alternative (if R2 setup too complex):**
Store small files (<10MB) as base64 in PostgreSQL:
```python
# Store Excel in DB temporarily
excel_data = base64.b64encode(file_bytes).decode()
db_record.excel_base64 = excel_data
```

---

## Complete Free Architecture

```python
# backend/app/main.py - Complete FastAPI app for Render free tier
from fastapi import FastAPI, UploadFile, Depends
from fastapi.middleware.cors import CORSMiddleware
import asyncio

app = FastAPI(title="Odoo Compression Test - Free Tier")

# CORS for Cloudflare Pages frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://your-app.pages.dev"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Reuse existing modules (copy from desktop app)
from modules.excel_handler import ExcelBebanProcessor
from modules.db_pool import get_db_pool
from modules.playwright_helpers import setup_driver

# Simple in-memory task queue (no Redis needed for <100 concurrent users)
from app.tasks.simple_queue import SimpleTaskQueue
task_queue = SimpleTaskQueue(db=get_db_pool())

@app.post("/api/excel/upload")
async def upload_excel(file: UploadFile):
    """Upload Excel file and queue processing"""
    # Save to temp location
    file_path = f"/tmp/{file.filename}"
    with open(file_path, "wb") as f:
        f.write(await file.read())
    
    # Queue background task
    task_id = await task_queue.enqueue(
        task_type='excel_process',
        payload={'file_path': file_path}
    )
    
    return {"task_id": task_id, "status": "queued"}

@app.post("/api/automation/start")
async def start_automation(excel_data: dict):
    """Start Playwright automation"""
    task_id = await task_queue.enqueue(
        task_type='playwright_automate',
        payload=excel_data
    )
    return {"task_id": task_id}

@app.get("/api/tasks/{task_id}")
async def get_task_status(task_id: int):
    """Check task status"""
    task = await task_queue.get_task(task_id)
    return {
        "status": task.status,
        "progress": task.progress,
        "result": task.result
    }

@app.get("/health")
async def health_check():
    """Keep Render instance warm (called by cron-job.org)"""
    return {"status": "ok"}

# Start background worker
@app.on_event("startup")
async def startup():
    asyncio.create_task(task_queue.worker_loop())
```

---

## Deployment Steps (30 Minutes Total)

### Step 1: Frontend (Cloudflare Pages) - 5 min

```bash
cd frontend
npm run build
npx wrangler pages deploy dist --project-name=odoo-app
# Done! https://odoo-app.pages.dev
```

### Step 2: Database (Neon) - 3 min

1. Visit https://neon.tech → Sign up
2. Create project → Copy connection string
3. Run migrations:
```bash
psql $DATABASE_URL < database_schema.sql
```

### Step 3: Redis (Upstash) - 2 min

1. Visit https://upstash.com → Sign up
2. Create database → Copy connection string
3. Add to `.env`

### Step 4: Backend (Render) - 10 min

1. Visit https://dashboard.render.com → Sign up
2. New Web Service → Connect GitHub repo
3. Configure:
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `uvicorn app.main:app --host 0.0.0.0 --port $PORT`
   - Add environment variables (DATABASE_URL, SESSION_TOKEN, etc.)
4. Deploy → Wait 5 min
5. Done! https://odoo-backend.onrender.com

### Step 5: Keep Alive (cron-job.org) - 5 min

1. Visit https://cron-job.org → Sign up
2. Create job:
   - URL: `https://odoo-backend.onrender.com/health`
   - Interval: Every 14 minutes
   - Enabled
3. Done! Server stays warm

### Step 6: Connect Frontend to Backend - 5 min

```typescript
// frontend/src/config.ts
export const API_URL = import.meta.env.PROD 
  ? 'https://odoo-backend.onrender.com'
  : 'http://localhost:8000';
```

Rebuild and redeploy frontend.

---

## Usage Limits & Warnings

### What Will Break Free Tier:

| Scenario | Impact | Workaround |
|----------|--------|------------|
| **>100 hours/month DB usage** | Neon suspends | Optimize queries, cache data |
| **>10,000 Redis cmds/day** | Upstash rejects | Use in-memory queue |
| **>512MB RAM on Render** | Process crashes | Optimize Excel processing, process smaller batches |
| **Excel file >10MB** | Slow processing | Reject large files with error message |
| **>30 concurrent automations** | Queue backlog | Implement rate limiting (1 automation/min) |

### Expected Performance:

| Operation | Free Tier | Paid Tier |
|-----------|-----------|-----------|
| Excel upload (5MB file) | 2-3 seconds | <1 second |
| Excel processing (100 rows) | 15-30 seconds | 5-10 seconds |
| Playwright automation (1 form) | 10-15 seconds | 5-8 seconds |
| Cold start (after 15 min idle) | 30-45 seconds | Instant |
| Page load (frontend) | <1 second (CDN) | <1 second |

### Usage Caps for $0/month:

- **Users:** ~50-100 active users/month
- **Excel processing:** ~200 files/month (if 5MB each)
- **Automation runs:** ~500 runs/month (10 min each)
- **Data storage:** ~500,000 records in PostgreSQL
- **File storage:** ~100 Excel files in R2 (10GB total)

**Beyond these limits:** Upgrade to paid tiers ($19-50/month)

---

## Free Tier Maximization Tips

### 1. Optimize for Cold Starts
```python
# app/main.py
import asyncio

# Preload heavy modules on startup
@app.on_event("startup")
async def warmup():
    # Initialize DB connection pool
    get_db_pool()
    
    # Preload Playwright browser (stays in memory)
    global browser
    browser = setup_driver_lightweight()
    
    # Cache frequently-used data
    await cache_mutu_data()
```

### 2. Implement Request Queuing
```python
# Limit concurrent automations to avoid RAM exhaustion
from fastapi import HTTPException

MAX_CONCURRENT_TASKS = 3
current_tasks = 0

@app.post("/api/automation/start")
async def start_automation(data: dict):
    global current_tasks
    if current_tasks >= MAX_CONCURRENT_TASKS:
        raise HTTPException(
            status_code=429,
            detail="Too many concurrent tasks. Please wait."
        )
    current_tasks += 1
    try:
        result = await run_automation(data)
    finally:
        current_tasks -= 1
    return result
```

### 3. Use Aggressive Caching
```python
# Cache processed Excel results in PostgreSQL
from functools import lru_cache

@lru_cache(maxsize=100)
def get_mutu_calculation(mutu: str, umur: int):
    # Expensive calculation cached in memory
    return calculate_beban(mutu, umur)
```

### 4. Compress Responses
```python
# Reduce bandwidth usage (stays under R2 limits)
from fastapi.middleware.gzip import GZipMiddleware

app.add_middleware(GZipMiddleware, minimum_size=1000)
```

---

## Migration from Desktop App (Minimal Changes)

**Reuse 90% of existing code:**

```python
# backend/app/services/excel_service.py
# Just wrap existing modules!

from modules.excel_handler import ExcelBebanProcessor
from modules.db_controller import upload_to_database

async def process_excel_file(file_path: str):
    # Reuse desktop app logic unchanged
    processor = ExcelBebanProcessor(file_path, sheet_name='ODOO')
    processor.process()
    
    # Upload results
    upload_to_database(processor.results)
    
    return {"status": "success", "rows": len(processor.results)}
```

**Only changes:**
1. Wrap in FastAPI routes (2-3 hours)
2. Replace GUI with React components (15-20 hours)
3. Add simple task queue (3-4 hours)

**Total migration effort:** 20-27 hours (vs 40-60h for Cloudflare Workers rewrite)

---

## Monitoring Your Free Tier Usage

```python
# app/monitoring.py
import psutil
from datetime import datetime

@app.get("/api/stats")
async def get_stats():
    """Monitor resource usage"""
    return {
        "memory_mb": psutil.virtual_memory().used / 1024 / 1024,
        "memory_percent": psutil.virtual_memory().percent,
        "cpu_percent": psutil.cpu_percent(interval=1),
        "disk_used_gb": psutil.disk_usage('/').used / 1024 / 1024 / 1024,
        "uptime_hours": (datetime.now() - app.start_time).total_seconds() / 3600
    }
```

Visit `/api/stats` weekly to check if approaching limits.

---

## When to Upgrade (Break-Even Analysis)

### Stay Free If:
- ✅ <100 users/month
- ✅ <500 automation runs/month
- ✅ <0.5GB database size
- ✅ Can tolerate 30s cold starts
- ✅ Processing <200 Excel files/month

### Upgrade to Paid ($20-50/month) When:
- ❌ >200 users/month
- ❌ Need <5s response times (no cold starts)
- ❌ >1,000 automation runs/month
- ❌ Database >0.5GB
- ❌ Business-critical (need uptime SLA)

**Upgrade path:**
1. Neon Pro: $19/mo (3GB DB, no compute limits)
2. Render Starter: $7/mo (1GB RAM, no sleep)
3. Upstash Pro: $10/mo (100K cmds/day)
**Total:** $36/mo for "no limits" tier

---

## Comparison: Free vs Paid

| Feature | Free Tier (This Plan) | Paid Cloud ($66/mo) |
|---------|----------------------|---------------------|
| **Cost** | $0/month | $66-157/month |
| **Setup Time** | 30 minutes | 3-4 hours |
| **Cold Starts** | Yes (30s) | No (instant) |
| **RAM** | 512MB | 8GB |
| **DB Storage** | 0.5GB | Unlimited |
| **Automation Speed** | 15s/run | 5s/run |
| **Concurrent Users** | ~50 | 1,000+ |
| **Uptime** | 99% (with cron ping) | 99.9% SLA |
| **Support** | Community | Priority |
| **Suitable For** | Personal, demo, <100 users | Production, business |

---

## Complete Free Stack Checklist

- [ ] **Frontend:** Deploy React to Cloudflare Pages (FREE ∞)
- [ ] **Backend:** Deploy FastAPI to Render.com (FREE 750h/mo)
- [ ] **Database:** Setup Neon PostgreSQL (FREE 0.5GB)
- [ ] **Redis:** Setup Upstash Redis (FREE 10K cmds/day)
- [ ] **Storage:** Enable Cloudflare R2 (FREE 10GB)
- [ ] **Monitoring:** Setup cron-job.org ping (FREE)
- [ ] **Domain:** Use Cloudflare DNS (FREE) or .pages.dev subdomain
- [ ] **SSL:** Auto-provisioned by Cloudflare (FREE)

**Total Setup Time:** 30-45 minutes  
**Monthly Cost:** $0.00  
**Maintenance:** 1 hour/month (check usage stats)

---

## Final Recommendation

**✅ YES, use this free stack if:**
- Personal project or demo
- <100 users/month
- Learning/experimentation
- Budget absolutely $0

**❌ NO, upgrade to paid if:**
- Business/production use
- Need reliability SLA
- >100 users/month
- Can't tolerate cold starts

**Next Step:** Follow deployment steps above. Entire setup in <1 hour.

**Need help?** Share any deployment errors and I'll debug.
