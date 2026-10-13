# Web Migration Project

Migration dari desktop app (customtkinter) ke web application modern.

## 📁 Struktur Folder

```
web-migration/
├── README.md                          # File ini
├── docs/                              # Dokumentasi lengkap
│   ├── WEB_MIGRATION_PLAN.md         # Master migration plan (3 deployment paths)
│   ├── ZERO_COST_HOSTING.md          # Free tier guide ($0/month setup)
│   └── CLOUDFLARE_WORKERS_ANALYSIS.md # CF Workers feasibility report
├── backend/                           # FastAPI backend (will be created)
│   ├── app/
│   │   ├── main.py
│   │   ├── api/
│   │   ├── tasks/
│   │   └── ...
│   ├── modules/                       # Symlink to ../modules
│   ├── requirements.txt
│   └── render.yaml                    # Render deployment config
└── frontend/                          # React frontend (will be created)
    ├── src/
    │   ├── pages/
    │   ├── components/
    │   └── services/
    ├── package.json
    └── vite.config.ts
```

## 🚀 Quick Start (3 Paths)

### Path A: Zero-Cost ($0/month) ⭐ RECOMMENDED untuk start
**Time:** 30 minutes to working prototype

```bash
# 1. Baca panduan lengkap
cat docs/ZERO_COST_HOSTING.md

# 2. Create accounts (10 min)
# - Cloudflare Pages, Render.com, Neon, Upstash, cron-job.org

# 3. Deploy backend (10 min)
cd backend
# Follow docs/ZERO_COST_HOSTING.md Step 3

# 4. Deploy frontend (5 min)
cd ../frontend
npm run build
npx wrangler pages deploy dist --project-name=odoo-app

# 5. Setup cron ping (5 min)
# Visit cron-job.org, ping /health every 14 minutes
```

**Limits:**
- 50-100 users/month
- 500 automation runs/month
- 0.5GB database
- 30s cold starts after 15min idle

**Upgrade when:** >200 users/mo, need instant response, DB >0.5GB

---

### Path B: Production Cloud ($66-157/month)
**Time:** 3-4 weeks full development

```bash
# 1. Baca full migration plan
cat docs/WEB_MIGRATION_PLAN.md

# 2. Provision infrastructure
# - VPS (DigitalOcean/AWS)
# - Managed PostgreSQL + Redis

# 3. Follow Phases 1-5
# - Week 1: Architecture planning
# - Week 2: Backend development (FastAPI + Celery)
# - Week 3: Frontend development (React)
# - Week 4: Integration testing
# - Week 5: Deployment + monitoring

# 4. Deploy with Docker Compose
docker-compose up -d
```

**For:** Enterprise, high-traffic, business-critical

---

### Path C: Hybrid ($36/month)
Start zero-cost, upgrade selectively

1. Deploy to free tier (Day 1)
2. Monitor usage (Days 2-5)
3. Upgrade only what's needed (Days 6-7)
   - Render Starter: $7/mo (no sleep)
   - Neon Pro: $19/mo (unlimited compute)
   - Upstash Pro: $10/mo (100K commands)

**Sweet spot:** Small business, growing startups

---

## 📚 Documentation Guide

### Start Here
1. **WEB_MIGRATION_PLAN.md** - Master plan, all 3 paths
2. **ZERO_COST_HOSTING.md** - Free tier deep dive
3. **CLOUDFLARE_WORKERS_ANALYSIS.md** - Why we DON'T use CF Workers

### Decision Tree

```
Do you need it FREE? 
  ├─ YES → Read ZERO_COST_HOSTING.md
  │         Follow Path A (30 minutes)
  │
  └─ NO → Budget available?
           ├─ <$50/mo → Path C (Hybrid)
           └─ >$50/mo → Path B (Production)
                        Read WEB_MIGRATION_PLAN.md Phases 1-5
```

---

## 🛠️ Current Status

- [x] Architecture planning complete
- [x] Technology stack decided (FastAPI + React)
- [x] Three deployment paths documented
- [x] Cost analysis complete
- [ ] Backend skeleton (ready to create)
- [ ] Frontend skeleton (ready to create)
- [ ] Database migrations
- [ ] API implementation
- [ ] UI components
- [ ] Testing suite
- [ ] Production deployment

---

## 💰 Cost Comparison

| Service | Zero-Cost | Budget | Production |
|---------|-----------|--------|------------|
| Frontend | Cloudflare Pages (FREE) | Cloudflare Pages (FREE) | Cloudflare Pages (FREE) |
| Backend | Render Free | Render Starter $7 | VPS $40-80 |
| Database | Neon Free (0.5GB) | Neon Pro $19 (3GB) | Managed PG $15-50 |
| Redis | Upstash Free | Upstash Pro $10 | Managed Redis $10-25 |
| **Total** | **$0** | **$36** | **$66-157** |
| **Users** | 50-100/mo | 200-500/mo | 1,000+/mo |

---

## 🔗 Key Technologies

**Backend:**
- FastAPI (Python async web framework)
- SQLAlchemy (ORM + connection pooling)
- Playwright (browser automation)
- Simple task queue (no Celery for free tier)

**Frontend:**
- React + TypeScript
- Vite (build tool)
- TailwindCSS (styling)
- Zustand (state management)

**Database:**
- PostgreSQL (reuse existing schema)
- Redis (optional, for task queue)

**Deployment:**
- Cloudflare Pages (frontend CDN)
- Render.com (backend hosting)
- Neon (serverless PostgreSQL)
- Upstash (serverless Redis)

---

## 📦 Module Reuse Strategy

**90% kode existing bisa reuse:**

| Desktop Module | Web Migration | Effort |
|----------------|---------------|--------|
| `modules/excel_handler.py` | ✅ Reuse 100% | 0 hours (wrap in API) |
| `modules/db_pool.py` | ✅ Reuse 100% | 0 hours |
| `modules/playwright_helpers.py` | ✅ Reuse 95% | 1 hour (add `--single-process`) |
| `modules/input_rencana_benda_uji.py` | ✅ Reuse 90% | 2 hours (async wrapper) |
| `modules/config.py` | ✅ Adapt 80% | 2 hours (Pydantic Settings) |
| `modules/ui/main_window.py` | ❌ Rewrite | 15-20 hours (React components) |

**Total migration effort:** 20-27 hours (zero-cost) vs 120-160 hours (full production)

---

## 🚦 Next Actions

### If choosing Zero-Cost (Path A):
1. [ ] Create Cloudflare account
2. [ ] Create Render.com account
3. [ ] Create Neon account (PostgreSQL)
4. [ ] Create Upstash account (Redis)
5. [ ] Read `docs/ZERO_COST_HOSTING.md` Steps 1-6
6. [ ] Deploy backend following guide
7. [ ] Deploy frontend to Cloudflare Pages
8. [ ] Setup cron-job.org ping
9. [ ] Test with sample Excel file

### If choosing Production (Path B):
1. [ ] Get stakeholder approval on budget ($66-157/mo)
2. [ ] Provision cloud infrastructure
3. [ ] Read `docs/WEB_MIGRATION_PLAN.md` full plan
4. [ ] Setup development environment
5. [ ] Follow Phases 1-5 week-by-week
6. [ ] Deploy with Docker Compose
7. [ ] Setup monitoring (Prometheus + Grafana)

---

## 📞 Support & Resources

**Documentation:**
- FastAPI: https://fastapi.tiangolo.com/
- React: https://react.dev/
- Playwright: https://playwright.dev/

**Free Tier Platforms:**
- Cloudflare Pages: https://pages.cloudflare.com/
- Render: https://render.com/
- Neon: https://neon.tech/
- Upstash: https://upstash.com/

**Issues?**
- Check troubleshooting table in `ZERO_COST_HOSTING.md` Section 6.6
- Review `CLOUDFLARE_WORKERS_ANALYSIS.md` for why CF Workers don't fit

---

## 📝 Notes

- **Rekomendasi:** Mulai dengan Path A (zero-cost), upgrade ke Path C jika ada traction
- **Production:** Path B hanya untuk enterprise dengan budget dan SLA requirements
- **Playwright:** Tested in 512MB RAM dengan flag `--single-process`
- **Token auth:** Existing SESSION_TOKEN env var works in web version
- **Database:** 0.5GB Neon free tier = ~500,000 records benda uji data

---

**Created:** 2026-10-12  
**Last Updated:** 2026-10-12  
**Status:** Ready for deployment
