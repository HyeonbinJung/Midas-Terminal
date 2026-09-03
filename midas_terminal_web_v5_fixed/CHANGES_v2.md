# Midas Terminal v2 — Stability hardening

Implemented for multi-user/ngrok use:

1. DOCX filename collision protection with timestamp + UUID.
2. Browser-side export lock to stop repeated clicks.
3. POST `/api/report/{ticker}` instead of a side-effecting GET.
4. Bounded report generation and same-ticker serialization.
5. TTL caches for market, research, peers, SEC and news.
6. Single-flight per cache key so simultaneous identical requests share one job.
7. Bounded Yahoo upstream concurrency.
8. Global SEC request pacing across worker threads.
9. Per-IP endpoint rate limits using forwarded ngrok client IP when present.
10. Lightweight peer fetch path that skips holders, estimates, quarterly data and analyst endpoints.
11. yfinance Industry/Sector constructor compatibility fix (removed unsupported `region=` argument).
12. US-listing peer cleanup for foreign exchange suffixes and OTC/Pink venues.
13. `PARTIAL` research status and visible warnings when SEC/peer enrichment is incomplete.
14. Research cache is reused by DOCX export instead of refetching the same upstream data.
15. DOCX is deleted after the HTTP response; stale leftovers are swept on startup.
16. `.env` is loaded before research-engine settings are initialized.
17. Favicon route added; duplicate HTML `download-report` ID removed.
18. API security headers and upstream timeout handling added.

This hardening is intended for a small internal group on one FastAPI process / PC / ngrok. For a large public service, move cache/rate-limit state to Redis and report jobs to a durable worker queue.
