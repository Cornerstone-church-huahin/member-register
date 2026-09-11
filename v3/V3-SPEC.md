# MEMBER REGISTER V3 — Implementation Spec (captured brief)

Source of truth: Google Sheet "ระบบทะเบียนสมาชิก V3 — Central Database"
`V3_SPREADSHEET_ID = 17j5hFBtoUk4gxz6ykYLKA7ve2fw9VUE8h_PpzPXdFYM` (Script Properties only —
never in frontend JS). Read in code via
(Script Properties; SCHEMA_VERSION=3.0, BACKEND_BUILD=v3-dev-01). Old AppSheet sheet = backup only.

## Hard constraints
- No visual redesign. Preserve layout, cards, groups, filters, classes, print CSS.
- No new Sheet tabs. Existing: Members, Attendance, Giving, EngagementLog, StatusTypes,
  MembershipMeta, Decisions, DecisionHistory, Accounts, AuditLog, Settings, ImportLog,
  UnmatchedRecords, README. Config goes in Settings rows.
- No secrets in frontend. No invented /exec URL (insert after real deploy).
- No localStorage/sessionStorage as database.
- Never guess: MemberID, gender, birthdate, intent, origin, decisions, money → UNKNOWN/review.

## Status model (3 layers)
Base/Registry → System Result → Effective Result. Never overwrite registry with calculation.
- FULL1: active full member (vote + nominate)
- FULL2: no worship ≥ `FULL2_ABSENCE_MONTHS` (12). Still registry member, no vote/nominate.
  Returns to FULL1 immediately on qualifying return. No 1-month cooldown.
- FULL1_BOARD: System=FULL2 but active board decision KEEP_FULL1_BY_BOARD. Has FULL1 rights,
  separately identifiable. Must not auto-reactivate for a NEW absence period.
- `FULL2_REVIEW_MONTHS` (36): flag "ควรนำเข้าพิจารณาเป็นสมาชิกสำรอง" only — never auto-RESERVE.
- `FOLLOWUP_WARNING_MONTHS` (6): analytics/pastoral flag only, no status/right change.
- PENDING_REVIEW: eligible by `COUNCIL_REVIEW_YEARS`(1) / `OUTSIDE_COUNCIL_REVIEW_YEARS`(2),
  never auto-FULL1.
- Total สมาชิกสมบูรณ์ทั้งหมด = FULL1 + FULL1_BOARD + FULL2 (no double count).

StatusTypes is the master: StatusCode is the stable key; all Thai labels come from DisplayName/
ShortName. No hard-coded Thai status strings.

Legacy mapping: สมบูรณ์ 1/สมาชิกสมบูรณ์ 1→FULL1, สมบูรณ์ 2/สมาชิกสมบูรณ์ 2→FULL2,
สำรอง/สมาชิกสำรอง→RESERVE, ผู้ร่วมนมัสการ→ATTENDER, else UNSPECIFIED.

## Age / ministry engine (derived, never stored per row)
Exact completed age from BirthDate, Asia/Bangkok. Settings-driven bands:
SUNDAY_SCHOOL_MIN_AGE=3 / MAX=10, YOUTH_MIN=11 / MAX=24, ADULT_MIN=25 / MAX=59, SENIOR_MIN=60.
0–2 เด็กเล็ก · 3–10 เด็กรวี · 11–24 อนุชน · 25–59 ผู้ใหญ่ (M บุรุษ / F สตรี / unknown
"ผู้ใหญ่ · ยังไม่ระบุกลุ่ม") · 60+ ผู้สูงวัย. Missing DOB → ยังไม่ระบุกลุ่มอายุ.
Validate no overlap/gap. Future: CalculatedGroup vs EffectiveGroup override (design only).

## Decisions
Decisions = current active decision; DecisionHistory = APPEND ONLY.
Actions: CREATE_DECISION, CHANGE_STATUS, KEEP_FULL1_BY_BOARD, APPROVE_FULL1, KEEP_PENDING,
MOVE_TO_RESERVE, TRANSFER, RESIGN, REMOVE_BY_BOARD, MARK_DECEASED, REVIEW, CLEAR_DECISION, OTHER.
Reasons: GIVING_CONTINUES, ACTIVITY_ACTIVE, SERVING_ACTIVE, CONTACT_ACTIVE, HEALTH_REASON,
WORK_REASON, LIVING_AWAY, AGE_NOT_READY, NO_MEMBERSHIP_INTENT, QUALIFICATION_INCOMPLETE,
NEED_MORE_FOLLOWUP, OTHER.
Record: SystemResultAtDecision, PreviousStatusCode, NewStatusCode, DecisionDate, NextReviewDate,
RecordedBy, RecordedAt.

## Roles
Admin (all) · Editor (permitted member fields) · ธรรมกิจ (decision-relevant) ·
ฝ่ายทะเบียน (READ ONLY, 3 groups: สมบูรณ์ 1 / สมบูรณ์ 2 / ผู้ร่วมนมัสการ; no notes, reasons,
amounts, admin, audit). Backend enforces; server-side field filtering mandatory — not CSS hiding.
Giving amounts only to authorized financial roles; never in other payloads.

## Backend module shape
getV3Spreadsheet_ · getV3Sheet_ · validateSchema_ (fail closed) · loadSettings_/getSetting_ ·
loadStatusTypes_ · normalizeLegacyStatus_ · buildMemberIndex_ · calculateAge_/calculateAgeGroup_ ·
buildAttendanceIndex_/getLastWorshipDate_ · calculateSystemStatus_ · getActiveDecision_ ·
calculateEffectiveStatus_ · authorizeRole_ · sanitizeMemberForRole_ · appendAuditLog_ ·
appendDecisionHistory_. LockService on mutations. AuditLog on important writes.

## Migrations
Dry-run first, log to ImportLog (rows read/inserted/updated/skipped/unmatched/errors).
Attendance name→MemberID only on unique unambiguous match; everything else → UnmatchedRecords.
Status-engine first pass is READ-ONLY preview (first ~20 members) — no mass write to Members.

## Frontend
Two-way: web edit → Apps Script → Sheet, survives refresh & other devices; sheet edit → refresh
→ web. "รีเฟรชข้อมูลล่าสุด" button; optional 30–60s auto-refresh that never wipes an open form
or modal. Save states: กำลังบันทึก… / บันทึกแล้ว / เกิดข้อผิดพลาด (never optimistic-fake).
Destructive actions require explicit confirmation naming the effect.

## Responsive
360×800, 390×844, 412×915, 844×390, 915×412 + tablet/desktop. No page-level horizontal scroll
(tables scroll inside their container). Touch-friendly, no overlap, modals inside viewport.
Must not alter desktop layout or @media print.

## Print
A4 landscape, print-only header, repeated thead, no row splitting, ~15–17 rows/page, no blank
cover page, controls hidden.

## Test matrix
Age boundaries 2/3/10/11/24/25/59/60, missing DOB, adult M/F/unknown, YOUTH_MIN 11→12 reclass.
Membership TEST 1–10 per brief. Web↔Sheet round-trips. Role security incl. denied mutations and
payload field absence.

## Deploy order
backend code → review → paste into V3 Apps Script → test → deploy Web App → obtain /exec →
wire frontend → live end-to-end.
