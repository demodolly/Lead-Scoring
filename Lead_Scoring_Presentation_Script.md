# Eloqua Lead Scoring Model — Presentation Script

**Presenter:** Amanda Chenery  
**Audience:** Sales, VDC, and marketing operations stakeholders  
**Suggested timing:** ~30 minutes (sales overview) | ~55 minutes (full technical flow)

---

## Slide 1 — Title / Welcome

Hello, and welcome to this overview of the Eloqua lead scoring model.

My name is Amanda Chenery, and today I'll walk you through three things: how our lead scoring infrastructure is configured, how customer interactions accumulate points over time, and how high-intent leads are validated and routed to the VDC.

By the end of this session, you'll understand not just *that* we score leads, but *why* a lead does or does not reach sales—and what happens at each step in between.

---

## Slide 2 — Objectives

Before we dive into the mechanics, let's set clear objectives.

**Purpose:** Explain how the lead scoring model is structured, and show the exact path a customer takes as their interactions turn into a high-scored lead passed to the VDC.

**By the end, you will understand:**
- Our **scoring matrix** — how profile fit and engagement intent combine
- Our **three-stage data processing pipeline** inside Eloqua
- The **qualification thresholds** that determine whether a lead is routed to sales or held for nurture

We'll cover the scoring logic, walk through the technical processing flow, and bring it to life with a real-world example: **Tony**, an IT manager, and his activity over a 30-day period.

---

## Slide 3 — Dual Matrix: Profile + Engagement

At the core of our framework is a **dual-matrix approach** that combines two types of data.

### Profile score (explicit data) — "Who is this person?"
Profile logic captures information the customer **intentionally gives us**:
- Job title and role level
- Department
- Company
- Email type (corporate vs. personal)

This is our **fit assessment**: how closely does this lead match our ideal customer profile (ICP)?

Profile is graded **A through D**, in roughly 25-point bands:
| Grade | Points | Meaning |
|-------|--------|---------|
| **A** | 76–100 | Excellent ICP fit |
| **B** | 50–75 | Good fit |
| **C** | 25–49 | Marginal fit |
| **D** | 0–24 | Poor fit |

### Engagement score (implicit data) — "What are they doing?"
Engagement captures **behavioral activity** we observe passively:
- Email click-throughs
- Landing page and website visits
- Webinar and video consumption
- Event registration and attendance
- Inbound calls and chats
- Form submissions (demo requests, contact sales, etc.)

This is our **interest assessment**: how much intent and active interest is this person showing right now?

Engagement is graded **1 through 4**:
| Grade | Points | Meaning |
|-------|--------|---------|
| **4** | 76–100 | Hot — immediate follow-up |
| **3** | 51–75 | Warm-high — prioritise |
| **2** | 26–50 | Warm — building intent |
| **1** | 0–25 | Cold — low activity |

**Key takeaway:** Both scores work together. A strong profile with low engagement is a nurture candidate. High engagement with a weak profile may still warrant review—but our routing rules focus on combinations that historically convert.

---

## Slide 4 — Profile Scoring Detail

Profile scoring uses **four categories**, each weighted at **25%** of the total profile score:

1. **Decision maker type** — derived in CDF (Customer Data Foundation) from job title, level, and department  
   - Example: **Technical Decision Maker (TDM)** scores 100% of category weight → 25 points  
   - **Business Decision Maker (BDM)** scores 50% → 12.5 points

2. **Job level** — Manager, Director, VP, etc.

3. **Job department** — e.g., MIS/IT, Network Management, Operations score higher than most other departments

4. **Email domain type** — corporate email scores significantly higher than personal/public email  
   - Corporate email: up to 25 points  
   - Public/personal email: as low as 1.25 points

**How the math works:**  
Each field value within a category receives a score (up to 100%). That score is multiplied by the category weight. All category totals are summed to produce the profile point total, which maps to grade A–D.

*Do not read every row in the table—highlight the principle: fit is measurable, weighted, and consistent.*

---

## Slide 5 — Engagement Scoring Detail

Engagement scoring works similarly but adds two critical dimensions that profile scoring does not use:

### Category weights
Activities are grouped into categories (demo scheduled, event attendance, form submits, email clicks, video views, inbound calls, etc.). Each category has a defined **maximum weight**. The highest-weight activities include **Demo Scheduled** and **Event Attendance** (each up to 25% of engagement score).

### Recency windows
Activity is scored based on **when** it happened:
- **Last 7 days** — highest points  
- **Last 14 days** — moderate points  
- **Last 30 days** — lowest points  

The more recent the activity, the higher the score.

### Frequency
Within each recency window, we also ask: did this happen **at least once**, or **more than once**? Multiple interactions in the same window score higher.

**Key takeaway:** Not all activity is equal. A demo request today outweighs an email click two weeks ago—and recency decay means older activity contributes less to the current score.

*Full weight tables and field-level values are maintained on our SharePoint site.*

---

## Slide 6 — Transaction Sources (What Gets Scored)

Everything that feeds the model comes from four buckets of marketing interaction:

**1. Event integrations**
- Cvent and RainFocus registrations, attendees, and session data
- ITN (now Stova) registrations
- Booth scans

**2. Form transactions**
- Demo and callback requests
- Webex event registrations
- Contact Sales forms
- Offer pages and platform-specific forms — each individually weighted

**3. General integrations**
- Manual uploads: third-party trade shows, webinars, paid media, content syndication
- PathFactory webinars, BrightTALK webinars, Brightcove videos
- Inbound calls and inbound chats

**4. Email and digital activity (within Eloqua)**
- Email click-throughs
- Landing page visits
- Website visits  
- *Note:* Tagged landing pages receive additional points when reached via email activity

All of this data enters Eloqua and is stored in **Custom Data Objects (CDOs)** named to reflect the transaction type—before it ever reaches the scoring engine.

---

## Slide 7 — High-Score Matrix & Routing Rules

Once profile and engagement grades are calculated, they combine on a **matrix grid** to determine routing.

**High-scored combinations (routed to VDC)** — shown in white on the deck:
- Profile **A** with Engagement **1, 2, or 3**
- Profile **B** with Engagement **1 or 2**
- Profile **C** with Engagement **1 or 2**

**Low-scored combinations (not routed)** — shown in gray:
- Most Profile **D** leads — historically low conversion; not sent to sales
- Profile A/B/C with engagement level 4 (cold activity)
- Other gray-cell combinations

Low-scored records either enter a **nurture program** or receive **no further action** until a new interaction pushes them through the scoring model again.

**Why these thresholds?** After several quarters of monitoring, we observed that converting leads cluster at engagement levels **1 and 2**, and that Profile **A** customers sometimes convert with slightly lower activity—so **A3** is also included. Profile **D** leads rarely justify sales outreach.

---

## Slide 8 — End-to-End Flow (Overview)

Here is the journey at a high level—left to right:

```
Marketing activity → Eloqua (CDOs) → Lead Scoring Engine → Validations → App Cloud Lite → Trey → CDF → Salesforce → VDC
```

**Step by step:**

1. **Marketing interactions** (email, forms, events, uploads) land in Eloqua CDOs  
   - Paid leads → Transaction CDO  
   - Webinars and trade shows → Registration CDO  

2. **Lead scoring engine** evaluates the contact. If below threshold → stop (nurture or no action).

3. **Additional suppression checks** before routing:
   - **Australian public sector** companies — blocked from BDC routing by policy  
   - **30-day VDC pause** — if the customer recently declined contact, the record is stopped; if an open lead already exists in VDC, the new score **updates** that existing record with additional context for the agent

4. **High-scored leads** pass to **App Cloud Lite** (Oracle database) for validation:
   - CCID / activity ID present  
   - PIPL (privacy) validations honored  
   - Country value present  
   - Transaction date present  
   - *Any missing field stops processing*

5. Validated records go to **Trey**, which:
   - Sets the lead as **Warm Lead / Non-Hand Raiser** (originated from lead scoring, not a direct hand-raise)
   - Appends CCIDs, Offer IDs, and Drive-to IDs
   - Enriches via the **CTT** (Campaign Tagging & Tracking) tool with activity name, campaign, program, and offer type

6. **CDF** performs contact and account matching — update existing or create new; applies **suppression rules** (employees, partners excluded)

7. **Opt-in / opt-out validation** by country:
   - **Opt-out countries (e.g., USA):** Can route unless email/phone permission is explicitly "No"  
   - **Opt-in countries (e.g., Germany):** Requires explicit **Yes** on email or phone — blank is not sufficient

8. **Salesforce** lead created → **LeanData** matches/merges → assigned to VDC agent

9. Standard **lead management processes** take over from there.

*For sales audiences: spend one minute here, then jump to the Tony example. Technical audiences continue to the next slide.*

---

## Slide 9 — Technical Processing Pipeline (Eloqua Detail)

Because the lead score model cannot read CDOs directly, Eloqua runs a **three-stage preparation pipeline** before scoring executes.

### Stage 1 — Capture (6 Custom Data Objects)
Transaction data from emails, forms, events, and manual uploads is captured across six CDOs in Eloqua.

### Stage 2 — Segment & External Activities (Program 30380)
- Data is pulled from transaction CDOs into **segments** using defined criteria  
- Segments enter **Program 30380**, which creates **External Activities**  
- *Why?* The lead score model can only monitor External Activities—not raw CDO records

### Stage 3 — Timeframe & Frequency (Shared Lists)
External activities enter another program that:
- Creates **shared lists** segmented by recency (**7 / 14 / 30 days**) and frequency (**1 time vs. more than 1 time**)
- **Updates contact fields** in Eloqua with the relevant timeframe and frequency values

### Scoring execution
- A segment built from those external activities feeds the **Lead Score Model**
- The model runs and outputs profile + engagement grades
- Results enter **Program 389** (or equivalent routing program) to determine next steps

### High-score output path
If the lead qualifies (per matrix + validation checks):
1. Eloqua creates a **blind form submit** (not counted as a new marketing transaction)
2. Data written to the **Scored Lead CDO**
3. Contact and transaction details pushed to the **Oracle Scored Lead table** (App Cloud Lite)
4. CCID validation and downstream routing to Trey → VDC

**Why all these steps?** They give us granular, auditable data at each stage—so we can analyze which transactions drive high scores and optimize the model quarterly.

---

## Slide 10 — Example: Tony the IT Manager (Profile)

Let's make this concrete.

**Tony** is an IT Manager on the security team at Intuitech Technology Ltd. He has engaged with Cisco previously and is now exploring solutions for his company.

### Profile breakdown

| Category | Value | Calculation | Points |
|----------|-------|-------------|--------|
| Job level | Manager | Scored & weighted | 6.25 |
| Decision maker type | Technical | 75% of 25 weight | 18.75 |
| Department | MIS/IT | 70% of 25 weight | 17.5 |
| Email type | Corporate | 100% of 25 weight | 25.0 |
| **Total** | | | **67.5** |

**Profile grade: B** — good ICP fit, not top-tier A, but solid enough to matter when engagement is present.

*Note: If the slide title says "Director" but the narrative says Manager, clarify that CDF maps his title to the Manager job level for scoring purposes.*

---

## Slide 11 — Example: Tony's 30-Day Engagement Journey

Tony's behavior over the last 30 days tells the intent story:

**~12 days ago:** Clicked a link in an email and landed on a tagged landing page — scored for **14-day recency**, **one occurrence**. Points are moderate and begin decaying as time passes.

**~6 days ago:** Returned independently, watched **50% of a video** — scored within the **7-day window**.

**Today:**  
- Watched a **full webinar** — 7-day recency, one occurrence → ~13.5 points  
- **Scheduled a demo** — highest-weight activity, 7-day recency → score of 90 at 25% weight = **22.5 points**

Even though a demo alone isn't always enough to auto-route, **combined with recent webinar, video, and earlier email activity**, Tony's engagement total reaches approximately **52 points**.

**Engagement grade: 2** (Warm — building intent)

### Matrix result: **B2 → High score → Route to VDC**

The demo request is the trigger that matters most—but it's the **pattern of escalating activity** over 30 days that pushes him over the threshold.

**Punchline for the audience:** *"Tony clicked an email and left. He came back, watched content, attended a webinar, and scheduled a demo. That's exactly the lead we want sales to call."*

---

## Slide 12 — Validations & Edge Cases (Quick Reference)

Use this slide if asked about exceptions:

| Check | Outcome |
|-------|---------|
| Below matrix threshold | Nurture or no action |
| Australian public sector | Blocked — policy restriction |
| Open lead in VDC (last 30 days) | Still routed — enriches existing record |
| Customer declined contact (last 30 days) | Blocked |
| Missing CCID, country, or transaction date | Stopped at App Cloud Lite |
| PIPL / privacy validation fail | Stopped |
| Employee or partner (CDF suppression) | Blocked |
| Opt-in country without explicit Yes | Blocked |

---

## Slide 13 — Supporting Resources & Close

We maintain detailed reference material so you don't need to memorise every weight:

- **SharePoint site** — full scoring weights, field values, and category definitions  
- **Lead Scoring Examples document** — additional walkthroughs like Tony's  
- **Quarterly Lead Scoring Analysis** — we review and optimise the model every quarter based on conversion data

**Closing:**

Thank you for your time. The lead scoring model exists to send sales **fewer, better leads**—leads where fit and intent align. If you have questions about a specific lead or score combination, the SharePoint documentation and quarterly analysis are the best starting points.

---

## Appendix — Delivery Tips

**Pacing**
- Slides 1–7: ~15 minutes for a mixed audience  
- Slide 8–9: Optional deep dive (+15 min for technical/ops)  
- Slides 10–11: ~8 minutes — always include Tony; this is what makes the model stick  

**Phrases to emphasise**
- *"Explicit fit + implicit intent"*  
- *"Recency and frequency decay"*  
- *"White cells go to VDC; gray cells wait"*  
- *"The model re-runs on every new interaction"*  

**Common questions**
- *"Why didn't my lead route?"* → Check matrix cell, AU public sector flag, opt-in status, 30-day VDC pause, missing CCID  
- *"Why did a low-activity A-profile lead route?"* → A3 is intentionally included based on conversion data  
- *"Does an email click alone create a VDC lead?"* → Rarely; single low-weight activity with time decay typically stays below threshold  

---

*Script derived from presentation transcript (11 Aug 2026) and aligned to the Eloqua Lead Scoring Training deck structure.*
