/* eslint-disable */
// @ts-nocheck
import React, { useState, useEffect, useMemo, useRef } from "react";
import type { IDataService } from "../services/IDataService";
import type { Role } from "../models/IRole";
import type { ICheckResult } from "../services/ConnectionDiagnosticsService";
import { REGIONS, normalizeRegion } from "../domain/referenceData";
import { formatLocalDate, normalizeDateOnly, addCalendarDays, addCalendarMonths, calendarDaysBetween, sameCalendarDay, isDateInYearThroughToday } from "../domain/calendarDates";
import { buildIssueTransition, reopenIssuePatch } from "../domain/issueLifecycle";
import { changedIssueFields } from "../domain/issueDraft";
import { IssueConflictError, IssueRefreshError } from "../services/issueErrors";
import "./QstarPrototype.scss";
import {
  LayoutDashboard, Inbox, ClipboardList, BellRing, ListChecks, Send,
  ShieldCheck, AlertTriangle, Clock, CheckCircle2, XCircle, Lock,
  Filter, ChevronRight, ArrowLeft, Plus, Search, Building2, Globe2,
  UserCircle2, CalendarClock, FileText, Paperclip, RefreshCw,
  Layers, Settings, Link2, Plug, Eye, FlaskConical, KeyRound, UserCog, Star, ArrowUp, ArrowDown, Minus, MessageSquare
} from "lucide-react";
import {
  ResponsiveContainer, BarChart, Bar, XAxis, YAxis, Tooltip, CartesianGrid,
  PieChart, Pie, Cell, ComposedChart, Line, Legend, LabelList
} from "recharts";

/* ============================================================
   Q-STAR ISSUE MANAGER  —  ISO 9001:2015 issue handling
   Test tool. Data persists in window.storage (swap for the
   SharePoint List when replicating into your repo).
   ============================================================ */

/* ---------- Reference data (mirrors the SharePoint List) ---------- */
const SEVERITIES = ["Critical", "High", "Medium", "Low"];
const SEVERITY_DUE_DAYS = { Critical: 7, High: 14, Medium: 30, Low: 60 };
const STATUSES = ["Created", "In Progress", "On Hold", "Closed", "Rejected"];
// NC-only effectiveness test state + the period (in months) an NC must sit under test before it can be closed
const NC_TEST = "Under Testing/Revision";
const NC_TEST_MONTHS = 2;
const ALL_STATUSES = ["Created", "In Progress", NC_TEST, "On Hold", "Closed", "Rejected"];
const statusOptionsFor = (i) => i && isNC(i) ? ["Created", "In Progress", NC_TEST, "On Hold", "Closed", "Rejected"] : STATUSES;
const TRANSFORM_TYPES = ["OFI", "NC Minor", "NC Major", "Only sent to Dept/BU for Action"];
const DEVIATION_TYPES = ["Communication", "Compliance", "Documentation", "Equipment", "Process", "Quality", "Safety", "System"];
const ORIGINS = ["Customer Complaints or Claims", "Internal Finding"];
const YESNO = ["Yes", "No"];
const BUSINESS_UNITS = [
  "BU Aftermarket", "BU Airlines", "BU Automotive", "BU Diplo & High Security",
  "BU High Tech & SemiCon", "BU Life Science", "Central Europe & Commercial Services",
  "Claims & Complaints", "Customer Solution & Business Development",
  "Digital Transformation & Data Management", "Finance & Controlling", "Human Resources",
  "IT", "Legal & Data Protection", "Marketing", "Network & Products", "Quality",
  "Risk Management", "Strategy & Transformation", "tmCT FRA", "tmCT MUC", "tmCT MEX/NLU", "tmCT PVG"
];
const OWNERS = ["A. Keller", "M. Dubois", "S. Tanaka", "L. Rossi", "P. Nguyen", "C. Gems", "J. Becker", "R. Okafor"];

/* ---------- Date helpers ---------- */
const DAY = 86400000;
const today = () => { const d = new Date(); d.setHours(0, 0, 0, 0); return d; };
const iso = (d) => d instanceof Date ? formatLocalDate(d) : normalizeDateOnly(d);
const addDays = addCalendarDays;
const addMonths = addCalendarMonths;
const daysBetween = calendarDaysBetween;
const fmtDate = (d) => {
  const date = iso(d);
  return date ? new Date(`${date}T12:00:00`).toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" }) : "—";
};
const fmtDateTime = (d) => new Date(d).toLocaleString("en-GB", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });

/* ---------- Derived state helpers ---------- */
const ACTIVE_STATUSES = ["Created", "In Progress", "On Hold", NC_TEST];
const isOpen = (i) => i.triaged && i.taskCreated === "Yes" && ACTIVE_STATUSES.includes(i.status);
const isNC = (i) => i.transformedInto === "NC Minor" || i.transformedInto === "NC Major";
const inNCTest = (i) => isNC(i) && i.status === NC_TEST;
// Overdue tracks the implementation deadline; an NC under effectiveness test is NOT "overdue" (work is done, it's being observed)
const isOverdue = (i) => isOpen(i) && i.status !== NC_TEST && i.status !== "On Hold" && i.dueDate && daysBetween(today(), i.dueDate) < 0;
const overdueDays = (i) => i.dueDate ? -daysBetween(today(), i.dueDate) : 0;
// NC effectiveness test window: starts at implementation, must run NC_TEST_MONTHS before the NC may be closed
const ncTestEnd = (i) => i.implementationDate ? addMonths(i.implementationDate, NC_TEST_MONTHS) : null;
const ncTestDaysLeft = (i) => { const e = ncTestEnd(i); return e ? daysBetween(today(), e) : null; };
const ncTestComplete = (i) => { const e = ncTestEnd(i); return !!e && daysBetween(today(), e) <= 0; };
// Category grouping for the issue-mix chart
const categoryOf = (i) => {
  if (isNC(i)) return "NC";
  if (i.transformedInto === "OFI") return "OFI";
  return "Other";
};
// Derive a Microsoft 365 email from an owner's display name (demo convention)

/* ---------- Style maps (Tailwind core classes only) ---------- */
const STATUS_STYLE = {
  "Created": "bg-slate-100 text-slate-700 ring-slate-200",
  "In Progress": "bg-blue-50 text-blue-700 ring-blue-200",
  [NC_TEST]: "bg-cyan-50 text-cyan-700 ring-cyan-200",
  "On Hold": "bg-amber-50 text-amber-700 ring-amber-200",
  "Closed": "bg-emerald-50 text-emerald-700 ring-emerald-200",
  "Rejected": "bg-rose-50 text-rose-700 ring-rose-200",
};
const SEVERITY_STYLE = {
  "Critical": "bg-rose-50 text-rose-700 ring-rose-200",
  "High": "bg-orange-50 text-orange-700 ring-orange-200",
  "Medium": "bg-amber-50 text-amber-700 ring-amber-200",
  "Low": "bg-slate-100 text-slate-600 ring-slate-200",
};
const TYPE_STYLE = {
  "OFI": "bg-sky-50 text-sky-700 ring-sky-200",
  "NC Minor": "bg-violet-50 text-violet-700 ring-violet-200",
  "NC Major": "bg-purple-100 text-purple-800 ring-purple-300",
  "Only sent to Dept/BU for Action": "bg-slate-100 text-slate-600 ring-slate-300",
};
const CHART = { teal: "#1B205C", slate: "#64748b", blue: "#3b82f6", amber: "#f59e0b", emerald: "#10b981", rose: "#f43f5e", violet: "#8b5cf6", orange: "#f97316", sky: "#0ea5e9", purple: "#9333ea", yellow: "#FBB900" };
/* time:matters brand palette: Space Cadet navy + Selective Yellow */
const BRAND = { navy: "#1B205C", navyHover: "#2a3170", navySoft: "#EEF0F7", yellow: "#FBB900", yellowSoft: "#FFF7E0", yellowRing: "#F6E3A8", yellowText: "#7A5B00" };

/* ---------- Seed data ---------- */
const T = today();
function seed() {
  const mk = (o) => ({
    attachments: [], additionalComments: "", followUp: "", progressLog: [],
    rootCause: "", correctiveAction: "", implementationDate: "", effectivenessCheck: "",
    verifiedBy: "", verifiedDate: "", closedDate: "", closedAt: "", holdReason: "", holdUntil: "", ownerUpdate: false, ownerUpdateAt: "", ownerUpdateText: "", ownerBU: o.departmentBU, ...o
  });
  let qs = 1040;
  const id = (() => { let n = 220; return () => ++n; })();
  const base = [
    // --- Untriaged candidates (sit in the QM triage queue) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: false, status: null, taskCreated: "No", transformedInto: null,
      shortSummary: "Mislabelled temperature-controlled shipment to Basel",
      description: "A Life Science consignment left FRA with a label showing the wrong temperature band (2–8°C printed as 15–25°C). Caught at the gateway before dispatch.",
      immediateAction: "Shipment held, relabelled, cold chain re-verified before release.",
      severity: "High", createdBy: "T. Schmidt", reportDate: addDays(T, -1),
      departmentBU: "BU Life Science", region: "Western Europe", alreadyInContact: "Yes",
      deviationType: "Documentation", issueOrigin: "Internal Finding", taskOwner: "", dueDate: "",
    }),
    mk({
      qsNumber: ++qs, id: id(), triaged: false, status: null, taskCreated: "No", transformedInto: null,
      shortSummary: "Customer claim: damaged AOG part on Airlines route",
      description: "Customer reports an AOG aircraft part arrived with a dented protective case. Goods usable but customer requests root-cause feedback.",
      immediateAction: "Acknowledged claim, requested photos from customer.",
      severity: "Medium", createdBy: "M. Albrecht", reportDate: addDays(T, -2),
      departmentBU: "BU Airlines", region: "Germany", alreadyInContact: "No",
      deviationType: "Process", issueOrigin: "Customer Complaints or Claims", taskOwner: "", dueDate: "",
    }),
    mk({
      qsNumber: ++qs, id: id(), triaged: false, status: null, taskCreated: "No", transformedInto: null,
      shortSummary: "Suggestion: add scan checkpoint at MUC inbound dock",
      description: "Operator suggests a barcode checkpoint at the MUC inbound dock to reduce manual reconciliation effort.",
      immediateAction: "—",
      severity: "Low", createdBy: "K. Wagner", reportDate: addDays(T, -3),
      departmentBU: "tmCT MUC", region: "Germany", alreadyInContact: "No",
      deviationType: "Process", issueOrigin: "Internal Finding", taskOwner: "", dueDate: "",
    }),

    // --- OVERDUE open NC Major (drives reminders + escalation) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "In Progress", taskCreated: "Yes", transformedInto: "NC Major",
      shortSummary: "Cold-chain breach: 3h excursion on pharma shipment",
      description: "Data logger shows a 3-hour temperature excursion above 8°C on a high-value pharma shipment. Product quarantined pending stability assessment.",
      immediateAction: "Shipment quarantined, customer QA notified, logger data secured.",
      severity: "Critical", createdBy: "T. Schmidt", reportDate: addDays(T, -22),
      departmentBU: "BU Life Science", region: "Western Europe", alreadyInContact: "Yes",
      deviationType: "Quality", issueOrigin: "Internal Finding",
      taskOwner: "M. Dubois", dueDate: addDays(T, -8),
      followUp: "Escalated to BU lead. Awaiting carrier corrective measures.",
      rootCause: "Pre-conditioned packaging held in ambient staging area beyond the validated 30-minute window during a gateway delay.",
      correctiveAction: "Introduce a hard 30-min staging timer with supervisor sign-off; retrain gateway team on excursion protocol.",
      implementationDate: "", effectivenessCheck: "", verifiedBy: "", verifiedDate: "",
      progressLog: [
        { ts: new Date(T.getTime() - 18 * DAY).toISOString(), author: "M. Dubois", text: "Pulled logger data and carrier handover times. Confirmed the delay happened at the gateway, not in transit." },
        { ts: new Date(T.getTime() - 11 * DAY).toISOString(), author: "M. Dubois", text: "Drafted staging-timer procedure. Waiting on supervisor roster to schedule retraining." },
      ],
    }),

    // --- OVERDUE OFI ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "Created", taskCreated: "Yes", transformedInto: "OFI",
      shortSummary: "Inconsistent CMR archiving across MEX hub",
      description: "Audit sampling found CMR documents archived under inconsistent naming, slowing retrieval during the customer audit.",
      immediateAction: "—",
      severity: "Medium", createdBy: "Internal Audit", reportDate: addDays(T, -40),
      departmentBU: "tmCT MEX/NLU", region: "Americas", alreadyInContact: "Yes",
      deviationType: "Documentation", issueOrigin: "Internal Finding",
      taskOwner: "P. Nguyen", dueDate: addDays(T, -10),
      followUp: "Owner assigned. No update received yet — reminder cadence active.",
      progressLog: [],
    }),

    // --- Due soon (within reminder window) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "In Progress", taskCreated: "Yes", transformedInto: "NC Minor",
      shortSummary: "Gate dimensions deviate from approved drawing",
      description: "Installed gate width deviates from the approved technical drawing by 12cm, blocking forklift access on one side.",
      immediateAction: "Access rerouted; contractor notified to provide as-built evidence.",
      severity: "High", createdBy: "L. Rossi", reportDate: addDays(T, -10),
      departmentBU: "Network & Products", region: "Western Europe", alreadyInContact: "Yes",
      deviationType: "Equipment", issueOrigin: "Internal Finding",
      taskOwner: "L. Rossi", dueDate: addDays(T, 2),
      rootCause: "Contractor used a superseded drawing revision; revision control not enforced on site.",
      correctiveAction: "Re-issue current revision, require contractor sign-off on drawing rev before works, rework gate to spec.",
      implementationDate: "", effectivenessCheck: "", verifiedBy: "", verifiedDate: "",
      followUp: "Photographic/video proof requested from site contractor.",
      progressLog: [
        { ts: new Date(T.getTime() - 4 * DAY).toISOString(), author: "L. Rossi", text: "Requested as-built photos and the drawing revision used on site. Contractor confirmed wrong revision." },
      ],
    }),

    // --- On Hold ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "On Hold", taskCreated: "Yes", transformedInto: "OFI",
      shortSummary: "Manual customs paperwork double-entry at PVG",
      description: "Same customs data entered twice in two systems at PVG, raising error risk. Improvement blocked pending IT integration roadmap.",
      immediateAction: "—",
      severity: "Low", createdBy: "S. Tanaka", reportDate: addDays(T, -30),
      departmentBU: "tmCT PVG", region: "China", alreadyInContact: "No",
      deviationType: "System", issueOrigin: "Internal Finding",
      taskOwner: "S. Tanaka", dueDate: addDays(T, 25),
      followUp: "On hold until IT integration window opens in next quarter.",
      progressLog: [
        { ts: new Date(T.getTime() - 12 * DAY).toISOString(), author: "S. Tanaka", text: "Confirmed dependency on the customs API project. Parking until IT roadmap slot is confirmed." },
      ],
    }),

    // --- Healthy in-progress (not overdue) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "In Progress", taskCreated: "Yes", transformedInto: "OFI",
      shortSummary: "Driver briefing checklist missing signature field",
      description: "High-security transport briefing checklist lacks a driver signature field, weakening evidence of acknowledgement.",
      immediateAction: "Interim manual sign-off added.",
      severity: "Medium", createdBy: "J. Becker", reportDate: addDays(T, -8),
      departmentBU: "BU Diplo & High Security", region: "Germany", alreadyInContact: "Yes",
      deviationType: "Documentation", issueOrigin: "Internal Finding",
      taskOwner: "J. Becker", dueDate: addDays(T, 20),
      progressLog: [
        { ts: new Date(T.getTime() - 3 * DAY).toISOString(), author: "J. Becker", text: "Updated the checklist template with a signature field. In review with the security lead." },
      ],
    }),

    // --- Closed, verified NC (good for effectiveness rate + time-to-close) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "Closed", taskCreated: "Yes", transformedInto: "NC Minor",
      shortSummary: "Forklift inspection log gap in Aftermarket warehouse",
      description: "Two weekly forklift inspection logs were missing for the Aftermarket warehouse.",
      immediateAction: "Forklifts re-inspected before further use.",
      severity: "Medium", createdBy: "Internal Audit", reportDate: addDays(T, -55),
      departmentBU: "BU Aftermarket", region: "Germany", alreadyInContact: "Yes",
      deviationType: "Safety", issueOrigin: "Internal Finding",
      taskOwner: "A. Keller", dueDate: addDays(T, -25),
      rootCause: "Inspection reminders relied on one person's calendar; no backup when on leave.",
      correctiveAction: "Moved inspection schedule to shared team planner with automated weekly reminder.",
      implementationDate: addDays(T, -30), effectivenessCheck: "8 consecutive weeks of complete logs verified in follow-up sampling.",
      verifiedBy: "A. Keller", verifiedDate: addDays(T, -5), closedDate: addDays(T, -5), closedAt: `${addDays(T, -5)}T16:05:00`,
      followUp: "Closed after effectiveness check passed.",
      progressLog: [
        { ts: new Date(T.getTime() - 40 * DAY).toISOString(), author: "A. Keller", text: "Set up the shared planner and automated reminder. Logs now auto-prompt." },
        { ts: new Date(T.getTime() - 8 * DAY).toISOString(), author: "A. Keller", text: "Sampled the last 8 weeks — all logs complete. Ready for QM verification." },
      ],
    }),

    // --- Closed OFI ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "Closed", taskCreated: "Yes", transformedInto: "OFI",
      shortSummary: "Add multilingual labels at FRA returns desk",
      description: "Returns desk signage was German-only; international staff requested multilingual labels.",
      immediateAction: "—",
      severity: "Low", createdBy: "R. Okafor", reportDate: addDays(T, -48),
      departmentBU: "tmCT FRA", region: "Germany", alreadyInContact: "No",
      deviationType: "Communication", issueOrigin: "Internal Finding",
      taskOwner: "R. Okafor", dueDate: addDays(T, -12),
      implementationDate: addDays(T, -14), closedDate: addDays(T, -10), closedAt: `${addDays(T, -10)}T11:20:00`,
      followUp: "Signage replaced with EN/DE/FR labels.",
      progressLog: [{ ts: new Date(T.getTime() - 14 * DAY).toISOString(), author: "R. Okafor", text: "New multilingual signage installed and photographed." }],
    }),

    // --- Rejected (declined at triage) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "Rejected", taskCreated: "No", transformedInto: null,
      shortSummary: "Report: coffee machine broken in HR office",
      description: "Reported via Q-Star but not a quality/process issue — facilities matter.",
      immediateAction: "—",
      severity: "Low", createdBy: "Anonymous", reportDate: addDays(T, -6),
      departmentBU: "Human Resources", region: "Head Office", alreadyInContact: "No",
      deviationType: "Process", issueOrigin: "Internal Finding",
      taskOwner: "", dueDate: "",
      followUp: "Out of scope for QM. Redirected to Facilities. Declined.",
    }),

    // --- NC Minor in the 2-month effectiveness TEST window (implemented ~5 weeks ago) ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: NC_TEST, taskCreated: "Yes", transformedInto: "NC Minor",
      shortSummary: "Calibration certificate expired on data logger fleet",
      description: "Three temperature data loggers in use past their calibration due date.",
      immediateAction: "Affected loggers pulled from service; spares deployed.",
      severity: "High", createdBy: "C. Gems", reportDate: addDays(T, -70),
      departmentBU: "Quality", region: "Western Europe", alreadyInContact: "Yes",
      deviationType: "Equipment", issueOrigin: "Internal Finding",
      taskOwner: "C. Gems", dueDate: addDays(T, -40),
      rootCause: "Calibration tracker not linked to procurement; renewals slipped.",
      correctiveAction: "Add calibration due dates to asset register with 30-day advance alert; assign procurement owner.",
      implementationDate: addDays(T, -35), effectivenessCheck: "", verifiedBy: "", verifiedDate: "",
      followUp: "Mitigation implemented. Now under 2-month effectiveness test before closure.",
      progressLog: [
        { ts: new Date(T.getTime() - 40 * DAY).toISOString(), author: "C. Gems", text: "Asset register updated with calibration due dates and the 30-day alert. Procurement owner assigned." },
        { ts: new Date(T.getTime() - 35 * DAY).toISOString(), author: "C. Gems", text: "Mitigation implemented — starting the effectiveness test period." },
      ],
    }),

    // --- NC Major, test window ENDED, awaiting QM verification & closure ---
    mk({
      qsNumber: ++qs, id: id(), triaged: true, status: NC_TEST, taskCreated: "Yes", transformedInto: "NC Major",
      shortSummary: "Repeated mis-routing of dangerous-goods paperwork at MUC",
      description: "DG declarations were attached to the wrong consignments twice in one week, risking non-compliant carriage.",
      immediateAction: "Affected consignments held and re-documented; DG desk double-check enforced.",
      severity: "Critical", createdBy: "Internal Audit", reportDate: addDays(T, -120),
      departmentBU: "tmCT MUC", region: "Germany", alreadyInContact: "Yes",
      deviationType: "Compliance", issueOrigin: "Internal Finding",
      taskOwner: "J. Becker", dueDate: addDays(T, -80),
      rootCause: "DG paperwork matched to consignments by hand at a shared desk with no barcode confirmation.",
      correctiveAction: "Introduce scan-to-match confirmation for every DG declaration; second-person sign-off at handover.",
      implementationDate: addDays(T, -65), effectivenessCheck: "", verifiedBy: "", verifiedDate: "",
      followUp: "Scan-to-match live. Test window has elapsed — ready for effectiveness verification.",
      progressLog: [
        { ts: new Date(T.getTime() - 65 * DAY).toISOString(), author: "J. Becker", text: "Scan-to-match deployed at the DG desk and staff briefed. Effectiveness test started." },
        { ts: new Date(T.getTime() - 10 * DAY).toISOString(), author: "J. Becker", text: "Two months of handovers sampled — zero mismatches recorded. Recommend closing." },
      ],
      ownerUpdate: true, ownerUpdateAt: new Date(T.getTime() - 10 * DAY).toISOString(),
    }),
  ];

  // --- Historical closed issues (deterministic) so the year-over-year / cumulative views have depth ---
  const HIST_TYPES = ["OFI", "OFI", "NC Minor", "NC Major", "Only sent to Dept/BU for Action"];
  let _s = 20260619;
  const rnd = () => { _s = (_s * 1103515245 + 12345) & 0x7fffffff; return _s / 0x7fffffff; };
  const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
  for (let k = 0; k < 28; k++) {
    const monthsAgo = 2 + Math.floor(rnd() * 30);        // spread across the past ~2.5 years
    const rd = addMonths(iso(T), -monthsAgo);
    const tt = pick(HIST_TYPES);
    base.push(mk({
      qsNumber: ++qs, id: id(), triaged: true, status: "Closed", taskCreated: "Yes", transformedInto: tt,
      shortSummary: `Archived ${tt} — ${pick(DEVIATION_TYPES).toLowerCase()} finding`,
      description: "Closed issue retained for trend history.", immediateAction: "Resolved and verified.",
      severity: pick(SEVERITIES), createdBy: "Archive", reportDate: rd,
      departmentBU: pick(BUSINESS_UNITS), region: pick(REGIONS), alreadyInContact: "Yes",
      deviationType: pick(DEVIATION_TYPES), issueOrigin: pick(ORIGINS),
      taskOwner: pick(OWNERS), dueDate: addDays(rd, 30),
      implementationDate: addDays(rd, 18), verifiedBy: pick(OWNERS), verifiedDate: addDays(rd, 38),
      closedDate: addDays(rd, 44), closedAt: `${addDays(rd, 44)}T${String(8 + (k % 9)).padStart(2, "0")}:${String((k * 7) % 60).padStart(2, "0")}:00`,
    }));
  }
  return base.map((issue) => ({ ...issue, region: normalizeRegion(issue.region) }));
}

/* ---------- Reminder engine (simulated) ---------- */
function buildReminders(issues) {
  const out = [];
  const now = today();
  issues.filter(isOpen).forEach((i) => {
    const push = (recipient, role, type, when, message, urgency) =>
      out.push({ issueId: i.id, qsNumber: i.qsNumber, summary: i.shortSummary, recipient, role, type, when, message, urgency, sent: daysBetween(now, when) <= 0 });

    // --- Task owner posted a comment or changed status: notify the Quality Team (email) ---
    if (i.ownerUpdate) {
      const txt = i.ownerUpdateText || ((i.progressLog || [])[(i.progressLog || []).length - 1] || {}).text || "";
      const snip = txt ? `: "${txt.slice(0, 90)}${txt.length > 90 ? "…" : ""}"` : ".";
      push("Quality Team", "Quality Manager", "New update from task owner", iso(today()),
        `${i.taskOwner || "The task owner"} updated "${i.shortSummary}" (QS-${i.qsNumber})${snip}`, "info");
    }

    // --- NC under effectiveness test: remind owner + QM near / at the end of the test window ---
    if (inNCTest(i)) {
      const end = ncTestEnd(i);
      if (!end) return;
      push(i.taskOwner, "Task Owner", "Effectiveness test ending soon", addDays(end, -7),
        `"${i.shortSummary}" (QS-${i.qsNumber}) finishes its ${NC_TEST_MONTHS}-month effectiveness test on ${fmtDate(end)}. Gather evidence the fix held.`, "info");
      push(i.taskOwner, "Task Owner", "Test complete — confirm effectiveness", end,
        `Effectiveness test complete for "${i.shortSummary}" (QS-${i.qsNumber}). Post your final observation so the QM can verify and close.`, "warn");
      push("Quality Team", "Quality Manager", "Verify effectiveness & close", end,
        `NC "${i.shortSummary}" (QS-${i.qsNumber}) has finished its ${NC_TEST_MONTHS}-month test. Verify effectiveness and close if the corrective action held.`, "warn");
      return;
    }

    // --- On hold: remind a week before the resume date and on the day itself ---
    if (i.status === "On Hold" && i.holdUntil) {
      const until = i.holdUntil;
      push(i.taskOwner, "Task Owner", "Hold ending soon", addDays(until, -7),
        `"${i.shortSummary}" (QS-${i.qsNumber}) comes off hold on ${fmtDate(until)} — prepare to resume the mitigation.`, "info");
      push(i.taskOwner, "Task Owner", "Resume work today", until,
        `Hold period over for "${i.shortSummary}" (QS-${i.qsNumber}). Resume work on the mitigation.`, "warn");
      push("Quality Team", "Quality Manager", "Hold ended — follow up", until,
        `"${i.shortSummary}" (QS-${i.qsNumber}) was on hold until ${fmtDate(until)}. Check it's back in progress.`, "warn");
      return;
    }

    const due = i.dueDate;
    if (!due) return;
    const dToDue = daysBetween(now, due); // negative => overdue

    // Owner: -3 days heads-up
    push(i.taskOwner, "Task Owner", "Heads-up (due in 3 days)", addDays(due, -3),
      `Reminder: "${i.shortSummary}" (QS-${i.qsNumber}) is due ${fmtDate(due)}. Please post a progress update.`,
      "info");
    // Owner: due date
    push(i.taskOwner, "Task Owner", "Due today", due,
      `"${i.shortSummary}" (QS-${i.qsNumber}) is due today. Update the status or request more time.`,
      "warn");
    // Owner: overdue nudges every 3 days
    if (dToDue < 0) {
      for (let k = 3; k <= Math.min(-dToDue, 30); k += 3) {
        push(i.taskOwner, "Task Owner", `Overdue nudge (+${k}d)`, addDays(due, k),
          `OVERDUE ${k} day(s): "${i.shortSummary}" (QS-${i.qsNumber}). Action required.`,
          "danger");
      }
      // QM overdue alert (at due+1)
      push("Quality Team", "Quality Manager", "Overdue alert", addDays(due, 1),
        `Task overdue: "${i.shortSummary}" (QS-${i.qsNumber}) — owner ${i.taskOwner}. Follow up.`,
        "danger");
      // BU escalation at +7
      if (-dToDue >= 7) {
        push(`${i.ownerBU} (BU lead)`, "BU Escalation", "Escalation (overdue 7d+)", addDays(due, 7),
          `Escalation: "${i.shortSummary}" (QS-${i.qsNumber}) overdue 7+ days. Owner ${i.taskOwner} unresponsive.`,
          "danger");
      }
    }
  });
  // sort: sent first (most recent), then upcoming
  return out.sort((a, b) => new Date(b.when) - new Date(a.when));
}

/* ============================================================
   Small UI atoms
   ============================================================ */
function Pill({ children, className = "" }) {
  return <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-medium ring-1 ring-inset ${className}`}>{children}</span>;
}
function StatusPill({ s }) { return s ? <Pill className={STATUS_STYLE[s]}>{s}</Pill> : <Pill className="bg-slate-100 text-slate-500 ring-slate-200">Awaiting triage</Pill>; }
function SeverityPill({ s }) { return <Pill className={SEVERITY_STYLE[s]}>{s}</Pill>; }
function TypePill({ t }) { return t ? <Pill className={TYPE_STYLE[t]}>{t}</Pill> : null; }

function Field({ label, children, hint }) {
  return (
    <label className="block">
      <span className="mb-1 flex items-center gap-1 text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
      {children}
      {hint && <span className="mt-1 block text-xs text-slate-400">{hint}</span>}
    </label>
  );
}
const inputCls = "w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none tm-input";
function TextInput(p) { return <input {...p} className={inputCls} />; }
function TextArea(p) { return <textarea {...p} className={`${inputCls} min-h-20 resize-y`} />; }
function Select({ value, onChange, options, placeholder, disabled }) {
  return (
    <select value={value || ""} onChange={onChange} disabled={disabled} className={inputCls}>
      <option value="">{placeholder || "Select…"}</option>
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
}
function Btn({ children, onClick, variant = "primary", disabled, className = "", type, style }) {
  const styles = {
    primary: "tm-btn-primary text-white",
    ghost: "bg-white text-slate-700 ring-1 ring-inset ring-slate-300 hover:bg-slate-50",
    danger: "bg-rose-600 text-white hover:bg-rose-700",
    subtle: "bg-slate-100 text-slate-700 hover:bg-slate-200",
  };
  return <button onClick={onClick} disabled={disabled} style={style} className={`inline-flex items-center justify-center gap-1.5 rounded-lg px-3.5 py-2 text-sm font-semibold transition ${styles[variant]} ${className}`}>{children}</button>;
}
function Card({ children, className = "", style }) { return <div style={style} className={`rounded-xl border border-slate-200 bg-white ${className}`}>{children}</div>; }
function SectionTitle({ icon: Icon, children, right }) {
  return (
    <div className="mb-3 flex items-center justify-between">
      <h3 className="flex items-center gap-2 text-sm font-bold text-slate-700">{Icon && <Icon size={16} style={{ color: BRAND.navy }} />}{children}</h3>
      {right}
    </div>
  );
}

/* ============================================================
   Dashboard
   ============================================================ */
function StatCard({ label, value, sub, tone = "slate", icon: Icon }) {
  const tones = { slate: "text-slate-900", teal: "tm-text-navy", rose: "text-rose-700", amber: "text-amber-700", emerald: "text-emerald-700", blue: "text-blue-700" };
  return (
    <Card className="p-4">
      <div className="flex items-start justify-between">
        <span className="text-xs font-semibold uppercase tracking-wide text-slate-500">{label}</span>
        {Icon && <Icon size={16} className="text-slate-400" />}
      </div>
      <div className={`mt-2 text-3xl font-bold ${tones[tone]}`}>{value}</div>
      {sub && <div className="mt-0.5 text-xs text-slate-500">{sub}</div>}
    </Card>
  );
}

function Dot({ color }) { return <span className="inline-block h-2.5 w-2.5 rounded-sm" style={{ backgroundColor: color }} />; }
function FlowTooltip({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;
  const d = payload[0].payload;
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-2.5 text-xs shadow-md">
      <div className="mb-1.5 font-semibold text-slate-700">{label}</div>
      <div className="space-y-1">
        <div className="flex items-center justify-between gap-4"><span className="flex items-center gap-1.5"><Dot color={CHART.teal} />Opened</span><strong className="text-slate-800">{d.Created}</strong></div>
        <div className="flex items-center justify-between gap-4"><span className="flex items-center gap-1.5"><Dot color={CHART.emerald} />Closed</span><strong className="text-slate-800">{d.Closed}</strong></div>
        <div className="flex items-center justify-between gap-4 border-t border-slate-100 pt-1"><span className="flex items-center gap-1.5"><Dot color={BRAND.yellow} />Open backlog</span><strong className="text-slate-800">{d.Backlog}</strong></div>
        <div className="flex items-center justify-between gap-4"><span className="text-slate-500">Net this month</span><strong className={d.Net > 0 ? "text-rose-600" : d.Net < 0 ? "text-emerald-600" : "text-slate-500"}>{d.Net > 0 ? "+" : ""}{d.Net}</strong></div>
      </div>
    </div>
  );
}

function Segmented({ options, value, onChange }) {
  return (
    <div className="inline-flex rounded-lg bg-slate-100 p-0.5 text-xs font-semibold">
      {options.map((o) => (
        <button key={o} onClick={() => onChange(o)}
          style={value === o ? { color: BRAND.navy } : undefined}
          className={`rounded-md px-2.5 py-1 transition ${value === o ? "bg-white shadow-sm" : "text-slate-500 hover:text-slate-700"}`}>
          {o}
        </button>
      ))}
    </div>
  );
}

function ChartNote({ children }) {
  return <p className="mt-3 border-t border-slate-100 pt-2 text-xs text-slate-400">{children}</p>;
}

export function Dashboard({ issues: allIssues }) {
  const [cumCat, setCumCat] = useState("All");
  const [cumStatus, setCumStatus] = useState("All");
  const [cumGran, setCumGran] = useState("Month");
  const [fDept, setFDept] = useState("");
  const [fRegion, setFRegion] = useState("");
  // Filter options from the full dataset (so options never disappear as you narrow)
  const deptOptions = [...new Set(allIssues.map((i) => i.departmentBU).filter(Boolean))].sort();
  const regionOptions = [...new Set(allIssues.map((i) => i.region).filter(Boolean))].sort();
  // Global dashboard filter — applies to every KPI and chart below
  const issues = allIssues.filter((i) => (!fDept || i.departmentBU === fDept) && (!fRegion || i.region === fRegion));
  const filterActive = !!(fDept || fRegion);
  const created = issues.filter((i) => i.triaged && i.status !== "Rejected");
  // Current-year scope for the dashboard graphs (the Cumulative chart deliberately spans all years)
  const nowD = today();
  const asOf = iso(nowD);
  const datedCreated = created.filter((i) => iso(i.reportDate) && iso(i.reportDate) <= asOf);
  const curYear = nowD.getFullYear();
  const yearStart = new Date(curYear, 0, 1);
  const inYear = (d) => isDateInYearThroughToday(d, nowD);
  const createdY = created.filter((i) => inYear(i.reportDate));
  const ytdRange = `Data calculated from ${fmtDate(yearStart)} to ${fmtDate(nowD)}.`;
  const cumFrom = datedCreated.length ? datedCreated.map((i) => iso(i.reportDate)).sort()[0] : nowD;
  const cumRange = `Data calculated from ${fmtDate(cumFrom)} to ${fmtDate(nowD)}.`;
  const open = issues.filter(isOpen);
  const closed = issues.filter((i) => i.status === "Closed");
  const overdue = issues.filter(isOverdue);
  const aging = { "0–7": 0, "8–30": 0, "30+": 0 };
  overdue.forEach((i) => { const d = overdueDays(i); if (d <= 7) aging["0–7"]++; else if (d <= 30) aging["8–30"]++; else aging["30+"]++; });

  const countBy = (key, universe = createdY) => {
    const m = {}; universe.forEach((i) => { const k = i[key] || "—"; m[k] = (m[k] || 0) + 1; });
    return Object.entries(m).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  };
  const byStatus = ALL_STATUSES.map((s) => ({ name: s, value: createdY.filter((i) => i.status === s).length }));
  const bySeverity = SEVERITIES.map((s) => ({ name: s, value: createdY.filter((i) => i.severity === s).length }));
  const byDeviation = countBy("deviationType");
  const byBU = countBy("departmentBU").slice(0, 7);
  const byRegion = countBy("region");
  const bySource = ORIGINS.map((o) => ({ name: o.includes("Customer") ? "Customer" : "Internal", value: createdY.filter((i) => i.issueOrigin === o).length }));
  // Issue category mix: NC / OFI / Other (current year)
  const CATS = ["NC", "OFI", "Other"];
  const byCategory = CATS.map((c) => ({ name: c, value: createdY.filter((i) => categoryOf(i) === c).length }));
  const catTotal = byCategory.reduce((s, d) => s + d.value, 0);
  const catColor = { NC: CHART.purple, OFI: CHART.sky, Other: CHART.slate };

  // Avg time to close
  const closedWithDates = closed.filter((i) => i.closedDate && i.reportDate);
  const avgTTC = closedWithDates.length ? Math.round(closedWithDates.reduce((s, i) => s + daysBetween(i.reportDate, i.closedDate), 0) / closedWithDates.length) : 0;

  // NC effectiveness verification rate
  const ncs = created.filter(isNC);
  const ncVerified = ncs.filter((i) => i.verifiedBy);
  const ncRate = ncs.length ? Math.round((ncVerified.length / ncs.length) * 100) : 0;

  // Monthly trend — current year (Jan → current month)
  const backlogAt = (end) => datedCreated.filter((i) => iso(i.reportDate) <= end && (i.status !== "Closed" || (iso(i.closedDate) && iso(i.closedDate) > end))).length;
  const months = [];
  for (let mo = 0; mo <= nowD.getMonth(); mo++) {
    const d = new Date(curYear, mo, 1);
    months.push({ key: d.toLocaleDateString("en-GB", { month: "short", year: "2-digit" }), month: iso(d).slice(0, 7), end: mo === nowD.getMonth() ? asOf : iso(new Date(curYear, mo + 1, 0)) });
  }
  const trend = months.map(({ key, month, end }) => {
    const c = datedCreated.filter((i) => iso(i.reportDate).slice(0, 7) === month).length;
    const cl = closed.filter((i) => iso(i.closedDate).slice(0, 7) === month && iso(i.closedDate) <= end).length;
    // Open backlog at month-end: created on/before end, and not yet closed by end (all-time, so the backlog is accurate)
    const backlog = backlogAt(end);
    return { name: key, Created: c, Closed: cl, Backlog: backlog, Net: c - cl };
  });
  const backlogChange = backlogAt(asOf) - backlogAt(addDays(yearStart, -1));

  const sevColors = { Critical: CHART.rose, High: CHART.orange, Medium: CHART.amber, Low: CHART.slate };
  const sourceColors = [CHART.violet, CHART.teal];

  // ---- Cumulative issues by category, RESET EACH YEAR (stacked) + in-year total line ----
  const cumCats = ["All", "NC", "OFI", "Other"];
  const cumStatuses = ["All", "Open", "Closed"];
  const cumGrans = ["Month", "Quarter", "Year"];
  const MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const catColorAll = { NC: CHART.purple, OFI: CHART.sky, Other: CHART.slate };
  const cumSeries = cumCat === "All" ? ["NC", "OFI", "Other"] : [cumCat];
  const byStatusSet = cumStatus === "Open" ? datedCreated.filter(isOpen) : cumStatus === "Closed" ? datedCreated.filter((i) => i.status === "Closed") : datedCreated;
  const cumFiltered = cumCat === "All" ? byStatusSet : byStatusSet.filter((i) => categoryOf(i) === cumCat);
  const gran = cumGran;
  const idxOf = (d) => {
    const [year, month] = iso(d).split("-").map(Number);
    return gran === "Year" ? year : gran === "Quarter" ? year * 4 + Math.floor((month - 1) / 3) : year * 12 + month - 1;
  };
  const yearOfIdx = (ix) => gran === "Year" ? ix : gran === "Quarter" ? Math.floor(ix / 4) : Math.floor(ix / 12);
  const labelOf = (ix) => gran === "Year" ? String(ix) : gran === "Quarter" ? `Q${(ix % 4) + 1} '${String(Math.floor(ix / 4)).slice(2)}` : `${MONTH_ABBR[ix % 12]} '${String(Math.floor(ix / 12)).slice(2)}`;
  const nowIdx = idxOf(nowD);
  const startIdx = datedCreated.length ? Math.min(...datedCreated.map((i) => idxOf(i.reportDate))) : nowIdx;
  const adds = {};
  cumFiltered.forEach((i) => { const ix = idxOf(i.reportDate); const c = categoryOf(i); adds[ix] = adds[ix] || {}; adds[ix][c] = (adds[ix][c] || 0) + 1; });
  const running = {}; cumSeries.forEach((c) => (running[c] = 0));
  let prevYear = null;
  const cumData = [];
  for (let ix = startIdx; ix <= nowIdx; ix++) {
    const yr = yearOfIdx(ix);
    if (prevYear !== null && yr !== prevYear) cumSeries.forEach((c) => (running[c] = 0)); // reset cumulative every new calendar year
    prevYear = yr;
    const row = { name: labelOf(ix) };
    cumSeries.forEach((c) => { running[c] += (adds[ix]?.[c] || 0); row[c] = running[c]; });
    row.Total = cumSeries.reduce((s, c) => s + row[c], 0);
    cumData.push(row);
  }
  // ---- Per-year totals + YoY trend (drives the arrow strip) ----
  const minYear = datedCreated.length ? Math.min(...datedCreated.map((i) => Number(iso(i.reportDate).slice(0, 4)))) : curYear;
  const maxYear = curYear;
  const yearTotalOf = (y) => cumFiltered.filter((i) => Number(iso(i.reportDate).slice(0, 4)) === y).length;
  const yearTrend = [];
  for (let y = minYear; y <= maxYear; y++) {
    const total = yearTotalOf(y);
    const prev = y > minYear ? yearTotalOf(y - 1) : null;
    const delta = prev == null ? null : total - prev;
    const dir = delta == null ? null : delta > 0 ? "up" : delta < 0 ? "down" : "flat";
    yearTrend.push({ year: y, total, delta, dir, ytd: y === maxYear });
  }

  return (
    <div className="space-y-5">
      <div className="grid grid-cols-2 gap-4 lg:grid-cols-4">
        <StatCard label="Open issues" value={open.length} sub={`${created.length} created total`} tone="teal" icon={ClipboardList} />
        <StatCard label="Overdue" value={overdue.length} sub={`${aging["0–7"]} new · ${aging["8–30"]} aging · ${aging["30+"]} critical`} tone={overdue.length ? "rose" : "slate"} icon={AlertTriangle} />
        <StatCard label="Avg time to close" value={`${avgTTC}d`} sub={`${closed.length} closed`} tone="blue" icon={Clock} />
        <StatCard label="NC effectiveness verified" value={`${ncRate}%`} sub={`${ncVerified.length}/${ncs.length} NCs verified`} tone={ncRate >= 60 ? "emerald" : "amber"} icon={ShieldCheck} />
      </div>

      {/* ---- Global dashboard filter: department and/or region ---- */}
      <Card className="p-3" style={filterActive ? { borderColor: BRAND.yellowRing } : undefined}>
        <div className="flex flex-wrap items-end gap-3">
          <div className="flex items-center gap-1.5 pb-2 text-sm font-semibold text-slate-600"><Filter size={15} style={{ color: BRAND.navy }} />Filter dashboard</div>
          <div className="min-w-0 grow basis-56">
            <label className="mb-1 flex items-center gap-1 text-xs font-medium text-slate-400"><Building2 size={12} />Department / BU</label>
            <Select value={fDept} onChange={(e) => setFDept(e.target.value)} options={deptOptions} placeholder="All departments" />
          </div>
          <div className="min-w-0 grow basis-48">
            <label className="mb-1 flex items-center gap-1 text-xs font-medium text-slate-400"><Globe2 size={12} />Region</label>
            <Select value={fRegion} onChange={(e) => setFRegion(e.target.value)} options={regionOptions} placeholder="All regions" />
          </div>
          {filterActive && (
            <button onClick={() => { setFDept(""); setFRegion(""); }} className="inline-flex items-center gap-1 rounded-lg px-2.5 py-2 text-sm font-medium text-slate-500 hover:bg-slate-100"><XCircle size={15} />Clear</button>
          )}
          <div className="ml-auto pb-2 text-xs text-slate-400">{filterActive ? `Showing ${issues.length} of ${allIssues.length} issues` : `All ${allIssues.length} issues`}</div>
        </div>
      </Card>

      {/* ---- Headline: issue flow (opened vs. closed) + open backlog ---- */}
      <Card className="p-4">
        <SectionTitle icon={RefreshCw} right={
          <span className={`inline-flex items-center gap-1 rounded-full px-2.5 py-0.5 text-xs font-semibold ring-1 ring-inset ${backlogChange > 0 ? "bg-rose-50 text-rose-700 ring-rose-200" : backlogChange < 0 ? "bg-emerald-50 text-emerald-700 ring-emerald-200" : "bg-slate-100 text-slate-600 ring-slate-200"}`}>
            Backlog {backlogChange > 0 ? `▲ +${backlogChange}` : backlogChange < 0 ? `▼ ${backlogChange}` : "flat"} year to date
          </span>
        }>Issue flow — opened vs. closed, with open backlog</SectionTitle>
        <p className="-mt-1 mb-4 text-xs leading-relaxed text-slate-500">
          Bars show how many issues were <span className="font-semibold tm-text-navy">opened</span> and <span className="font-semibold text-emerald-700">closed</span> each month. The <span className="font-semibold text-amber-600">backlog line</span> is how many remain open at month-end — when it climbs, issues are arriving faster than they're being resolved.
        </p>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={trend} margin={{ top: 18, right: 6, left: -6, bottom: 0 }} barGap={2} barCategoryGap="24%">
            <CartesianGrid strokeDasharray="3 3" stroke="#eef2f6" vertical={false} />
            <XAxis dataKey="name" tick={{ fontSize: 11, fill: "#64748b" }} axisLine={{ stroke: "#e2e8f0" }} tickLine={false} />
            <YAxis yAxisId="left" allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} axisLine={false} tickLine={false} label={{ value: "Per month", angle: -90, position: "insideLeft", offset: 18, style: { fontSize: 10, fill: "#94a3b8" } }} />
            <YAxis yAxisId="right" orientation="right" allowDecimals={false} tick={{ fontSize: 11, fill: "#b45309" }} axisLine={false} tickLine={false} label={{ value: "Open backlog", angle: 90, position: "insideRight", offset: 14, style: { fontSize: 10, fill: "#b45309" } }} />
            <Tooltip content={<FlowTooltip />} cursor={{ fill: "#f1f5f9" }} />
            <Legend wrapperStyle={{ fontSize: 12, paddingTop: 8 }} iconType="circle" />
            <Bar yAxisId="left" dataKey="Created" name="Opened" fill={CHART.teal} radius={[3, 3, 0, 0]} maxBarSize={26}>
              <LabelList dataKey="Created" position="top" formatter={(v) => v || ""} style={{ fontSize: 10, fill: BRAND.navy, fontWeight: 600 }} />
            </Bar>
            <Bar yAxisId="left" dataKey="Closed" name="Closed" fill={CHART.emerald} radius={[3, 3, 0, 0]} maxBarSize={26}>
              <LabelList dataKey="Closed" position="top" formatter={(v) => v || ""} style={{ fontSize: 10, fill: "#047857", fontWeight: 600 }} />
            </Bar>
            <Line yAxisId="right" type="monotone" dataKey="Backlog" name="Open backlog" stroke={BRAND.yellow} strokeWidth={3} dot={{ r: 3, fill: BRAND.yellow, stroke: "#fff", strokeWidth: 1 }} activeDot={{ r: 5 }}>
              <LabelList dataKey="Backlog" position="top" formatter={(v) => v || ""} style={{ fontSize: 10, fill: "#b45309", fontWeight: 700 }} />
            </Line>
          </ComposedChart>
        </ResponsiveContainer>
        <ChartNote>{ytdRange}</ChartNote>
      </Card>

      {/* ---- Cumulative issues by category: resets each year + per-year trend arrows ---- */}
      <Card className="p-4">
        <SectionTitle icon={Layers}>Cumulative issues by category — year by year</SectionTitle>
        <p className="-mt-1 mb-3 text-xs leading-relaxed text-slate-500">
          Issues by category, cumulated <strong>within each calendar year</strong> (the running total resets every January) so you can compare one year to the next. The navy line is the in-year total. {cumStatus === "All" ? "Showing all issues" : `Showing ${cumStatus.toLowerCase()} issues`}. The arrows compare each year's total with the year before — <span className="font-semibold text-rose-600">red ▲ = more issues</span>, <span className="font-semibold text-emerald-600">green ▼ = fewer</span>, navy = flat.
        </p>

        <div className="mb-3 flex flex-wrap items-center gap-x-5 gap-y-2">
          <div className="flex items-center gap-2"><span className="text-xs font-medium text-slate-400">Category</span><Segmented options={cumCats} value={cumCat} onChange={setCumCat} /></div>
          <div className="flex items-center gap-2"><span className="text-xs font-medium text-slate-400">Status</span><Segmented options={cumStatuses} value={cumStatus} onChange={setCumStatus} /></div>
          <div className="flex items-center gap-2"><span className="text-xs font-medium text-slate-400">Date</span><Segmented options={cumGrans} value={cumGran} onChange={setCumGran} /></div>
        </div>

        {/* Year trend strip */}
        <div className="mb-4 flex flex-wrap items-stretch gap-2">
          {yearTrend.map((y) => {
            const color = y.dir === "up" ? "#e11d48" : y.dir === "down" ? "#059669" : BRAND.navy;
            return (
              <div key={y.year} className="flex items-center gap-2.5 rounded-lg border border-slate-200 bg-white px-3 py-2">
                <div className="leading-tight">
                  <div className="text-xs font-semibold text-slate-500">{y.year}{y.ytd ? " · YTD" : ""}</div>
                  <div className="text-lg font-bold tm-text-navy">{y.total}</div>
                </div>
                {y.dir && (
                  <div className="flex items-center gap-0.5 text-sm font-bold" style={{ color }}>
                    {y.dir === "up" ? <ArrowUp size={16} /> : y.dir === "down" ? <ArrowDown size={16} /> : <Minus size={16} />}
                    <span className="text-xs">{y.delta > 0 ? `+${y.delta}` : y.delta}</span>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        <ResponsiveContainer width="100%" height={320}>
          <ComposedChart data={cumData} margin={{ top: 10, right: 8, left: -8, bottom: 0 }} barCategoryGap="12%">
            <CartesianGrid strokeDasharray="3 3" stroke="#eef2f6" vertical={false} />
            <XAxis dataKey="name" tick={{ fontSize: 10, fill: "#64748b" }} interval="preserveStartEnd" minTickGap={6} axisLine={{ stroke: "#e2e8f0" }} tickLine={false} />
            <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} axisLine={false} tickLine={false} label={{ value: "Cumulative (in-year)", angle: -90, position: "insideLeft", offset: 16, style: { fontSize: 10, fill: "#94a3b8" } }} />
            <Tooltip />
            <Legend wrapperStyle={{ fontSize: 12, paddingTop: 8 }} iconType="circle" />
            {cumSeries.map((c, ci) => (
              <Bar key={c} dataKey={c} name={c} stackId="cat" fill={catColorAll[c]} maxBarSize={48} radius={ci === cumSeries.length - 1 ? [3, 3, 0, 0] : [0, 0, 0, 0]} />
            ))}
            <Line type="monotone" dataKey="Total" name="In-year total" stroke={BRAND.navy} strokeWidth={2.5} dot={{ r: 2.5, fill: BRAND.navy }} activeDot={{ r: 4 }} />
          </ComposedChart>
        </ResponsiveContainer>
        <ChartNote>{cumRange}</ChartNote>
      </Card>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <Card className="p-4">
          <SectionTitle icon={Layers} right={<span className="text-xs text-slate-400">{catTotal} issues</span>}>Issue category mix</SectionTitle>
          <ResponsiveContainer width="100%" height={170}>
            <PieChart>
              <Pie data={byCategory.filter((d) => d.value > 0)} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={44} outerRadius={70} paddingAngle={2}>
                {byCategory.filter((d) => d.value > 0).map((e) => <Cell key={e.name} fill={catColor[e.name]} />)}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
          <div className="mt-2 grid grid-cols-2 gap-1.5">
            {byCategory.map((c) => (
              <div key={c.name} className="flex items-center justify-between rounded-md bg-slate-50 px-2.5 py-1.5 text-xs">
                <span className="flex items-center gap-1.5"><Dot color={catColor[c.name]} /><span className="font-medium text-slate-700">{c.name}</span></span>
                <span className="text-slate-500">{c.value} · {catTotal ? Math.round((c.value / catTotal) * 100) : 0}%</span>
              </div>
            ))}
          </div>
          <p className="mt-2 text-xs text-slate-400">NC combines NC Minor + NC Major.</p>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
        <Card className="p-4">
          <SectionTitle icon={Filter}>By status</SectionTitle>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={byStatus} margin={{ top: 5, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef2f6" />
              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} interval={0} angle={-20} textAnchor="end" height={60} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} />
              <Tooltip />
              <Bar dataKey="value" fill={CHART.teal} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
        <Card className="p-4">
          <SectionTitle icon={AlertTriangle}>By severity</SectionTitle>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={bySeverity} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={42} outerRadius={72} paddingAngle={2}>
                {bySeverity.map((e) => <Cell key={e.name} fill={sevColors[e.name]} />)}
              </Pie>
              <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
      </div>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-3">
        <Card className="p-4">
          <SectionTitle icon={Send}>Source split</SectionTitle>
          <ResponsiveContainer width="100%" height={200}>
            <PieChart>
              <Pie data={bySource} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={72} paddingAngle={2}>
                {bySource.map((e, idx) => <Cell key={e.name} fill={sourceColors[idx]} />)}
              </Pie>
              <Tooltip /><Legend wrapperStyle={{ fontSize: 11 }} />
            </PieChart>
          </ResponsiveContainer>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
        <Card className="p-4">
          <SectionTitle icon={FileText}>By deviation type</SectionTitle>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart layout="vertical" data={byDeviation} margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
              <XAxis type="number" allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} />
              <YAxis type="category" dataKey="name" width={90} tick={{ fontSize: 10, fill: "#64748b" }} />
              <Tooltip />
              <Bar dataKey="value" fill={CHART.violet} radius={[0, 4, 4, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
        <Card className="p-4">
          <SectionTitle icon={Globe2}>By region</SectionTitle>
          <ResponsiveContainer width="100%" height={200}>
            <BarChart data={byRegion} margin={{ top: 5, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#eef2f6" />
              <XAxis dataKey="name" tick={{ fontSize: 9, fill: "#64748b" }} interval={0} angle={-20} textAnchor="end" height={60} />
              <YAxis allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} />
              <Tooltip />
              <Bar dataKey="value" fill={CHART.amber} radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
          <ChartNote>{ytdRange}</ChartNote>
        </Card>
      </div>

      <Card className="p-4">
        <SectionTitle icon={Building2}>By business unit (top 7)</SectionTitle>
        <ResponsiveContainer width="100%" height={230}>
          <BarChart layout="vertical" data={byBU} margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
            <XAxis type="number" allowDecimals={false} tick={{ fontSize: 11, fill: "#64748b" }} />
            <YAxis type="category" dataKey="name" width={200} tick={{ fontSize: 11, fill: "#64748b" }} />
            <Tooltip />
            <Bar dataKey="value" fill={CHART.sky} radius={[0, 4, 4, 0]} maxBarSize={26} />
          </BarChart>
        </ResponsiveContainer>
        <ChartNote>{ytdRange}</ChartNote>
      </Card>
    </div>
  );
}

/* ============================================================
   Issue list row + Register
   ============================================================ */
function IssueRow({ i, onOpen }) {
  const od = isOverdue(i);
  return (
    <button onClick={() => onOpen(i.id)} className="flex w-full items-center gap-3 border-b border-slate-100 px-4 py-3 text-left hover:bg-slate-50">
      <div className="w-16 shrink-0 font-mono text-xs text-slate-400">QS-{i.qsNumber}</div>
      <div className="min-w-0 flex-1">
        <div className="truncate text-sm font-semibold text-slate-800">{i.shortSummary}</div>
        <div className="mt-0.5 flex flex-wrap items-center gap-1.5 text-xs text-slate-500">
          <span>{i.departmentBU}</span><span>·</span><span>{i.region}</span>
          {i.taskOwner && <><span>·</span><span className="inline-flex items-center gap-1"><UserCircle2 size={12} />{i.taskOwner}</span></>}
          {i.status === "On Hold" && i.holdUntil && <><span>·</span><span className="inline-flex items-center gap-1 font-medium text-amber-600"><Clock size={12} />On hold · resumes {fmtDate(i.holdUntil)}</span></>}
          {inNCTest(i) && <><span>·</span><span className={`inline-flex items-center gap-1 font-medium ${ncTestComplete(i) ? "text-rose-600" : "text-cyan-700"}`}><FlaskConical size={12} />{ncTestComplete(i) ? `Test period over (${fmtDate(ncTestEnd(i))}) — verify` : `Under test · due ${fmtDate(ncTestEnd(i))}`}</span></>}
          {i.ownerUpdate && <><span>·</span><span className="inline-flex items-center gap-1 font-semibold text-amber-600"><MessageSquare size={12} />New update</span></>}
        </div>
      </div>
      <div className="hidden shrink-0 items-center gap-2 sm:flex">
        <SeverityPill s={i.severity} />
        <TypePill t={i.transformedInto} />
      </div>
      <div className="w-28 shrink-0 text-right">
        {od ? <Pill className="bg-rose-50 text-rose-700 ring-rose-200"><Clock size={11} />{overdueDays(i)}d overdue</Pill> : <StatusPill s={i.status} />}
      </div>
      <ChevronRight size={16} className="shrink-0 text-slate-300" />
    </button>
  );
}

function Register({ issues, onOpen, filter, setFilter }) {
  const { q, statuses, type, bu, overdueOnly } = filter;
  const set = (patch) => setFilter((f) => ({ ...f, ...patch }));
  const toggleStatus = (s) => setFilter((f) => ({ ...f, statuses: f.statuses.includes(s) ? f.statuses.filter((x) => x !== s) : [...f.statuses, s] }));
  const anyActive = q || statuses.length || type || bu || overdueOnly;
  const clearAll = () => setFilter({ q: "", statuses: [], type: "", bu: "", overdueOnly: false });

  const list = issues.filter((i) => i.triaged).filter((i) => {
    if (overdueOnly && !isOverdue(i)) return false;
    if (statuses.length && !statuses.includes(i.status)) return false;
    if (type && i.transformedInto !== type) return false;
    if (bu && i.departmentBU !== bu) return false;
    if (q) { const hay = `${i.qsNumber} ${i.shortSummary} ${i.description} ${i.taskOwner}`.toLowerCase(); if (!hay.includes(q.toLowerCase())) return false; }
    return true;
  }).sort((a, b) => (isOverdue(b) - isOverdue(a)) || new Date(b.reportDate) - new Date(a.reportDate));

  return (
    <div className="space-y-4">
      <Card className="p-3">
        <div className="flex flex-wrap items-center gap-2">
          <div className="relative min-w-48 flex-1">
            <Search size={15} className="absolute left-3 top-2.5 text-slate-400" />
            <input value={q} onChange={(e) => set({ q: e.target.value })} placeholder="Search summary, owner, ID…" className={`${inputCls} pl-9`} />
          </div>
          <select value={type} onChange={(e) => set({ type: e.target.value })} className={`${inputCls} w-auto`}><option value="">All types</option>{TRANSFORM_TYPES.map((s) => <option key={s}>{s}</option>)}</select>
          <select value={bu} onChange={(e) => set({ bu: e.target.value })} className={`${inputCls} w-auto`}><option value="">All BUs</option>{BUSINESS_UNITS.map((s) => <option key={s}>{s}</option>)}</select>
          <button onClick={() => set({ overdueOnly: !overdueOnly })} className={`rounded-lg px-3 py-2 text-sm font-semibold ring-1 ring-inset ${overdueOnly ? "bg-rose-600 text-white ring-rose-600" : "bg-white text-slate-700 ring-slate-300 hover:bg-slate-50"}`}>Overdue only</button>
          {anyActive ? <button onClick={clearAll} className="inline-flex items-center gap-1 rounded-lg px-2.5 py-2 text-sm font-medium text-slate-500 hover:bg-slate-100"><XCircle size={15} />Clear</button> : null}
        </div>
        <div className="mt-2.5 flex flex-wrap items-center gap-1.5">
          <span className="mr-0.5 text-xs font-medium text-slate-400">Status:</span>
          {ALL_STATUSES.map((s) => {
            const on = statuses.includes(s);
            return (
              <button key={s} onClick={() => toggleStatus(s)}
                style={on ? { backgroundColor: BRAND.navy, color: "#fff", borderColor: BRAND.navy } : undefined}
                className={`rounded-full border px-2.5 py-1 text-xs font-medium transition ${on ? "" : "border-slate-300 text-slate-600 hover:bg-slate-50"}`}>
                {s}
              </button>
            );
          })}
          {statuses.length > 0 && <button onClick={() => set({ statuses: [] })} className="ml-0.5 text-xs font-medium text-slate-400 hover:text-slate-600">reset</button>}
        </div>
      </Card>
      <Card>
        <div className="border-b border-slate-100 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-slate-500">{list.length} issue{list.length !== 1 ? "s" : ""}{statuses.length ? ` · ${statuses.length} status filter${statuses.length !== 1 ? "s" : ""}` : ""}</div>
        {list.length === 0 ? <div className="px-4 py-10 text-center text-sm text-slate-400">No issues match these filters.</div> : list.map((i) => <IssueRow key={i.id} i={i} onOpen={onOpen} />)}
      </Card>
    </div>
  );
}

/* ============================================================
   Progress log (append-only, immutable)
   ============================================================ */
export function ProgressLog({ entries, canAdd, author, onAdd, disabled = false }) {
  const [text, setText] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const submit = async () => {
    if (disabled || saving || !text.trim()) return;
    setSaving(true);
    setError("");
    try {
      await onAdd({ ts: new Date().toISOString(), author, text: text.trim() });
      setText("");
    } catch (failure) {
      setError(failure instanceof Error ? failure.message : String(failure));
    } finally { setSaving(false); }
  };
  const sorted = [...(entries || [])].sort((a, b) => new Date(b.ts) - new Date(a.ts));
  return (
    <div>
      <SectionTitle icon={ListChecks} right={<span className="inline-flex items-center gap-1 text-xs text-slate-400"><Lock size={12} />Timestamped · cannot be edited or deleted</span>}>Progress log</SectionTitle>
      {canAdd && (
        <div className="mb-3 rounded-lg border border-slate-200 bg-slate-50 p-3">
          <TextArea disabled={disabled || saving} value={text} onChange={(e) => setText(e.target.value)} placeholder="What did you do? What are the next steps or blockers?" />
          <div className="mt-2 flex items-center justify-between">
            <span className="text-xs text-slate-400">Posting as {author} · {fmtDateTime(new Date())}</span>
            <Btn disabled={disabled || !text.trim() || saving} onClick={submit}><Plus size={15} />{saving ? "Saving…" : "Add update"}</Btn>
          </div>
          {error && <p role="alert" className="mt-2 text-sm text-rose-700">{error}</p>}
        </div>
      )}
      {sorted.length === 0 ? (
        <div className="rounded-lg border border-dashed border-slate-200 px-4 py-6 text-center text-sm text-slate-400">No updates yet.</div>
      ) : (
        <ol className="relative space-y-3 border-l border-slate-200 pl-4">
          {sorted.map((e, idx) => (
            <li key={idx} className="relative">
              <span className="absolute -left-[21px] top-1.5 h-2.5 w-2.5 rounded-full ring-4 ring-white" style={{ backgroundColor: BRAND.navy }} />
              <div className="rounded-lg bg-white p-3 ring-1 ring-slate-100">
                <div className="mb-1 flex items-center justify-between">
                  <span className="text-xs font-semibold text-slate-700">{e.author}</span>
                  <span className="font-mono text-xs text-slate-400">{fmtDateTime(e.ts)}</span>
                </div>
                <p className="whitespace-pre-wrap text-sm text-slate-700">{e.text}</p>
              </div>
            </li>
          ))}
        </ol>
      )}
    </div>
  );
}

/* ============================================================
   Read-only intake summary block
   ============================================================ */
function ReadRow({ label, value }) {
  return (
    <div className="grid grid-cols-3 gap-2 py-1.5">
      <dt className="text-xs font-semibold uppercase tracking-wide text-slate-400">{label}</dt>
      <dd className="col-span-2 text-sm text-slate-700">{value || "—"}</dd>
    </div>
  );
}
function IntakeSummary({ i }) {
  return (
    <Card className="p-4">
      <SectionTitle icon={Inbox}>Reported information (Q-Star)</SectionTitle>
      <dl className="divide-y divide-slate-100">
        <ReadRow label="Description" value={<span className="whitespace-pre-wrap">{i.description}</span>} />
        <ReadRow label="Immediate action" value={<span className="whitespace-pre-wrap">{i.immediateAction}</span>} />
        <ReadRow label="Reported by" value={i.createdBy} />
        <ReadRow label="Report date" value={fmtDate(i.reportDate)} />
        <ReadRow label="Severity" value={<SeverityPill s={i.severity} />} />
        <ReadRow label="Department / BU" value={i.departmentBU} />
        <ReadRow label="Region" value={i.region} />
        <ReadRow label="Deviation type" value={i.deviationType} />
        <ReadRow label="Origin" value={i.issueOrigin} />
        <ReadRow label="Already in contact w/ dept" value={i.alreadyInContact} />
        <ReadRow label="Attachments" value={i.attachments?.length ? i.attachments.join(", ") : <span className="inline-flex items-center gap-1 text-slate-400"><Paperclip size={12} />none</span>} />
        <ReadRow label="Additional comments" value={i.additionalComments} />
      </dl>
    </Card>
  );
}

/* ============================================================
   On-hold banner + required-fields dialog
   ============================================================ */
function HoldBanner({ i }) {
  if (i.status !== "On Hold" || !i.holdUntil) return null;
  const left = daysBetween(today(), i.holdUntil);
  const when = left > 0 ? `in ${left} day${left === 1 ? "" : "s"}` : left === 0 ? "today" : `${-left} day${-left === -1 ? "" : "s"} ago`;
  return (
    <Card className="p-4" style={{ backgroundColor: "#fffbeb", borderColor: "#fde68a" }}>
      <div className="flex items-center gap-2 text-sm font-bold text-amber-700"><Clock size={16} />On hold — resumes {fmtDate(i.holdUntil)} ({when})</div>
      {i.holdReason && <p className="mt-1.5 whitespace-pre-wrap text-sm text-slate-600"><span className="font-medium text-slate-500">Reason: </span>{i.holdReason}</p>}
      <p className="mt-1.5 text-xs text-slate-400">The owner and Quality Team are reminded a week before the resume date and again on the day.</p>
    </Card>
  );
}

function HoldDialog({ open, initialReason, initialUntil, onCancel, onConfirm, saving, needsReload }) {
  const [reason, setReason] = useState("");
  const [until, setUntil] = useState("");
  useEffect(() => { if (open) { setReason(initialReason || ""); setUntil(initialUntil || ""); } }, [open]);
  if (!open) return null;
  const valid = reason.trim() && until;
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 p-4" onClick={saving ? undefined : onCancel}>
      <div className="w-full max-w-md rounded-xl bg-white p-5 shadow-2xl" onClick={(e) => e.stopPropagation()}>
        <div className="mb-1 flex items-center gap-2"><Clock size={18} className="text-amber-500" /><h3 className="text-base font-bold text-slate-900">Put issue on hold</h3></div>
        <p className="mb-3 text-sm text-slate-500">Both fields are required. The owner and Quality Team will be reminded a week before the resume date and again on the day.</p>
        <Field label="Reason for hold *"><TextArea disabled={saving} value={reason} onChange={(e) => setReason(e.target.value)} placeholder="Why is this issue being paused?" /></Field>
        <div className="mt-3"><Field label="Resume work on *" hint="When work on the mitigation will continue"><TextInput disabled={saving} type="date" value={until} min={iso(today())} onChange={(e) => setUntil(e.target.value)} /></Field></div>
        <div className="mt-4 flex justify-end gap-2">
          <Btn variant="ghost" disabled={saving} onClick={onCancel}>Cancel</Btn>
          <Btn variant="primary" onClick={() => onConfirm(reason.trim(), until)} disabled={!valid || saving || needsReload} style={valid ? { background: "#d97706" } : undefined}><Clock size={15} />Put on hold</Btn>
        </div>
        {!valid && <p className="mt-2 text-right text-xs text-slate-400">Enter a reason and a resume date to continue.</p>}
      </div>
    </div>
  );
}

function IssueSaveNotice({ error }) {
  if (!error) return null;
  const saved = error instanceof IssueRefreshError;
  return (
    <div role={saved ? "status" : "alert"} className={`rounded-lg p-3 text-sm ${saved ? "bg-amber-50 text-amber-800" : "bg-rose-50 text-rose-800"}`}>
      {error instanceof Error ? error.message : String(error)}{!saved && " Your draft has been kept."}
    </div>
  );
}

/* ============================================================
   NC effectiveness test-period banner (NC-only 2-month gate)
   ============================================================ */
function NCTestBanner({ i }) {
  if (!isNC(i) || i.status === "Closed" || i.status === "Rejected") return null;
  const end = ncTestEnd(i);
  const left = ncTestDaysLeft(i);
  if (inNCTest(i)) {
    const totalDays = daysBetween(i.implementationDate, end) || (NC_TEST_MONTHS * 30);
    const done = left <= 0;
    const pct = Math.max(0, Math.min(100, Math.round(((totalDays - Math.max(0, left)) / totalDays) * 100)));
    return (
      <Card className="p-4" style={{ borderColor: done ? "#6ee7b7" : "#67e8f9", backgroundColor: done ? "#ecfdf5" : "#ecfeff" }}>
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-sm font-bold" style={{ color: done ? "#047857" : "#0e7490" }}>
            <FlaskConical size={16} />{done ? "Effectiveness test complete" : "Under effectiveness test"}
          </div>
          <span className="text-xs font-semibold" style={{ color: done ? "#047857" : "#0e7490" }}>{done ? "Ready to verify & close" : `${left} day${left === 1 ? "" : "s"} left`}</span>
        </div>
        <div className="mt-2 h-2 w-full overflow-hidden rounded-full bg-white ring-1 ring-inset ring-cyan-100">
          <div className="h-full rounded-full" style={{ width: `${pct}%`, backgroundColor: done ? "#10b981" : "#06b6d4" }} />
        </div>
        <p className="mt-2 text-xs text-slate-600">Mitigation implemented {fmtDate(i.implementationDate)} → {NC_TEST_MONTHS}-month test ends <strong>{fmtDate(end)}</strong>. The NC can only be closed after this period, once the Owner and QM confirm the fix is effective.</p>
      </Card>
    );
  }
  // NC not yet under test
  return (
    <Card className="p-4" style={{ backgroundColor: BRAND.yellowSoft, borderColor: BRAND.yellowRing }}>
      <div className="flex items-center gap-2 text-sm tm-text-navy">
        <FlaskConical size={16} />
        {i.implementationDate
          ? <span>Mitigation implemented {fmtDate(i.implementationDate)}. Start the {NC_TEST_MONTHS}-month effectiveness test to begin the closure clock.</span>
          : <span>This NC isn't closeable directly. Once the mitigation is implemented it enters a {NC_TEST_MONTHS}-month effectiveness test, after which the Owner and QM confirm the fix worked.</span>}
      </div>
    </Card>
  );
}


/* ============================================================
   QM Issue detail (full edit)
   ============================================================ */
export function QMIssueDetail({ issue, onBack, onUpdate, onAddProgress, onReload, author, issueBusy = false }) {
  const [d, setD] = useState(issue);
  const [baseline, setBaseline] = useState(issue);
  const [saving, setSaving] = useState(false);
  const [closing, setClosing] = useState(false);
  const [error, setError] = useState("");
  const [needsReload, setNeedsReload] = useState(!issue.eTag);
  const [holdOpen, setHoldOpen] = useState(false);
  const [hadUpdate, setHadUpdate] = useState(issue.ownerUpdate); // header badge, snapshot per opened issue
  useEffect(() => { setD(issue); setBaseline(issue); setHadUpdate(issue.ownerUpdate); setError(""); setNeedsReload(!issue.eTag); }, [issue.id]);
  const set = (k, v) => setD((p) => ({ ...p, [k]: v }));
  const dirty = Object.keys(changedIssueFields(baseline, d)).length > 0;
  const nc = isNC(d);
  const testEnd = ncTestEnd(d);
  const inTest = inNCTest(d);
  const testLeft = ncTestDaysLeft(d);
  const testDone = ncTestComplete(d);

  const acceptSaved = (saved) => { setD(saved); setBaseline(saved); setHadUpdate(saved.ownerUpdate); };
  const save = async (extra = {}) => {
    if (saving || needsReload) return false;
    setSaving(true);
    setError("");
    try {
      const next = { ...d, ...extra };
      if (next.taskOwner && !next.taskOwnerEmail) throw new Error("Enter the task owner's Microsoft 365 email.");
      if (next.verifiedBy && !next.verifiedByEmail && !next.verifiedById) throw new Error("Enter the verifier's Microsoft 365 email.");
      const patch = buildIssueTransition(baseline, changedIssueFields(baseline, next));
      setClosing(next.status === "Closed");
      acceptSaved(await onUpdate(issue.id, patch, baseline.eTag));
      return true;
    } catch (failure) {
      setError(failure);
      if (failure instanceof IssueConflictError || failure instanceof IssueRefreshError) setNeedsReload(true);
      return false;
    } finally { setSaving(false); setClosing(false); }
  };
  const putOnHold = async (holdReason, holdUntil) => { if (await save({ status: "On Hold", holdReason, holdUntil })) setHoldOpen(false); };
  const startTest = () => save({ implementationDate: d.implementationDate || iso(today()), status: NC_TEST });
  const close = () => save({ status: "Closed" });
  const reload = async () => {
    setSaving(true);
    try { acceptSaved(await onReload(issue.id)); setNeedsReload(false); setError(""); }
    catch (failure) { setError(failure); }
    finally { setSaving(false); }
  };

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-slate-800"><ArrowLeft size={15} />Back to register</button>
      <IssueSaveNotice error={error} />
      {needsReload && <Btn disabled={saving} variant="ghost" onClick={reload}>Reload latest and discard draft</Btn>}
      <Card className="p-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div className="min-w-0">
            <div className="font-mono text-xs text-slate-400">QS-{d.qsNumber}</div>
            <h2 className="text-lg font-bold text-slate-900">{d.shortSummary}</h2>
            <div className="mt-1.5 flex flex-wrap items-center gap-2">
              <StatusPill s={d.status} /><SeverityPill s={d.severity} /><TypePill t={d.transformedInto} />
              {isOverdue(d) && <Pill className="bg-rose-50 text-rose-700 ring-rose-200"><Clock size={11} />{overdueDays(d)}d overdue</Pill>}
              {hadUpdate && <Pill className="bg-amber-50 text-amber-700 ring-amber-200"><MessageSquare size={11} />New update from task owner</Pill>}
            </div>
          </div>
          <div className="text-right text-xs text-slate-500">
            <div className="flex items-center justify-end gap-1"><CalendarClock size={13} />Due {fmtDate(d.dueDate)}</div>
            <div className="mt-1 flex items-center justify-end gap-1"><UserCircle2 size={13} />{d.taskOwner || "Unassigned"}</div>
          </div>
        </div>
      </Card>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-5">
        <div className="space-y-4 lg:col-span-3">
          {nc && <NCTestBanner i={d} />}
          <HoldBanner i={d} />
          <IntakeSummary i={d} />

          <Card className="p-4">
            <SectionTitle icon={ClipboardList}>QM assessment</SectionTitle>
            <div className="grid grid-cols-2 gap-3">
              <Field label="Status"><Select disabled={saving} value={d.status} onChange={(e) => set("status", e.target.value)} options={statusOptionsFor(d)} /></Field>
              <Field label="Transformed into"><Select disabled={saving} value={d.transformedInto} onChange={(e) => set("transformedInto", e.target.value)} options={TRANSFORM_TYPES} /></Field>
              <Field label="Task owner"><TextInput disabled={saving} value={d.taskOwner || ""} onChange={(e) => set("taskOwner", e.target.value)} placeholder="Full name" /></Field>
              <Field label="Task owner Microsoft 365 email"><TextInput disabled={saving} type="email" value={d.taskOwnerEmail || ""} onChange={(e) => set("taskOwnerEmail", e.target.value)} placeholder="owner@company.com" /></Field>
              <Field label="Escalation BU"><Select disabled={saving} value={d.ownerBU} onChange={(e) => set("ownerBU", e.target.value)} options={BUSINESS_UNITS} /></Field>
              <Field label="Due date" hint={`Default for ${d.severity}: ${SEVERITY_DUE_DAYS[d.severity]} days`}><TextInput disabled={saving} type="date" value={d.dueDate || ""} onChange={(e) => set("dueDate", e.target.value)} /></Field>
              <Field label="Task created"><Select disabled={saving} value={d.taskCreated} onChange={(e) => set("taskCreated", e.target.value)} options={YESNO} /></Field>
            </div>
            <div className="mt-3"><Field label="Follow up (Quality Team notes)"><TextArea disabled={saving} value={d.followUp} onChange={(e) => set("followUp", e.target.value)} /></Field></div>
            {d.status === "Closed" && (d.closedAt || d.closedDate) && (
              <div className="mt-3 flex items-center gap-2 rounded-lg bg-emerald-50 px-3 py-2 text-sm font-medium text-emerald-800 ring-1 ring-inset ring-emerald-200">
                <CheckCircle2 size={15} />Closed on {d.closedAt ? fmtDateTime(d.closedAt) : fmtDate(d.closedDate)}
              </div>
            )}
          </Card>

          {nc && (
            <Card className="border-violet-200 p-4">
              <SectionTitle icon={ShieldCheck}><span className="text-violet-700">Corrective action — ISO 9001 §10.2</span></SectionTitle>
              <div className="space-y-3">
                <Field label="Root cause"><TextArea disabled={saving} value={d.rootCause} onChange={(e) => set("rootCause", e.target.value)} placeholder="Why did this happen? (5 Whys / Ishikawa)" /></Field>
                <Field label="Corrective action taken"><TextArea disabled={saving} value={d.correctiveAction} onChange={(e) => set("correctiveAction", e.target.value)} /></Field>
                <div className="grid grid-cols-2 gap-3">
                  <Field label="Implementation date"><TextInput disabled={saving} type="date" value={d.implementationDate || ""} onChange={(e) => set("implementationDate", e.target.value)} /></Field>
                  <Field label="Verified by"><TextInput disabled={saving} value={d.verifiedBy || ""} onChange={(e) => set("verifiedBy", e.target.value)} placeholder="Full name" /></Field>
                  <Field label="Verifier Microsoft 365 email"><TextInput disabled={saving} type="email" value={d.verifiedByEmail || ""} onChange={(e) => set("verifiedByEmail", e.target.value)} placeholder="verifier@company.com" /></Field>
                </div>
                <Field label="Effectiveness check"><TextArea disabled={saving} value={d.effectivenessCheck} onChange={(e) => set("effectivenessCheck", e.target.value)} placeholder="Evidence the action worked and the issue has not recurred." /></Field>
                {d.verifiedBy && <Field label="Verification date"><TextInput disabled={saving} type="date" value={d.verifiedDate || ""} onChange={(e) => set("verifiedDate", e.target.value)} /></Field>}
              </div>
            </Card>
          )}

          <Card className="p-4"><ProgressLog entries={issue.progressLog} canAdd disabled={closing} author={author} onAdd={(entry) => onAddProgress(issue.id, entry)} /></Card>
        </div>

        <div className="space-y-4 lg:col-span-2">
          <Card className="p-4">
            <SectionTitle icon={CheckCircle2}>Actions</SectionTitle>
            <div className="flex flex-col gap-2">
              <Btn onClick={() => save()} disabled={!dirty || saving || needsReload || (issueBusy && d.status === "Closed")}><CheckCircle2 size={15} />{saving ? "Saving…" : "Save changes"}</Btn>
              {nc && !inTest && d.status !== "Closed" && (
                <Btn variant="primary" disabled={saving || needsReload} onClick={startTest} style={{ background: "#0891b2" }} className="hover:opacity-90"><FlaskConical size={15} />Start {NC_TEST_MONTHS}-month effectiveness test</Btn>
              )}
              <Btn variant="ghost" disabled={saving || needsReload} onClick={() => setHoldOpen(true)}><Clock size={15} />Put on hold</Btn>
              <Btn variant="primary" onClick={close} disabled={saving || needsReload || issueBusy || (nc && !testDone)} style={{ background: nc && !testDone ? "#94a3b8" : "#059669" }} className="hover:opacity-90"><ShieldCheck size={15} />Verify &amp; close</Btn>
              <Btn variant="danger" disabled={saving || needsReload} onClick={() => save({ status: "Rejected", taskCreated: "No" })}><XCircle size={15} />Reject issue</Btn>
              {hadUpdate && <Btn variant="ghost" disabled={saving || needsReload} onClick={() => save({ ownerUpdate: false })}>Acknowledge owner update</Btn>}
            </div>
            {issueBusy && !saving && <p className="mt-2 text-xs text-slate-500">Wait for this issue's pending work to finish before closing.</p>}
            {nc && !testDone && d.status !== "Closed" && <p className="mt-2 text-xs text-cyan-700">{inTest ? `Closure unlocks when the test ends (${fmtDate(testEnd)}).` : "Closure unlocks after the effectiveness test completes."}</p>}
            {dirty && <p className="mt-2 text-xs text-amber-600">Unsaved changes.</p>}
          </Card>
          <Card className="p-4">
            <SectionTitle icon={BellRing}>Reminder status</SectionTitle>
            {inTest ? (
              <div className="space-y-2 text-sm text-slate-600">
                <p>Effectiveness test running. Owner <strong>{d.taskOwner || "—"}</strong> and the QM are reminded 7 days before and on the test end date.</p>
                <Pill className="bg-cyan-50 text-cyan-700 ring-cyan-200"><FlaskConical size={11} />{testDone ? "Test complete — verify & close" : `Test ends ${fmtDate(testEnd)}`}</Pill>
              </div>
            ) : isOpen(d) ? (
              <div className="space-y-2 text-sm text-slate-600">
                <p>Owner <strong>{d.taskOwner || "—"}</strong> is reminded at −3d, on the due date, then every 3 days while overdue.</p>
                {isOverdue(d) ? <Pill className="bg-rose-50 text-rose-700 ring-rose-200">Overdue {overdueDays(d)}d — QM alerted{overdueDays(d) >= 7 ? " + BU escalated" : ""}</Pill>
                  : <Pill className="bg-emerald-50 text-emerald-700 ring-emerald-200">On track — due {fmtDate(d.dueDate)}</Pill>}
              </div>
            ) : <p className="text-sm text-slate-400">No active reminders (issue is {d.status?.toLowerCase() || "untriaged"}).</p>}
          </Card>
        </div>
      </div>
      <HoldDialog saving={saving} needsReload={needsReload} open={holdOpen} initialReason={d.holdReason} initialUntil={d.holdUntil} onCancel={() => setHoldOpen(false)} onConfirm={putOnHold} />
    </div>
  );
}

/* ============================================================
   Triage queue + triage form
   ============================================================ */
function TriageQueue({ issues, onOpen }) {
  const pending = issues.filter((i) => !i.triaged);
  return (
    <div className="space-y-4">
      <Card className="p-4" style={{ backgroundColor: BRAND.yellowSoft, borderColor: BRAND.yellowRing }}>
        <div className="flex items-center gap-2 text-sm tm-text-navy"><Inbox size={16} /><strong>{pending.length}</strong> candidate{pending.length !== 1 ? "s" : ""} awaiting your assessment. Create an issue (classify + assign) or reject.</div>
      </Card>
      {pending.length === 0 ? <Card className="px-4 py-10 text-center text-sm text-slate-400">Triage queue is clear. 🎉</Card> :
        <Card>{pending.map((i) => (
          <button key={i.id} onClick={() => onOpen(i.id)} className="flex w-full items-center gap-3 border-b border-slate-100 px-4 py-3 text-left hover:bg-slate-50">
            <div className="w-16 shrink-0 font-mono text-xs text-slate-400">QS-{i.qsNumber}</div>
            <div className="min-w-0 flex-1">
              <div className="truncate text-sm font-semibold text-slate-800">{i.shortSummary}</div>
              <div className="mt-0.5 text-xs text-slate-500">{i.departmentBU} · {i.region} · reported {fmtDate(i.reportDate)} by {i.createdBy}</div>
            </div>
            <SeverityPill s={i.severity} />
            <Pill className="bg-slate-100 text-slate-600 ring-slate-200">{i.issueOrigin?.includes("Customer") ? "Customer" : "Internal"}</Pill>
            <ChevronRight size={16} className="shrink-0 text-slate-300" />
          </button>
        ))}</Card>}
    </div>
  );
}

export function TriageForm({ issue, onBack, onTriage, onReload }) {
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [needsReload, setNeedsReload] = useState(!issue.eTag);
  const baseline = useRef(issue);
  const persist = async (patch) => {
    if (saving || needsReload) return;
    setSaving(true); setError("");
    try { await onTriage(issue.id, patch, baseline.current.eTag); }
    catch (failure) {
      setError(failure);
      if (failure instanceof IssueConflictError || failure instanceof IssueRefreshError) setNeedsReload(true);
    } finally { setSaving(false); }
  };
  const reload = async () => {
    setSaving(true);
    try { await onReload(issue.id); onBack(); }
    catch (failure) { setError(failure); }
    finally { setSaving(false); }
  };
  const [transformedInto, setT] = useState("OFI");
  const [taskOwner, setOwner] = useState("");
  const [taskOwnerEmail, setOwnerEmail] = useState("");
  const [ownerBU, setBU] = useState(issue.departmentBU);
  const [dueDate, setDue] = useState(addDays(issue.reportDate, SEVERITY_DUE_DAYS[issue.severity]));
  const [followUp, setFollow] = useState("");
  const nc = transformedInto === "NC Minor" || transformedInto === "NC Major";

  const create = () => {
    if (!taskOwner || !taskOwnerEmail) { setError("Assign a task owner and their Microsoft 365 email so permissions and reminders have a stable recipient."); return; }
    persist({ triaged: true, status: "Created", transformedInto, taskOwner, taskOwnerEmail, ownerBU, dueDate, followUp, taskCreated: transformedInto === "Only sent to Dept/BU for Action" ? "No" : "Yes" });
  };
  const reject = () => persist({ triaged: true, status: "Rejected", taskCreated: "No", followUp: followUp || "Declined at triage — no quality action required." });

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-slate-800"><ArrowLeft size={15} />Back to triage queue</button>
      <IssueSaveNotice error={error} />
      {needsReload && <Btn onClick={reload} disabled={saving}>Reload latest and return to queue</Btn>}
      <div className="grid grid-cols-1 gap-4 lg:grid-cols-5">
        <div className="lg:col-span-3"><IntakeSummary i={issue} /></div>
        <div className="lg:col-span-2">
          <Card className="p-4">
            <SectionTitle icon={ClipboardList}>Assess & decide</SectionTitle>
            <div className="space-y-3">
              <Field label="Transform into"><Select disabled={saving} value={transformedInto} onChange={(e) => setT(e.target.value)} options={TRANSFORM_TYPES} /></Field>
              {nc && <p className="rounded-lg bg-violet-50 px-3 py-2 text-xs text-violet-700">Nonconformity selected — §10.2 corrective-action fields will open on the issue once created.</p>}
              <Field label="Task owner (gets reminders)"><TextInput disabled={saving} value={taskOwner} onChange={(e) => setOwner(e.target.value)} placeholder="Full name" /></Field>
              <Field label="Task owner Microsoft 365 email"><TextInput disabled={saving} type="email" value={taskOwnerEmail} onChange={(e) => setOwnerEmail(e.target.value)} placeholder="owner@company.com" /></Field>
              <Field label="Escalation BU"><Select disabled={saving} value={ownerBU} onChange={(e) => setBU(e.target.value)} options={BUSINESS_UNITS} /></Field>
              <Field label="Due date" hint={`Auto from ${issue.severity} severity (${SEVERITY_DUE_DAYS[issue.severity]}d) — override if needed`}><TextInput disabled={saving} type="date" value={dueDate} onChange={(e) => setDue(e.target.value)} /></Field>
              <Field label="Follow up note (optional)"><TextArea disabled={saving} value={followUp} onChange={(e) => setFollow(e.target.value)} /></Field>
              <div className="flex flex-col gap-2 pt-1">
                <Btn onClick={create} disabled={saving || needsReload}><CheckCircle2 size={15} />Create issue</Btn>
                <Btn variant="danger" onClick={reject} disabled={saving || needsReload}><XCircle size={15} />Reject (no action)</Btn>
              </div>
            </div>
          </Card>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   Reporter intake form (mirrors Q-Star)
   ============================================================ */
export function ReporterForm({ onSubmit, settings, reporterName }) {
  const blank = { shortSummary: "", description: "", immediateAction: "", severity: "", createdBy: reporterName, departmentBU: "", region: "", alreadyInContact: "", deviationType: "", issueOrigin: "", additionalComments: "", attachments: [] };
  const [f, setF] = useState(blank);
  const [done, setDone] = useState(null);
  const [submitting, setSubmitting] = useState(false);
  const [submitError, setSubmitError] = useState("");
  const set = (k, v) => setF((p) => ({ ...p, [k]: v }));
  const valid = f.shortSummary && f.description && f.severity && f.createdBy && f.departmentBU && f.region && f.deviationType && f.issueOrigin;
  const formUrl = settings?.msFormUrl;

  const submit = async () => {
    if (submitting || !valid) return;
    setSubmitting(true);
    setSubmitError("");
    try {
      const qs = await onSubmit(f);
      setDone(qs);
      setF(blank);
    } catch (error) {
      setSubmitError(error instanceof Error ? error.message : String(error));
    } finally {
      setSubmitting(false);
    }
  };

  if (done) return (
    <Card className="mx-auto max-w-xl p-8 text-center">
      <CheckCircle2 size={40} className="mx-auto text-emerald-500" />
      <h2 className="mt-3 text-lg font-bold text-slate-900">Report submitted</h2>
      <p className="mt-1 text-sm text-slate-600">Reference <strong className="font-mono">QS-{done}</strong>. The Quality Team will assess it and, if a quality issue is confirmed, create a tracked issue with an owner and due date.</p>
      <div className="mt-4"><Btn variant="ghost" onClick={() => setDone(null)}><Plus size={15} />Report another</Btn></div>
    </Card>
  );

  return (
    <Card className="mx-auto max-w-2xl p-5">
      <SectionTitle icon={Send}>Report an issue — Q-Star</SectionTitle>
      <p className="mb-4 text-sm text-slate-500">Anyone can report a potential quality issue. The Quality Team reviews every submission.</p>
      {formUrl && (
        <a href={formUrl} target="_blank" rel="noreferrer" className="mb-4 flex items-center justify-between rounded-lg p-3 text-sm tm-text-navy ring-1 ring-inset" style={{ backgroundColor: BRAND.navySoft, borderColor: "transparent" }}>
          <span className="flex items-center gap-2"><Link2 size={15} />Open the live Q-Star form (Microsoft Forms)</span>
          <ChevronRight size={15} />
        </a>
      )}
      <div className="space-y-3">
        <Field label="Short summary *"><TextInput disabled={submitting} value={f.shortSummary} onChange={(e) => set("shortSummary", e.target.value)} placeholder="One line describing the issue" /></Field>
        <Field label="Description *"><TextArea disabled={submitting} value={f.description} onChange={(e) => set("description", e.target.value)} placeholder="What happened, where, and when?" /></Field>
        <Field label="Immediate action taken"><TextArea disabled={submitting} value={f.immediateAction} onChange={(e) => set("immediateAction", e.target.value)} placeholder="Any containment already done?" /></Field>
        <div className="grid grid-cols-2 gap-3">
          <Field label="Severity *"><Select disabled={submitting} value={f.severity} onChange={(e) => set("severity", e.target.value)} options={SEVERITIES} /></Field>
          <Field label="Reported by" hint="Your signed-in Microsoft 365 identity"><TextInput value={reporterName} readOnly /></Field>
          <Field label="Department / Business Unit *"><Select disabled={submitting} value={f.departmentBU} onChange={(e) => set("departmentBU", e.target.value)} options={BUSINESS_UNITS} /></Field>
          <Field label="Region *"><Select disabled={submitting} value={f.region} onChange={(e) => set("region", e.target.value)} options={REGIONS} /></Field>
          <Field label="Deviation type *"><Select disabled={submitting} value={f.deviationType} onChange={(e) => set("deviationType", e.target.value)} options={DEVIATION_TYPES} /></Field>
          <Field label="Where does it come from? *"><Select disabled={submitting} value={f.issueOrigin} onChange={(e) => set("issueOrigin", e.target.value)} options={ORIGINS} /></Field>
          <Field label="Already in contact with the dept?"><Select disabled={submitting} value={f.alreadyInContact} onChange={(e) => set("alreadyInContact", e.target.value)} options={YESNO} /></Field>
        </div>
        <Field label="Additional comments (optional)"><TextArea disabled={submitting} value={f.additionalComments} onChange={(e) => set("additionalComments", e.target.value)} /></Field>
        <Field label="Attachment" hint="Attachments are not supported by this form yet."><div className="flex items-center gap-2 rounded-lg border border-dashed border-slate-300 px-3 py-3 text-sm text-slate-400"><Paperclip size={15} />File upload unavailable</div></Field>
        {submitError && <p className="text-sm text-rose-700">Could not submit: {submitError}</p>}
        <div className="flex justify-end pt-1"><Btn onClick={submit} disabled={!valid || submitting}><Send size={15} />{submitting ? "Submitting…" : "Submit report"}</Btn></div>
        {!valid && <p className="text-right text-xs text-slate-400">Fields marked * are required.</p>}
      </div>
    </Card>
  );
}

/* ============================================================
   Admin · IT settings (connect MS Form + SharePoint)
   ============================================================ */
export function SettingsView({ settings, onSave, onRunDiagnostics, connection }) {
  const [saving, setSaving] = useState(false);
  const [saveError, setSaveError] = useState("");
  const save = async () => {
    if (saving) return;
    setSaving(true); setSaveError("");
    try { await onSave(s); }
    catch (error) { setSaveError(error instanceof Error ? error.message : String(error)); }
    finally { setSaving(false); }
  };
  const [s, setS] = useState(settings);
  const [running, setRunning] = useState(false);
  const [results, setResults] = useState<ICheckResult[] | null>(null);
  const [diagnosticError, setDiagnosticError] = useState("");
  useEffect(() => setS(settings), [settings]);
  const set = (k, v) => setS((p) => ({ ...p, [k]: v }));
  const dirty = JSON.stringify(s) !== JSON.stringify(settings);

  const runDiagnostics = async () => {
    setRunning(true);
    setDiagnosticError("");
    try {
      setResults(await onRunDiagnostics());
    } catch (error) {
      setDiagnosticError(error instanceof Error ? error.message : String(error));
    } finally {
      setRunning(false);
    }
  };

  return (
    <div className="mx-auto max-w-3xl space-y-4">
      <Card className="p-4" style={{ backgroundColor: BRAND.navySoft }}>
        <div className="flex items-center gap-2 text-sm tm-text-navy">
          <KeyRound size={16} />
          <strong>Admin · IT settings.</strong>
          Authentication uses the signed-in Microsoft 365 identity. Roles are maintained in the four Q-Star SharePoint groups.
        </div>
      </Card>

      <Card className="p-5">
        <SectionTitle icon={ShieldCheck}>Production access model</SectionTitle>
        <p className="text-sm text-slate-600">
          Q-Star Admins and Quality Managers can edit the register, Readers are read-only, and assigned Task Owners receive item-level edit permission through Power Automate. There is no email-to-role table or separate Q-Star login.
        </p>
      </Card>

      <Card className="p-5">
        <SectionTitle icon={Link2}>Reporter intake — Microsoft Forms</SectionTitle>
        <div className="space-y-3">
          <Field label="Q-Star Microsoft Form URL" hint="The shareable link employees use to report issues">
            <TextInput disabled={saving} value={s.msFormUrl || ""} onChange={(e) => set("msFormUrl", e.target.value)} placeholder="https://forms.office.com/r/XXXXXXXXXX" />
          </Field>
          <Field label="Form responses → SharePoint flow ID (optional)">
            <TextInput disabled={saving} value={s.flowId || ""} onChange={(e) => set("flowId", e.target.value)} />
          </Field>
        </div>
      </Card>

      <Card className="p-5">
        <SectionTitle icon={Building2}>SharePoint connection</SectionTitle>
        <p className="mb-3 text-sm text-slate-600">
          The web part uses its configured site and Lists through the signed-in user's SharePoint session. No client secret or app registration is needed for same-site data.
        </p>
        <dl className="divide-y divide-slate-100">
          <ReadRow label="SharePoint site URL" value={connection?.siteUrl || "Local development"} />
          <ReadRow label="Issues list" value={connection?.issuesListName || "Q-Star Issues"} />
          <ReadRow label="Progress list" value={connection?.progressListName || "Q-Star Progress Log"} />
          <ReadRow label="Access mode" value={connection?.betaAccessMode ? "Beta pilot" : "Production groups"} />
        </dl>
        <p className="mt-3 text-xs text-slate-500">To change the connection, edit the SharePoint page and open this web part's properties. Saving settings below changes only the Forms link and flow ID.</p>
      </Card>

      <Card className="p-5">
        <SectionTitle icon={Plug}>Connection diagnostics</SectionTitle>
        <p className="mb-3 text-sm text-slate-600">Checks site access, the signed-in user, List schema, indexes, and a create/update/delete round trip.</p>
        <Btn onClick={runDiagnostics} disabled={running}><Plug size={15} />{running ? "Running…" : "Run connection test"}</Btn>
        {diagnosticError && <p className="mt-3 text-sm text-rose-700">{diagnosticError}</p>}
        {results && (
          <div className="mt-3 divide-y divide-slate-100 rounded-lg border border-slate-200">
            {results.map((result, index) => (
              <div key={index} className="px-3 py-2 text-sm">
                <strong>{result.status === "pass" ? "✓" : result.status === "warn" ? "!" : "×"} {result.name}</strong>
                <span className="ml-2 text-slate-600">{result.message}</span>
              </div>
            ))}
          </div>
        )}
      </Card>

      <div className="flex items-center gap-2">
        <Btn onClick={save} disabled={!dirty || saving}><CheckCircle2 size={15} />{saving ? "Saving…" : "Save settings"}</Btn>
        {saveError && <p role="alert" className="text-sm text-rose-700">{saveError}</p>}
      </div>
    </div>
  );
}

/* ============================================================
   Read-only issue detail (Owner viewing others' issues)
   ============================================================ */
function ReadOnlyIssueDetail({ issue, onBack, onReopen }) {
  const closed = issue.status === "Closed";
  return (
    <div className="space-y-4">
      <button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-slate-800"><ArrowLeft size={15} />Back to register</button>
      <Card className="p-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <div className="font-mono text-xs text-slate-400">QS-{issue.qsNumber}</div>
            <h2 className="text-lg font-bold text-slate-900">{issue.shortSummary}</h2>
            <div className="mt-1.5 flex flex-wrap items-center gap-2"><StatusPill s={issue.status} /><SeverityPill s={issue.severity} /><TypePill t={issue.transformedInto} /></div>
          </div>
          <Pill className="bg-slate-100 text-slate-500 ring-slate-200"><Eye size={11} />{onReopen ? "Closed · read-only" : "Read-only"}</Pill>
        </div>
      </Card>
      {onReopen && (
        <Card className="p-4" style={{ borderColor: BRAND.yellowRing, backgroundColor: BRAND.yellowSoft }}>
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-start gap-2 text-sm tm-text-navy">
              <RefreshCw size={16} className="mt-0.5 shrink-0" />
              <span><strong>This issue is closed.</strong> The full record and history below are read-only. Re-opening restarts the corrective-action cycle — status returns to In Progress and effectiveness must be verified again before it can be closed.</span>
            </div>
            <Btn variant="primary" onClick={onReopen}><RefreshCw size={15} />Re-open issue</Btn>
          </div>
        </Card>
      )}
      {isNC(issue) && <NCTestBanner i={issue} />}
      <HoldBanner i={issue} />
      <IntakeSummary i={issue} />
      <Card className="p-4">
        <SectionTitle icon={ClipboardList}>QM assessment</SectionTitle>
        <dl className="divide-y divide-slate-100">
          <ReadRow label="Task owner" value={issue.taskOwner} />
          <ReadRow label="Due date" value={fmtDate(issue.dueDate)} />
          {issue.status === "On Hold" && issue.holdUntil && <ReadRow label="On hold until" value={fmtDate(issue.holdUntil)} />}
          {issue.status === "On Hold" && issue.holdReason && <ReadRow label="Hold reason" value={<span className="whitespace-pre-wrap">{issue.holdReason}</span>} />}
          <ReadRow label="Follow up" value={<span className="whitespace-pre-wrap">{issue.followUp}</span>} />
          {issue.rootCause && <ReadRow label="Root cause" value={<span className="whitespace-pre-wrap">{issue.rootCause}</span>} />}
          {issue.correctiveAction && <ReadRow label="Corrective action" value={<span className="whitespace-pre-wrap">{issue.correctiveAction}</span>} />}
          {issue.implementationDate && <ReadRow label="Implemented" value={fmtDate(issue.implementationDate)} />}
          {issue.verifiedBy && <ReadRow label="Verified by" value={`${issue.verifiedBy} · ${fmtDate(issue.verifiedDate)}`} />}
          {issue.status === "Closed" && (issue.closedAt || issue.closedDate) && <ReadRow label="Closed on" value={issue.closedAt ? fmtDateTime(issue.closedAt) : fmtDate(issue.closedDate)} />}
        </dl>
      </Card>
      <Card className="p-4"><ProgressLog entries={issue.progressLog} canAdd={false} /></Card>
    </div>
  );
}

/* ============================================================
   Reminders panel
   ============================================================ */
function RemindersView({ issues }) {
  const reminders = useMemo(() => buildReminders(issues), [issues]);
  const todayStr = iso(today());
  const dueToday = reminders.filter((r) => sameCalendarDay(r.when, todayStr));
  const sent = reminders.filter((r) => r.sent && r.when !== todayStr);
  const upcoming = reminders.filter((r) => !r.sent).sort((a, b) => new Date(a.when) - new Date(b.when));
  const urgencyCls = { info: "bg-blue-50 text-blue-700 ring-blue-200", warn: "bg-amber-50 text-amber-700 ring-amber-200", danger: "bg-rose-50 text-rose-700 ring-rose-200" };
  const roleIcon = { "Task Owner": UserCircle2, "Quality Manager": ShieldCheck, "BU Escalation": Building2 };

  const Item = ({ r }) => {
    const Icon = roleIcon[r.role] || BellRing;
    return (
      <div className="flex items-start gap-3 border-b border-slate-100 px-4 py-3">
        <div className={`mt-0.5 rounded-md p-1.5 ring-1 ring-inset ${urgencyCls[r.urgency]}`}><Icon size={15} /></div>
        <div className="min-w-0 flex-1">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-sm font-semibold text-slate-800">{r.type}</span>
            <Pill className="bg-slate-100 text-slate-600 ring-slate-200">{r.role}</Pill>
            <span className="text-xs text-slate-400">→ {r.recipient}</span>
          </div>
          <p className="mt-0.5 text-sm text-slate-600">{r.message}</p>
        </div>
        <div className="shrink-0 text-right text-xs text-slate-400">{fmtDate(r.when)}</div>
      </div>
    );
  };

  return (
    <div className="space-y-4">
      <Card className="p-4" style={{ backgroundColor: BRAND.yellowSoft, borderColor: BRAND.yellowRing }}>
        <div className="flex items-center gap-2 text-sm tm-text-navy"><BellRing size={16} /><strong>{dueToday.length}</strong> reminder{dueToday.length !== 1 ? "s" : ""} would be sent today. This panel simulates the notification engine — connect Power Automate / Microsoft Graph in production to actually deliver them.</div>
      </Card>

      <Card>
        <div className="border-b border-slate-100 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide tm-text-navy">Going out today</div>
        {dueToday.length ? dueToday.map((r, k) => <Item key={k} r={r} />) : <div className="px-4 py-6 text-center text-sm text-slate-400">Nothing scheduled for today.</div>}
      </Card>

      <Card>
        <div className="border-b border-slate-100 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Upcoming</div>
        {upcoming.length ? upcoming.slice(0, 12).map((r, k) => <Item key={k} r={r} />) : <div className="px-4 py-6 text-center text-sm text-slate-400">No upcoming reminders.</div>}
      </Card>

      <Card>
        <div className="border-b border-slate-100 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Already sent</div>
        {sent.length ? sent.slice(0, 12).map((r, k) => <Item key={k} r={r} />) : <div className="px-4 py-6 text-center text-sm text-slate-400">No reminders sent yet.</div>}
      </Card>
    </div>
  );
}

/* ============================================================
   Task Owner views
   ============================================================ */
function OwnerTasks({ issues, owner, ownerEmailAddress, onOpen }) {
  const mine = issues.filter((i) => isOpen(i) && (
    i.taskOwner === owner || (i.taskOwnerEmail && i.taskOwnerEmail.toLowerCase() === ownerEmailAddress.toLowerCase())
  )).sort((a, b) => new Date(a.dueDate) - new Date(b.dueDate));
  const od = mine.filter(isOverdue);
  return (
    <div className="space-y-4">
      <div className="grid grid-cols-3 gap-4">
        <StatCard label="My open tasks" value={mine.length} tone="teal" icon={ClipboardList} />
        <StatCard label="Overdue" value={od.length} tone={od.length ? "rose" : "slate"} icon={AlertTriangle} />
        <StatCard label="Due in 7 days" value={mine.filter((i) => { const d = daysBetween(today(), i.dueDate); return d >= 0 && d <= 7; }).length} tone="amber" icon={CalendarClock} />
      </div>
      <Card>
        <div className="border-b border-slate-100 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-slate-500">Tasks assigned to {owner}</div>
        {mine.length ? mine.map((i) => <IssueRow key={i.id} i={i} onOpen={onOpen} />) : <div className="px-4 py-10 text-center text-sm text-slate-400">No open tasks assigned to you.</div>}
      </Card>
    </div>
  );
}

export function OwnerIssueDetail({ issue, owner, onBack, onUpdate, onAddProgress, onReload }) {
  const [baseline, setBaseline] = useState(issue);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [needsReload, setNeedsReload] = useState(!issue.eTag);
  const [status, setStatus] = useState(issue.status);
  const [impl, setImpl] = useState(issue.implementationDate || "");
  const [holdOpen, setHoldOpen] = useState(false);
  useEffect(() => { setBaseline(issue); setStatus(issue.status); setImpl(issue.implementationDate || ""); setNeedsReload(!issue.eTag); setError(""); }, [issue.id]);
  const accept = (saved) => { setBaseline(saved); setStatus(saved.status); setImpl(saved.implementationDate || ""); };
  const persist = async (patch) => {
    if (saving || needsReload) return false;
    setSaving(true); setError("");
    try { accept(await onUpdate(issue.id, buildIssueTransition(baseline, patch), baseline.eTag)); return true; }
    catch (failure) {
      setError(failure);
      if (failure instanceof IssueConflictError || failure instanceof IssueRefreshError) setNeedsReload(true);
      return false;
    } finally { setSaving(false); }
  };
  const reload = async () => {
    setSaving(true);
    try { accept(await onReload(issue.id)); setNeedsReload(false); setError(""); }
    catch (failure) { setError(failure); }
    finally { setSaving(false); }
  };
  const nc = isNC(issue);
  const inTest = inNCTest(issue);
  const ownerOpts = issue.status === "On Hold" ? ["On Hold", "Created", "In Progress"] : ["Created", "In Progress"];
  const putOnHold = async (holdReason, holdUntil) => { if (await persist({ status: "On Hold", holdReason, holdUntil })) setHoldOpen(false); };
  const implementMitigation = () => persist({ status: NC_TEST, implementationDate: impl || iso(today()) });

  return (
    <div className="space-y-4">
      <button onClick={onBack} className="inline-flex items-center gap-1 text-sm font-medium text-slate-500 hover:text-slate-800"><ArrowLeft size={15} />Back to my tasks</button>
      <IssueSaveNotice error={error} />
      {needsReload && <Btn onClick={reload} disabled={saving}>Reload latest and discard draft</Btn>}
      <Card className="p-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <div className="font-mono text-xs text-slate-400">QS-{issue.qsNumber}</div>
            <h2 className="text-lg font-bold text-slate-900">{issue.shortSummary}</h2>
            <div className="mt-1.5 flex flex-wrap items-center gap-2"><StatusPill s={issue.status} /><SeverityPill s={issue.severity} /><TypePill t={issue.transformedInto} /></div>
          </div>
          <div className="text-right text-xs">
            {inTest ? <Pill className="bg-cyan-50 text-cyan-700 ring-cyan-200"><FlaskConical size={11} />Under test</Pill>
              : isOverdue(issue) ? <Pill className="bg-rose-50 text-rose-700 ring-rose-200"><Clock size={11} />{overdueDays(issue)}d overdue</Pill>
              : <div className="flex items-center justify-end gap-1 text-slate-500"><CalendarClock size={13} />Due {fmtDate(issue.dueDate)}</div>}
          </div>
        </div>
      </Card>

      <div className="grid grid-cols-1 gap-4 lg:grid-cols-5">
        <div className="space-y-4 lg:col-span-3">
          {nc && <NCTestBanner i={issue} />}
          <HoldBanner i={issue} />
          <Card className="p-4">
            <SectionTitle icon={FileText}>What you need to do</SectionTitle>
            <dl className="divide-y divide-slate-100">
              <ReadRow label="Issue" value={<span className="whitespace-pre-wrap">{issue.description}</span>} />
              {issue.correctiveAction && <ReadRow label="Corrective action" value={<span className="whitespace-pre-wrap">{issue.correctiveAction}</span>} />}
              {issue.followUp && <ReadRow label="QM follow-up note" value={<span className="whitespace-pre-wrap">{issue.followUp}</span>} />}
            </dl>
          </Card>
          <Card className="p-4"><ProgressLog entries={issue.progressLog} canAdd author={owner} onAdd={(e) => onAddProgress(issue.id, e)} /></Card>
        </div>

        <div className="lg:col-span-2">
          <Card className="p-4">
            <SectionTitle icon={CheckCircle2}>Update your task</SectionTitle>
            {inTest ? (
              <div className="space-y-3">
                <p className="text-sm text-slate-600">Your mitigation is in its {NC_TEST_MONTHS}-month effectiveness test. Keep adding observations in the progress log. The QM verifies and closes once the test period ends.</p>
              </div>
            ) : nc ? (
              <div className="space-y-3">
                <Field label="Status"><Select disabled={saving} value={status} onChange={(e) => setStatus(e.target.value)} options={ownerOpts} /></Field>
                <Field label="Implementation date" hint="When you put the corrective action in place"><TextInput disabled={saving} type="date" value={impl} onChange={(e) => setImpl(e.target.value)} /></Field>
                <Btn variant="ghost" onClick={() => persist({ status, implementationDate: impl })} disabled={saving || needsReload || (status === baseline.status && impl === (baseline.implementationDate || ""))}><CheckCircle2 size={15} />Save progress</Btn>
                {issue.status !== "On Hold" && <Btn variant="ghost" disabled={saving || needsReload} onClick={() => setHoldOpen(true)}><Clock size={15} />Put on hold</Btn>}
                <Btn variant="primary" disabled={saving || needsReload} onClick={implementMitigation} style={{ background: "#0891b2" }} className="hover:opacity-90"><FlaskConical size={15} />Mitigation implemented — start {NC_TEST_MONTHS}-month test</Btn>
                <p className="text-xs text-slate-400">Marking the mitigation as implemented moves this NC into the {NC_TEST_MONTHS}-month effectiveness test. It can only be closed by the QM afterwards.</p>
              </div>
            ) : (
              <div className="space-y-3">
                <Field label="Status"><Select disabled={saving} value={status} onChange={(e) => setStatus(e.target.value)} options={ownerOpts} /></Field>
                <Btn onClick={() => persist({ status, implementationDate: impl })} disabled={saving || needsReload || status === baseline.status}><CheckCircle2 size={15} />Save</Btn>
                {issue.status !== "On Hold" && <Btn variant="ghost" disabled={saving || needsReload} onClick={() => setHoldOpen(true)}><Clock size={15} />Put on hold</Btn>}
                <p className="text-xs text-slate-400">Closing and effectiveness verification are done by the Quality Team. Add a progress note to let them know when you're ready.</p>
              </div>
            )}
          </Card>
        </div>
      </div>
      <HoldDialog saving={saving} needsReload={needsReload} open={holdOpen} initialReason={issue.holdReason} initialUntil={issue.holdUntil} onCancel={() => setHoldOpen(false)} onConfirm={putOnHold} />
    </div>
  );
}

/* ============================================================
   Root app
   ============================================================ */
const DEFAULT_SETTINGS = {
  msFormUrl: "",
  flowId: "",
  access: [],
};

export interface IQstarPrototypeProps {
  dataService: IDataService;
  profile: Role;
  userDisplayName: string;
  userEmail: string;
  developmentMode: boolean;
  onRunDiagnostics: () => Promise<ICheckResult[]>;
  connection?: { siteUrl: string; issuesListName: string; progressListName: string; betaAccessMode: boolean };
}

export default function App({
  dataService,
  profile,
  userDisplayName,
  userEmail,
  developmentMode,
  onRunDiagnostics,
  connection,
}: IQstarPrototypeProps) {
  const [issues, setIssues] = useState(null);
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [loadError, setLoadError] = useState("");
  const [saveError, setSaveError] = useState("");
  const [saveWarning, setSaveWarning] = useState(null);
  const [reloadingSaved, setReloadingSaved] = useState(false);
  const [reloadToken, setReloadToken] = useState(0);
  const [tab, setTab] = useState(profile === "owner" ? "mytasks" : "dashboard");
  const pendingOperations = useRef(new Map());
  const [busyIssueIds, setBusyIssueIds] = useState([]);
  const [openId, setOpenId] = useState(null);
  const [regFilter, setRegFilter] = useState({ q: "", statuses: [], type: "", bu: "", overdueOnly: false });

  useEffect(() => {
    let cancelled = false;
    setLoadError("");
    setIssues(null);
    Promise.all([dataService.loadIssues(), dataService.loadSettings()])
      .then(async ([loadedIssues, loadedSettings]) => {
        const initialIssues = !loadedIssues.length && developmentMode && dataService.initializeIssues
          ? await dataService.initializeIssues(seed()) : loadedIssues;
        if (cancelled) return;
        setIssues(initialIssues);
        setSettings({ ...DEFAULT_SETTINGS, ...loadedSettings, access: [] });
      })
      .catch((error) => {
        if (cancelled) return;
        setLoadError(error instanceof Error ? error.message : String(error));
        setIssues([]);
      });
    return () => { cancelled = true; };
  }, [dataService, developmentMode, reloadToken]);

  const runIssueOperation = (id, operation) => {
    const previous = pendingOperations.current.get(id) || Promise.resolve();
    const pending = previous.catch(() => undefined).then(operation);
    pendingOperations.current.set(id, pending);
    setBusyIssueIds(Array.from(pendingOperations.current.keys()));
    return pending.finally(() => {
      if (pendingOperations.current.get(id) === pending) {
        pendingOperations.current.delete(id);
        setBusyIssueIds(Array.from(pendingOperations.current.keys()));
      }
    });
  };
  const acceptIssue = (saved) => {
    setIssues((items) => items.map((item) => item.id === saved.id ? saved : item));
    setSaveWarning((warning) => saved.saveWarning ? { message: saved.saveWarning, issueId: saved.id } : warning?.issueId === saved.id ? null : warning);
    return saved;
  };
  const reloadIssue = (id) => runIssueOperation(id, async () => acceptIssue(await dataService.getIssue(id)));
  const reloadSavedIssue = async () => {
    if (reloadingSaved) return;
    setReloadingSaved(true);
    try { await reloadIssue(saveWarning.issueId); }
    catch (error) { setSaveWarning({ ...saveWarning, message: `Your change was saved, but reloading is still unavailable: ${error instanceof Error ? error.message : String(error)}. Do not submit it again.` }); }
    finally { setReloadingSaved(false); }
  };
  const updateIssue = (id, patch, expectedETag) => {
    if (patch.status === "Closed" && pendingOperations.current.has(id)) {
      throw new Error("Wait for this issue's pending work to finish before closing. Try closing again once it finishes.");
    }
    return runIssueOperation(id, async () => {
      const previous = issues.find((issue) => issue.id === id);
      if (!previous) throw new Error("Issue is no longer available. Reload the register.");
      if (!previous.eTag) throw new Error("Reload this issue before saving so its current version can be checked.");
      const normalized = buildIssueTransition(previous, patch);
      setSaveError("");
      try {
        return acceptIssue(await dataService.updateIssue(id, normalized, expectedETag || previous.eTag));
      } catch (error) {
        if (error instanceof IssueRefreshError) {
          setIssues((items) => items.map((item) => item.id === id ? { ...item, eTag: undefined } : item));
          setSaveWarning({ message: error.message, issueId: id });
        } else setSaveError(error instanceof Error ? error.message : String(error));
        throw error;
      }
    });
  };

  const addProgress = (id, entry) => runIssueOperation(id, async () => {
    const signedEntry = { ...entry, author: userDisplayName, authorEmail: userEmail };
    setSaveError("");
    try {
      const saved = await dataService.addProgressLogEntry(id, signedEntry);
      setIssues((items) => items.map((item) => item.id === id
        ? { ...item, progressLog: [...(item.progressLog || []), saved] } : item));
      if (saved.saveWarning) setSaveWarning({ message: saved.saveWarning, issueId: id });
      return saved;
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error));
      throw error;
    }
  });

  // Flow B consumes accepted journal rows directly; no second write can make
  // an accepted comment appear failed and invite a duplicate submission.
  const ownerAddProgress = addProgress;
  const ownerUpdateTask = (id, patch, expectedETag) => {
    const current = issues.find((issue) => issue.id === id);
    const statusChanged = current && patch.status && patch.status !== current.status;
    const extra = statusChanged ? {
      ownerUpdate: true,
      ownerUpdateAt: new Date().toISOString(),
      ownerUpdateText: `Status changed from "${current.status}" to "${patch.status}".`,
    } : {};
    return updateIssue(id, { ...patch, ...extra }, expectedETag);
  };
  const triage = async (id, patch, expectedETag) => {
    const saved = await updateIssue(id, patch, expectedETag);
    setOpenId(null);
    return saved;
  };

  const addIntake = async (form) => {
    setSaveError("");
    const created = await dataService.createIssue({
      ...form,
      reportDate: iso(today()),
      createdBy: userDisplayName,
      createdByEmail: userEmail,
      triaged: false,
      status: undefined,
      transformedInto: undefined,
      taskCreated: "No",
      taskOwner: "",
      ownerBU: form.departmentBU,
      dueDate: "",
      followUp: "",
      rootCause: "",
      correctiveAction: "",
      implementationDate: "",
      effectivenessCheck: "",
      verifiedBy: "",
      verifiedDate: "",
      closedDate: "",
      closedAt: "",
      holdReason: "",
      holdUntil: "",
      ownerUpdate: false,
      ownerUpdateAt: "",
      ownerUpdateText: "",
      additionalComments: form.additionalComments || "",
    });
    setIssues((currentIssues) => [...currentIssues, created]);
    if (created.saveWarning) setSaveWarning({ message: created.saveWarning, issueId: created.id });
    return created.qsNumber;
  };

  const saveSettings = async (next) => {
    const productionSettings = { msFormUrl: next.msFormUrl, flowId: next.flowId, access: [] };
    setSaveError("");
    try {
      await dataService.saveSettings(productionSettings);
      setSettings(productionSettings);
    } catch (error) {
      setSaveError(error instanceof Error ? error.message : String(error));
      throw error;
    }
  };

  const resetDemo = async () => {
    if (developmentMode && dataService.initializeIssues && confirm("Reset the saved demo data?")) {
      try { setIssues(await dataService.initializeIssues(seed(), true)); setOpenId(null); }
      catch (error) { setSaveError(error instanceof Error ? error.message : String(error)); }
    }
  };

  if (!issues) {
    return (
      <div className="qstar-app">
        <div className="flex h-64 items-center justify-center text-sm text-slate-500">
          <RefreshCw size={16} className="mr-2 animate-spin" />Loading issue register…
        </div>
      </div>
    );
  }

  const current = openId != null ? issues.find((issue) => issue.id === openId) : null;
  const pending = issues.filter((issue) => !issue.triaged).length;
  const owner = userDisplayName;

  const PROFILE_META = {
    admin: { label: "Admin", icon: UserCog, access: "Full + IT" },
    qm: { label: "Quality Manager", icon: ShieldCheck, access: "Full (no IT)" },
    owner: { label: "Task Owner", icon: UserCircle2, access: "Assigned issues" },
    reader: { label: "Reader", icon: Eye, access: "Read only" },
  };
  const T_DASH = { id: "dashboard", label: "Dashboard", icon: LayoutDashboard };
  const T_TRIAGE = { id: "triage", label: "Triage queue", icon: Inbox, badge: pending };
  const T_REG = { id: "register", label: "Issue register", icon: ClipboardList };
  const T_REM = { id: "reminders", label: "Reminders", icon: BellRing };
  const T_REPORT = { id: "report", label: "Report", icon: Send };
  const T_MYTASKS = { id: "mytasks", label: "My tasks", icon: ListChecks };
  const T_SETTINGS = { id: "settings", label: "IT settings", icon: Settings };
  const TABS = {
    admin: [T_DASH, T_TRIAGE, T_REG, T_REM, T_REPORT, T_SETTINGS],
    qm: [T_DASH, T_TRIAGE, T_REG, T_REM, T_REPORT],
    owner: [T_MYTASKS, T_DASH, T_REG, T_REM],
    reader: [T_DASH],
  };
  const tabs = TABS[profile];
  const activeProfile = PROFILE_META[profile];
  const defaultTab = profile === "owner" ? "mytasks" : "dashboard";
  const activeTab = tabs.some((item) => item.id === tab) ? tab : defaultTab;

  const reopen = async (id) => {
    const issue = issues.find((candidate) => candidate.id === id);
    if (!issue) return;
    try {
      await updateIssue(id, reopenIssuePatch(issue), issue.eTag);
    } catch (error) { return; }
    try {
      await addProgress(id, {
        ts: new Date().toISOString(), author: userDisplayName, authorEmail: userEmail,
        text: `Issue re-opened — corrective-action cycle restarted. Previous close: ${issue.closedAt ? fmtDateTime(issue.closedAt) : fmtDate(issue.closedDate)}${issue.verifiedBy ? `, effectiveness verified by ${issue.verifiedBy}` : ""}.`,
      });
    } catch (error) {
      setSaveError("");
      setSaveWarning({ message: "Issue was re-opened, but its journal note could not be saved. Add the note from the progress log.", issueId: id });
    }
  };

  const back = () => setOpenId(null);
  const ownsCurrent = current && (
    current.taskOwnerEmail ? current.taskOwnerEmail.toLowerCase() === userEmail.toLowerCase() :
    current.taskOwner === userDisplayName
  );
  let body;
  if (current) {
    const isClosed = current.status === "Closed";
    if (profile === "reader") {
      body = <ReadOnlyIssueDetail issue={current} onBack={back} />;
    } else if (profile === "owner") {
      body = (!isClosed && ownsCurrent)
        ? <OwnerIssueDetail issue={current} owner={owner} onBack={back} onUpdate={ownerUpdateTask} onReload={reloadIssue} onAddProgress={ownerAddProgress} />
        : <ReadOnlyIssueDetail issue={current} onBack={back} />;
    } else if (isClosed) {
      body = <ReadOnlyIssueDetail issue={current} onBack={back} onReopen={() => reopen(current.id)} />;
    } else {
      body = current.triaged
        ? <QMIssueDetail issue={current} issueBusy={busyIssueIds.includes(current.id)} onBack={back} onUpdate={updateIssue} onAddProgress={addProgress} onReload={reloadIssue} author={userDisplayName} />
        : <TriageForm issue={current} onBack={back} onTriage={triage} onReload={reloadIssue} />;
    }
  } else if (activeTab === "dashboard") body = <Dashboard issues={issues} />;
  else if (activeTab === "triage") body = <TriageQueue issues={issues} onOpen={setOpenId} />;
  else if (activeTab === "register") body = <Register issues={issues} onOpen={setOpenId} filter={regFilter} setFilter={setRegFilter} />;
  else if (activeTab === "reminders") body = <RemindersView issues={issues} />;
  else if (activeTab === "report") body = <ReporterForm onSubmit={addIntake} settings={settings} reporterName={userDisplayName} />;
  else if (activeTab === "mytasks") body = <OwnerTasks issues={issues} owner={owner} ownerEmailAddress={userEmail} onOpen={setOpenId} />;
  else if (activeTab === "settings") body = <SettingsView connection={connection} settings={settings} onSave={saveSettings} onRunDiagnostics={onRunDiagnostics} />;

  return (
    <div className="qstar-app">
      <div className="min-h-screen bg-slate-50 text-slate-900" style={{ fontFamily: "ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial" }}>
        <style>{`
          .qstar-app .tm-text-navy { color: ${BRAND.navy}; }
          .qstar-app .tm-btn-primary { background: ${BRAND.navy}; }
          .qstar-app .tm-btn-primary:hover { background: ${BRAND.navyHover}; }
          .qstar-app .tm-btn-primary:disabled { background: #cbd5e1; }
          .qstar-app .tm-input:focus { border-color: ${BRAND.navy}; box-shadow: 0 0 0 3px rgba(27,32,92,.15); }
        `}</style>
        <div className="h-1 w-full" style={{ backgroundColor: BRAND.yellow }} />
        <header className="sticky top-0 z-10 border-b border-slate-200 bg-white/90 backdrop-blur">
          <div className="mx-auto flex max-w-7xl flex-wrap items-center justify-between gap-3 px-4 py-3">
            <div className="flex items-center gap-2.5">
              <div className="relative inline-flex h-9 w-9 items-center justify-center" aria-label="Q-Star logo">
                <span className="font-black leading-none" style={{ color: BRAND.navy, fontSize: 30 }}>Q</span>
                <Star size={15} strokeWidth={0} fill={BRAND.yellow} style={{ position: "absolute", top: 0, right: 1 }} />
              </div>
              <div className="text-xs font-medium text-slate-500">Q-Star Issue Manager · ISO 9001:2015</div>
            </div>
            <div className="flex items-center gap-2">
              <div className="hidden text-right text-xs text-slate-500 sm:block">
                <div>{userDisplayName}</div><div>{userEmail}</div>
              </div>
              <div className="flex items-center gap-1.5 rounded-lg px-2.5 py-1 text-xs font-semibold" style={{ backgroundColor: BRAND.navySoft, color: BRAND.navy }}>
                <KeyRound size={12} />{activeProfile.label} · {activeProfile.access}
              </div>
              {developmentMode && <button onClick={resetDemo} title="Reset demo data" className="rounded-lg p-2 text-slate-500 hover:bg-slate-100"><RefreshCw size={15} /></button>}
            </div>
          </div>
          {!current && (
            <div className="border-t border-slate-100 bg-white">
              <div className="mx-auto flex max-w-7xl flex-wrap gap-1 px-3">
                {tabs.map((item) => (
                  <button key={item.id} onClick={() => setTab(item.id)}
                    style={activeTab === item.id ? { borderBottomColor: BRAND.yellow, color: BRAND.navy } : undefined}
                    className={`flex items-center gap-1.5 border-b-2 px-3 py-2.5 text-sm font-medium transition ${activeTab === item.id ? "" : "border-transparent text-slate-500 hover:text-slate-800"}`}>
                    <item.icon size={15} />{item.label}
                    {item.badge ? <span className="ml-0.5 rounded-full px-1.5 text-xs font-bold" style={{ backgroundColor: BRAND.yellow, color: BRAND.navy }}>{item.badge}</span> : null}
                  </button>
                ))}
              </div>
            </div>
          )}
        </header>

        <main className="mx-auto max-w-7xl px-4 py-5">
          {loadError && <div className="mb-4 rounded-lg bg-rose-50 p-3 text-sm text-rose-800">Could not load SharePoint data: {loadError} <button className="ml-2 underline" onClick={() => setReloadToken((value) => value + 1)}>Retry</button></div>}
          {saveError && <div role="alert" className="mb-4 rounded-lg bg-rose-50 p-3 text-sm text-rose-800">A SharePoint change failed and was not saved: {saveError}</div>}
          {saveWarning && <div role="status" className="mb-4 rounded-lg bg-amber-50 p-3 text-sm text-amber-800">
            Saved with a warning: {saveWarning.message}
            {(!current || current.id !== saveWarning.issueId || current.eTag || current.status === "Closed") && <button disabled={reloadingSaved} className="ml-2 underline" onClick={reloadSavedIssue}>{reloadingSaved ? "Reloading…" : "Reload saved issue"}</button>}
          </div>}
          {activeTab === "report" && !current && <div className="mb-4 text-sm text-slate-500">Submit a potential quality issue. It enters the Quality Team's triage queue for assessment.</div>}
          {body}
        </main>

        <footer className="mx-auto max-w-7xl px-4 pb-8 pt-2 text-center text-xs text-slate-500">
          Production SPFx build · authenticated by Microsoft 365 · data stored in the Q-Star SharePoint site.
        </footer>
      </div>
    </div>
  );
}
