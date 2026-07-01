import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  LayoutDashboard, CalendarDays, ClipboardList, Lightbulb, Telescope, Megaphone,
  Route, Trophy, Building2, NotebookPen, FileText, Archive, Settings as SettingsIcon,
  Plus, Pencil, Trash2, CheckCircle2, RotateCcw, Download, Upload, Filter, X,
  ChevronLeft, ChevronRight, Menu, Database, Search, ArrowUpDown, Inbox, ExternalLink, Clock,
  Sun, Moon, Share2, Image as ImageIcon, Link2, Maximize2, MessageSquare, BookOpen, Sparkles,
  Coffee, Cookie, Beef, Flame, Utensils, CupSoda,
} from "lucide-react";
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, Cell } from "recharts";
import Papa from "papaparse";
import * as XLSX from "xlsx";

/* ------------------------------------------------------------------ */
/* Defaults                                                            */
/* ------------------------------------------------------------------ */
const DEFAULT_SETTINGS = {
  brands: [
    "Meat at Tony's", "Trapani", "Masgouf", "Okra", "Melenzane",
    "Recco", "The Grove", "The Social Table", "Secco", "Faye", "Cafeteria Al Dhahiya",
  ],
  priorities: ["High", "Medium", "Low"],
  statuses: ["Not Started", "In Progress", "Waiting", "Completed", "On Hold", "Cancelled", "Saved for Later", "Backlog", "Archived"],
  recStatuses: ["Draft", "Shared", "Under Review", "Approved", "Implemented", "Rejected", "On Hold", "Backlog", "Archived"],
  categories: ["Brand", "Marketing", "Commercial", "Menu", "Product", "CRM", "Customer Experience", "Delivery", "Operations", "Research", "Meeting", "Campaign", "Follow-up", "Pricing", "Promotion", "Competitor Research", "Review Analysis", "Product Development", "Presentation", "Report"],
  platforms: ["Talabat", "Deliveroo", "Jahez", "Website", "Instagram", "TikTok", "Snapchat", "Google", "Meta Ads", "WhatsApp", "Dine-in", "Pickup", "Delivery", "Catering", "Events", "Offline", "Other"],
  channels: ["Delivery App", "Social Media", "Paid Media", "CRM", "Website", "In-store", "PR", "Partnership", "Influencer", "Corporate Sales", "Catering", "Other"],
  campaignTypes: ["Awareness", "Sales", "Retention", "Launch", "Offer", "Seasonal", "Product Push", "Brand Building", "Customer Experience", "Other"],
  impactTypes: ["Revenue", "Profitability", "Customer Experience", "Retention", "Brand Awareness", "Operational Efficiency", "Product Innovation", "Menu Improvement", "Delivery Improvement", "Cost Control", "Other"],
  importanceLevels: ["Critical", "High", "Medium", "Low"],
  eventTypes: ["Task", "Meeting", "Follow-up", "Brand Review", "Campaign", "Deadline", "Report", "Presentation", "Quarterly Initiative"],
  brandStatuses: ["Active", "Paused", "Watch", "Inactive"],
};

const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const MONTHS_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
const QUARTERS = ["Q1", "Q2", "Q3", "Q4"];

/* ------------------------------------------------------------------ */
/* Collection configs                                                  */
/* field: n=name l=label t=type o=optionsKey/array col=show in table   */
/* ------------------------------------------------------------------ */
const COLLECTIONS = {
  work: {
    key: "seazen_work", label: "Work Log", icon: ClipboardList, primaryDate: "date", title: "project",
    restoreStatus: "In Progress",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "time", l: "Time", t: "time", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "project", l: "Project", t: "text", col: true },
      { n: "category", l: "Category", t: "select", o: "categories", col: true },
      { n: "priority", l: "Priority", t: "select", o: "priorities", col: true },
      { n: "status", l: "Status", t: "select", o: "statuses", col: true },
      { n: "owner", l: "Owner", t: "text" },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "outcome", l: "Outcome", t: "textarea" },
      { n: "linkedEvent", l: "Linked Calendar Event", t: "text" },
    ],
  },
  ideas: {
    key: "seazen_ideas", label: "Ideas Bank", icon: Lightbulb, primaryDate: "date", title: "idea",
    restoreStatus: "In Progress",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "idea", l: "Idea", t: "text", col: true },
      { n: "category", l: "Category", t: "select", o: "categories", col: true },
      { n: "impact", l: "Impact", t: "select", o: "impactTypes", col: true },
      { n: "priority", l: "Priority", t: "select", o: "priorities", col: true },
      { n: "status", l: "Status", t: "select", o: "statuses", col: true },
      { n: "owner", l: "Owner", t: "text" },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "nextStep", l: "Next Step", t: "textarea" },
      { n: "reviewDate", l: "Review Date", t: "date" },
    ],
  },
  findings: {
    key: "seazen_findings", label: "Findings & Insights", icon: Telescope, primaryDate: "date", title: "finding",
    restoreStatus: "In Progress",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "finding", l: "Finding", t: "text", col: true },
      { n: "source", l: "Source", t: "text", col: true },
      { n: "category", l: "Category", t: "select", o: "categories", col: true },
      { n: "importance", l: "Importance", t: "select", o: "importanceLevels", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "relatedLink", l: "Related File or Link", t: "text" },
    ],
  },
  recommendations: {
    key: "seazen_recommendations", label: "Recommendations", icon: Megaphone, primaryDate: "date", title: "recommendation",
    statusField: "status", statusOpt: "recStatuses", restoreStatus: "Shared",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "recommendation", l: "Recommendation", t: "text", col: true },
      { n: "expectedImpact", l: "Expected Impact", t: "select", o: "impactTypes", col: true },
      { n: "priority", l: "Priority", t: "select", o: "priorities", col: true },
      { n: "status", l: "Status", t: "select", o: "recStatuses", col: true },
      { n: "decision", l: "Decision", t: "text" },
      { n: "sharedWith", l: "Shared With", t: "text" },
      { n: "followUpDate", l: "Follow-up Date", t: "date", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
    ],
  },
  wins: {
    key: "seazen_wins", label: "Wins & Achievements", icon: Trophy, primaryDate: "date", title: "achievement",
    restoreStatus: "Completed",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "achievement", l: "Achievement", t: "text", col: true },
      { n: "businessImpact", l: "Business Impact", t: "select", o: "impactTypes", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "relatedProject", l: "Related Project", t: "text" },
    ],
  },
  meetings: {
    key: "seazen_meetings", label: "Meeting Notes", icon: NotebookPen, primaryDate: "date", title: "meetingTitle",
    restoreStatus: "In Progress",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "time", l: "Time", t: "time" },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "meetingTitle", l: "Meeting Title", t: "text", col: true },
      { n: "attendees", l: "Attendees", t: "text", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "decisions", l: "Decisions", t: "textarea" },
      { n: "actionItems", l: "Action Items", t: "textarea" },
      { n: "followUpDate", l: "Follow-up Date", t: "date", col: true },
      { n: "status", l: "Status", t: "select", o: "statuses", col: true },
    ],
  },
  files: {
    key: "seazen_files", label: "Files & Links", icon: FileText, primaryDate: "date", title: "fileName",
    restoreStatus: "In Progress",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "fileName", l: "File Name", t: "text", col: true },
      { n: "link", l: "Link", t: "url", col: true },
      { n: "type", l: "Type", t: "text", col: true },
      { n: "category", l: "Category", t: "select", o: "categories", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
      { n: "relatedProject", l: "Related Project", t: "text" },
    ],
  },
  events: {
    key: "seazen_events", label: "Calendar", icon: CalendarDays, primaryDate: "date", title: "title",
    restoreStatus: "Not Started",
    fields: [
      { n: "date", l: "Date", t: "date", col: true },
      { n: "time", l: "Time", t: "time" },
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "platform", l: "Platform", t: "select", o: "platforms" },
      { n: "channel", l: "Channel", t: "select", o: "channels" },
      { n: "title", l: "Title", t: "text", col: true },
      { n: "eventType", l: "Event Type", t: "select", o: "eventTypes", col: true },
      { n: "category", l: "Category", t: "select", o: "categories" },
      { n: "priority", l: "Priority", t: "select", o: "priorities", col: true },
      { n: "status", l: "Status", t: "select", o: "statuses", col: true },
      { n: "relatedProject", l: "Related Project", t: "text" },
      { n: "followUpRequired", l: "Follow-up Required", t: "checkbox" },
      { n: "followUpDate", l: "Follow-up Date", t: "date" },
      { n: "repeat", l: "Repeat", t: "select", o: ["None", "Daily", "Weekly", "Monthly"] },
      { n: "notes", l: "Notes", t: "textarea" },
    ],
  },
  brandHub: {
    key: "seazen_brandhub", label: "Brand Hub", icon: Building2, primaryDate: "lastReview", title: "brand",
    restoreStatus: "Active",
    fields: [
      { n: "brand", l: "Brand", t: "select", o: "brands", col: true },
      { n: "status", l: "Status", t: "select", o: "brandStatuses", col: true },
      { n: "lastReview", l: "Last Review", t: "date", col: true },
      { n: "monthFocus", l: "Current Month Focus", t: "text", col: true },
      { n: "quarterFocus", l: "Current Quarter Focus", t: "text" },
      { n: "yearFocus", l: "Yearly Focus", t: "text" },
      { n: "activePlatforms", l: "Active Platforms", t: "text" },
      { n: "platformsToImprove", l: "Key Platforms to Improve", t: "text" },
      { n: "keyIssue", l: "Key Issue", t: "textarea" },
      { n: "keyOpportunity", l: "Key Opportunity", t: "textarea" },
      { n: "mainOpportunities", l: "Main Opportunities", t: "textarea" },
      { n: "nextAction", l: "Next Action", t: "text", col: true },
      { n: "priority", l: "Priority", t: "select", o: "priorities", col: true },
      { n: "notes", l: "Notes", t: "textarea" },
    ],
  },
};

const COLLECTION_KEYS = Object.keys(COLLECTIONS);
const SETTINGS_KEY = "seazen_settings";
const ROADMAP_KEY = "seazen_roadmap";

/* Collections shown in the unified Backlog & Archive view */
const AGG_COLLECTIONS = ["work", "ideas", "findings", "recommendations", "wins", "meetings", "files", "events"];

/* One-time seed: Q3 2026 marketing calendar (loads once, never duplicates) */
const SEED_FLAG = "seazen_seed_q3_2026";
const THEME_KEY = "seazen_theme";
const IDEAMAP_KEY = "seazen_ideamap";
const TRENDS_KEY = "seazen_trends";
const INSPO_KEY = "seazen_inspo";
const SEED_TRENDS = [
  { title: "Smash burgers everywhere", tag: "Format", blurb: "Thin patties with crisp browned edges are one of the most ordered items this year. A strong fit for a burger menu and a reliable seller on delivery.", source: "National Restaurant Association", url: "https://restaurant.org/education-and-resources/resource-library/what%E2%80%99s-hot-in-2026-comfort-health-and-value/", keyword: "smashburger,cheeseburger", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4d/Cheeseburger.jpg/640px-Cheeseburger.jpg" },
  { title: "Dubai chocolate", tag: "Dessert", blurb: "Pistachio and kunafa filled chocolate keeps pulling crowds across the Gulf. A ready made limited time dessert or add on for the cafes.", source: "National Restaurant Association", url: "https://restaurant.org/education-and-resources/resource-library/what%E2%80%99s-hot-in-2026-comfort-health-and-value/", keyword: "chocolate,pistachio,dessert", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e6/Dubai_chocolate_on_a_plate_01.jpg/640px-Dubai_chocolate_on_a_plate_01.jpg" },
  { title: "Matcha on every menu", tag: "Beverage", blurb: "Matcha lattes keep growing with younger guests who care how a drink looks and tastes. Works hot or iced and travels well on delivery.", source: "Datassential", url: "https://datassential.com/resource/new-classics-2026-food-beverage-trends/", keyword: "matcha,latte", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4e/Matcha_latte.jpg/640px-Matcha_latte.jpg" },
  { title: "Pistachio on everything", tag: "Flavor", blurb: "Pistachio runs across lattes, ice cream, and pastries right now. An easy flavor to add to desserts and drinks across the brands.", source: "National Restaurant Association", url: "https://restaurant.org/education-and-resources/resource-library/what%E2%80%99s-hot-in-2026-comfort-health-and-value/", keyword: "pistachio,icecream", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/a/ae/Pistachio_ice_cream_with_whipped_mascarpone_and_pistachio_biscotti.jpg/640px-Pistachio_ice_cream_with_whipped_mascarpone_and_pistachio_biscotti.jpg" },
  { title: "Birria and the dip taco", tag: "Global", blurb: "Slow cooked beef birria and the dip style taco went global through social video. A bold limited run idea that films well.", source: "Datassential", url: "https://datassential.com/resource/new-classics-2026-food-beverage-trends/", keyword: "tacos,beef", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a3/Taco_de_birria.jpg/640px-Taco_de_birria.jpg" },
  { title: "Texture led drinks", tag: "Texture", blurb: "Drinks with boba, jelly, and foam win because texture makes them fun to share. A simple way to add a signature non alcoholic drink.", source: "Datassential", url: "https://datassential.com/resource/new-classics-2026-food-beverage-trends/", keyword: "bubbletea,boba", image: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/0c/Popping_boba_w_Bubble_Tea.jpg/640px-Popping_boba_w_Bubble_Tea.jpg" },
];
const MIG_OKRA = "seazen_mig_okra";
const Q3_SEED = [{"date":"2026-07-01","brand":"Trapani","title":"[QAT] New menu full experience campaign with drinks (Vendome Mall)","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Qatar. Remark: not the right time for a new menu launch; move campaign to September. Prepare a summer activation plan for mid-July to August."},{"date":"2026-08-01","brand":"Trapani","title":"[QAT] Brand collab (TBC), delivery focus","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Qatar."},{"date":"2026-09-01","brand":"Trapani","title":"[QAT] New menu launch campaign (Mshreib) + delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Qatar. Runs September into October."},{"date":"2026-06-01","brand":"Melenzane","title":"[UAE] Forno thin crust pizza experience + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane UAE. Runs June into July. Prepare a full campaign presentation for MND approval. Obtain approval for monthly delivery activities (CPC, rankings, banners, with targets). Coordinate with aggregators on best-selling ideas and opportunities."},{"date":"2026-08-01","brand":"Melenzane","title":"[UAE] Home delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane UAE."},{"date":"2026-09-01","brand":"Melenzane","title":"[UAE] New menu campaign + Dubai events participation + delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane UAE. Runs September into October."},{"date":"2026-06-01","brand":"Melenzane","title":"[OMN] Forno thin crust pizza experience + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Oman. Runs June into July. Prepare a full campaign presentation for MND approval. Obtain approval for monthly delivery activities (CPC, rankings, banners, with targets). Coordinate with aggregators on best-selling ideas and opportunities."},{"date":"2026-08-01","brand":"Melenzane","title":"[OMN] Home delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Oman."},{"date":"2026-09-01","brand":"Melenzane","title":"[OMN] New menu campaign + delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Oman. Runs September into October."},{"date":"2026-06-01","brand":"Melenzane","title":"[BAH] Forno thin crust pizza experience + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Bahrain. Runs June into July. Prepare a full campaign presentation for MND approval. Obtain approval for monthly delivery activities (CPC, rankings, banners, with targets). Coordinate with aggregators on best-selling ideas and opportunities."},{"date":"2026-08-01","brand":"Melenzane","title":"[BAH] Home delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Bahrain."},{"date":"2026-09-01","brand":"Melenzane","title":"[BAH] New menu campaign + delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Bahrain. Runs September into October."},{"date":"2026-06-01","brand":"Melenzane","title":"[QAT] Forno thin crust pizza experience + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Qatar. Runs June into July. Prepare a full campaign presentation for MND approval. Obtain approval for monthly delivery activities (CPC, rankings, banners, with targets). Coordinate with aggregators on best-selling ideas and opportunities."},{"date":"2026-08-01","brand":"Melenzane","title":"[QAT] Home delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Qatar."},{"date":"2026-09-01","brand":"Melenzane","title":"[QAT] New menu campaign: MZ experience, cart in the mall + delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Qatar. Runs September into October."},{"date":"2026-07-01","brand":"Secco","title":"[BAH] Takeaway/pickup focus + summer dessert/gelato launch + delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Secco Bahrain. Runs July into August. Focus on takeaway activation (arch and storefront/window branding). Prepare a full activation plan and a summer campaign with a new summer dessert offering."},{"date":"2026-09-01","brand":"Secco","title":"[BAH] Delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Secco Bahrain."},{"date":"2026-10-01","brand":"Secco","title":"[BAH] New menu campaign + delivery activities","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Secco Bahrain."},{"date":"2026-07-01","brand":"Melenzane","title":"[KW] Forno thin crust: Avenues branch positioning content; blogger narrative (Avenues/Kout/Mutla) + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Kuwait."},{"date":"2026-08-01","brand":"Melenzane","title":"[KW] Delivery: Ciabatta","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Kuwait."},{"date":"2026-09-01","brand":"Melenzane","title":"[KW] Facelift","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Melenzane Kuwait."},{"date":"2026-07-01","brand":"The Grove","title":"[KW] Awareness campaign","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"The Grove Kuwait."},{"date":"2026-08-01","brand":"The Grove","title":"[KW] Delivery collab with healthy brand","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"The Grove Kuwait."},{"date":"2026-10-01","brand":"The Grove","title":"[KW] Campaign: Le Pecan delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"The Grove Kuwait."},{"date":"2026-07-01","brand":"The Social Table","title":"[KW] Matcha but protein campaign, home delivery focus + catering","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"The Social Table Kuwait."},{"date":"2026-08-01","brand":"The Social Table","title":"[KW] Glasshouse: catering + home delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"The Social Table Kuwait."},{"date":"2026-07-01","brand":"Secco","title":"[KW] Summer ice cream campaign","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Secco Kuwait. Runs July into August."},{"date":"2026-07-01","brand":"Trapani","title":"[KW] Assima: La Dolce Sip new strategy","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Kuwait."},{"date":"2026-08-01","brand":"Trapani","title":"[KW] La Dolce Sip: Avenues","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Kuwait."},{"date":"2026-09-01","brand":"Trapani","title":"[KW] La Dolce Sip: Muruj","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Trapani Kuwait."},{"date":"2026-07-01","brand":"Masgouf","title":"[KW] Chef Aziz breakfast menu collab (all branches) + summer catering","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Masgouf Kuwait."},{"date":"2026-08-01","brand":"Masgouf","title":"[KW] Chef Aziz lunch menu collab (all branches)","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Masgouf Kuwait."},{"date":"2026-07-01","brand":"Cafeteria Al Dhahiya","title":"[KW] Summer catering + delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Cafeteria Al Dhahiya Kuwait."},{"date":"2026-08-01","brand":"Cafeteria Al Dhahiya","title":"[KW] Delivery + catering","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Cafeteria Al Dhahiya Kuwait."},{"date":"2026-07-01","brand":"Okra","title":"[KW] Okra experience campaign (modern Arabic theme) + delivery","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Okra Kuwait."},{"date":"2026-07-01","brand":"Meat at Tony's","title":"[KW] Elamigos collab + Talabat exclusivity","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Meat at Tony's Kuwait. Runs July into August."},{"date":"2026-09-01","brand":"Meat at Tony's","title":"[KW] Bus activation","eventType":"Campaign","category":"Campaign","priority":"Medium","status":"Not Started","notes":"Meat at Tony's Kuwait."}];

/* ------------------------------------------------------------------ */
/* Storage layer (persistent, survives refresh / restart)              */
/* ------------------------------------------------------------------ */
const mem = {}; // fallback if window.storage is unavailable
const hasStore = () => typeof window !== "undefined" && window.storage;

async function loadKey(key, fallback) {
  if (!hasStore()) return mem[key] !== undefined ? mem[key] : fallback;
  try {
    const r = await window.storage.get(key);
    return r && r.value != null ? JSON.parse(r.value) : fallback;
  } catch {
    return fallback;
  }
}
async function saveKey(key, value) {
  mem[key] = value;
  if (!hasStore()) return;
  try {
    await window.storage.set(key, JSON.stringify(value));
  } catch (e) {
    console.error("save failed", key, e);
  }
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */
const uid = () => "id_" + Date.now().toString(36) + "_" + Math.random().toString(36).slice(2, 8);
const todayKey = () => new Date().toISOString().slice(0, 10);

function parseDate(s) {
  if (!s) return null;
  const d = new Date(s + "T00:00:00");
  return isNaN(d.getTime()) ? null : d;
}
function isoWeek(d) {
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  return Math.ceil(((t - yearStart) / 86400000 + 1) / 7);
}
function derive(dateStr) {
  const d = parseDate(dateStr);
  if (!d) return { month: "", quarter: "", year: "", week: "" };
  return {
    month: MONTHS[d.getMonth()],
    quarter: "Q" + (Math.floor(d.getMonth() / 3) + 1),
    year: String(d.getFullYear()),
    week: "W" + isoWeek(d),
  };
}
const dKey = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;

function currentPeriod() {
  const n = new Date();
  return { month: MONTHS[n.getMonth()], quarter: "Q" + (Math.floor(n.getMonth() / 3) + 1), year: String(n.getFullYear()), week: "W" + isoWeek(n) };
}

const DONE = ["Completed", "Implemented", "Approved", "Cancelled", "Rejected"];
const PARKED = ["Backlog", "Saved for Later"];

const STATUS_STYLES = {
  "Not Started": "bg-gray-100 text-gray-700",
  "In Progress": "bg-blue-100 text-blue-700",
  "Waiting": "bg-amber-100 text-amber-700",
  "Completed": "bg-emerald-100 text-emerald-700",
  "Implemented": "bg-emerald-100 text-emerald-700",
  "Approved": "bg-emerald-100 text-emerald-700",
  "On Hold": "bg-orange-100 text-orange-700",
  "Cancelled": "bg-red-100 text-red-600",
  "Rejected": "bg-red-100 text-red-600",
  "Saved for Later": "bg-violet-100 text-violet-700",
  "Backlog": "bg-slate-200 text-slate-700",
  "Archived": "bg-zinc-200 text-zinc-600",
  "Draft": "bg-gray-100 text-gray-600",
  "Shared": "bg-indigo-100 text-indigo-700",
  "Under Review": "bg-amber-100 text-amber-700",
  "Active": "bg-emerald-100 text-emerald-700",
  "Paused": "bg-amber-100 text-amber-700",
  "Watch": "bg-orange-100 text-orange-700",
  "Inactive": "bg-gray-100 text-gray-500",
};
const PRIORITY_STYLES = {
  High: "bg-red-100 text-red-700", Medium: "bg-amber-100 text-amber-700", Low: "bg-emerald-100 text-emerald-700",
  Critical: "bg-red-100 text-red-700",
};
function chipClass(v) {
  return STATUS_STYLES[v] || PRIORITY_STYLES[v] || "bg-gray-100 text-gray-600";
}

function download(filename, blob) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

/* ------------------------------------------------------------------ */
/* Small UI primitives                                                 */
/* ------------------------------------------------------------------ */
function Chip({ value }) {
  if (!value) return <span className="text-gray-300">-</span>;
  return <span className={`inline-block px-2 py-0.5 rounded-full text-xs font-medium ${chipClass(value)}`}>{value}</span>;
}

function Btn({ children, onClick, variant = "ghost", className = "", type = "button", title }) {
  const styles = {
    primary: "bg-emerald-600 text-white hover:bg-emerald-700",
    soft: "bg-emerald-50 text-emerald-700 hover:bg-emerald-100",
    ghost: "bg-white text-gray-600 border border-gray-200 hover:bg-gray-50",
    danger: "bg-white text-red-600 border border-red-200 hover:bg-red-50",
    plain: "text-gray-500 hover:text-gray-900 hover:bg-gray-100",
  };
  return (
    <button type={type} title={title} onClick={onClick}
      className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium transition-colors ${styles[variant]} ${className}`}>
      {children}
    </button>
  );
}

function Modal({ open, onClose, title, children, wide }) {
  if (!open) return null;
  return (
    <div className="fixed inset-0 z-50 flex items-start justify-center p-3 sm:p-6 bg-black/30 overflow-y-auto" onMouseDown={onClose}>
      <div className={`bg-white rounded-2xl shadow-xl w-full ${wide ? "max-w-4xl" : "max-w-2xl"} my-4`} onMouseDown={(e) => e.stopPropagation()}>
        <div className="flex items-center justify-between px-5 py-3.5 border-b border-gray-100">
          <h3 className="font-semibold text-gray-900">{title}</h3>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700"><X size={20} /></button>
        </div>
        <div className="p-5">{children}</div>
      </div>
    </div>
  );
}

function Confirm({ state, onCancel }) {
  if (!state) return null;
  return (
    <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-black/40" onMouseDown={onCancel}>
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm p-5" onMouseDown={(e) => e.stopPropagation()}>
        <h4 className="font-semibold text-gray-900 mb-1">{state.title}</h4>
        <p className="text-sm text-gray-600 mb-4">{state.message}</p>
        <div className="flex justify-end gap-2">
          <Btn variant="ghost" onClick={onCancel}>Cancel</Btn>
          <Btn variant={state.danger ? "danger" : "primary"} onClick={state.onConfirm}>{state.confirmText || "Confirm"}</Btn>
        </div>
      </div>
    </div>
  );
}

function FieldInput({ field, value, onChange, settings }) {
  const base = "w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200 focus:border-emerald-400";
  if (field.t === "textarea")
    return <textarea rows={3} className={base} value={value || ""} onChange={(e) => onChange(e.target.value)} />;
  if (field.t === "checkbox")
    return (
      <label className="inline-flex items-center gap-2 text-sm text-gray-700">
        <input type="checkbox" checked={!!value} onChange={(e) => onChange(e.target.checked)} className="rounded accent-emerald-600 w-4 h-4" />
        Yes
      </label>
    );
  if (field.t === "select") {
    const opts = Array.isArray(field.o) ? field.o : settings[field.o] || [];
    return (
      <select className={base} value={value || ""} onChange={(e) => onChange(e.target.value)}>
        <option value="">Select...</option>
        {opts.map((o) => <option key={o} value={o}>{o}</option>)}
      </select>
    );
  }
  const type = field.t === "url" ? "text" : field.t;
  return <input type={type} className={base} value={value || ""} onChange={(e) => onChange(e.target.value)} />;
}

function RecordForm({ config, record, settings, onSave, onCancel }) {
  const [form, setForm] = useState(() => {
    const init = {};
    config.fields.forEach((f) => (init[f.n] = record ? record[f.n] : ""));
    if (!record && config.fields.some((f) => f.n === "date")) init.date = todayKey();
    return init;
  });
  const set = (k, v) => setForm((p) => ({ ...p, [k]: v }));
  const d = derive(form[config.primaryDate]);
  return (
    <div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-3">
        {config.fields.map((f) => (
          <div key={f.n} className={f.t === "textarea" ? "sm:col-span-2" : ""}>
            <label className="block text-xs font-medium text-gray-500 mb-1">{f.l}</label>
            <FieldInput field={f} value={form[f.n]} onChange={(v) => set(f.n, v)} settings={settings} />
          </div>
        ))}
      </div>
      {d.year && (
        <div className="mt-4 flex flex-wrap gap-2 text-xs text-gray-500">
          <span className="px-2 py-1 bg-gray-50 rounded">Auto: {d.month}</span>
          <span className="px-2 py-1 bg-gray-50 rounded">{d.quarter}</span>
          <span className="px-2 py-1 bg-gray-50 rounded">{d.year}</span>
          <span className="px-2 py-1 bg-gray-50 rounded">{d.week}</span>
        </div>
      )}
      <div className="flex justify-end gap-2 mt-5">
        <Btn variant="ghost" onClick={onCancel}>Cancel</Btn>
        <Btn variant="primary" onClick={() => onSave(form)}><CheckCircle2 size={16} />Save</Btn>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Filter bar                                                          */
/* ------------------------------------------------------------------ */
function FilterBar({ config, settings, filters, setFilters, extra }) {
  const fieldNames = config.fields.map((f) => f.n);
  const has = (n) => fieldNames.includes(n);
  const sel = "border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white";
  const opt = (key) => (Array.isArray(key) ? key : settings[key] || []);
  const set = (k, v) => setFilters((p) => ({ ...p, [k]: v }));
  const dropdowns = [
    has("brand") && ["brand", "Brand", settings.brands],
    (has("status") || config.statusField) && ["status", "Status", config.statusOpt ? settings[config.statusOpt] : settings.statuses],
    has("priority") && ["priority", "Priority", settings.priorities],
    has("category") && ["category", "Category", settings.categories],
    has("platform") && ["platform", "Platform", settings.platforms],
    has("channel") && ["channel", "Channel", settings.channels],
  ].filter(Boolean);
  return (
    <div className="flex flex-wrap items-center gap-2 mb-3">
      <div className="relative">
        <Search size={15} className="absolute left-2.5 top-2.5 text-gray-400" />
        <input placeholder="Search..." value={filters.q || ""} onChange={(e) => set("q", e.target.value)}
          className="pl-8 pr-3 py-1.5 border border-gray-200 rounded-lg text-sm w-44 focus:outline-none focus:ring-2 focus:ring-emerald-200" />
      </div>
      {dropdowns.map(([k, label, options]) => (
        <select key={k} className={sel} value={filters[k] || ""} onChange={(e) => set(k, e.target.value)}>
          <option value="">{label}: All</option>
          {(options || []).map((o) => <option key={o} value={o}>{o}</option>)}
        </select>
      ))}
      <select className={sel} value={filters.quarter || ""} onChange={(e) => set("quarter", e.target.value)}>
        <option value="">Quarter: All</option>{QUARTERS.map((q) => <option key={q} value={q}>{q}</option>)}
      </select>
      <select className={sel} value={filters.year || ""} onChange={(e) => set("year", e.target.value)}>
        <option value="">Year: All</option>{["2024", "2025", "2026", "2027"].map((y) => <option key={y} value={y}>{y}</option>)}
      </select>
      <input type="date" className={sel} value={filters.from || ""} onChange={(e) => set("from", e.target.value)} title="From date" />
      <input type="date" className={sel} value={filters.to || ""} onChange={(e) => set("to", e.target.value)} title="To date" />
      {Object.values(filters).some(Boolean) && (
        <Btn variant="plain" onClick={() => setFilters({})}><X size={14} />Clear</Btn>
      )}
      {extra}
    </div>
  );
}

function applyFilters(rows, config, filters) {
  const q = (filters.q || "").toLowerCase().trim();
  return rows.filter((r) => {
    if (filters.brand && r.brand !== filters.brand) return false;
    if (filters.status && r.status !== filters.status) return false;
    if (filters.priority && r.priority !== filters.priority) return false;
    if (filters.category && r.category !== filters.category) return false;
    if (filters.platform && r.platform !== filters.platform) return false;
    if (filters.channel && r.channel !== filters.channel) return false;
    if (filters.quarter && r.quarter !== filters.quarter) return false;
    if (filters.year && r.year !== filters.year) return false;
    if (filters.from && (r[config.primaryDate] || "") < filters.from) return false;
    if (filters.to && (r[config.primaryDate] || "") > filters.to) return false;
    if (q) {
      const hay = config.fields.map((f) => r[f.n]).join(" ").toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

/* ------------------------------------------------------------------ */
/* Generic collection page                                             */
/* ------------------------------------------------------------------ */
function CollectionPage({ ckey, data, settings, api, toast, confirm }) {
  const config = COLLECTIONS[ckey];
  const [filters, setFilters] = useState({});
  const [showTrash, setShowTrash] = useState(false);
  const [editing, setEditing] = useState(null); // record or "new"
  const [sort, setSort] = useState({ k: config.primaryDate, dir: "desc" });
  const fileRef = useRef(null);

  const rows = useMemo(() => {
    let r = (data[ckey] || []).filter((x) => !!x.deleted === showTrash);
    r = applyFilters(r, config, filters);
    r = [...r].sort((a, b) => {
      const av = (a[sort.k] || "").toString(), bv = (b[sort.k] || "").toString();
      return sort.dir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    });
    return r;
  }, [data, ckey, filters, sort, showTrash]);

  const cols = config.fields.filter((f) => f.col);

  const save = (form) => {
    if (editing === "new") api.add(ckey, form);
    else api.update(ckey, editing.id, form);
    setEditing(null);
    toast("Saved");
  };
  const quickStatus = (rec, status) => { api.update(ckey, rec.id, { ...rec, status, ...(status === "Backlog" ? { movedToBacklog: todayKey() } : {}) }); toast(status === "Backlog" ? "Moved to backlog" : status === "Archived" ? "Archived" : "Updated"); };
  const softDelete = (rec) => { api.softDelete(ckey, rec.id); toast("Moved to trash"); };
  const restore = (rec) => { api.restore(ckey, rec.id); toast("Restored"); };
  const permaDelete = (rec) => {
    confirm({
      title: "Delete permanently?", message: "This item will be removed. You will be asked once more.", danger: true, confirmText: "Continue",
      onConfirm: () => confirm({
        title: "Are you absolutely sure?", message: "This cannot be undone.", danger: true, confirmText: "Delete forever",
        onConfirm: () => { api.hardDelete(ckey, rec.id); toast("Deleted permanently"); },
      }),
    });
  };

  const exportRows = (fmt) => exportData(`${config.label}`, rows, config, fmt);
  const onImport = (e) => {
    const f = e.target.files?.[0];
    if (f) importData(f, config, (recs) => { api.bulkAdd(ckey, recs); toast(`Imported ${recs.length} rows`); });
    e.target.value = "";
  };

  return (
    <div>
      <PageHead title={config.label} count={rows.length} onAdd={() => setEditing("new")}>
        <Btn variant="ghost" onClick={() => fileRef.current?.click()}><Upload size={15} />Import</Btn>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={onImport} />
        <Btn variant="ghost" onClick={() => exportRows("csv")}><Download size={15} />CSV</Btn>
        <Btn variant="ghost" onClick={() => exportRows("xlsx")}><Download size={15} />Excel</Btn>
      </PageHead>

      <FilterBar config={config} settings={settings} filters={filters} setFilters={setFilters}
        extra={<Btn variant={showTrash ? "soft" : "plain"} onClick={() => setShowTrash((s) => !s)}><Trash2 size={14} />{showTrash ? "Viewing trash" : "Trash"}</Btn>} />

      <div className="bg-white border border-gray-200 rounded-xl overflow-x-auto">
        <table className="w-full text-sm">
          <thead>
            <tr className="text-left text-gray-500 border-b border-gray-100 bg-gray-50/60">
              {cols.map((c) => (
                <th key={c.n} className="px-3 py-2.5 font-medium whitespace-nowrap cursor-pointer select-none"
                  onClick={() => setSort((s) => ({ k: c.n, dir: s.k === c.n && s.dir === "asc" ? "desc" : "asc" }))}>
                  <span className="inline-flex items-center gap-1">{c.l}<ArrowUpDown size={11} className="text-gray-300" /></span>
                </th>
              ))}
              <th className="px-3 py-2.5 font-medium text-right">Actions</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && (
              <tr><td colSpan={cols.length + 1} className="px-3 py-10 text-center text-gray-400">
                {showTrash ? "Trash is empty." : "Nothing here yet. Add your first entry."}
              </td></tr>
            )}
            {rows.map((r) => (
              <tr key={r.id} className="border-b border-gray-50 hover:bg-gray-50/60">
                {cols.map((c) => (
                  <td key={c.n} className="px-3 py-2.5 align-top">{renderCell(r, c)}</td>
                ))}
                <td className="px-3 py-2.5 text-right whitespace-nowrap">
                  <div className="inline-flex items-center gap-0.5">
                    {showTrash ? (
                      <>
                        <IconBtn title="Restore" onClick={() => restore(r)}><RotateCcw size={15} /></IconBtn>
                        <IconBtn title="Delete forever" danger onClick={() => permaDelete(r)}><Trash2 size={15} /></IconBtn>
                      </>
                    ) : (
                      <>
                        <IconBtn title="Edit" onClick={() => setEditing(r)}><Pencil size={15} /></IconBtn>
                        <IconBtn title="Mark complete" onClick={() => quickStatus(r, "Completed")}><CheckCircle2 size={15} /></IconBtn>
                        <IconBtn title="Move to backlog" onClick={() => quickStatus(r, "Backlog")}><Inbox size={15} /></IconBtn>
                        <IconBtn title="Archive" onClick={() => quickStatus(r, "Archived")}><Archive size={15} /></IconBtn>
                        <IconBtn title="Move to trash" danger onClick={() => softDelete(r)}><Trash2 size={15} /></IconBtn>
                      </>
                    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <Modal open={!!editing} onClose={() => setEditing(null)} wide
        title={editing === "new" ? `New ${config.label.replace(/s$/, "")}` : `Edit ${config.label.replace(/s$/, "")}`}>
        {editing && <RecordForm config={config} record={editing === "new" ? null : editing} settings={settings} onSave={save} onCancel={() => setEditing(null)} />}
      </Modal>
    </div>
  );
}

function renderCell(r, c) {
  const v = r[c.n];
  if (!v) return <span className="text-gray-300">-</span>;
  if (c.t === "select" && (STATUS_STYLES[v] || PRIORITY_STYLES[v])) return <Chip value={v} />;
  if (c.t === "url" || c.n === "link")
    return <a href={v} target="_blank" rel="noreferrer" className="text-emerald-700 hover:underline inline-flex items-center gap-1">Open<ExternalLink size={12} /></a>;
  if (c.n === r.__titleField || ["project", "idea", "finding", "recommendation", "achievement", "meetingTitle", "fileName", "title", "brand"].includes(c.n))
    return <span className="font-medium text-gray-900">{v}</span>;
  return <span className="text-gray-600">{String(v).length > 60 ? String(v).slice(0, 60) + "..." : v}</span>;
}

function IconBtn({ children, onClick, title, danger }) {
  return (
    <button title={title} onClick={onClick}
      className={`p-1.5 rounded-lg ${danger ? "text-gray-400 hover:text-red-600 hover:bg-red-50" : "text-gray-400 hover:text-emerald-700 hover:bg-emerald-50"}`}>
      {children}
    </button>
  );
}

function PageHead({ title, count, onAdd, children }) {
  return (
    <div className="flex flex-wrap items-center justify-between gap-3 mb-4">
      <div className="flex items-center gap-2.5">
        <h2 className="text-lg font-semibold text-gray-900">{title}</h2>
        {count != null && <span className="text-xs font-medium px-2 py-0.5 bg-gray-100 text-gray-500 rounded-full">{count}</span>}
      </div>
      <div className="flex flex-wrap items-center gap-2">
        {children}
        {onAdd && <Btn variant="primary" onClick={onAdd}><Plus size={16} />Add</Btn>}
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Import / export                                                     */
/* ------------------------------------------------------------------ */
function recordToRow(r, config) {
  const row = {};
  config.fields.forEach((f) => (row[f.l] = r[f.n] ?? ""));
  row["Month"] = r.month || ""; row["Quarter"] = r.quarter || ""; row["Year"] = r.year || ""; row["Week"] = r.week || "";
  row["Status Flag"] = r.deleted ? "Trash" : "Active";
  row["Created Date"] = r.createdDate || ""; row["Last Updated"] = r.updatedDate || "";
  return row;
}
function exportData(name, rows, config, fmt) {
  const data = rows.map((r) => recordToRow(r, config));
  if (fmt === "csv") {
    download(`${name}.csv`, new Blob([Papa.unparse(data)], { type: "text/csv" }));
  } else {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data), name.slice(0, 28) || "Sheet1");
    download(`${name}.xlsx`, new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })], { type: "application/octet-stream" }));
  }
}
function importData(file, config, cb) {
  const labelToName = {};
  config.fields.forEach((f) => (labelToName[f.l.toLowerCase()] = f.n));
  const mapRow = (raw) => {
    const rec = {};
    Object.keys(raw).forEach((k) => {
      const name = labelToName[k.toLowerCase().trim()];
      if (name) rec[name] = raw[k];
    });
    return rec;
  };
  const ext = file.name.split(".").pop().toLowerCase();
  if (ext === "csv") {
    Papa.parse(file, { header: true, skipEmptyLines: true, complete: (res) => cb(res.data.map(mapRow).filter((r) => Object.keys(r).length)) });
  } else {
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const json = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
      cb(json.map(mapRow).filter((r) => Object.keys(r).length));
    };
    reader.readAsArrayBuffer(file);
  }
}

/* ------------------------------------------------------------------ */
/* Calendar                                                            */
/* ------------------------------------------------------------------ */
function expandOccurrences(events, start, end) {
  const out = [];
  events.forEach((ev) => {
    const base = parseDate(ev.date);
    if (!base) return;
    const rep = ev.repeat;
    if (!rep || rep === "None") {
      if (base >= start && base <= end) out.push({ ...ev, _date: dKey(base) });
      return;
    }
    let cur = new Date(base), guard = 0;
    while (cur <= end && guard < 500) {
      if (cur >= start) out.push({ ...ev, _date: dKey(cur) });
      if (rep === "Daily") cur.setDate(cur.getDate() + 1);
      else if (rep === "Weekly") cur.setDate(cur.getDate() + 7);
      else cur.setMonth(cur.getMonth() + 1);
      guard++;
    }
  });
  return out;
}

const BRAND_PALETTE = ["#059669", "#2563eb", "#7c3aed", "#d97706", "#e11d48", "#0891b2", "#ea580c", "#4f46e5", "#0d9488", "#db2777", "#65a30d", "#0ea5e9", "#9333ea", "#475569"];
const STATUS_HEX = { "Completed": "#059669", "Implemented": "#059669", "Approved": "#059669", "In Progress": "#2563eb", "Shared": "#4f46e5", "Under Review": "#d97706", "Waiting": "#d97706", "On Hold": "#ea580c", "Not Started": "#94a3b8", "Cancelled": "#dc2626", "Rejected": "#dc2626", "Backlog": "#64748b", "Saved for Later": "#7c3aed", "Archived": "#a1a1aa" };
function hexToRgba(hex, a) {
  if (typeof hex !== "string" || hex[0] !== "#" || hex.length < 7) return `rgba(100,116,139,${a})`;
  const h = hex.replace("#", "");
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  return `rgba(${r},${g},${b},${a})`;
}
function eventColor(ev, colorBy, settings) {
  if (colorBy === "priority") return ev.priority === "High" ? "#dc2626" : ev.priority === "Low" ? "#059669" : "#d97706";
  if (colorBy === "status") return STATUS_HEX[ev.status] || "#94a3b8";
  const list = colorBy === "eventType" ? settings.eventTypes : settings.brands;
  const key = colorBy === "eventType" ? ev.eventType : ev.brand;
  const i = list.indexOf(key);
  return BRAND_PALETTE[(i < 0 ? 0 : i) % BRAND_PALETTE.length];
}

const DARK_CSS = `
.theme-dark{background:#0b1220}
.theme-dark .bg-white{background:#131c30 !important}
.theme-dark .bg-gray-50{background:#0e1626 !important}
.theme-dark .bg-gray-100{background:#1c2740 !important}
.theme-dark .bg-gray-50\\/40{background:rgba(20,28,46,.4)!important}
.theme-dark .bg-gray-50\\/60{background:rgba(20,28,46,.6)!important}
.theme-dark .bg-gray-50\\/80{background:rgba(20,28,46,.8)!important}
.theme-dark .border-gray-200{border-color:#2a3650 !important}
.theme-dark .border-gray-100{border-color:#1e293b !important}
.theme-dark .border-gray-50{border-color:#18223a !important}
.theme-dark .border-gray-300{border-color:#374563 !important}
.theme-dark .text-gray-900{color:#e7ebf3 !important}
.theme-dark .text-gray-800{color:#d4dae6 !important}
.theme-dark .text-gray-700{color:#c4ccda !important}
.theme-dark .text-gray-600{color:#9aa6b8 !important}
.theme-dark .text-gray-500{color:#8a96a8 !important}
.theme-dark .text-gray-400{color:#6b7890 !important}
.theme-dark .text-gray-300{color:#4a5468 !important}
.theme-dark .hover\\:bg-gray-100:hover{background:#243150 !important}
.theme-dark .hover\\:bg-gray-50:hover{background:#1a2336 !important}
.theme-dark .hover\\:bg-gray-50\\/60:hover{background:rgba(30,41,59,.55)!important}
.theme-dark .hover\\:bg-gray-50\\/80:hover{background:rgba(30,41,59,.7)!important}
.theme-dark .bg-emerald-50{background:rgba(16,185,129,.14)!important}
.theme-dark .bg-emerald-50\\/50{background:rgba(16,185,129,.10)!important}
.theme-dark .text-emerald-700{color:#34d399 !important}
.theme-dark input,.theme-dark select,.theme-dark textarea{background:#0e1626 !important;color:#e7ebf3 !important;border-color:#2a3650 !important}
.theme-dark input::placeholder,.theme-dark textarea::placeholder{color:#5b6678 !important}
`;

const CAL_CSS = `
@keyframes calFade{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:none}}
.cal-fade{animation:calFade .18s ease-out}
@keyframes calSlide{from{transform:translateX(18px);opacity:.4}to{transform:none;opacity:1}}
.cal-slide{animation:calSlide .22s cubic-bezier(.2,.7,.3,1)}
.cal-scroll::-webkit-scrollbar{width:8px}
.cal-scroll::-webkit-scrollbar-thumb{background:#e5e7eb;border-radius:8px}
`;

function CalEventCard({ ev, color, onEdit, onComplete, onDelete }) {
  return (
    <div onClick={onEdit} style={{ borderLeft: `3px solid ${color}` }}
      className="group bg-white border border-gray-200 rounded-xl p-2.5 mb-2 cursor-pointer transition hover:shadow-sm hover:border-gray-300">
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <div className="flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full shrink-0" style={{ background: color }} />
            <span className="text-sm font-medium text-gray-900 truncate">{ev.title}</span>
          </div>
          <div className="flex flex-wrap items-center gap-2 mt-1 text-xs text-gray-400">
            {ev.time && <span className="inline-flex items-center gap-0.5"><Clock size={11} />{ev.time}</span>}
            {ev.brand && <span>{ev.brand}</span>}
          </div>
        </div>
        <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition">
          <IconBtn title="Mark complete" onClick={(e) => { e.stopPropagation(); onComplete(); }}><CheckCircle2 size={14} /></IconBtn>
          <IconBtn title="Move to trash" danger onClick={(e) => { e.stopPropagation(); onDelete(); }}><Trash2 size={14} /></IconBtn>
        </div>
      </div>
      {(ev.eventType || ev.status) && (
        <div className="flex flex-wrap gap-1.5 mt-1.5">
          {ev.eventType && <span className="text-xs px-1.5 py-0.5 rounded bg-gray-100 text-gray-500">{ev.eventType}</span>}
          {ev.status && <Chip value={ev.status} />}
        </div>
      )}
    </div>
  );
}

function QuickAdd({ dateKey, settings, onAdd, onFull }) {
  const [title, setTitle] = useState("");
  const [brand, setBrand] = useState("");
  const [type, setType] = useState("Task");
  const [time, setTime] = useState("");
  const submit = () => { if (!title.trim()) return; onAdd({ title: title.trim(), brand, eventType: type, time }); setTitle(""); setTime(""); };
  const inp = "w-full border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200";
  return (
    <div className="bg-gray-50 border border-gray-100 rounded-xl p-3 mb-4">
      <input className={inp + " mb-2"} placeholder="Add an event on this day" value={title}
        onChange={(e) => setTitle(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} autoFocus />
      <div className="grid grid-cols-2 gap-2 mb-2">
        <select className={inp} value={brand} onChange={(e) => setBrand(e.target.value)}>
          <option value="">Brand</option>{settings.brands.map((b) => <option key={b} value={b}>{b}</option>)}
        </select>
        <select className={inp} value={type} onChange={(e) => setType(e.target.value)}>
          {settings.eventTypes.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <input type="time" className={inp} value={time} onChange={(e) => setTime(e.target.value)} />
        <Btn variant="primary" onClick={submit}><Plus size={15} />Add</Btn>
      </div>
      <button onClick={onFull} className="text-xs text-emerald-700 hover:underline">Open full form for more fields</button>
    </div>
  );
}

function DayPanel({ date, occ, settings, colorBy, onClose, onAdd, onFull, onEdit, onComplete, onDelete }) {
  const list = [...occ].sort((a, b) => (a.time || "~").localeCompare(b.time || "~"));
  return (
    <div className="fixed inset-0 z-40 flex justify-end" onMouseDown={onClose}>
      <div className="flex-1 bg-black/20" />
      <div className="w-full max-w-sm h-full bg-white shadow-2xl p-5 overflow-y-auto cal-slide cal-scroll" onMouseDown={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-4">
          <div>
            <div className="text-xs font-medium text-emerald-700 uppercase tracking-wide">{date.toLocaleDateString("en-US", { weekday: "long" })}</div>
            <div className="text-2xl font-semibold text-gray-900">{date.toLocaleDateString("en-US", { month: "long", day: "numeric" })}</div>
            <div className="text-sm text-gray-400">{date.getFullYear()}</div>
          </div>
          <button onClick={onClose} className="text-gray-400 hover:text-gray-700 p-1"><X size={20} /></button>
        </div>
        <QuickAdd dateKey={dKey(date)} settings={settings} onAdd={onAdd} onFull={onFull} />
        <div className="text-xs font-medium text-gray-400 uppercase tracking-wide mb-2">{list.length} event{list.length === 1 ? "" : "s"}</div>
        {list.length === 0 && <p className="text-sm text-gray-400 py-6 text-center">Nothing scheduled. Add something above.</p>}
        {list.map((ev, i) => (
          <CalEventCard key={ev.id + i} ev={ev} color={eventColor(ev, colorBy, settings)}
            onEdit={() => onEdit(ev)} onComplete={() => onComplete(ev)} onDelete={() => onDelete(ev)} />
        ))}
      </div>
    </div>
  );
}

function CalendarPage({ data, settings, api, toast }) {
  const [cursor, setCursor] = useState(() => { const n = new Date(); return new Date(n.getFullYear(), n.getMonth(), n.getDate()); });
  const [view, setView] = useState("month");
  const [filters, setFilters] = useState({});
  const [colorBy, setColorBy] = useState("brand");
  const [editing, setEditing] = useState(null);
  const [panelDate, setPanelDate] = useState(null);
  const config = COLLECTIONS.events;
  const fileRef = useRef(null);
  const dragRef = useRef(null);
  const suppressClick = useRef(false);
  const [dragGhost, setDragGhost] = useState(null);
  const [dragOverDay, setDragOverDay] = useState(null);

  const onImport = (e) => {
    const f = e.target.files?.[0];
    if (f) importData(f, config, (recs) => { api.bulkAdd("events", recs); toast(`Imported ${recs.length} events`); });
    e.target.value = "";
  };

  const events = useMemo(() => (data.events || []).filter((e) => !e.deleted), [data]);
  const filtered = useMemo(() => events.filter((e) =>
    (!filters.brand || e.brand === filters.brand) &&
    (!filters.status || e.status === filters.status) &&
    (!filters.priority || e.priority === filters.priority) &&
    (!filters.eventType || e.eventType === filters.eventType)
  ), [events, filters]);

  const baseOf = (occ) => events.find((e) => e.id === occ.id) || occ;
  const openEdit = (occ) => { setPanelDate(null); setEditing(baseOf(occ)); };
  const complete = (occ) => { const b = baseOf(occ); api.update("events", b.id, { ...b, status: "Completed" }); toast("Marked complete"); };
  const remove = (occ) => { api.softDelete("events", baseOf(occ).id); toast("Moved to trash"); };
  const quickAdd = (date) => (fields) => { api.add("events", { date: dKey(date), status: "Not Started", priority: "Medium", category: "Campaign", repeat: "None", ...fields }); toast("Event added"); };

  const saveModal = (form) => {
    if (editing && editing.__new) api.add("events", form);
    else api.update("events", editing.id, form);
    setEditing(null); toast("Saved");
  };

  const step = (n) => { const d = new Date(cursor); if (view === "day") d.setDate(d.getDate() + n); else if (view === "week") d.setDate(d.getDate() + n * 7); else d.setMonth(d.getMonth() + n); setCursor(d); };
  const goToday = () => { const n = new Date(); setCursor(new Date(n.getFullYear(), n.getMonth(), n.getDate())); };

  const title = view === "day"
    ? cursor.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric" })
    : view === "agenda" ? "Upcoming" : `${MONTHS[cursor.getMonth()]} ${cursor.getFullYear()}`;

  // month grid
  const monthGrid = useMemo(() => {
    const first = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
    const start = new Date(first); start.setDate(1 - ((first.getDay() + 7) % 7));
    return Array.from({ length: 42 }, (_, i) => { const d = new Date(start); d.setDate(start.getDate() + i); return d; });
  }, [cursor]);
  const monthOcc = useMemo(() => expandOccurrences(filtered, monthGrid[0], monthGrid[41]), [filtered, monthGrid]);
  const byDay = useMemo(() => { const m = {}; monthOcc.forEach((o) => { (m[o._date] = m[o._date] || []).push(o); }); return m; }, [monthOcc]);

  const weekCells = useMemo(() => { const d = new Date(cursor); d.setDate(d.getDate() - ((d.getDay() + 7) % 7)); return Array.from({ length: 7 }, (_, i) => { const x = new Date(d); x.setDate(d.getDate() + i); return x; }); }, [cursor]);

  const agenda = useMemo(() => {
    const start = new Date(); start.setHours(0, 0, 0, 0);
    const end = new Date(start); end.setDate(end.getDate() + 120);
    const occ = expandOccurrences(filtered, start, end).sort((a, b) => a._date.localeCompare(b._date) || (a.time || "").localeCompare(b.time || ""));
    const groups = {};
    occ.forEach((o) => { (groups[o._date] = groups[o._date] || []).push(o); });
    return Object.keys(groups).sort().map((k) => ({ date: parseDate(k), items: groups[k] }));
  }, [filtered]);

  // pointer-based drag that works on mouse and touch
  const startPillDrag = (e, ev) => {
    if (e.button != null && e.button > 0) return;
    const startX = e.clientX, startY = e.clientY;
    const st = { id: ev.id, moved: false };
    dragRef.current = st;
    const color = eventColor(ev, colorBy, settings);
    const move = (me) => {
      const cx = me.clientX, cy = me.clientY;
      if (!st.moved && Math.hypot(cx - startX, cy - startY) > 6) st.moved = true;
      if (st.moved) {
        if (me.cancelable) me.preventDefault();
        setDragGhost({ x: cx, y: cy, title: ev.title, color });
        const el = document.elementFromPoint(cx, cy);
        const cell = el && el.closest ? el.closest("[data-cal-day]") : null;
        setDragOverDay(cell ? cell.getAttribute("data-cal-day") : null);
      }
    };
    const up = (ue) => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
      setDragGhost(null); setDragOverDay(null);
      const moved = st.moved; dragRef.current = null;
      if (moved) {
        suppressClick.current = true;
        setTimeout(() => { suppressClick.current = false; }, 350);
        const el = document.elementFromPoint(ue.clientX, ue.clientY);
        const cell = el && el.closest ? el.closest("[data-cal-day]") : null;
        const key = cell && cell.getAttribute("data-cal-day");
        const base = events.find((x) => x.id === ev.id);
        if (key && base) {
          if (!base.repeat || base.repeat === "None") { api.update("events", base.id, { ...base, date: key }); toast("Rescheduled"); }
          else toast("Recurring events can't be dragged");
        }
      } else {
        openEdit(ev);
      }
    };
    window.addEventListener("pointermove", move, { passive: false });
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
  };

  const seg = (v, label) => (
    <button key={v} onClick={() => setView(v)}
      className={`px-3 py-1 rounded-md text-sm font-medium capitalize transition ${view === v ? "bg-white shadow-sm text-gray-900" : "text-gray-500 hover:text-gray-800"}`}>{label}</button>
  );
  const todayK = dKey(new Date());

  return (
    <div>
      <style>{CAL_CSS}</style>

      {/* header */}
      <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
        <div className="flex items-center gap-1.5">
          <button onClick={() => step(-1)} className="p-1.5 rounded-lg text-gray-500 hover:bg-gray-100"><ChevronLeft size={18} /></button>
          <h2 className="text-xl font-semibold text-gray-900 tracking-tight min-w-[170px]">{title}</h2>
          <button onClick={() => step(1)} className="p-1.5 rounded-lg text-gray-500 hover:bg-gray-100"><ChevronRight size={18} /></button>
          <Btn variant="ghost" className="ml-1" onClick={goToday}>Today</Btn>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <Btn variant="ghost" onClick={() => fileRef.current?.click()} title="Import events"><Upload size={15} /></Btn>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" className="hidden" onChange={onImport} />
          <Btn variant="ghost" onClick={() => exportData("Calendar", events, config, "csv")} title="Export CSV"><Download size={15} /></Btn>
          <Btn variant="primary" onClick={() => setEditing({ __new: true, date: todayK })}><Plus size={16} />New event</Btn>
        </div>
      </div>

      {/* controls */}
      <div className="flex flex-wrap items-center justify-between gap-2 mb-3">
        <div className="inline-flex bg-gray-100 rounded-lg p-0.5">
          {seg("month", "Month")}{seg("week", "Week")}{seg("day", "Day")}{seg("agenda", "Agenda")}
        </div>
        <div className="flex flex-wrap items-center gap-2">
          {[["brand", "Brand", settings.brands], ["status", "Status", settings.statuses], ["priority", "Priority", settings.priorities], ["eventType", "Type", settings.eventTypes]].map(([k, l, o]) => (
            <select key={k} className="border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white" value={filters[k] || ""} onChange={(e) => setFilters((p) => ({ ...p, [k]: e.target.value }))}>
              <option value="">{l}: All</option>{o.map((x) => <option key={x} value={x}>{x}</option>)}
            </select>
          ))}
          <select className="border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white" value={colorBy} onChange={(e) => setColorBy(e.target.value)} title="Color events by">
            <option value="brand">Color: Brand</option><option value="priority">Color: Priority</option><option value="status">Color: Status</option><option value="eventType">Color: Type</option>
          </select>
          {Object.values(filters).some(Boolean) && <Btn variant="plain" onClick={() => setFilters({})}><X size={14} />Clear</Btn>}
        </div>
      </div>

      {/* drag/tap hint + brand legend (also filters) */}
      <p className="text-xs text-gray-400 mb-2">Tap a day to add, tap an event to edit, and in Month view drag an event to another day to reschedule.</p>
      {colorBy === "brand" && (
        <div className="flex flex-wrap gap-1.5 mb-3">
          {settings.brands.map((b) => {
            const c = eventColor({ brand: b }, "brand", settings);
            const active = filters.brand === b;
            return (
              <button key={b} onClick={() => setFilters((p) => ({ ...p, brand: active ? "" : b }))}
                className={`inline-flex items-center gap-1.5 px-2 py-1 rounded-full text-xs font-medium border transition ${active ? "border-gray-300 bg-gray-50" : "border-transparent hover:bg-gray-50"}`}>
                <span className="w-2.5 h-2.5 rounded-full" style={{ background: c }} />{b}
              </button>
            );
          })}
        </div>
      )}

      {/* MONTH */}
      {view === "month" && (
        <div className="bg-white border border-gray-200 rounded-2xl overflow-hidden shadow-sm cal-fade" key={"m" + cursor.getMonth() + cursor.getFullYear()}>
          <div className="grid grid-cols-7 text-xs font-semibold text-gray-400 border-b border-gray-100">
            {["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].map((d) => <div key={d} className="px-2 py-2.5 text-center uppercase tracking-wide">{d}</div>)}
          </div>
          <div className="grid grid-cols-7">
            {monthGrid.map((d, i) => {
              const inMonth = d.getMonth() === cursor.getMonth();
              const key = dKey(d), isToday = key === todayK;
              const list = byDay[key] || [];
              return (
                <div key={i} data-cal-day={key} style={{ minHeight: 118 }}
                  onClick={() => { if (suppressClick.current) { suppressClick.current = false; return; } setPanelDate(d); }}
                  className={`group relative border-b border-r border-gray-100 p-1.5 cursor-pointer transition-colors ${inMonth ? "hover:bg-gray-50/80" : "bg-gray-50/40"} ${dragOverDay === key ? "ring-2 ring-inset ring-emerald-400 bg-emerald-50/50" : ""}`}>
                  <div className="flex items-center justify-between mb-1">
                    <span className={`inline-flex items-center justify-center w-6 h-6 rounded-full text-xs font-semibold ${isToday ? "text-white" : inMonth ? "text-gray-700" : "text-gray-300"}`} style={isToday ? { background: "#059669" } : {}}>{d.getDate()}</span>
                    <button onClick={(e) => { e.stopPropagation(); setPanelDate(d); }} className="opacity-0 group-hover:opacity-100 text-gray-300 hover:text-emerald-600 transition"><Plus size={15} /></button>
                  </div>
                  <div className="space-y-1">
                    {list.slice(0, 3).map((ev, j) => {
                      const c = eventColor(ev, colorBy, settings);
                      return (
                        <div key={ev.id + j}
                          onPointerDown={(e) => { e.stopPropagation(); startPillDrag(e, ev); }}
                          onClick={(e) => { e.stopPropagation(); }}
                          style={{ background: hexToRgba(c, 0.12), borderLeft: `3px solid ${c}`, touchAction: "none" }}
                          className="text-xs leading-tight text-gray-800 rounded px-1.5 py-1 truncate transition hover:shadow-sm cursor-grab active:cursor-grabbing select-none" title={ev.title}>
                          {ev.time ? <span className="text-gray-500">{ev.time} </span> : null}{ev.title}
                        </div>
                      );
                    })}
                    {list.length > 3 && (
                      <button onClick={(e) => { e.stopPropagation(); setPanelDate(d); }} className="text-xs text-gray-400 hover:text-emerald-600 px-1 font-medium">+{list.length - 3} more</button>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* WEEK */}
      {view === "week" && (
        <div className="grid grid-cols-1 sm:grid-cols-7 gap-2 cal-fade" key={"w" + dKey(weekCells[0])}>
          {weekCells.map((d, i) => {
            const list = expandOccurrences(filtered, d, d).sort((a, b) => (a.time || "~").localeCompare(b.time || "~"));
            const isToday = dKey(d) === todayK;
            return (
              <div key={i} className="bg-white border border-gray-200 rounded-2xl p-2.5 shadow-sm" style={{ minHeight: 180 }} onClick={() => setPanelDate(d)}>
                <div className="flex items-center justify-between mb-2.5 cursor-pointer">
                  <div>
                    <div className="text-xs text-gray-400 uppercase tracking-wide">{["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"][d.getDay()]}</div>
                    <div className={`text-lg font-semibold ${isToday ? "text-emerald-600" : "text-gray-800"}`}>{d.getDate()}</div>
                  </div>
                  {list.length > 0 && <span className="text-xs text-gray-300">{list.length}</span>}
                </div>
                {list.map((ev, j) => (
                  <CalEventCard key={ev.id + j} ev={ev} color={eventColor(ev, colorBy, settings)} onEdit={() => openEdit(ev)} onComplete={() => complete(ev)} onDelete={() => remove(ev)} />
                ))}
              </div>
            );
          })}
        </div>
      )}

      {/* DAY */}
      {view === "day" && (
        <div className="bg-white border border-gray-200 rounded-2xl p-4 shadow-sm cal-fade max-w-2xl" key={"d" + dKey(cursor)}>
          {(() => {
            const list = expandOccurrences(filtered, cursor, cursor).sort((a, b) => (a.time || "~").localeCompare(b.time || "~"));
            if (list.length === 0) return (
              <div className="text-center py-12">
                <p className="text-gray-400 text-sm mb-3">Nothing scheduled for this day.</p>
                <Btn variant="soft" onClick={() => setPanelDate(cursor)}><Plus size={15} />Add an event</Btn>
              </div>
            );
            return list.map((ev, j) => (
              <CalEventCard key={ev.id + j} ev={ev} color={eventColor(ev, colorBy, settings)} onEdit={() => openEdit(ev)} onComplete={() => complete(ev)} onDelete={() => remove(ev)} />
            ));
          })()}
        </div>
      )}

      {/* AGENDA */}
      {view === "agenda" && (
        <div className="cal-fade max-w-2xl">
          {agenda.length === 0 && <div className="bg-white border border-gray-200 rounded-2xl p-10 text-center text-gray-400 text-sm">No upcoming events in the next 120 days.</div>}
          {agenda.map((g, i) => {
            const isToday = dKey(g.date) === todayK;
            return (
              <div key={i} className="mb-5">
                <div className="flex items-center gap-2 mb-2">
                  <span className={`text-sm font-semibold ${isToday ? "text-emerald-600" : "text-gray-800"}`}>{g.date.toLocaleDateString("en-US", { weekday: "short", month: "short", day: "numeric" })}</span>
                  {isToday && <span className="text-xs px-1.5 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-medium">Today</span>}
                  <span className="text-xs text-gray-300">{g.items.length} event{g.items.length === 1 ? "" : "s"}</span>
                </div>
                {g.items.map((ev, j) => (
                  <CalEventCard key={ev.id + j} ev={ev} color={eventColor(ev, colorBy, settings)} onEdit={() => openEdit(ev)} onComplete={() => complete(ev)} onDelete={() => remove(ev)} />
                ))}
              </div>
            );
          })}
        </div>
      )}

      {/* day panel */}
      {panelDate && (
        <DayPanel date={panelDate} occ={expandOccurrences(filtered, panelDate, panelDate)} settings={settings} colorBy={colorBy}
          onClose={() => setPanelDate(null)} onAdd={quickAdd(panelDate)}
          onFull={() => { const d = panelDate; setPanelDate(null); setEditing({ __new: true, date: dKey(d) }); }}
          onEdit={openEdit} onComplete={complete} onDelete={remove} />
      )}

      {/* floating drag preview */}
      {dragGhost && (
        <div style={{ position: "fixed", left: dragGhost.x + 10, top: dragGhost.y + 10, zIndex: 90, pointerEvents: "none", background: dragGhost.color }}
          className="text-xs text-white px-2 py-1 rounded-lg shadow-lg max-w-[200px] truncate">{dragGhost.title}</div>
      )}

      {/* full add / edit modal */}
      <Modal open={!!editing} onClose={() => setEditing(null)} wide title={editing && !editing.__new ? "Edit event" : "New event"}>
        {editing && (
          <RecordForm config={config} record={editing.__new ? { date: editing.date } : editing} settings={settings} onSave={saveModal} onCancel={() => setEditing(null)} />
        )}
        {editing && !editing.__new && (
          <div className="mt-3 -mb-2"><Btn variant="danger" onClick={() => { api.softDelete("events", editing.id); setEditing(null); toast("Moved to trash"); }}><Trash2 size={15} />Delete event</Btn></div>
        )}
      </Modal>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Roadmap                                                             */
/* ------------------------------------------------------------------ */
function RoadmapPage({ data, settings, api, toast }) {
  const [year, setYear] = useState(String(new Date().getFullYear()));
  const [gran, setGran] = useState("quarter");
  const roadmap = data.roadmap || [];
  const cols = gran === "quarter" ? QUARTERS : MONTHS_SHORT;
  const colKeys = gran === "quarter" ? ["q1", "q2", "q3", "q4"] : MONTHS.map((m) => m.toLowerCase().slice(0, 3));

  const rowFor = (brand) => roadmap.find((r) => r.brand === brand && r.year === year) || { brand, year };
  const setCell = (brand, key, val) => {
    const existing = roadmap.find((r) => r.brand === brand && r.year === year);
    if (existing) api.update("roadmap", existing.id, { ...existing, [key]: val });
    else api.add("roadmap", { brand, year, [key]: val });
  };

  return (
    <div>
      <PageHead title="Quarterly Roadmap">
        <select className="border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white" value={year} onChange={(e) => setYear(e.target.value)}>
          {["2024", "2025", "2026", "2027"].map((y) => <option key={y} value={y}>{y}</option>)}
        </select>
        {["quarter", "month"].map((g) => (
          <button key={g} onClick={() => setGran(g)} className={`px-3 py-1.5 rounded-lg text-sm font-medium capitalize ${gran === g ? "bg-emerald-600 text-white" : "bg-white border border-gray-200 text-gray-600"}`}>{g}s</button>
        ))}
      </PageHead>
      <p className="text-xs text-gray-400 mb-3">Type directly into a cell to plan. Changes save when you click away.</p>
      <div className="bg-white border border-gray-200 rounded-xl overflow-x-auto">
        <table className="w-full text-sm border-collapse">
          <thead>
            <tr className="bg-gray-50 text-gray-500 text-left">
              <th className="px-3 py-2.5 font-medium sticky left-0 bg-gray-50 min-w-[150px]">Brand</th>
              {cols.map((c) => <th key={c} className="px-3 py-2.5 font-medium min-w-[140px]">{c}</th>)}
              <th className="px-3 py-2.5 font-medium min-w-[160px]">Notes</th>
            </tr>
          </thead>
          <tbody>
            {settings.brands.map((b) => {
              const row = rowFor(b);
              return (
                <tr key={b} className="border-t border-gray-100">
                  <td className="px-3 py-2 font-medium text-gray-800 sticky left-0 bg-white border-r border-gray-100">{b}</td>
                  {colKeys.map((k) => (
                    <td key={k} className="p-0 border-r border-gray-50 align-top">
                      <textarea defaultValue={row[k] || ""} onBlur={(e) => { if (e.target.value !== (row[k] || "")) { setCell(b, k, e.target.value); toast("Saved"); } }}
                        rows={2} className="w-full h-full px-2 py-1.5 text-xs resize-none focus:outline-none focus:bg-emerald-50/50" placeholder="-" />
                    </td>
                  ))}
                  <td className="p-0 align-top">
                    <textarea defaultValue={row.notes || ""} onBlur={(e) => { if (e.target.value !== (row.notes || "")) { setCell(b, "notes", e.target.value); toast("Saved"); } }}
                      rows={2} className="w-full px-2 py-1.5 text-xs resize-none focus:outline-none focus:bg-emerald-50/50" placeholder="-" />
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Backlog & Archive (aggregated)                                      */
/* ------------------------------------------------------------------ */
function BacklogPage({ data, settings, api, toast, confirm }) {
  const [tab, setTab] = useState("backlog");
  const [filters, setFilters] = useState({});

  const all = useMemo(() => {
    const out = [];
    AGG_COLLECTIONS.forEach((ck) => {
      (data[ck] || []).forEach((r) => out.push({ ...r, __ck: ck, __type: COLLECTIONS[ck].label, __title: r[COLLECTIONS[ck].title] || "(untitled)" }));
    });
    return out;
  }, [data]);

  const rows = useMemo(() => {
    let r = all.filter((x) => {
      if (tab === "backlog") return !x.deleted && PARKED.includes(x.status);
      if (tab === "archived") return !x.deleted && x.status === "Archived";
      return x.deleted;
    });
    const q = (filters.q || "").toLowerCase();
    return r.filter((x) =>
      (!filters.brand || x.brand === filters.brand) &&
      (!filters.priority || x.priority === filters.priority) &&
      (!filters.quarter || x.quarter === filters.quarter) &&
      (!filters.year || x.year === filters.year) &&
      (!filters.type || x.__type === filters.type) &&
      (!q || (x.__title + " " + (x.notes || "")).toLowerCase().includes(q))
    );
  }, [all, tab, filters]);

  const restore = (r) => { api.update(r.__ck, r.id, { ...r, status: COLLECTIONS[r.__ck].restoreStatus, deleted: false }); toast("Restored to active"); };
  const archive = (r) => { api.update(r.__ck, r.id, { ...r, status: "Archived" }); toast("Archived"); };
  const perma = (r) => confirm({
    title: "Delete permanently?", message: "You will be asked once more.", danger: true, confirmText: "Continue",
    onConfirm: () => confirm({ title: "Are you absolutely sure?", message: "This cannot be undone.", danger: true, confirmText: "Delete forever", onConfirm: () => { api.hardDelete(r.__ck, r.id); toast("Deleted permanently"); } }),
  });
  const exportAll = (fmt) => {
    const data2 = rows.map((r) => ({ Type: r.__type, Title: r.__title, Brand: r.brand || "", Status: r.status || "", Priority: r.priority || "", Category: r.category || "", Quarter: r.quarter || "", Year: r.year || "", Notes: r.notes || "", "Last Updated": r.updatedDate || "" }));
    if (fmt === "csv") download(`Backlog-${tab}.csv`, new Blob([Papa.unparse(data2)], { type: "text/csv" }));
    else { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data2), tab); download(`Backlog-${tab}.xlsx`, new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })], { type: "application/octet-stream" })); }
  };

  return (
    <div>
      <PageHead title="Backlog & Archive">
        <Btn variant="ghost" onClick={() => exportAll("csv")}><Download size={15} />CSV</Btn>
        <Btn variant="ghost" onClick={() => exportAll("xlsx")}><Download size={15} />Excel</Btn>
      </PageHead>
      <div className="flex gap-2 mb-3">
        {[["backlog", "Backlog"], ["archived", "Archived"], ["trash", "Trash"]].map(([k, l]) => (
          <button key={k} onClick={() => setTab(k)} className={`px-3.5 py-1.5 rounded-lg text-sm font-medium ${tab === k ? "bg-emerald-600 text-white" : "bg-white border border-gray-200 text-gray-600"}`}>
            {l} <span className="opacity-70">({all.filter((x) => k === "backlog" ? !x.deleted && PARKED.includes(x.status) : k === "archived" ? !x.deleted && x.status === "Archived" : x.deleted).length})</span>
          </button>
        ))}
      </div>
      <div className="flex flex-wrap items-center gap-2 mb-3">
        <div className="relative">
          <Search size={15} className="absolute left-2.5 top-2.5 text-gray-400" />
          <input placeholder="Search..." value={filters.q || ""} onChange={(e) => setFilters((p) => ({ ...p, q: e.target.value }))} className="pl-8 pr-3 py-1.5 border border-gray-200 rounded-lg text-sm w-44 focus:outline-none focus:ring-2 focus:ring-emerald-200" />
        </div>
        <select className="border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white" value={filters.type || ""} onChange={(e) => setFilters((p) => ({ ...p, type: e.target.value }))}>
          <option value="">Type: All</option>{AGG_COLLECTIONS.map((c) => <option key={c} value={COLLECTIONS[c].label}>{COLLECTIONS[c].label}</option>)}
        </select>
        {[["brand", "Brand", settings.brands], ["priority", "Priority", settings.priorities], ["quarter", "Quarter", QUARTERS], ["year", "Year", ["2024", "2025", "2026", "2027"]]].map(([k, l, o]) => (
          <select key={k} className="border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm bg-white" value={filters[k] || ""} onChange={(e) => setFilters((p) => ({ ...p, [k]: e.target.value }))}>
            <option value="">{l}: All</option>{o.map((x) => <option key={x} value={x}>{x}</option>)}
          </select>
        ))}
      </div>
      <div className="bg-white border border-gray-200 rounded-xl overflow-x-auto">
        <table className="w-full text-sm">
          <thead><tr className="text-left text-gray-500 border-b border-gray-100 bg-gray-50/60">
            {["Type", "Title", "Brand", "Status", "Priority", "Quarter", "Year", "Actions"].map((h) => <th key={h} className={`px-3 py-2.5 font-medium ${h === "Actions" ? "text-right" : ""}`}>{h}</th>)}
          </tr></thead>
          <tbody>
            {rows.length === 0 && <tr><td colSpan={8} className="px-3 py-10 text-center text-gray-400">Nothing in {tab}.</td></tr>}
            {rows.map((r) => (
              <tr key={r.__ck + r.id} className="border-b border-gray-50 hover:bg-gray-50/60">
                <td className="px-3 py-2.5 text-gray-500">{r.__type}</td>
                <td className="px-3 py-2.5 font-medium text-gray-900">{r.__title}</td>
                <td className="px-3 py-2.5">{r.brand || "-"}</td>
                <td className="px-3 py-2.5"><Chip value={r.status} /></td>
                <td className="px-3 py-2.5"><Chip value={r.priority} /></td>
                <td className="px-3 py-2.5 text-gray-500">{r.quarter || "-"}</td>
                <td className="px-3 py-2.5 text-gray-500">{r.year || "-"}</td>
                <td className="px-3 py-2.5 text-right whitespace-nowrap">
                  <IconBtn title="Restore to active" onClick={() => restore(r)}><RotateCcw size={15} /></IconBtn>
                  {tab !== "archived" && tab !== "trash" && <IconBtn title="Archive" onClick={() => archive(r)}><Archive size={15} /></IconBtn>}
                  <IconBtn title="Delete permanently" danger onClick={() => perma(r)}><Trash2 size={15} /></IconBtn>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Settings                                                            */
/* ------------------------------------------------------------------ */
const SETTING_LABELS = {
  brands: "Brands", priorities: "Priorities", statuses: "Statuses", recStatuses: "Recommendation Statuses",
  categories: "Categories", platforms: "Platforms", channels: "Channels", campaignTypes: "Campaign Types",
  impactTypes: "Impact Types", importanceLevels: "Importance Levels", eventTypes: "Event Types", brandStatuses: "Brand Statuses",
};
function SettingsPage({ settings, saveSettings, toast }) {
  const update = (key, arr) => saveSettings({ ...settings, [key]: arr });
  return (
    <div>
      <PageHead title="Reference Lists & Settings" />
      <p className="text-sm text-gray-500 mb-4">These lists feed every dropdown in the hub. Edit them once and they update everywhere.</p>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {Object.keys(SETTING_LABELS).map((key) => (
          <ListEditor key={key} label={SETTING_LABELS[key]} items={settings[key] || []} onChange={(arr) => { update(key, arr); toast("Saved"); }} />
        ))}
      </div>
    </div>
  );
}
function ListEditor({ label, items, onChange }) {
  const [val, setVal] = useState("");
  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <h3 className="font-semibold text-gray-900 mb-3 text-sm">{label}</h3>
      <div className="flex flex-wrap gap-1.5 mb-3">
        {items.map((it, i) => (
          <span key={i} className="inline-flex items-center gap-1 bg-gray-100 text-gray-700 text-xs px-2 py-1 rounded-lg">
            {it}<button onClick={() => onChange(items.filter((_, j) => j !== i))} className="text-gray-400 hover:text-red-600"><X size={12} /></button>
          </span>
        ))}
        {items.length === 0 && <span className="text-xs text-gray-400">Empty list</span>}
      </div>
      <div className="flex gap-2">
        <input value={val} onChange={(e) => setVal(e.target.value)} onKeyDown={(e) => { if (e.key === "Enter" && val.trim()) { onChange([...items, val.trim()]); setVal(""); } }}
          placeholder={`Add to ${label.toLowerCase()}`} className="flex-1 border border-gray-200 rounded-lg px-2.5 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200" />
        <Btn variant="soft" onClick={() => { if (val.trim()) { onChange([...items, val.trim()]); setVal(""); } }}><Plus size={15} /></Btn>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Home dashboard                                                      */
/* ------------------------------------------------------------------ */
function Kpi({ label, value, accent }) {
  return (
    <div className="bg-white border border-gray-200 rounded-xl px-3.5 py-3">
      <div className={`text-2xl font-semibold tabular-nums ${accent ? "text-emerald-600" : "text-gray-900"}`}>{value}</div>
      <div className="text-xs text-gray-500 mt-0.5 leading-tight">{label}</div>
    </div>
  );
}

function iconForTag(tag) {
  const t = (tag || "").toLowerCase();
  if (/bever|drink|tea|coffee|matcha|texture/.test(t)) return Coffee;
  if (/dessert|sweet|choc|pastry|bake/.test(t)) return Cookie;
  if (/format|burger|meat|beef|protein/.test(t)) return Beef;
  if (/flavor|spice|hot|viral|snack/.test(t)) return Flame;
  if (/soda|juice|shake/.test(t)) return CupSoda;
  return Utensils;
}

function TrendImage({ image, color, tag }) {
  const [err, setErr] = useState(false);
  if (image && !err) {
    return <img src={image} alt={tag || ""} onError={() => setErr(true)} className="w-full h-full object-cover" loading="lazy" />;
  }
  const Icon = iconForTag(tag);
  return (
    <div className="w-full h-full relative overflow-hidden" style={{ background: `linear-gradient(135deg, ${color} 0%, ${hexToRgba(color, 0.5)} 100%)` }}>
      <div className="absolute inset-0 flex items-center justify-center" style={{ opacity: 0.22 }}>
        <Icon size={104} color="#ffffff" strokeWidth={1.25} />
      </div>
      <div className="absolute -right-8 -top-10 rounded-full" style={{ width: 130, height: 130, background: "rgba(255,255,255,0.10)" }} />
    </div>
  );
}

function Home({ data, settings, go, saveTrends, saveInspo }) {
  const cur = currentPeriod();
  const live = (ck) => (data[ck] || []).filter((r) => !r.deleted);
  const allLive = AGG_COLLECTIONS.flatMap((ck) => live(ck));

  const openItems = allLive.filter((r) => r.status && !DONE.includes(r.status) && !PARKED.includes(r.status) && r.status !== "Archived").length;
  const completed = allLive.filter((r) => DONE.includes(r.status)).length;
  const backlog = allLive.filter((r) => PARKED.includes(r.status)).length;
  const archived = allLive.filter((r) => r.status === "Archived").length;
  const activeBrands = live("brandHub").filter((b) => b.status === "Active").length || new Set(allLive.map((r) => r.brand).filter(Boolean)).size;

  const inMonth = (r) => r.month === cur.month && r.year === cur.year;
  const inQuarter = (r) => r.quarter === cur.quarter && r.year === cur.year;
  const inYear = (r) => r.year === cur.year;

  const kpis = [
    ["Work entries", live("work").length], ["Ideas logged", live("ideas").length],
    ["Findings recorded", live("findings").length], ["Recommendations", live("recommendations").length],
    ["Wins achieved", live("wins").length], ["Open items", openItems, true],
    ["Completed items", completed], ["Backlog items", backlog],
    ["Archived items", archived], ["Active brands", activeBrands],
    ["Work this month", live("work").filter(inMonth).length], ["Work this quarter", live("work").filter(inQuarter).length],
    ["Work this year", live("work").filter(inYear).length], ["Ideas this month", live("ideas").filter(inMonth).length],
    ["Recommendations this quarter", live("recommendations").filter(inQuarter).length],
    ["Completed this year", allLive.filter((r) => DONE.includes(r.status) && r.year === cur.year).length],
  ];

  // follow-ups due this week
  const weekStart = new Date(); weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() + 7) % 7));
  const weekEnd = new Date(weekStart); weekEnd.setDate(weekStart.getDate() + 6);
  const followUps = allLive.filter((r) => { const d = parseDate(r.followUpDate); return d && d >= weekStart && d <= weekEnd; });

  const today = todayKey();
  const todaysTasks = [...live("events").filter((e) => e.date === today), ...live("work").filter((w) => w.date === today)];
  const recent = [...allLive].filter((r) => r.updatedDate).sort((a, b) => (b.updatedDate || "").localeCompare(a.updatedDate || "")).slice(0, 8);

  const byBrand = settings.brands.map((b) => ({ name: b.length > 10 ? b.slice(0, 9) + "…" : b, full: b, count: allLive.filter((r) => r.brand === b).length })).filter((x) => x.count > 0).sort((a, b) => b.count - a.count).slice(0, 8);

  // system briefing: turn the data into plain-language callouts
  const liveEvents = live("events");
  const todayMid = new Date(today + "T00:00:00");
  const overdue = liveEvents.filter((e) => e.date && e.date < today && !DONE.includes(e.status) && ["Task", "Deadline", "Follow-up", "Report", "Presentation"].includes(e.eventType));
  const todayEvents = liveEvents.filter((e) => e.date === today);
  const upcomingWeek = liveEvents.filter((e) => { const d = parseDate(e.date); return d && d > todayMid && d <= weekEnd; });
  const campMonth = liveEvents.filter((e) => e.eventType === "Campaign" && e.month === cur.month && e.year === cur.year);
  const nm = new Date(); nm.setMonth(nm.getMonth() + 1);
  const nextM = MONTHS[nm.getMonth()], nextY = String(nm.getFullYear());
  const campNext = liveEvents.filter((e) => e.eventType === "Campaign" && e.month === nextM && e.year === nextY);
  const brandsMonth = [...new Set(campMonth.map((e) => e.brand).filter(Boolean))];
  const ideasReview = live("ideas").filter((i) => i.reviewDate && i.reviewDate <= today && !DONE.includes(i.status));
  const recsAwaiting = live("recommendations").filter((r) => ["Shared", "Under Review"].includes(r.status));
  const backlogCount = allLive.filter((r) => PARKED.includes(r.status)).length;
  const winsQ = live("wins").filter((w) => w.quarter === cur.quarter && w.year === cur.year).length;
  const titlesOf = (arr) => arr.slice(0, 2).map((x) => x.recommendation || x.title || x.meetingTitle || x.idea).filter(Boolean).join(", ");
  const plural = (n) => (n === 1 ? "" : "s");
  const B = [];
  if (overdue.length) B.push({ tone: "alert", section: "events", text: `${overdue.length} calendar item${plural(overdue.length)} ${overdue.length === 1 ? "is" : "are"} past due and not marked done.` });
  if (followUps.length) B.push({ tone: "alert", section: "recommendations", text: `${followUps.length} follow-up${plural(followUps.length)} due this week${titlesOf(followUps) ? `: ${titlesOf(followUps)}` : ""}.` });
  if (todayEvents.length) B.push({ tone: "info", section: "events", text: `${todayEvents.length} thing${plural(todayEvents.length)} scheduled today: ${titlesOf(todayEvents)}.` });
  if (upcomingWeek.length) B.push({ tone: "info", section: "events", text: `${upcomingWeek.length} event${plural(upcomingWeek.length)} coming up later this week.` });
  if (campMonth.length) B.push({ tone: "info", section: "events", text: `${campMonth.length} campaign${plural(campMonth.length)} live this ${cur.month} across ${brandsMonth.length} brand${plural(brandsMonth.length)}.` });
  if (campNext.length) B.push({ tone: "info", section: "events", text: `Heads up: ${campNext.length} campaign${plural(campNext.length)} scheduled for ${nextM}.` });
  if (ideasReview.length) B.push({ tone: "info", section: "ideas", text: `${ideasReview.length} idea${plural(ideasReview.length)} flagged for review.` });
  if (recsAwaiting.length) B.push({ tone: "info", section: "recommendations", text: `${recsAwaiting.length} recommendation${plural(recsAwaiting.length)} still awaiting a decision.` });
  if (backlogCount) B.push({ tone: "info", section: "backlog", text: `${backlogCount} item${plural(backlogCount)} parked in the backlog, worth a review.` });
  if (winsQ) B.push({ tone: "good", section: "wins", text: `You have logged ${winsQ} win${plural(winsQ)} this ${cur.quarter}.` });
  const headline = (overdue.length || followUps.length) ? "A few things need your attention." : B.length ? "Here is what is moving right now." : "All clear. Nothing pressing at the moment.";
  const briefing = B.slice(0, 6);
  if (!briefing.length) briefing.push({ tone: "good", section: null, text: "No overdue items and nothing due this week. Log work, ideas, or events to start tracking." });

  // ---- trending F&B board ----
  const savedFresh = (data.trends && data.trends.v === 2 && Array.isArray(data.trends.items) && data.trends.items.length) ? data.trends.items : null;
  const [mix, setMix] = useState(() => savedFresh || SEED_TRENDS);
  const [loadingTrends, setLoadingTrends] = useState(false);
  const [trendNote, setTrendNote] = useState(null);
  const inspo = data.inspo || [];
  const [addInspo, setAddInspo] = useState(null);

  const TREND_PALETTE = ["#059669", "#0891b2", "#7c3aed", "#d97706", "#db2777", "#ea580c", "#0d9488", "#4f46e5"];

  useEffect(() => {
    if (data.trends && data.trends.v === 2 && Array.isArray(data.trends.items) && data.trends.items.length) setMix(data.trends.items);
  }, [data.trends]);

  const refreshTrends = async () => {
    if (loadingTrends) return;
    setLoadingTrends(true); setTrendNote(null);
    try {
      const res = await fetch("/api/claude", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 1000,
          tools: [{ type: "web_search_20250305", name: "web_search" }],
          messages: [{
            role: "user",
            content: "Search the web for the newest food and beverage trends spiking right now this season, the specific items, flavors, drinks, or formats people are talking about on social media and in restaurants in the last few months. Be specific and name actual things, not broad evergreen themes. Do not include generic categories such as comfort food, health and wellness, sustainability, value, or protein. Pick 6 fresh, specific ones. For each return: title (the specific trend, short and punchy), tag (one short word like Flavor, Dessert, Drink, Format, Snack, Viral), blurb (one plain sentence on why it matters for a multi brand restaurant group, no fluff), source (the publication name), url (the article link), keyword (two lowercase words comma separated no spaces describing the food for a photo). HARD RULES: no alcohol, beer, wine, cocktails, bars, or pork, bacon, ham anywhere. No em dashes. Reply with ONLY a JSON array of 6 objects with keys title, tag, blurb, source, url, keyword. No other text.",
          }],
        }),
      });
      if (!res.ok) throw new Error(String(res.status));
      const data2 = await res.json();
      const raw = (data2.content || []).filter((b) => b.type === "text").map((b) => b.text || "").join("").trim();
      const a = raw.indexOf("["), z = raw.lastIndexOf("]");
      if (a < 0 || z < 0) throw new Error("parse");
      const items = JSON.parse(raw.slice(a, z + 1)).slice(0, 6).map((t) => ({ ...t, image: t.keyword ? `https://loremflickr.com/640/420/${String(t.keyword).replace(/\s+/g, "")}` : "" }));
      if (!items.length) throw new Error("empty");
      setMix(items);
      saveTrends({ items, v: 2, updatedAt: Date.now() });
      setTrendNote({ ok: true, text: "Pulled a fresh set from the web." });
    } catch (err) {
      setTrendNote({ ok: false, text: "Could not reach the web just now. Showing the last set. Try again in a moment." });
    }
    setLoadingTrends(false);
  };

  const didAuto = useRef(false);
  useEffect(() => {
    if (didAuto.current) return;
    didAuto.current = true;
    const fresh = data.trends && data.trends.v === 2 && Array.isArray(data.trends.items) && data.trends.items.length && (Date.now() - (data.trends.updatedAt || 0) < 21600000);
    if (!fresh) refreshTrends();
  }, []);

  const featured = mix[0];
  const rest = mix.slice(1);

  return (
    <div className="space-y-6">
      {/* hero greeting */}
      <div className="flex items-start justify-between gap-3 flex-wrap">
        <div>
          <h2 className="text-xl font-semibold text-gray-900 tracking-tight">What's cooking, Suleman</h2>
          <p className="text-sm text-gray-500">{cur.month} {cur.year}. Fresh inspiration and where things stand.</p>
        </div>
        <Btn variant="primary" onClick={() => setAddInspo({ title: "", note: "", img: "" })}><Plus size={15} />Pin inspiration</Btn>
      </div>

      {/* TRENDING BOARD */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <div className="flex items-center gap-2">
            <h3 className="font-semibold text-gray-900">Trending in F&B</h3>
            <span className="text-xs font-medium px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">Worldwide</span>
          </div>
          <Btn variant="ghost" onClick={refreshTrends} className={loadingTrends ? "opacity-60" : ""}>
            <RotateCcw size={14} className={loadingTrends ? "animate-spin" : ""} />{loadingTrends ? "Searching" : "Refresh"}
          </Btn>
        </div>

        {trendNote && (
          <p className={`text-xs mb-3 ${trendNote.ok ? "text-emerald-700" : "text-amber-600"}`}>{trendNote.text}</p>
        )}

        <div className="space-y-4">
          {/* featured */}
          {featured && (
            <div key={featured.title} className="relative rounded-2xl overflow-hidden border border-gray-200" style={{ height: 260 }}>
              <div className="absolute inset-0"><TrendImage image={featured.image} idx={0} color={TREND_PALETTE[0]} tag={featured.tag} /></div>
              <div className="absolute inset-0" style={{ background: "linear-gradient(to top, rgba(0,0,0,.82), rgba(0,0,0,.15) 55%, rgba(0,0,0,0))" }} />
              <div className="absolute bottom-0 left-0 right-0 p-5">
                <span className="inline-block text-xs font-semibold px-2 py-0.5 rounded-full mb-2 text-white" style={{ background: TREND_PALETTE[0] }}>{featured.tag}</span>
                <h4 className="text-white text-2xl font-semibold leading-tight mb-1">{featured.title}</h4>
                <p className="text-white/85 text-sm max-w-2xl leading-snug">{featured.blurb}</p>
                {featured.source && (
                  <a href={featured.url || undefined} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-white/70 hover:text-white text-xs mt-2">
                    via {featured.source}{featured.url && <ExternalLink size={11} />}
                  </a>
                )}
              </div>
            </div>
          )}
          {/* grid */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
            {rest.map((t, i) => {
              const c = TREND_PALETTE[(i + 1) % TREND_PALETTE.length];
              return (
                <div key={t.title + i} className="bg-white border border-gray-200 rounded-2xl overflow-hidden hover:shadow-md transition group">
                  <div className="overflow-hidden" style={{ height: 150 }}>
                    <div className="w-full h-full group-hover:scale-105 transition-transform duration-500">
                      <TrendImage image={t.image} idx={i + 1} color={c} tag={t.tag} />
                    </div>
                  </div>
                  <div className="p-3.5">
                    <span className="inline-block text-xs font-semibold px-2 py-0.5 rounded-full mb-1.5" style={{ background: hexToRgba(c, 0.12), color: c }}>{t.tag}</span>
                    <h4 className="font-semibold text-gray-900 text-sm leading-tight mb-1">{t.title}</h4>
                    <p className="text-xs text-gray-500 leading-relaxed">{t.blurb}</p>
                    {t.source && (
                      <a href={t.url || undefined} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1 text-gray-400 hover:text-emerald-700 text-xs mt-1.5">
                        via {t.source}{t.url && <ExternalLink size={10} />}
                      </a>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          <p className="text-xs text-gray-400">Refresh pulls fresh trends from the web with clickable sources. Photos sit on your Pin board below, where uploads display reliably.</p>
        </div>
      </div>

      {/* YOUR INSPIRATION */}
      <div>
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold text-gray-900">Your inspiration</h3>
          <Btn variant="primary" onClick={() => setAddInspo({ title: "", note: "", img: "" })}><Plus size={15} />Pin something</Btn>
        </div>
        {inspo.length === 0 ? (
          <div className="bg-white border border-gray-200 border-dashed rounded-2xl p-8 text-center">
            <ImageIcon className="mx-auto mb-2 text-gray-300" />
            <p className="text-sm text-gray-400">Pin brands, campaigns, or references that inspire you. Upload a photo and it displays right here.</p>
          </div>
        ) : (
          <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
            {inspo.map((p) => (
              <div key={p.id} className="bg-white border border-gray-200 rounded-2xl overflow-hidden group relative">
                <button onClick={() => saveInspo(inspo.filter((x) => x.id !== p.id))}
                  className="absolute top-2 right-2 z-10 w-6 h-6 rounded-full bg-black/50 text-white flex items-center justify-center opacity-0 group-hover:opacity-100 transition"><X size={13} /></button>
                <div className="bg-gray-100" style={{ height: 130 }}>
                  <InspoImg src={p.img} alt={p.title} />
                </div>
                <div className="p-3">
                  <h4 className="font-medium text-gray-900 text-sm leading-tight">{p.title || "Untitled"}</h4>
                  {p.note && <p className="text-xs text-gray-500 mt-0.5 leading-snug">{p.note}</p>}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* compact briefing */}
      <div className="bg-white border border-gray-200 rounded-2xl p-4 sm:p-5">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-semibold text-gray-900 flex items-center gap-2">
            <span className="w-1.5 h-5 rounded-full" style={{ background: "#059669" }} />{headline}
          </h3>
          <span className="text-xs text-gray-400 shrink-0">Your briefing</span>
        </div>
        <ul className="space-y-2">
          {briefing.map((l, i) => (
            <li key={i} className="flex items-start gap-2.5">
              <span className="w-2 h-2 rounded-full mt-1.5 shrink-0" style={{ background: l.tone === "alert" ? "#d97706" : l.tone === "good" ? "#059669" : "#3b82f6" }} />
              <span className="text-sm text-gray-700 flex-1">{l.text}</span>
              {l.section && (
                <button onClick={() => go(l.section)} className="text-xs font-medium text-emerald-700 hover:underline shrink-0 mt-0.5 inline-flex items-center gap-0.5">View<ChevronRight size={12} /></button>
              )}
            </li>
          ))}
        </ul>
      </div>

      {/* compact stat strip */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-8 gap-2.5">
        {kpis.slice(0, 8).map(([l, v, a]) => <Kpi key={l} label={l} value={v} accent={a} />)}
      </div>

      {addInspo && <InspoModal state={addInspo} setState={setAddInspo} onSave={(p) => { saveInspo([{ id: uid(), ...p }, ...inspo]); setAddInspo(null); }} />}
    </div>
  );
}

function InspoImg({ src, alt }) {
  const [err, setErr] = useState(false);
  if (!src || err) return <div className="w-full h-full flex items-center justify-center text-gray-300"><ImageIcon /></div>;
  return <img src={src} alt={alt || ""} onError={() => setErr(true)} className="w-full h-full object-cover" />;
}

function InspoModal({ state, setState, onSave }) {
  const fileRef = useRef(null);
  const inp = "w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200";
  const onFile = (e) => { const f = e.target.files?.[0]; e.target.value = ""; if (f) compressImage(f, (d) => d && setState((p) => ({ ...p, img: d }))); };
  return (
    <Modal open onClose={() => setState(null)} title="Pin inspiration">
      <label className="block text-xs font-medium text-gray-500 mb-1">Title</label>
      <input className={inp + " mb-3"} value={state.title} onChange={(e) => setState((p) => ({ ...p, title: e.target.value }))} autoFocus />
      <label className="block text-xs font-medium text-gray-500 mb-1">Note (optional)</label>
      <input className={inp + " mb-3"} value={state.note} onChange={(e) => setState((p) => ({ ...p, note: e.target.value }))} />
      <label className="block text-xs font-medium text-gray-500 mb-1">Image</label>
      <Btn variant="primary" onClick={() => fileRef.current?.click()}><Upload size={15} />Upload a photo</Btn>
      <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={onFile} />
      <p className="text-xs text-gray-400 mt-1.5 mb-2">Uploading is the reliable way to see your photo here. A pasted link may not display.</p>
      <input className={inp} value={state.img && state.img.startsWith("data:") ? "" : state.img} placeholder="Or paste an image URL" onChange={(e) => setState((p) => ({ ...p, img: e.target.value }))} />
      {state.img && <button onClick={() => setState((p) => ({ ...p, img: "" }))} className="text-xs text-gray-400 hover:text-red-600 mt-2">Remove image</button>}
      {state.img && state.img.startsWith("data:") && <img src={state.img} alt="" className="w-full rounded-lg object-cover border border-gray-200 mt-3" style={{ maxHeight: 160 }} />}
      <div className="flex justify-end gap-2 mt-3">
        <Btn variant="ghost" onClick={() => setState(null)}>Cancel</Btn>
        <Btn variant="primary" onClick={() => onSave({ title: state.title.trim(), note: state.note.trim(), img: state.img })}><CheckCircle2 size={16} />Pin</Btn>
      </div>
    </Modal>
  );
}

function Panel({ title, badge, onMore, children }) {
  return (
    <div className="bg-white border border-gray-200 rounded-xl p-4">
      <div className="flex items-center justify-between mb-2.5">
        <h3 className="font-semibold text-gray-900 text-sm flex items-center gap-2">{title}{badge != null && <span className="text-xs font-medium px-1.5 py-0.5 bg-emerald-50 text-emerald-700 rounded-full">{badge}</span>}</h3>
        {onMore && <button onClick={onMore} className="text-xs text-emerald-700 hover:underline">View all</button>}
      </div>
      <div>{children}</div>
    </div>
  );
}
function Line({ title, meta, chip }) {
  return (
    <div className="flex items-center justify-between py-1.5 border-b border-gray-50 last:border-0 gap-2">
      <span className="text-sm text-gray-800 truncate flex-1">{title}</span>
      <span className="flex items-center gap-2 shrink-0">{chip && <Chip value={chip} />}<span className="text-xs text-gray-400">{meta}</span></span>
    </div>
  );
}
function Empty({ text }) { return <p className="text-sm text-gray-400 py-3 text-center">{text}</p>; }

/* ------------------------------------------------------------------ */
/* Error boundary                                                      */
/* ------------------------------------------------------------------ */
class ErrorBoundary extends React.Component {
  constructor(p) { super(p); this.state = { err: null }; }
  static getDerivedStateFromError(err) { return { err }; }
  render() {
    if (this.state.err) return (
      <div className="p-8 text-center">
        <p className="text-gray-700 font-medium mb-1">Something hiccuped on this view.</p>
        <p className="text-sm text-gray-500 mb-4">Your data is safe. Reload to continue.</p>
        <Btn variant="primary" onClick={() => this.setState({ err: null })}>Try again</Btn>
      </div>
    );
    return this.props.children;
  }
}

/* ------------------------------------------------------------------ */
/* App shell                                                           */
/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/* Idea Map: animated force-graph of nodes (with images) and links      */
/* ------------------------------------------------------------------ */
const MAP_PALETTE = ["#10b981", "#3b82f6", "#8b5cf6", "#f59e0b", "#ef4444", "#06b6d4", "#ec4899", "#f97316", "#14b8a6", "#a3e635"];
const NODE_R = 38; // node radius

// shrink any uploaded image to a small thumbnail so it always stores and loads fast
function compressImage(file, cb) {
  const reader = new FileReader();
  reader.onload = () => {
    const img = new Image();
    img.onload = () => {
      const max = 240;
      let { width, height } = img;
      if (width > height && width > max) { height = Math.round(height * max / width); width = max; }
      else if (height > max) { width = Math.round(width * max / height); height = max; }
      const c = document.createElement("canvas");
      c.width = width; c.height = height;
      c.getContext("2d").drawImage(img, 0, 0, width, height);
      try { cb(c.toDataURL("image/jpeg", 0.8)); }
      catch { cb(reader.result); }
    };
    img.onerror = () => cb(reader.result);
    img.src = reader.result;
  };
  reader.onerror = () => cb(null);
  reader.readAsDataURL(file);
}

function NodeEditor({ node, settings, onSave, onDelete, onConnect, onCancel, toast }) {
  const [label, setLabel] = useState(node.label || "");
  const [brand, setBrand] = useState(node.brand || "");
  const [type, setType] = useState(node.type || "");
  const [desc, setDesc] = useState(node.description || "");
  const [img, setImg] = useState(node.img || "");
  const [pickColor, setPickColor] = useState(node.color || MAP_PALETTE[0]);
  const [polishing, setPolishing] = useState(false);
  const fileRef = useRef(null);
  const color = brand ? eventColor({ brand }, "brand", settings) : pickColor;
  const onFile = (e) => {
    const f = e.target.files?.[0]; e.target.value = "";
    if (!f) return;
    compressImage(f, (dataUrl) => { if (dataUrl) setImg(dataUrl); else toast && toast("Could not read that image"); });
  };
  const polish = async () => {
    if (!desc.trim() || polishing) return;
    setPolishing(true);
    try {
      const res = await fetch("/api/claude", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 1000,
          system: "You rewrite short notes for a marketing idea board. Rewrite the description so it reads clear, specific, and natural. Rules: never use em dashes; never use words or phrases that sound like AI filler such as dive in, unlock, leverage, elevate, seamless, robust, supercharge, game-changer, in today's landscape, it's worth noting; write like a sharp human colleague; keep it concise, a few sentences at most; no preamble, no quotes, no markdown; return only the rewritten description.",
          messages: [{ role: "user", content: `Title: ${label || "(untitled)"}\nBrand: ${brand || "(none)"}\nType: ${type || "(none)"}\n\nDescription to rewrite:\n${desc}` }],
        }),
      });
      if (!res.ok) throw new Error("api");
      const d = await res.json();
      const out = (d.content || []).map((b) => b.text || "").join("").trim();
      if (out) setDesc(out);
    } catch { toast && toast("Could not reach the AI. Try again."); }
    setPolishing(false);
  };
  const inp = "w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200";
  return (
    <div>
      <label className="block text-xs font-medium text-gray-500 mb-1">Title</label>
      <input className={inp + " mb-3"} value={label} onChange={(e) => setLabel(e.target.value)} autoFocus />
      <div className="grid grid-cols-2 gap-3 mb-3">
        <div>
          <label className="block text-xs font-medium text-gray-500 mb-1">Brand</label>
          <select className={inp} value={brand} onChange={(e) => setBrand(e.target.value)}>
            <option value="">No brand</option>
            {settings.brands.map((b) => <option key={b} value={b}>{b}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-500 mb-1">Type</label>
          <select className={inp} value={type} onChange={(e) => setType(e.target.value)}>
            {["", "Idea", "Campaign", "Brand", "Finding", "Reference", "Note"].map((t) => <option key={t} value={t}>{t || "None"}</option>)}
          </select>
        </div>
      </div>
      <div className="flex items-center justify-between mb-1">
        <label className="block text-xs font-medium text-gray-500">Description</label>
        <button onClick={polish} disabled={polishing || !desc.trim()}
          className="inline-flex items-center gap-1 text-xs font-medium text-emerald-700 hover:text-emerald-800 disabled:opacity-40">
          <Sparkles size={13} />{polishing ? "Polishing..." : "Polish with AI"}
        </button>
      </div>
      <textarea rows={4} className={inp} value={desc} onChange={(e) => setDesc(e.target.value)} placeholder="Write a few lines. Then let AI tighten it up." />
      <p className="text-[11px] text-gray-400 mt-1 mb-3">AI rewrites this in plain, human language. No em dashes, no fluff.</p>

      <label className="block text-xs font-medium text-gray-500 mb-1">Image</label>
      <input className={inp} value={img && img.startsWith("data:") ? "" : img} placeholder="Paste an image URL, or upload below" onChange={(e) => setImg(e.target.value)} />
      <div className="flex items-center gap-2 my-2">
        <Btn variant="ghost" onClick={() => fileRef.current?.click()}><ImageIcon size={15} />Upload image</Btn>
        <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={onFile} />
        {img && <button onClick={() => setImg("")} className="text-xs text-gray-400 hover:text-red-600">Remove image</button>}
      </div>
      {img && <img src={img} alt="" className="w-16 h-16 rounded-lg object-cover border border-gray-200 mb-3" />}

      <div className="flex items-center gap-2 mb-4">
        <span className="text-xs font-medium text-gray-500">Color</span>
        <span className="w-5 h-5 rounded-full border border-gray-200" style={{ background: color }} />
        {brand ? <span className="text-xs text-gray-400">follows {brand}</span> : (
          <div className="flex flex-wrap gap-1.5">
            {MAP_PALETTE.map((c) => (
              <button key={c} onClick={() => setPickColor(c)} className={`w-6 h-6 rounded-full border-2 ${pickColor === c ? "border-gray-900" : "border-transparent"}`} style={{ background: c }} />
            ))}
          </div>
        )}
      </div>

      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex gap-2">
          {!node.__new && <Btn variant="danger" onClick={onDelete}><Trash2 size={15} />Delete</Btn>}
          {!node.__new && <Btn variant="ghost" onClick={onConnect}><Link2 size={15} />Connect</Btn>}
        </div>
        <div className="flex gap-2">
          <Btn variant="ghost" onClick={onCancel}>Cancel</Btn>
          <Btn variant="primary" onClick={() => onSave({ label: label.trim() || "Untitled", brand, type, description: desc.trim(), img, color })}><CheckCircle2 size={16} />Save</Btn>
        </div>
      </div>
    </div>
  );
}

function NodeSummary({ node, settings, nodes, edges, onEdit, onConnect, onDelete, onClose }) {
  const color = node.brand ? eventColor({ brand: node.brand }, "brand", settings) : (node.color || MAP_PALETTE[0]);
  const connected = edges
    .filter((e) => e.a === node.id || e.b === node.id)
    .map((e) => nodes.find((n) => n.id === (e.a === node.id ? e.b : e.a)))
    .filter(Boolean);
  return (
    <div>
      <div className="rounded-xl overflow-hidden mb-4" style={{ background: hexToRgba(color, 0.1) }}>
        {node.img
          ? <img src={node.img} alt="" className="w-full object-cover" style={{ maxHeight: 200 }} />
          : <div className="w-full flex items-center justify-center" style={{ height: 120 }}>
              <div className="w-20 h-20 rounded-2xl flex items-center justify-center text-white text-3xl font-semibold" style={{ background: color }}>{(node.label || "?").slice(0, 1).toUpperCase()}</div>
            </div>}
      </div>
      <div className="flex items-start gap-2 mb-2">
        <span className="w-3 h-3 rounded-full mt-1 shrink-0" style={{ background: color }} />
        <h3 className="text-lg font-semibold text-gray-900 leading-tight">{node.label}</h3>
      </div>
      <div className="flex flex-wrap gap-1.5 mb-3">
        {node.brand && <span className="text-xs font-medium px-2 py-0.5 rounded-full" style={{ background: hexToRgba(color, 0.14), color }}>{node.brand}</span>}
        {node.type && <span className="text-xs px-2 py-0.5 rounded-full bg-gray-100 text-gray-600">{node.type}</span>}
      </div>
      {node.description
        ? <p className="text-sm text-gray-600 leading-relaxed whitespace-pre-line mb-4">{node.description}</p>
        : <p className="text-sm text-gray-400 italic mb-4">No description yet. Tap Edit to add one and let AI tidy it up.</p>}
      {connected.length > 0 && (
        <div className="mb-4">
          <div className="text-xs font-medium text-gray-400 uppercase tracking-wide mb-1.5">Connects to</div>
          <div className="flex flex-wrap gap-1.5">
            {connected.map((c) => (
              <span key={c.id} className="inline-flex items-center gap-1.5 text-xs px-2 py-1 rounded-lg bg-gray-50 border border-gray-100 text-gray-700">
                <span className="w-2 h-2 rounded-full" style={{ background: c.brand ? eventColor({ brand: c.brand }, "brand", settings) : (c.color || MAP_PALETTE[0]) }} />{c.label}
              </span>
            ))}
          </div>
        </div>
      )}
      <div className="flex flex-wrap items-center justify-between gap-2 pt-1">
        <div className="flex gap-2">
          <Btn variant="danger" onClick={onDelete}><Trash2 size={15} />Delete</Btn>
          <Btn variant="ghost" onClick={onConnect}><Link2 size={15} />Connect</Btn>
        </div>
        <div className="flex gap-2">
          <Btn variant="ghost" onClick={onClose}>Close</Btn>
          <Btn variant="primary" onClick={onEdit}><Pencil size={15} />Edit</Btn>
        </div>
      </div>
    </div>
  );
}

function IdeaMapPage({ map, ideas, settings, save, toast, confirm }) {
  const [connectMode, setConnectMode] = useState(false);
  const [linkSource, setLinkSource] = useState(null);
  const [editing, setEditing] = useState(null);
  const [viewing, setViewing] = useState(null);
  const [view, setView] = useState({ zoom: 1, x: 0, y: 0 });
  const [, forceTick] = useState(0);

  const wrapRef = useRef(null);
  const simRef = useRef({});        // live positions {id:{x,y,vx,vy}}
  const draggingRef = useRef(null); // {id} while dragging a node
  const pinnedRef = useRef({});     // nodes held by drag
  const rafRef = useRef(null);
  const panRef = useRef(null);
  const viewRef = useRef(view);
  viewRef.current = view;

  const nodes = map.nodes || [];
  const edges = map.edges || [];
  const nodesRef = useRef(nodes); nodesRef.current = nodes;
  const edgesRef = useRef(edges); edgesRef.current = edges;

  // seed simulation positions from stored coords (or scatter)
  useEffect(() => {
    const sim = simRef.current || {};
    nodes.forEach((n, i) => {
      if (!sim[n.id]) {
        const angle = i * 2.39996, rad = 60 + i * 26;
        sim[n.id] = {
          x: typeof n.x === "number" ? n.x : 400 + Math.cos(angle) * rad,
          y: typeof n.y === "number" ? n.y : 300 + Math.sin(angle) * rad,
          vx: 0, vy: 0,
        };
      }
    });
    Object.keys(sim).forEach((id) => { if (!nodes.find((n) => n.id === id)) delete sim[id]; });
    simRef.current = sim;
  }, [nodes]);

  // physics loop
  useEffect(() => {
    let running = true;
    const tick = () => {
      if (!running) return;
      const sim = simRef.current;
      const ns = nodesRef.current, es = edgesRef.current;
      const ids = ns.map((n) => n.id);
      const cx = 460, cy = 320;
      // repulsion
      for (let i = 0; i < ids.length; i++) {
        const a = sim[ids[i]]; if (!a) continue;
        for (let j = i + 1; j < ids.length; j++) {
          const b = sim[ids[j]]; if (!b) continue;
          let dx = a.x - b.x, dy = a.y - b.y;
          let dist2 = dx * dx + dy * dy; if (dist2 < 1) dist2 = 1;
          const dist = Math.sqrt(dist2);
          const force = 5200 / dist2;
          const fx = (dx / dist) * force, fy = (dy / dist) * force;
          a.vx += fx; a.vy += fy; b.vx -= fx; b.vy -= fy;
        }
        // gentle pull to center
        a.vx += (cx - a.x) * 0.0016;
        a.vy += (cy - a.y) * 0.0016;
      }
      // spring along edges
      es.forEach((ed) => {
        const a = sim[ed.a], b = sim[ed.b]; if (!a || !b) return;
        const dx = b.x - a.x, dy = b.y - a.y;
        const dist = Math.sqrt(dx * dx + dy * dy) || 1;
        const target = 150;
        const k = (dist - target) * 0.015;
        const fx = (dx / dist) * k, fy = (dy / dist) * k;
        a.vx += fx; a.vy += fy; b.vx -= fx; b.vy -= fy;
      });
      // integrate
      let moving = false;
      ids.forEach((id) => {
        const p = sim[id]; if (!p) return;
        if (pinnedRef.current[id]) { p.vx = 0; p.vy = 0; return; }
        p.vx *= 0.86; p.vy *= 0.86;
        if (p.vx > 30) p.vx = 30; if (p.vx < -30) p.vx = -30;
        if (p.vy > 30) p.vy = 30; if (p.vy < -30) p.vy = -30;
        p.x += p.vx; p.y += p.vy;
        if (Math.abs(p.vx) + Math.abs(p.vy) > 0.4) moving = true;
      });
      forceTick((t) => (t + 1) % 1000000);
      rafRef.current = requestAnimationFrame(tick);
    };
    rafRef.current = requestAnimationFrame(tick);
    return () => { running = false; cancelAnimationFrame(rafRef.current); };
  }, []);

  // persist positions (debounced) when sim settles or on unmount
  const persistPositions = () => {
    const sim = simRef.current;
    const updated = nodesRef.current.map((n) => sim[n.id] ? { ...n, x: Math.round(sim[n.id].x), y: Math.round(sim[n.id].y) } : n);
    save({ ...map, nodes: updated });
  };

  const screenToWorld = (clientX, clientY) => {
    const rect = wrapRef.current.getBoundingClientRect();
    const v = viewRef.current;
    return {
      x: (clientX - rect.left - v.x) / v.zoom,
      y: (clientY - rect.top - v.y) / v.zoom,
    };
  };

  // node drag (pointer events: works on touch + mouse)
  const startNodeDrag = (e, node) => {
    e.stopPropagation();
    if (connectMode) { handleTap(node); return; }
    if (e.button != null && e.button > 0) return;
    const sim = simRef.current;
    let moved = false;
    const startW = screenToWorld(e.clientX, e.clientY);
    const offX = sim[node.id].x - startW.x, offY = sim[node.id].y - startW.y;
    pinnedRef.current[node.id] = true;
    draggingRef.current = node.id;
    const move = (me) => {
      const w = screenToWorld(me.clientX, me.clientY);
      const nx = w.x + offX, ny = w.y + offY;
      if (!moved && Math.hypot(nx - sim[node.id].x, ny - sim[node.id].y) > 3) moved = true;
      if (moved && me.cancelable) me.preventDefault();
      sim[node.id].x = nx; sim[node.id].y = ny; sim[node.id].vx = 0; sim[node.id].vy = 0;
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      window.removeEventListener("pointercancel", up);
      delete pinnedRef.current[node.id];
      draggingRef.current = null;
      if (moved) persistPositions();
      else handleTap(node);
    };
    window.addEventListener("pointermove", move, { passive: false });
    window.addEventListener("pointerup", up);
    window.addEventListener("pointercancel", up);
  };

  const handleTap = (node) => {
    if (connectMode) {
      if (!linkSource) { setLinkSource(node.id); return; }
      if (linkSource === node.id) { setLinkSource(null); return; }
      const exists = edges.some((ed) => (ed.a === linkSource && ed.b === node.id) || (ed.a === node.id && ed.b === linkSource));
      if (!exists) { save({ ...map, edges: [...edges, { id: uid(), a: linkSource, b: node.id }] }); toast("Connected"); }
      setLinkSource(null);
      return;
    }
    setViewing(node);
  };

  // background pan
  const startPan = (e) => {
    if (linkSource) { setLinkSource(null); return; }
    const v = viewRef.current;
    const sx = e.clientX, sy = e.clientY, ox = v.x, oy = v.y;
    let moved = false;
    const move = (me) => {
      if (!moved && Math.hypot(me.clientX - sx, me.clientY - sy) > 3) moved = true;
      if (moved) setView((p) => ({ ...p, x: ox + (me.clientX - sx), y: oy + (me.clientY - sy) }));
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };

  const onWheel = (e) => {
    e.preventDefault();
    const rect = wrapRef.current.getBoundingClientRect();
    const mx = e.clientX - rect.left, my = e.clientY - rect.top;
    setView((p) => {
      const z = Math.min(2.2, Math.max(0.4, p.zoom * (e.deltaY < 0 ? 1.1 : 0.9)));
      const k = z / p.zoom;
      return { zoom: z, x: mx - (mx - p.x) * k, y: my - (my - p.y) * k };
    });
  };
  const zoomBy = (f) => setView((p) => {
    const z = Math.min(2.2, Math.max(0.4, p.zoom * f));
    const rect = wrapRef.current?.getBoundingClientRect();
    const mx = rect ? rect.width / 2 : 300, my = rect ? rect.height / 2 : 250;
    const k = z / p.zoom;
    return { zoom: z, x: mx - (mx - p.x) * k, y: my - (my - p.y) * k };
  });
  const resetView = () => setView({ zoom: 1, x: 0, y: 0 });

  const addNode = () => {
    const c = screenToWorld((wrapRef.current?.getBoundingClientRect().width || 600) / 2 + (wrapRef.current?.getBoundingClientRect().left || 0),
      (wrapRef.current?.getBoundingClientRect().height || 500) / 2 + (wrapRef.current?.getBoundingClientRect().top || 0));
    setEditing({ __new: true, x: c.x, y: c.y, color: MAP_PALETTE[Math.floor(Math.random() * MAP_PALETTE.length)] });
  };
  const saveNode = (vals) => {
    if (editing.__new) {
      const id = uid();
      simRef.current[id] = { x: editing.x, y: editing.y, vx: 0, vy: 0 };
      save({ ...map, nodes: [...nodes, { id, x: Math.round(editing.x), y: Math.round(editing.y), ...vals }] });
    } else {
      save({ ...map, nodes: nodes.map((n) => (n.id === editing.id ? { ...n, ...vals } : n)) });
    }
    setEditing(null); toast("Saved");
  };
  const deleteNode = () => {
    const id = editing.id;
    delete simRef.current[id];
    save({ ...map, nodes: nodes.filter((n) => n.id !== id), edges: edges.filter((ed) => ed.a !== id && ed.b !== id) });
    setEditing(null); toast("Node deleted");
  };
  const deleteEdge = (id) => confirm({
    title: "Remove this connection?", message: "The two nodes stay, only the link is removed.", confirmText: "Remove", danger: true,
    onConfirm: () => { save({ ...map, edges: edges.filter((ed) => ed.id !== id) }); toast("Connection removed"); },
  });
  const importIdeas = () => {
    const have = new Set(nodes.map((n) => (n.label || "").toLowerCase()));
    const fresh = (ideas || []).filter((i) => !i.deleted && i.idea && !have.has(i.idea.toLowerCase()));
    if (!fresh.length) { toast("No new ideas to add"); return; }
    const startN = nodes.length;
    const added = fresh.slice(0, 40).map((i, k) => {
      const id = uid();
      const ang = (startN + k) * 2.39996, rad = 70 + (startN + k) * 24;
      const x = 460 + Math.cos(ang) * rad, y = 320 + Math.sin(ang) * rad;
      simRef.current[id] = { x, y, vx: 0, vy: 0 };
      return { id, label: i.idea, brand: i.brand || "", type: "Idea", description: i.notes || "", img: "", color: i.brand ? eventColor({ brand: i.brand }, "brand", settings) : MAP_PALETTE[(startN + k) % MAP_PALETTE.length], x: Math.round(x), y: Math.round(y) };
    });
    save({ ...map, nodes: [...nodes, ...added] });
    toast(`Added ${added.length} idea${added.length === 1 ? "" : "s"}`);
  };

  const sim = simRef.current || {};
  const posOf = (id) => sim[id] || { x: 0, y: 0 };

  return (
    <div>
      <PageHead title="Idea Map">
        <Btn variant={connectMode ? "soft" : "ghost"} onClick={() => { setConnectMode((m) => !m); setLinkSource(null); }}><Link2 size={15} />{connectMode ? "Connecting" : "Connect"}</Btn>
        <Btn variant="ghost" onClick={importIdeas}><Lightbulb size={15} />Add from Ideas</Btn>
        <Btn variant="primary" onClick={addNode}><Plus size={16} />Add node</Btn>
      </PageHead>
      <p className="text-xs text-gray-400 mb-3">
        {connectMode ? "Connect mode on. Tap one node, then another, to link them. Tap a line to remove it." : "Drag a node to pull it around, the others settle around it. Drag the empty board to pan, scroll or use the buttons to zoom. Tap a node to edit or add an image."}
      </p>

      <div ref={wrapRef} onPointerDown={startPan} onWheel={onWheel}
        className="relative overflow-hidden rounded-2xl border border-gray-200 select-none"
        style={{ height: 600, background: "#0b1220", touchAction: "none", cursor: "grab" }}>

        {/* grid that pans/zooms with the view */}
        <div className="absolute inset-0" style={{
          backgroundImage: "linear-gradient(rgba(255,255,255,.05) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.05) 1px,transparent 1px)",
          backgroundSize: `${40 * view.zoom}px ${40 * view.zoom}px`,
          backgroundPosition: `${view.x}px ${view.y}px`,
        }} />

        {/* zoom controls */}
        <div className="absolute top-3 right-3 z-20 flex flex-col gap-1">
          <button onClick={() => zoomBy(1.2)} className="w-8 h-8 rounded-lg bg-white/10 hover:bg-white/20 text-white flex items-center justify-center backdrop-blur"><Plus size={16} /></button>
          <button onClick={() => zoomBy(0.83)} className="w-8 h-8 rounded-lg bg-white/10 hover:bg-white/20 text-white flex items-center justify-center backdrop-blur text-lg leading-none">−</button>
          <button onClick={resetView} title="Reset view" className="w-8 h-8 rounded-lg bg-white/10 hover:bg-white/20 text-white flex items-center justify-center backdrop-blur"><Maximize2 size={14} /></button>
        </div>

        {nodes.length === 0 && (
          <div className="absolute inset-0 flex items-center justify-center pointer-events-none">
            <div className="text-center" style={{ color: "#7c89a3" }}>
              <Share2 className="mx-auto mb-2" />
              <p className="text-sm">Your map is empty. Add a node, or pull in your ideas.</p>
            </div>
          </div>
        )}

        {/* world layer: transformed by pan + zoom */}
        <div className="absolute" style={{ left: 0, top: 0, transform: `translate(${view.x}px,${view.y}px) scale(${view.zoom})`, transformOrigin: "0 0" }}>
          <svg style={{ position: "absolute", left: -2000, top: -2000, overflow: "visible", pointerEvents: "none" }} width="4000" height="4000">
            <g transform="translate(2000,2000)">
              {edges.map((ed) => {
                const a = posOf(ed.a), b = posOf(ed.b);
                const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2 - 22;
                return (
                  <g key={ed.id}>
                    <path d={`M ${a.x} ${a.y} Q ${mx} ${my} ${b.x} ${b.y}`} fill="none" stroke="rgba(120,170,255,.45)" strokeWidth={1.6} />
                    <path d={`M ${a.x} ${a.y} Q ${mx} ${my} ${b.x} ${b.y}`} fill="none" stroke="transparent" strokeWidth={18}
                      style={{ pointerEvents: "stroke", cursor: "pointer" }} onClick={(e) => { e.stopPropagation(); deleteEdge(ed.id); }} />
                  </g>
                );
              })}
            </g>
          </svg>

          {nodes.map((n) => {
            const p = posOf(n.id);
            const isSrc = linkSource === n.id;
            const nc = (n.brand && settings && settings.brands ? eventColor({ brand: n.brand }, "brand", settings) : n.color) || MAP_PALETTE[0];
            return (
              <div key={n.id} onPointerDown={(e) => startNodeDrag(e, n)}
                style={{ position: "absolute", left: p.x, top: p.y, transform: "translate(-50%,-50%)", touchAction: "none" }}
                className="cursor-grab active:cursor-grabbing">
                <div className="rounded-2xl overflow-hidden"
                  style={{
                    width: NODE_R * 2, height: NODE_R * 2,
                    border: `2px solid ${nc}`,
                    boxShadow: isSrc ? `0 0 0 4px ${nc}, 0 0 26px ${nc}` : `0 0 22px ${hexToRgba(nc, 0.55)}, 0 4px 14px rgba(0,0,0,.5)`,
                  }}>
                  {n.img
                    ? <img src={n.img} alt="" className="w-full h-full object-cover" draggable={false} />
                    : <div className="w-full h-full flex items-center justify-center text-white text-xl font-semibold" style={{ background: nc }}>{(n.label || "?").slice(0, 1).toUpperCase()}</div>}
                </div>
                <div className="text-center mt-1.5 leading-tight whitespace-nowrap" style={{ transform: "translateX(-50%)", marginLeft: NODE_R, position: "absolute", left: 0 }}>
                  <div className="text-xs font-medium" style={{ color: "#e7ebf3", textShadow: "0 1px 4px rgba(0,0,0,.9)" }}>{n.label}</div>
                  {n.brand && <div className="text-[10px]" style={{ color: hexToRgba("#ffffff", 0.6), textShadow: "0 1px 4px rgba(0,0,0,.9)" }}>{n.brand}</div>}
                </div>
              </div>
            );
          })}
        </div>
      </div>

      <div className="flex items-center justify-between mt-2 text-xs text-gray-400">
        <span>{nodes.length} node{nodes.length === 1 ? "" : "s"} · {edges.length} connection{edges.length === 1 ? "" : "s"} · {Math.round(view.zoom * 100)}%</span>
        <span>Drag the board to pan · scroll to zoom</span>
      </div>

      <Modal open={!!viewing} onClose={() => setViewing(null)} title="">
        {viewing && (
          <NodeSummary node={viewing} settings={settings} nodes={nodes} edges={edges}
            onEdit={() => { const n = viewing; setViewing(null); setEditing(n); }}
            onConnect={() => { setConnectMode(true); setLinkSource(viewing.id); setViewing(null); toast("Now tap another node to link"); }}
            onDelete={() => { const id = viewing.id; delete simRef.current[id]; save({ ...map, nodes: nodes.filter((n) => n.id !== id), edges: edges.filter((ed) => ed.a !== id && ed.b !== id) }); setViewing(null); toast("Node deleted"); }}
            onClose={() => setViewing(null)} />
        )}
      </Modal>

      <Modal open={!!editing} onClose={() => setEditing(null)} title={editing && editing.__new ? "New node" : "Edit node"}>
        {editing && (
          <NodeEditor node={editing} settings={settings} onSave={saveNode} onDelete={deleteNode} toast={toast}
            onConnect={() => { setConnectMode(true); setLinkSource(editing.id); setEditing(null); toast("Now tap another node to link"); }}
            onCancel={() => setEditing(null)} />
        )}
      </Modal>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* Glossary                                                             */
/* ------------------------------------------------------------------ */
const GLOSSARY = [
  { term: "Home", cat: "Sections", def: "The command center. Read-only overview pulling counts and recent items from every section. Shows your daily briefing, KPI cards, brand activity chart, and quick panels for follow-ups and recent activity." },
  { term: "Calendar", cat: "Sections", def: "Anything time-bound: tasks, meetings, campaign dates, deadlines, presentations. Tap a day to add, drag an event to reschedule, and switch between Month, Week, Day, and Agenda views." },
  { term: "Work Log", cat: "Sections", def: "A record of what you actually worked on. One entry per piece of work. Use it to track execution and build a searchable history of your output across brands and time." },
  { term: "Ideas Bank", cat: "Sections", def: "Where ideas live before they become commitments. Opportunities, observations, what-ifs. Convert promising ideas into Recommendations or Calendar events directly from the table." },
  { term: "Findings & Insights", cat: "Sections", def: "Things discovered during analysis, reviews, or research. A finding is something you learned, not something you did (Work) or might do (Ideas)." },
  { term: "Recommendations", cat: "Sections", def: "Formal proposals shared with the team or leadership. Tracks status from Draft to Implemented and records the actual decision made, so nothing falls through." },
  { term: "Quarterly Roadmap", cat: "Sections", def: "Forward planning as a brand-by-month/quarter grid. Type directly into cells. Plan at the initiative level. Your Q3 marketing calendar is seeded here." },
  { term: "Wins & Achievements", cat: "Sections", def: "Completed work that created value. Your track record for performance reviews and for knowing what actually worked across brands and quarters." },
  { term: "Brand Hub", cat: "Sections", def: "One row per brand showing current state: status, monthly focus, key issue, key opportunity, and next action. A living scorecard, not a log." },
  { term: "Meeting Notes", cat: "Sections", def: "Structured record of meetings: attendees, decisions, action items, follow-up dates. Keeps ownership and outcomes clear after the call ends." },
  { term: "Files & Links", cat: "Sections", def: "Registry of where things live: reports, decks, Drive links, references. Find them without searching your inbox or Slack." },
  { term: "Backlog & Archive", cat: "Sections", def: "Items from all sections that were parked (Backlog/Saved for Later) or archived. Includes a Trash tab. Restore anything, export everything, review by quarter." },
  { term: "Idea Map", cat: "Sections", def: "A visual canvas of ideas as draggable nodes you connect with lines. Each node can carry an image (upload or URL). Good for mapping campaign structures or brand directions." },
  { term: "Glossary", cat: "Sections", def: "This page. Definitions for every section, field, status, and concept in the hub. Searchable." },
  { term: "Settings", cat: "Sections", def: "Manage all dropdown lists: brands, categories, statuses, platforms, channels, and more. Edit once and every section updates automatically." },
  { term: "Hub Assistant", cat: "Sections", def: "The AI chatbot (green button, bottom right). Understands plain language, adds records to the right section, navigates the hub, and accepts image and link attachments as context." },
  { term: "Not Started", cat: "Statuses", def: "The item exists but work has not begun." },
  { term: "In Progress", cat: "Statuses", def: "Actively being worked on right now." },
  { term: "Waiting", cat: "Statuses", def: "Your part is done. Waiting on someone else or an external dependency." },
  { term: "Completed", cat: "Statuses", def: "Done. Counts toward the completed total on the Home dashboard." },
  { term: "On Hold", cat: "Statuses", def: "Paused intentionally, with the expectation of resuming." },
  { term: "Cancelled", cat: "Statuses", def: "Will not happen. Closed without completion." },
  { term: "Saved for Later", cat: "Statuses", def: "Not urgent but worth keeping. Appears in the Backlog & Archive view." },
  { term: "Backlog", cat: "Statuses", def: "Parked for now. Not abandoned, just not active. Review quarterly." },
  { term: "Archived", cat: "Statuses", def: "Closed and stored for reference. Still searchable, hidden from active views." },
  { term: "Draft", cat: "Recommendation Statuses", def: "Recommendation written but not yet shared with anyone." },
  { term: "Shared", cat: "Recommendation Statuses", def: "Submitted to stakeholders. Awaiting response." },
  { term: "Under Review", cat: "Recommendation Statuses", def: "Actively being evaluated by decision-makers." },
  { term: "Approved", cat: "Recommendation Statuses", def: "Greenlit. Not yet executed." },
  { term: "Implemented", cat: "Recommendation Statuses", def: "Approved and fully executed." },
  { term: "Rejected", cat: "Recommendation Statuses", def: "Reviewed and declined by stakeholders." },
  { term: "Soft Delete", cat: "Concepts", def: "Deleting an item moves it to Trash, not permanent removal. You can restore it at any time from the Trash tab in Backlog & Archive. Nothing is truly lost by default." },
  { term: "Permanent Delete", cat: "Concepts", def: "Final removal of a record. Requires two separate confirmation clicks. Cannot be undone." },
  { term: "Quarter", cat: "Concepts", def: "Q1 = January to March, Q2 = April to June, Q3 = July to September, Q4 = October to December. Auto-calculated from the date on every entry." },
  { term: "ISO Week Number", cat: "Concepts", def: "The standard week number of the year, calculated automatically from the entry date. W1 starts on the Monday of the first week containing a Thursday of the year." },
  { term: "Auto-derived Fields", cat: "Concepts", def: "Month, Quarter, Year, and Week Number are calculated from the Date field you enter. You never need to fill them in manually. They appear as gray chips in the form preview." },
  { term: "Persistent Storage", cat: "Concepts", def: "Your data is saved to the artifact's built-in storage and survives refresh, logout, and restart. It does not reset between sessions." },
  { term: "Backup", cat: "Concepts", def: "A JSON file downloaded via the sidebar. Contains all data across every section and settings. Use it before opening an updated version of the hub." },
  { term: "Restore", cat: "Concepts", def: "Loading a backup JSON file to bring your data into a new version of the hub. Replaces all current data." },
  { term: "Seed Data", cat: "Concepts", def: "Pre-loaded content that appears on first open. The Q3 2026 marketing calendar from your planning sheet is seeded automatically once. The brand hub is seeded with your 11 brands." },
  { term: "Conversation History", cat: "Concepts", def: "The Hub Assistant remembers the last few exchanges in a session so follow-up messages make sense. For example: 'Add a Secco idea' then 'Make it high priority' works as expected." },
  { term: "Brand", cat: "Fields", def: "Which Seazen brand this entry belongs to. Defined in Settings and shared across all sections." },
  { term: "Platform", cat: "Fields", def: "The delivery, social, or digital surface involved: Talabat, Deliveroo, Instagram, TikTok, Google, and so on." },
  { term: "Channel", cat: "Fields", def: "The broader marketing channel type: Delivery App, Social Media, Paid Media, CRM, In-store, etc." },
  { term: "Category", cat: "Fields", def: "The type of work or topic: Marketing, Commercial, Menu, Campaign, Research, Report, and so on." },
  { term: "Priority", cat: "Fields", def: "High, Medium, or Low. Color-coded red, amber, and green. Used to filter and sort across every section." },
  { term: "Impact", cat: "Fields", def: "The expected business outcome: Revenue, Retention, Brand Awareness, Operational Efficiency, etc. Used in Ideas Bank and Wins." },
  { term: "Importance", cat: "Fields", def: "In Findings, how significant a discovery is: Critical, High, Medium, or Low." },
  { term: "Owner", cat: "Fields", def: "The person responsible for this item. Defaults to you but can be assigned to any team member." },
  { term: "Follow-up Date", cat: "Fields", def: "A specific date by which a follow-up action is needed. Items with upcoming follow-up dates appear in the Home briefing under 'Follow-ups due this week'." },
  { term: "Outcome", cat: "Fields", def: "In the Work Log, what actually resulted from the work. The 'so what' of the entry." },
  { term: "Expected Impact", cat: "Fields", def: "In Recommendations, the category of benefit you expect if the recommendation is implemented." },
  { term: "Next Step", cat: "Fields", def: "In Ideas Bank, the single clearest next action needed to move this idea forward." },
  { term: "Review Date", cat: "Fields", def: "In Ideas Bank, a date to revisit the idea and decide whether to advance, park, or drop it." },
  { term: "Decision", cat: "Fields", def: "In Recommendations, the actual decision made by stakeholders once the recommendation was reviewed." },
  { term: "Shared With", cat: "Fields", def: "In Recommendations, who the recommendation was presented or sent to (name, team, or role)." },
  { term: "Repeat", cat: "Fields", def: "In Calendar events, whether the event recurs: None, Daily, Weekly, or Monthly. Recurring events expand across the calendar but can only be edited at the series level." },
  { term: "Event Type", cat: "Fields", def: "The kind of calendar entry: Task, Meeting, Follow-up, Brand Review, Campaign, Deadline, Report, Presentation, or Quarterly Initiative." },
  { term: "Related Project", cat: "Fields", def: "Free-text field linking an entry to a broader project or initiative name." },
  { term: "Business Impact", cat: "Fields", def: "In Wins, the category of value the achievement delivered: Revenue, Retention, Brand Awareness, etc." },
  { term: "Linked Event", cat: "Fields", def: "In the Work Log, a reference to a related Calendar event (type the event title)." },
  { term: "Action Items", cat: "Fields", def: "In Meeting Notes, the tasks assigned during the meeting. Usually formatted as bullet points with owners." },
  { term: "Source", cat: "Fields", def: "In Findings, where the insight came from: a review, a meeting, research, a report, a platform dashboard, etc." },
];

function GlossaryPage() {
  const [q, setQ] = useState("");
  const filtered = q.trim()
    ? GLOSSARY.filter(t => (t.term + " " + t.def + " " + t.cat).toLowerCase().includes(q.toLowerCase().trim()))
    : GLOSSARY;
  const cats = [...new Set(GLOSSARY.map(t => t.cat))];
  return (
    <div>
      <PageHead title="Glossary" />
      <p className="text-sm text-gray-500 mb-4">{GLOSSARY.length} definitions covering every section, field, status, and concept in the hub.</p>
      <div className="relative mb-5 max-w-md">
        <Search size={15} className="absolute left-3 top-2.5 text-gray-400" />
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="Search terms..."
          className="w-full pl-9 pr-8 py-2 border border-gray-200 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-emerald-200" />
        {q && <button onClick={() => setQ("")} className="absolute right-2.5 top-2.5 text-gray-400 hover:text-gray-700"><X size={14} /></button>}
      </div>
      {filtered.length === 0 && <p className="text-gray-400 text-sm text-center py-10">No terms match that search.</p>}
      {cats.map(cat => {
        const items = filtered.filter(t => t.cat === cat);
        if (!items.length) return null;
        return (
          <div key={cat} className="mb-5">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-emerald-700 mb-2 px-1">{cat}</h3>
            <div className="bg-white border border-gray-200 rounded-xl divide-y divide-gray-100">
              {items.map((t, i) => (
                <div key={i} className="px-4 py-3 hover:bg-gray-50/60">
                  <div className="font-semibold text-gray-900 text-sm">{t.term}</div>
                  <div className="text-sm text-gray-500 mt-0.5 leading-relaxed">{t.def}</div>
                </div>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* AI Chatbot                                                           */
/* ------------------------------------------------------------------ */
const BOT_ACTIONS_DOC = `
AVAILABLE ACTIONS — respond with strict JSON only, no markdown, no backticks, no text outside the JSON object:

{
  "action": "add|navigate|complete|archive|backlog|update|answer|setting_add",
  "collection": "work|ideas|findings|recommendations|wins|meetings|files|events|brandHub|null",
  "data": { field values — omit what you don't know },
  "itemRef": "partial title or ID of existing item for complete/archive/backlog/update, else null",
  "reply": "One short friendly sentence confirming what you did or answering the question",
  "navigate": "section id to open after action, or null",
  "settingKey": "brands|categories|platforms|channels|priorities|statuses|null",
  "settingValue": "value to add to that settings list, else null"
}

ACTION RULES:
- add: create a new record in collection. Set navigate = collection so user can see it.
- navigate: go to a section. Set navigate = section id. No collection needed.
- complete: mark an existing item Completed. Use itemRef to identify it.
- archive: mark an existing item Archived. Use itemRef to identify it.
- backlog: move an existing item to Backlog status. Use itemRef to identify it.
- update: update fields of an existing item. Use itemRef to identify it, put new values in data.
- answer: answer a question or clarify. No data change.
- setting_add: add a new value (new brand, category, platform, etc.) to settings.

SECTION IDs FOR NAVIGATE: home, events, work, ideas, findings, recommendations, roadmap, wins, brandHub, meetings, files, backlog, settings, ideamap, glossary
`;

function ChatBot({ data, settings, api, go, saveSettings, toast }) {
  const [open, setOpen] = useState(false);
  const [msgs, setMsgs] = useState([{
    role: "bot",
    text: "Hi Suleman. Tell me what to add, log, or do — or ask me anything. You can also attach images as visual context before sending.",
  }]);
  const [input, setInput] = useState("");
  const [busy, setBusy] = useState(false);
  const [attachments, setAttachments] = useState([]);
  const bottomRef = useRef(null);
  const fileRef = useRef(null);
  const inputRef = useRef(null);
  const historyRef = useRef([]);

  useEffect(() => {
    if (open) setTimeout(() => bottomRef.current?.scrollIntoView({ behavior: "smooth" }), 60);
  }, [msgs, open]);

  const SECTION_LABELS = {
    home: "Home", events: "Calendar", work: "Work Log", ideas: "Ideas Bank",
    findings: "Findings", recommendations: "Recommendations", roadmap: "Roadmap",
    wins: "Wins", brandHub: "Brand Hub", meetings: "Meeting Notes",
    files: "Files & Links", backlog: "Backlog", settings: "Settings",
    ideamap: "Idea Map", glossary: "Glossary",
  };

  const buildSysPrompt = () => {
    const counts = {};
    Object.keys(COLLECTIONS).forEach(ck => { counts[ck] = (data[ck] || []).filter(r => !r.deleted).length; });
    const snap = (ck, fields) => (data[ck] || []).filter(r => !r.deleted).slice(0, 8)
      .map(r => `[${r.id}] ${fields.map(f => r[f] || "").filter(Boolean).join(" | ")}`).join("\n");
    return `You are the AI assistant inside Suleman Alhabashi's Seazen Work Hub dashboard.
Suleman is Senior Marketing Lead at Seazen Group (Kuwait), managing 11 F&B brands across UAE, Oman, Bahrain, Qatar, and Kuwait.
Today: ${todayKey()}

BRANDS: ${settings.brands.join(", ")}
CATEGORIES: ${settings.categories.join(", ")}
PLATFORMS: ${settings.platforms.join(", ")}
CHANNELS: ${settings.channels.join(", ")}
PRIORITIES: High, Medium, Low
STATUSES: Not Started, In Progress, Waiting, Completed, On Hold, Saved for Later, Backlog, Archived
REC STATUSES: Draft, Shared, Under Review, Approved, Implemented, Rejected
EVENT TYPES: Task, Meeting, Follow-up, Brand Review, Campaign, Deadline, Report, Presentation, Quarterly Initiative
IMPACT TYPES: Revenue, Profitability, Customer Experience, Retention, Brand Awareness, Operational Efficiency, Product Innovation, Menu Improvement, Delivery Improvement, Cost Control, Other
IMPORTANCE: Critical, High, Medium, Low

RECORD COUNTS: ${JSON.stringify(counts)}

RECENT WORK LOG:
${snap("work", ["project", "brand", "date", "status"]) || "none"}

RECENT IDEAS:
${snap("ideas", ["idea", "brand", "priority", "status"]) || "none"}

RECENT FINDINGS:
${snap("findings", ["finding", "brand", "importance"]) || "none"}

RECENT RECOMMENDATIONS:
${snap("recommendations", ["recommendation", "brand", "status"]) || "none"}

UPCOMING CALENDAR EVENTS:
${(data.events || []).filter(r => !r.deleted && r.date >= todayKey()).slice(0, 8).map(r => `[${r.id}] ${r.title || ""} (${r.brand || ""}, ${r.date || ""})`).join("\n") || "none"}

SECTION FIELDS:
work: date, time, brand, platform, channel, project, category, priority, status, owner, notes, outcome
ideas: date, brand, platform, channel, idea, category, impact, priority, status, owner, notes, nextStep, reviewDate
findings: date, brand, platform, channel, finding, source, category, importance, notes, relatedLink
recommendations: date, brand, platform, channel, recommendation, expectedImpact, priority, status, sharedWith, followUpDate, notes
wins: date, brand, achievement, businessImpact, notes, relatedProject
meetings: date, time, brand, meetingTitle, attendees, notes, decisions, actionItems, followUpDate, status
files: date, brand, fileName, link, type, category, notes, relatedProject
events: date, time, brand, platform, channel, title, eventType, category, priority, status, notes, followUpDate, repeat
brandHub: brand, status, monthFocus, quarterFocus, yearFocus, keyIssue, keyOpportunity, nextAction, priority, notes

${BOT_ACTIONS_DOC}

ADDITIONAL RULES:
- Always set date = today (${todayKey()}) if date not specified
- Default priority: Medium. Default status: Not Started (tasks), In Progress (active work)
- When adding, set navigate = collection id so user can view the new record
- If an image is attached: describe what you see, extract any relevant campaign/brand/creative info, incorporate into the record notes field
- If a URL is in the message: use it as the link or relatedLink field. If "add this link to files" → files collection with that URL
- When image shows a creative reference, competitor, or inspiration: add a meaningful finding or idea note
- For complete/archive/backlog/update: search the snapshot above for a matching itemRef (partial title works)
- For brandHub updates: action = "add", collection = "brandHub", include brand name in data
- If attaching multiple images: describe each briefly and synthesize into one record

STATUS AND SUMMARY QUESTIONS:
- If asked where things stand, for a status update, what is going on, what is due, or to summarize a brand, a section, or the whole hub, use action "answer" and write a genuinely useful summary grounded in the data snapshot above (counts, recent items, upcoming events).
- For these answer/summary requests the reply can be several sentences or short line-separated points. Lead with the headline, then the specifics. Be concrete: name brands, counts, and dates from the data.
- For every other action (add, navigate, complete, etc.) keep the reply to one short sentence.

VOICE (applies to every reply):
- Never use em dashes anywhere.
- Never use words or phrases that sound like AI filler: dive in, unlock, leverage, elevate, seamless, robust, supercharge, game-changer, in today's landscape, it's worth noting, navigate the complexities, at the end of the day.
- Write like a sharp, direct human colleague. Plain words. No hype. No preamble.
- If truly unclear, ask one specific clarifying question via action "answer"`;
  };

  const findItem = (collection, ref) => {
    if (!ref || !collection || !COLLECTIONS[collection]) return null;
    const arr = (data[collection] || []).filter(r => !r.deleted);
    const refL = ref.toLowerCase();
    return arr.find(r => {
      const title = (r.project || r.idea || r.finding || r.recommendation || r.achievement || r.title || r.meetingTitle || r.fileName || r.brand || "").toLowerCase();
      return title.includes(refL) || r.id === ref;
    }) || arr.find(r => {
      const all = Object.values(r).join(" ").toLowerCase();
      return refL.split(" ").filter(w => w.length > 3).some(w => all.includes(w));
    });
  };

  const executeAction = (parsed, urlsFromText) => {
    const action = parsed.action;
    const col = parsed.collection;
    try {
      if (action === "add" && col && COLLECTIONS[col]) {
        const rec = { date: todayKey(), ...parsed.data };
        if (urlsFromText.length) {
          const urlStr = urlsFromText.join(", ");
          if (col === "files" && !rec.link) rec.link = urlsFromText[0];
          else if (col === "findings" && !rec.relatedLink) rec.relatedLink = urlsFromText[0];
          else if (col === "recommendations" && !rec.notes) rec.notes = "Refs: " + urlStr;
          else if (rec.notes) rec.notes += "\nRefs: " + urlStr;
          else rec.notes = "Refs: " + urlStr;
        }
        api.add(col, rec);
        if (parsed.navigate) go(parsed.navigate);
      } else if ((action === "complete" || action === "archive" || action === "backlog") && col) {
        const statusMap = { complete: "Completed", archive: "Archived", backlog: "Backlog" };
        const item = findItem(col, parsed.itemRef);
        if (item) api.update(col, item.id, { ...item, status: statusMap[action] });
      } else if (action === "update" && col) {
        const item = findItem(col, parsed.itemRef);
        if (item && parsed.data) api.update(col, item.id, { ...item, ...parsed.data });
      } else if (action === "navigate" && parsed.navigate) {
        go(parsed.navigate);
      } else if (action === "setting_add" && parsed.settingKey && parsed.settingValue) {
        const cur = settings[parsed.settingKey] || [];
        if (!cur.includes(parsed.settingValue)) saveSettings({ ...settings, [parsed.settingKey]: [...cur, parsed.settingValue] });
      } else if (action === "add" && !col && parsed.navigate) {
        go(parsed.navigate);
      }
    } catch (e) {
      console.error("bot action error", e);
    }
  };

  const onFile = (e) => {
    Array.from(e.target.files || []).forEach(f => {
      if (f.size > 4000000) { toast("Keep images under 4MB"); return; }
      const r = new FileReader();
      r.onload = () => setAttachments(p => [...p, { type: f.type || "image/jpeg", data: r.result, name: f.name }].slice(0, 4));
      r.readAsDataURL(f);
    });
    e.target.value = "";
    inputRef.current?.focus();
  };

  const urlRe = /https?:\/\/[^\s]+/g;

  const send = async () => {
    const text = input.trim();
    if ((!text && !attachments.length) || busy) return;
    setInput("");

    const urls = text.match(urlRe) || [];
    const userMsg = { role: "user", text: text || "(attached image)", attachments: [...attachments] };
    setMsgs(p => [...p, userMsg]);
    const snap = [...attachments];
    setAttachments([]);
    setBusy(true);

    const content = [];
    snap.forEach(a => {
      if (a.data && a.data.startsWith("data:image")) {
        const b64 = a.data.split(",")[1];
        const mt = (a.type || "image/jpeg");
        content.push({ type: "image", source: { type: "base64", media_type: mt, data: b64 } });
      }
    });
    content.push({ type: "text", text: text || "Please analyze the attached image and add the relevant information to the hub." });

    const newHistory = [...historyRef.current, { role: "user", content }].slice(-8);
    historyRef.current = newHistory;

    try {
      const res = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1000,
          system: buildSysPrompt(),
          messages: newHistory,
        }),
      });
      if (!res.ok) {
        let detail = "";
        try { const err = await res.json(); detail = err.error?.message || JSON.stringify(err).slice(0, 160); }
        catch { detail = await res.text().catch(() => ""); }
        throw new Error(`(${res.status}) ${detail || "request rejected"}`);
      }
      const d = await res.json();
      const raw = (d.content || []).map(b => b.text || "").join("").trim();
      let parsed = null;
      try { parsed = JSON.parse(raw.replace(/^```json\n?|^```\n?|```$/gm, "").trim()); } catch {}

      const botReply = parsed?.reply || (typeof raw === "string" && raw && !raw.startsWith("{") ? raw : "Done.");
      historyRef.current = [...historyRef.current, { role: "assistant", content: botReply }].slice(-8);

      if (parsed && parsed.action && parsed.action !== "answer") {
        executeAction(parsed, urls);
        if (parsed.action !== "navigate") toast(botReply.slice(0, 72));
      }
      setMsgs(p => [...p, {
        role: "bot",
        text: botReply,
        navigate: parsed?.navigate || (parsed?.action === "add" ? parsed?.collection : null),
        action: parsed?.action,
      }]);
    } catch (err) {
      const m = err.message || "";
      const friendly = /401|403|auth|credential/i.test(m)
        ? "The assistant could not authenticate. The AI connection may need reconnecting in this chat."
        : /429|rate/i.test(m)
        ? "Rate limited for a moment. Wait a few seconds and try again."
        : /Failed to fetch|NetworkError/i.test(m)
        ? "Could not reach the AI service. Check your connection and try again."
        : `The AI request failed: ${m}`;
      setMsgs(p => [...p, { role: "bot", text: friendly }]);
    }
    setBusy(false);
  };

  const CHIPS = [
    "I have an idea for a campaign",
    "Log work I did today",
    "Record a meeting I just had",
    "Go to the Calendar",
    "How many ideas do I have?",
  ];

  const panelH = 520;

  return (
    <>
      <button onClick={() => setOpen(o => !o)} title="Hub Assistant"
        style={{ width: 52, height: 52, background: open ? "#1f2937" : "#059669", boxShadow: "0 4px 18px rgba(0,0,0,.28)" }}
        className="fixed bottom-5 right-5 z-[80] rounded-full text-white flex items-center justify-center transition-all hover:scale-105">
        {open ? <X size={21} /> : <MessageSquare size={21} />}
      </button>

      {open && (
        <div className="fixed z-[80] bg-white border border-gray-200 rounded-2xl shadow-2xl flex flex-col cal-slide overflow-hidden"
          style={{
            width: Math.min(390, (typeof window !== "undefined" ? window.innerWidth : 400) - 20),
            height: panelH,
            bottom: 76,
            right: 20,
          }}>

          {/* header */}
          <div className="flex items-center justify-between px-4 py-3 shrink-0" style={{ background: "#059669" }}>
            <div className="flex items-center gap-2 text-white">
              <MessageSquare size={16} />
              <span className="font-semibold text-sm">Hub Assistant</span>
              <span className="text-[11px] bg-white/20 px-1.5 py-0.5 rounded-full">Powered by Claude</span>
            </div>
            <button onClick={() => setOpen(false)} className="text-white/70 hover:text-white transition"><X size={17} /></button>
          </div>

          {/* messages */}
          <div className="flex-1 overflow-y-auto px-3 py-3 space-y-3 cal-scroll">
            {msgs.map((m, i) => (
              <div key={i} className={`flex items-end gap-2 ${m.role === "user" ? "justify-end" : "justify-start"}`}>
                {m.role === "bot" && (
                  <div className="w-6 h-6 rounded-full shrink-0 mb-0.5 flex items-center justify-center" style={{ background: "#059669" }}>
                    <MessageSquare size={12} className="text-white" />
                  </div>
                )}
                <div style={{ maxWidth: "83%" }}>
                  {m.attachments?.filter(a => a.data?.startsWith("data:image")).map((a, j) => (
                    <img key={j} src={a.data} alt={a.name}
                      className="rounded-xl mb-1.5 object-cover border border-gray-200"
                      style={{ maxHeight: 130, maxWidth: "100%" }} />
                  ))}
                  <div className={`rounded-2xl px-3.5 py-2.5 text-sm leading-snug ${m.role === "user" ? "text-white" : "bg-gray-100 text-gray-800"}`}
                    style={m.role === "user" ? { background: "#059669", borderBottomRightRadius: 6, whiteSpace: "pre-line" } : { borderBottomLeftRadius: 6, whiteSpace: "pre-line" }}>
                    {m.text}
                    {m.role === "bot" && m.navigate && SECTION_LABELS[m.navigate] && (
                      <button onClick={() => { go(m.navigate); setOpen(false); }}
                        className="block mt-1.5 text-xs font-semibold hover:underline" style={{ color: "#059669" }}>
                        View in {SECTION_LABELS[m.navigate]} →
                      </button>
                    )}
                  </div>
                </div>
              </div>
            ))}

            {busy && (
              <div className="flex items-end gap-2">
                <div className="w-6 h-6 rounded-full flex items-center justify-center shrink-0" style={{ background: "#059669" }}>
                  <MessageSquare size={12} className="text-white" />
                </div>
                <div className="bg-gray-100 rounded-2xl px-4 py-3 flex items-center gap-1.5">
                  {[0, 140, 280].map(d => (
                    <span key={d} className="w-1.5 h-1.5 bg-gray-400 rounded-full animate-bounce" style={{ animationDelay: d + "ms" }} />
                  ))}
                </div>
              </div>
            )}

            {msgs.length === 1 && !busy && (
              <div className="flex flex-wrap gap-1.5 pt-1">
                {CHIPS.map((s, i) => (
                  <button key={i} onClick={() => { setInput(s); inputRef.current?.focus(); }}
                    className="text-xs px-2.5 py-1.5 rounded-full border border-gray-200 text-gray-600 hover:bg-emerald-50 hover:border-emerald-300 hover:text-emerald-700 transition">
                    {s}
                  </button>
                ))}
              </div>
            )}
            <div ref={bottomRef} />
          </div>

          {/* attachment previews */}
          {attachments.length > 0 && (
            <div className="px-3 py-2 border-t border-gray-100 flex flex-wrap gap-2 shrink-0">
              {attachments.map((a, i) => (
                <div key={i} className="relative group">
                  <img src={a.data} alt={a.name}
                    className="h-12 w-12 object-cover rounded-xl border border-gray-200" />
                  <button onClick={() => setAttachments(p => p.filter((_, j) => j !== i))}
                    className="absolute -top-1.5 -right-1.5 w-4 h-4 rounded-full bg-gray-800 text-white flex items-center justify-center opacity-0 group-hover:opacity-100 transition">
                    <X size={9} />
                  </button>
                </div>
              ))}
            </div>
          )}

          {/* input */}
          <div className="p-3 border-t border-gray-100 shrink-0">
            <div className="flex items-end gap-2">
              <div className="flex-1 border border-gray-200 rounded-xl focus-within:ring-2 focus-within:ring-emerald-200 overflow-hidden bg-white">
                <textarea ref={inputRef} rows={2} value={input}
                  onChange={e => setInput(e.target.value)}
                  onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); send(); } }}
                  placeholder="Tell me what to add, do, or ask... (Enter sends)"
                  disabled={busy}
                  className="w-full px-3 pt-2.5 pb-1 text-sm resize-none focus:outline-none bg-white text-gray-900" />
                <div className="flex items-center gap-1 px-2 pb-2">
                  <button onClick={() => fileRef.current?.click()} title="Attach image"
                    className="p-1.5 rounded-lg text-gray-400 hover:text-emerald-600 hover:bg-emerald-50 transition">
                    <ImageIcon size={15} />
                  </button>
                  <input ref={fileRef} type="file" accept="image/*" multiple className="hidden" onChange={onFile} />
                  <span className="text-[11px] text-gray-300">Attach images as context</span>
                </div>
              </div>
              <button onClick={send}
                disabled={busy || (!input.trim() && !attachments.length)}
                className="p-2.5 rounded-xl text-white transition disabled:opacity-40 hover:opacity-90 shrink-0"
                style={{ background: "#059669" }}>
                <ChevronRight size={20} />
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}


const NAV = [
  { id: "home", label: "Home", icon: LayoutDashboard },
  { id: "events", label: "Calendar", icon: CalendarDays },
  { id: "work", label: "Work Log", icon: ClipboardList },
  { id: "ideas", label: "Ideas Bank", icon: Lightbulb },
  { id: "ideamap", label: "Idea Map", icon: Share2 },
  { id: "findings", label: "Findings", icon: Telescope },
  { id: "recommendations", label: "Recommendations", icon: Megaphone },
  { id: "roadmap", label: "Roadmap", icon: Route },
  { id: "wins", label: "Wins", icon: Trophy },
  { id: "brandHub", label: "Brand Hub", icon: Building2 },
  { id: "meetings", label: "Meeting Notes", icon: NotebookPen },
  { id: "files", label: "Files & Links", icon: FileText },
  { id: "backlog", label: "Backlog & Archive", icon: Archive },
  { id: "glossary", label: "Glossary", icon: BookOpen },
  { id: "settings", label: "Settings", icon: SettingsIcon },
];

export default function App() {
  const [loaded, setLoaded] = useState(false);
  const [data, setData] = useState({});
  const [settings, setSettings] = useState(DEFAULT_SETTINGS);
  const [section, setSection] = useState("home");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [toastMsg, setToastMsg] = useState(null);
  const [confirmState, setConfirmState] = useState(null);
  const [dark, setDark] = useState(false);
  const [now, setNow] = useState(new Date());
  const backupRef = useRef(null);

  const toast = (m) => { setToastMsg(m); setTimeout(() => setToastMsg(null), 1800); };
  const confirm = (cfg) => setConfirmState({ ...cfg, onConfirm: () => { setConfirmState(null); cfg.onConfirm(); } });
  useEffect(() => { const t = setInterval(() => setNow(new Date()), 30000); return () => clearInterval(t); }, []);
  const toggleDark = () => setDark((d) => { const nd = !d; saveKey(THEME_KEY, nd); return nd; });

  // initial load + seed
  useEffect(() => {
    (async () => {
      const s = await loadKey(SETTINGS_KEY, null);
      const settingsVal = s || DEFAULT_SETTINGS;
      if (!s) await saveKey(SETTINGS_KEY, DEFAULT_SETTINGS);
      const next = {};
      for (const ck of COLLECTION_KEYS) next[ck] = await loadKey(COLLECTIONS[ck].key, []);
      next.roadmap = await loadKey(ROADMAP_KEY, []);
      next.ideamap = await loadKey(IDEAMAP_KEY, { nodes: [], edges: [] });
      next.trends = await loadKey(TRENDS_KEY, { items: [], updatedAt: 0 });
      next.inspo = await loadKey(INSPO_KEY, []);
      const themeDark = await loadKey(THEME_KEY, false);
      // seed brand hub on first run
      if ((next.brandHub || []).length === 0) {
        next.brandHub = settingsVal.brands.map((b) => ({ id: uid(), brand: b, status: "Active", createdDate: todayKey(), updatedDate: todayKey(), deleted: false }));
        await saveKey(COLLECTIONS.brandHub.key, next.brandHub);
      }
      // one-time Q3 marketing calendar seed
      const seeded = await loadKey(SEED_FLAG, false);
      if (!seeded) {
        const seedEv = Q3_SEED.map((e) => ({
          ...e, ...derive(e.date), id: uid(), createdDate: todayKey(), updatedDate: todayKey(),
          deleted: false, repeat: e.repeat || "None", status: e.status || "Not Started",
        }));
        next.events = [...seedEv, ...(next.events || [])];
        await saveKey(COLLECTIONS.events.key, next.events);
        await saveKey(SEED_FLAG, true);
      }
      // one-time migration: drop "by Masgouf" from Okra in stored data
      const migOkra = await loadKey(MIG_OKRA, false);
      if (!migOkra) {
        if (Array.isArray(settingsVal.brands)) settingsVal.brands = settingsVal.brands.map((b) => (b === "Okra by Masgouf" ? "Okra" : b));
        await saveKey(SETTINGS_KEY, settingsVal);
        for (const ck of COLLECTION_KEYS) {
          let changed = false;
          const arr = (next[ck] || []).map((r) => {
            if (r.brand === "Okra by Masgouf") { changed = true; return { ...r, brand: "Okra", notes: typeof r.notes === "string" ? r.notes.split("Okra by Masgouf").join("Okra") : r.notes }; }
            return r;
          });
          if (changed) { next[ck] = arr; await saveKey(COLLECTIONS[ck].key, arr); }
        }
        if (Array.isArray(next.roadmap)) {
          const rm = next.roadmap.map((r) => (r.brand === "Okra by Masgouf" ? { ...r, brand: "Okra" } : r));
          next.roadmap = rm; await saveKey(ROADMAP_KEY, rm);
        }
        await saveKey(MIG_OKRA, true);
      }
      setDark(themeDark);
      setSettings(settingsVal);
      setData(next);
      setLoaded(true);
    })();
  }, []);

  const persist = (ck, arr) => {
    const key = ck === "roadmap" ? ROADMAP_KEY : COLLECTIONS[ck].key;
    saveKey(key, arr);
    setData((p) => ({ ...p, [ck]: arr }));
  };
  const saveIdeaMap = (next) => { saveKey(IDEAMAP_KEY, next); setData((p) => ({ ...p, ideamap: next })); };
  const saveTrends = (next) => { saveKey(TRENDS_KEY, next); setData((p) => ({ ...p, trends: next })); };
  const saveInspo = (next) => { saveKey(INSPO_KEY, next); setData((p) => ({ ...p, inspo: next })); };
  const normalize = (ck, form, existing) => {
    const config = COLLECTIONS[ck];
    const d = config && config.primaryDate ? derive(form[config.primaryDate]) : {};
    return { ...form, ...d, id: existing?.id || uid(), createdDate: existing?.createdDate || todayKey(), updatedDate: todayKey(), deleted: existing?.deleted || false };
  };
  const api = {
    add: (ck, form) => persist(ck, [normalize(ck, form), ...(data[ck] || [])]),
    bulkAdd: (ck, forms) => persist(ck, [...forms.map((f) => normalize(ck, f)), ...(data[ck] || [])]),
    update: (ck, id, form) => persist(ck, (data[ck] || []).map((r) => (r.id === id ? normalize(ck, form, r) : r))),
    softDelete: (ck, id) => persist(ck, (data[ck] || []).map((r) => (r.id === id ? { ...r, deleted: true, updatedDate: todayKey() } : r))),
    restore: (ck, id) => persist(ck, (data[ck] || []).map((r) => (r.id === id ? { ...r, deleted: false, updatedDate: todayKey() } : r))),
    hardDelete: (ck, id) => persist(ck, (data[ck] || []).filter((r) => r.id !== id)),
  };
  const saveSettings = (s) => { setSettings(s); saveKey(SETTINGS_KEY, s); };

  const backup = () => {
    const blob = new Blob([JSON.stringify({ version: 1, exportedAt: new Date().toISOString(), settings, data }, null, 2)], { type: "application/json" });
    download(`seazen-hub-backup-${todayKey()}.json`, blob);
    toast("Backup downloaded");
  };
  const restore = (e) => {
    const f = e.target.files?.[0]; e.target.value = "";
    if (!f) return;
    confirm({
      title: "Restore from backup?", message: "This replaces all current data with the backup file.", confirmText: "Restore", danger: true,
      onConfirm: () => {
        const reader = new FileReader();
        reader.onload = (ev) => {
          try {
            const parsed = JSON.parse(ev.target.result);
            if (parsed.settings) saveSettings(parsed.settings);
            const d = parsed.data || {};
            Object.keys(d).forEach((ck) => persist(ck, d[ck]));
            toast("Restored from backup");
          } catch { toast("Could not read that file"); }
        };
        reader.readAsText(f);
      },
    });
  };
  const exportEverything = () => {
    const wb = XLSX.utils.book_new();
    COLLECTION_KEYS.forEach((ck) => {
      const config = COLLECTIONS[ck];
      const rows = (data[ck] || []).map((r) => recordToRow(r, config));
      if (rows.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), config.label.slice(0, 28));
    });
    if (!wb.SheetNames.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([{ Info: "No data yet" }]), "Info");
    download(`seazen-hub-all-${todayKey()}.xlsx`, new Blob([XLSX.write(wb, { bookType: "xlsx", type: "array" })], { type: "application/octet-stream" }));
    toast("Exported workbook");
  };

  if (!loaded) return (
    <div className="min-h-screen flex items-center justify-center bg-gray-50 text-gray-400">
      <div className="text-center"><Database className="mx-auto mb-2 animate-pulse" /><p className="text-sm">Loading your hub...</p></div>
    </div>
  );

  const navItem = (n) => (
    <button key={n.id} onClick={() => { setSection(n.id); setSidebarOpen(false); }}
      className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm font-medium transition-colors ${section === n.id ? "bg-emerald-50 text-emerald-700" : "text-gray-600 hover:bg-gray-100"}`}>
      <n.icon size={17} className={section === n.id ? "text-emerald-600" : "text-gray-400"} />{n.label}
    </button>
  );

  const Sidebar = (
    <div className="flex flex-col h-full">
      <div className="px-4 py-4 border-b border-gray-100">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-emerald-600 text-white flex items-center justify-center font-semibold text-sm">S</div>
          <div><div className="font-semibold text-gray-900 text-sm leading-tight">Seazen Work Hub</div><div className="text-[11px] text-gray-400">Suleman Alhabashi</div></div>
        </div>
      </div>
      <nav className="flex-1 overflow-y-auto p-2.5 space-y-0.5">{NAV.map(navItem)}</nav>
      <div className="p-2.5 border-t border-gray-100 space-y-1">
        <button onClick={exportEverything} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-gray-600 hover:bg-gray-100"><Download size={16} className="text-gray-400" />Export all (Excel)</button>
        <button onClick={backup} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-gray-600 hover:bg-gray-100"><Database size={16} className="text-gray-400" />Backup data</button>
        <button onClick={() => backupRef.current?.click()} className="w-full flex items-center gap-2 px-3 py-2 rounded-lg text-sm text-gray-600 hover:bg-gray-100"><Upload size={16} className="text-gray-400" />Restore backup</button>
        <input ref={backupRef} type="file" accept=".json" className="hidden" onChange={restore} />
      </div>
    </div>
  );

  return (
    <div className={`min-h-screen bg-gray-50 text-gray-900 ${dark ? "theme-dark" : ""}`} style={{ fontFamily: "ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, sans-serif" }}>
      <style>{DARK_CSS}</style>
      {/* mobile top bar */}
      <div className="md:hidden sticky top-0 z-30 flex items-center justify-between bg-white border-b border-gray-200 px-3 py-2.5">
        <button onClick={() => setSidebarOpen(true)} className="p-1.5 text-gray-600"><Menu size={22} /></button>
        <span className="font-semibold text-gray-900">Seazen Work Hub</span>
        <button onClick={toggleDark} className="p-1.5 text-gray-600">{dark ? <Sun size={20} /> : <Moon size={20} />}</button>
      </div>

      {/* mobile drawer */}
      {sidebarOpen && (
        <div className="md:hidden fixed inset-0 z-40 flex">
          <div className="w-72 bg-white h-full shadow-xl">{Sidebar}</div>
          <div className="flex-1 bg-black/30" onClick={() => setSidebarOpen(false)} />
        </div>
      )}

      <div className="flex">
        <aside className="hidden md:block w-60 shrink-0 h-screen sticky top-0 bg-white border-r border-gray-200">{Sidebar}</aside>
        <main className="flex-1 min-w-0 p-4 sm:p-6 max-w-[1400px]">
          <div className="flex items-center justify-between mb-4">
            <div className="text-sm">
              <span className="font-medium text-gray-700">{now.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric" })}</span>
              <span className="ml-2 text-gray-400">{now.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" })}</span>
            </div>
            <button onClick={toggleDark} title="Toggle dark mode" className="hidden md:inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-sm font-medium bg-white border border-gray-200 text-gray-600 hover:bg-gray-50">
              {dark ? <Sun size={15} /> : <Moon size={15} />}{dark ? "Light" : "Dark"}
            </button>
          </div>
          <ErrorBoundary>
            {section === "home" && <Home data={data} settings={settings} go={setSection} saveTrends={saveTrends} saveInspo={saveInspo} />}
            {section === "events" && <CalendarPage data={data} settings={settings} api={api} toast={toast} />}
            {section === "ideamap" && <IdeaMapPage map={data.ideamap || { nodes: [], edges: [] }} ideas={data.ideas || []} settings={settings} save={saveIdeaMap} toast={toast} confirm={confirm} />}
            {section === "roadmap" && <RoadmapPage data={data} settings={settings} api={api} toast={toast} />}
            {section === "backlog" && <BacklogPage data={data} settings={settings} api={api} toast={toast} confirm={confirm} />}
            {section === "settings" && <SettingsPage settings={settings} saveSettings={saveSettings} toast={toast} />}
            {section === "glossary" && <GlossaryPage />}
            {COLLECTIONS[section] && section !== "events" && (
              <CollectionPage ckey={section} data={data} settings={settings} api={api} toast={toast} confirm={confirm} />
            )}
          </ErrorBoundary>
        </main>
      </div>

      {toastMsg && (
        <div className="fixed bottom-5 left-1/2 -translate-x-1/2 z-[70] bg-gray-900 text-white text-sm px-4 py-2 rounded-lg shadow-lg flex items-center gap-2">
          <CheckCircle2 size={16} className="text-emerald-400" />{toastMsg}
        </div>
      )}
      <ChatBot data={data} settings={settings} api={api} go={setSection} saveSettings={saveSettings} toast={toast} />
      <Confirm state={confirmState} onCancel={() => setConfirmState(null)} />
    </div>
  );
}
