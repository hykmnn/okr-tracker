// Connection config. Shares the Entra app with TA Performance Dashboard.
// OKR data lives in 3 SharePoint lists on the ORG HR Admin site; each item:
// Title = document id, OKRData = JSON of the objective / KPI / check-in.
export const CONFIG = {
  clientId: "acd70ab1-98df-4f7f-97bd-f32f2149c59c",
  tenantId: "1fd983f3-dc44-42f8-ad66-7972e9d94659",
  siteHost: "eastgatesoftware.sharepoint.com",
  sitePath: "/sites/ORGHRAdmin",
  lists: {
    objectives: "OKR Objectives",
    kpis: "OKR KPIs",
    checkins: "OKR Checkins",
  },
  // JSON field in each list.
  dataField: "OKRData",
  // Re-fetch others' edits every N ms. Only while the tab is visible.
  pollMs: 60000,
  // ?demo=1 on localhost → run on fake data in localStorage, no sign-in needed.
  isDemo() {
    const local = ["localhost", "127.0.0.1"].includes(location.hostname);
    return local && new URLSearchParams(location.search).has("demo");
  },
};
