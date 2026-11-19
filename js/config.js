// Cấu hình kết nối — dùng chung Entra app với TA Performance Dashboard.
// Dữ liệu OKR nằm trong 3 list SharePoint ở site ORG HR Admin; mỗi item:
// Title = id tài liệu, OKRData = JSON của objective / KPI / check-in.
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
  // Trường JSON trong mỗi list.
  dataField: "OKRData",
  // Tự tải lại dữ liệu người khác vừa sửa (ms). Chỉ chạy khi tab đang mở.
  pollMs: 60000,
  // ?demo=1 trên localhost → chạy dữ liệu giả trong localStorage, không cần đăng nhập.
  isDemo() {
    const local = ["localhost", "127.0.0.1"].includes(location.hostname);
    return local && new URLSearchParams(location.search).has("demo");
  },
};
