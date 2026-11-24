/**
 * CNBU4 Production Value Dashboard - configuration
 *
 * NOTHING in this file is a secret. The dashboard is a browser SPA ("public client"):
 * tenantId / clientId / site names are identifiers, not credentials. There is NO client
 * secret anywhere in this project and none must ever be added.
 */
export const CONFIG = {
  title: 'CNBU4 PRODUCTION VALUE',
  logoUrl: '',                 // optional: e.g. './assets/logo.png'

  /** 'sample'     = read ./data/Rolling_Estimation_2611.xlsx from this web server (prototype / offline)
   *  'sharepoint' = read the Excel from SharePoint via Microsoft Graph (requires Entra ID login) */
  dataSource: 'sharepoint',

  sample: {
    url: './data/Rolling_Estimation_2611.xlsx',
  },

  /** Fill these in when IT provides them (see docs/DEPLOYMENT.md, section 2). */
  sharepoint: {
    // --- Microsoft Entra ID app registration (Single-page application) ---
    tenantId: '19f25823-17ff-421f-ad4e-8fed035aedda',
    clientId: '9b2520c6-ad32-41f9-b45e-2e828ac9994c',
    redirectUri: window.location.origin + window.location.pathname.replace(/[^/]*$/, ''),

    // --- Where the Excel lives. Use ONE of the two options. ---
    // Option A (recommended, human friendly): host + site path + library + file path
    siteHostname: 'deltao365.sharepoint.com',
    sitePath: '/sites/CNBU4.MFG',
    libraryName: 'Documents',            // document library display name
    filePath: 'Rolling Estimation 2611.xlsx', // path INSIDE the library; replace this file every day
    // Option B (fastest at runtime): fixed IDs. If set they override Option A.
    siteId: '',
    driveId: '',

    graphBaseUrl: 'https://graph.microsoft.com/v1.0',
    // Delegated, read-only. Access is still limited by the signed-in user's own SharePoint permission.
    scopes: ['User.Read', 'Files.Read.All', 'Sites.Read.All'],
  },

  /** 0 = manual refresh only. >0 = poll file metadata every N minutes and reload only when the file changed. */
  autoRefreshMinutes: 10,

  /**
   * KPI / chart rules. Anything marked ASSUMPTION could not be confirmed from the Excel file alone;
   * change it here (single place) once the Power BI definition is confirmed.
   */
  rules: {
    // Confirmed from Excel + reference image:
    //   Production Value = SUM(Daily[AMT])            (Daily.AMT already includes the 'Check' flag)
    //   Rolling(MUSD)    = SUM(Rolling[Rolling AMT])
    //   Ach(%)           = Production Value / Rolling
    //   GAP by MRP       = PV(Act) - Rolling AMT, per MRP
    achGoodThreshold: 1.0,     // Ach >= 100% shown green, otherwise red

    // ASSUMPTION (chosen by user): how 'Forecast Cumulative Rolling' is spread over days.
    //   'dps'    = each model's Rolling AMT is spread over days in proportion to its DPS plan quantity
    //   'linear' = Rolling total / number of days in month, every calendar day
    //   (the reference screenshot looks closer to 'linear' - see README "Open questions")
    rollingDistribution: 'dps',

    // ASSUMPTION: Rolling AMT is a monthly figure, so the KPI/GAP/comparison use the FULL Rolling AMT
    // regardless of the date-range filter (same as a Power BI table with no date relationship).
    rollingFollowsDateRange: false,

    // ASSUMPTION: DPS_With_PV = DPS plan value on days that already have PV(Act) (<= last actual date)
    //             DPS_Without_PV = DPS plan value on days after the last actual date.
    // ASSUMPTION: Cumulative PV(DPS) = running total of the DPS plan value;
    //             Forecast Cumulative PV = cumulative PV(Act) up to last actual date, then + DPS plan afterwards.
  },
};
