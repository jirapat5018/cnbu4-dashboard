/**
 * SharePoint connector - reads the Excel file through Microsoft Graph with the SIGNED-IN USER's token
 * (delegated permission), so SharePoint permissions decide who can see the data.
 * No secret, no public link. Path-based addressing => always the CURRENT version of the file,
 * so replacing the Excel in the same location is picked up on the next refresh.
 *
 * Connector contract (shared with sampleConnector):
 *   getMeta()  -> { version, lastModified, name }        cheap call, used for change detection
 *   fetchFile()-> { buffer:ArrayBuffer, meta }
 */
export class ConnectorError extends Error {
  constructor(message, { status, code, cause } = {}) { super(message); this.name = 'ConnectorError'; this.status = status; this.code = code; this.cause = cause; }
}

const enc = (p) => p.split('/').filter(Boolean).map(encodeURIComponent).join('/');

export class SharePointConnector {
  /**
   * @param cfg           CONFIG.sharepoint
   * @param getToken      async () => accessToken
   * @param fetchImpl     injectable for tests
   */
  constructor(cfg, getToken, fetchImpl = (...a) => fetch(...a)) {
    this.cfg = cfg; this.getToken = getToken; this.fetch = fetchImpl; this._driveBase = null;
  }

  async _graph(path, { raw = false } = {}) {
    const token = await this.getToken();
    let res;
    try {
      res = await this.fetch(this.cfg.graphBaseUrl + path, { headers: { Authorization: `Bearer ${token}` }, cache: 'no-store' });
    } catch (e) {
      throw new ConnectorError('เชื่อมต่อ Microsoft Graph ไม่ได้ (เครือข่ายถูกบล็อกหรือออฟไลน์)', { code: 'network', cause: e });
    }
    if (!res.ok) throw await this._httpError(res, path);
    return raw ? res : res.json();
  }

  async _httpError(res, path) {
    let detail = '';
    try { const j = await res.clone().json(); detail = j?.error?.message || ''; } catch { /* ignore */ }
    const map = {
      401: 'Login หมดอายุหรือไม่ถูกต้อง — กรุณา Sign in ใหม่',
      403: 'บัญชีนี้ไม่มีสิทธิ์เข้าถึงไฟล์ใน SharePoint (หรือแอปยังไม่ได้รับ Admin consent)',
      404: 'ไม่พบไฟล์/Site/Library ใน SharePoint — ตรวจสอบ sitePath, libraryName, filePath ใน config.js',
      429: 'Microsoft Graph จำกัดความถี่การเรียก — ลองใหม่ในอีกสักครู่',
    };
    return new ConnectorError(`${map[res.status] || 'SharePoint/Graph error ' + res.status}${detail ? ` (${detail})` : ''}`, { status: res.status, code: 'http' });
  }

  /** Resolve /drives/{id} once (site -> library). */
  async _drive() {
    if (this._driveBase) return this._driveBase;
    const c = this.cfg;
    if (c.driveId) { this._driveBase = `/drives/${c.driveId}`; return this._driveBase; }
    let siteId = c.siteId;
    if (!siteId) {
      const site = await this._graph(`/sites/${c.siteHostname}:${c.sitePath}`);
      siteId = site.id;
    }
    const drives = await this._graph(`/sites/${siteId}/drives?$select=id,name`);
    const d = (drives.value || []).find((x) => x.name === (c.libraryName || 'Documents')) || (drives.value || [])[0];
    if (!d) throw new ConnectorError('ไม่พบ Document Library ใน Site นี้', { status: 404, code: 'http' });
    this._driveBase = `/drives/${d.id}`;
    return this._driveBase;
  }

  async _item() {
    const base = await this._drive();
    return this._graph(`${base}/root:/${enc(this.cfg.filePath)}?$select=id,name,eTag,cTag,lastModifiedDateTime,size,@microsoft.graph.downloadUrl`);
  }

  _meta(item) {
    return { version: item.cTag || item.eTag, lastModified: item.lastModifiedDateTime, name: item.name, size: item.size };
  }

  async getMeta() { return this._meta(await this._item()); }

  async fetchFile() {
    const item = await this._item();                      // fresh metadata + fresh pre-authenticated URL every time
    let res;
    const url = item['@microsoft.graph.downloadUrl'];
    try {
      if (url) res = await this.fetch(url, { cache: 'no-store' });   // pre-authenticated URL: must NOT send the bearer token
      else res = await this._graph(`/drives/${item.parentReference?.driveId || ''}/items/${item.id}/content`, { raw: true });
    } catch (e) {
      if (e instanceof ConnectorError) throw e;
      throw new ConnectorError('ดาวน์โหลดไฟล์จาก SharePoint ไม่สำเร็จ', { code: 'network', cause: e });
    }
    if (!res.ok) throw await this._httpError(res, 'download');
    return { buffer: await res.arrayBuffer(), meta: this._meta(item) };
  }
}
