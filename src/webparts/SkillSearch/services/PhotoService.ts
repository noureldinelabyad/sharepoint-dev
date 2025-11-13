// src/webparts/SkillSearch/services/PhotoService.ts
import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import type { Person } from './models';
import { colorFromString, getInitials } from './utils';


type Options = {
  noPhotoTtlMs?: number;
  concurrency?: number;        // keep concurrency here; size now per call
  defaultSize?: number;        // fallback if caller omits size (uses hero)
};

class PhotosService {
  private okCache = new Map<string, string | null>(); // upn/id -> dataURL|null
  private inflight = new Map<string, Promise<string | null>>();
  private queue: Promise<any>[] = [];
  private ttlMs: number;
  private MAX: number;
  private defaultSize: number;

  constructor(private client: MSGraphClientV3, opts: Options = {}) {
    this.ttlMs = opts.noPhotoTtlMs ?? 24 * 60 * 60 * 1000;
    this.MAX = Math.max(1, opts.concurrency ?? 4);
    this.defaultSize = opts.defaultSize ?? 96; // safe default
  }

  /** Get photo; size can vary (hero vs list) without duplicating constants. */
  async getUrl(
    user: Pick<Person, 'id' | 'userPrincipalName'>,
    size?: number
  ): Promise<string | null> {
    const keyRaw = user.userPrincipalName || user.id;
    const key = encodeURIComponent(keyRaw);
    const px = size ?? this.defaultSize;

    const memKey = `${key}:${px}`;
    if (this.okCache.has(memKey)) return this.okCache.get(memKey)!;
    if (this.inflight.has(memKey)) return this.inflight.get(memKey)!;
    if (this.isNoPhotoFresh(key, px)) {
       this.okCache.set(memKey, null);
       return null;
      }

    const job = this.limit(async () => {
      try {
        const sized = await this.getPhotoDataUrl(`/users/${key}/photos/${px}x${px}/$value`);
        const url = sized ?? await this.getPhotoDataUrl(`/users/${key}/photo/$value`);
        if (url) {
          this.okCache.set(memKey, url);
          return url;
        } else {
          this.markNoPhoto(key, px);
          this.okCache.set(memKey, null);
          return null;
        }
      } catch {
        this.okCache.set(memKey, null);
        return null;
      } finally {
        this.inflight.delete(memKey);
      }
    });

    this.inflight.set(memKey, job);
    return job;
  }

  prefetch(users: Array<Pick<Person, 'id' | 'userPrincipalName'>>, size?: number): void {
    for (const u of users) this.getUrl(u, size).catch(() => void 0);
  }

  private limit<T>(task: () => Promise<T>) {
    const p = (async () => {
      while (this.queue.length >= this.MAX) await Promise.race(this.queue);
      const run = task().finally(() => {
        const i = this.queue.indexOf(run);
        if (i >= 0) this.queue.splice(i, 1);
      });
      this.queue.push(run);
      return run;
    })();
    return p;
  }

  private async getPhotoDataUrl(path: string): Promise<string | undefined> {
    try {
      const blob: Blob = await this.client.api(path)
      .responseType(ResponseType.BLOB)
      .get();
      
      const r = new FileReader();
      return await new Promise<string>((res, rej) => {
        r.onloadend = () => res(r.result as string);
        r.onerror = rej;
        r.readAsDataURL(blob);
      });
    } catch { return undefined; }
  }

  private lsKey(key: string, px: number) { return `ss:noPhoto:${px}:${key}`; }

  private isNoPhotoFresh(key: string, px: number): boolean {
    try {
      const raw = localStorage.getItem(this.lsKey(key, px));
      if (!raw) return false;
      const until = parseInt(raw, 10);
      return Number.isFinite(until) && Date.now() < until;
    } catch { return false; }
  }

  private markNoPhoto(key: string, px: number) {
    try {
       localStorage.setItem(this.lsKey(key, px), String(Date.now() + this.ttlMs)); 
    } catch {
      // ignore
    }
  }
}

// ---- singleton factory (per page) ----
let _svc: PhotosService | null = null;
export async function getPhotosService(msGraphClientFactory: any, opts?: Options): Promise<PhotosService> {
  if (_svc) return _svc;
  const client: MSGraphClientV3 = await msGraphClientFactory.getClient('3');
  _svc = new PhotosService(client, opts);
  return _svc;
}

/**
 * Lightweight inline SVG avatar (two-letter initials) for instant placeholders.
 * Returns a data URL that can be assigned directly to an <img src>.
 */
export function makeInlineInitialsPlaceholder(name: string, size = 96): string {
  const safeName = name || " ";
  const initials = getInitials(safeName);
  const bg = colorFromString(safeName);
  const fontSize = Math.round(size * 0.44);

  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" role="img" aria-label="${initials}">` +
      `<rect width="100%" height="100%" fill="${bg}"/>` +
      `<text x="50%" y="55%" text-anchor="middle" dominant-baseline="middle" fill="#fff" font-family="Segoe UI, system-ui, -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif" font-size="${fontSize}" font-weight="700">${initials}</text>` +
    `</svg>`;

  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}
