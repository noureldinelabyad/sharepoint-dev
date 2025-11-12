// src/webparts/SkillSearch/services/PhotoService.ts
// Lazy, concurrency-limited photo fetcher with positive/negative caching.

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import type { Person } from './models';

type Options = {
  /** Re-check users with no photo after this many ms (default: 24h). */
  noPhotoTtlMs?: number;
  /** Preferred size for thumbnails (also used in the localStorage key). */
  preferSize?: number;
  /** Max parallel photo fetches (default: 4) */
  concurrency?: number;
};

class PhotosService {
  private okCache = new Map<string, string | null>();     // upn/id -> data URL|null (positive+negative)
  private inflight = new Map<string, Promise<string | null>>();
  private queue: Promise<any>[] = [];
  private ttlMs: number;
  private size: number;
  private lsPrefix: string;
  private MAX: number;

  constructor(private client: MSGraphClientV3, opts: Options = {}) {
    this.ttlMs = opts.noPhotoTtlMs ?? 24 * 60 * 60 * 1000; // 24h
    this.size  = opts.preferSize ?? 72;
    this.lsPrefix = `ss:noPhoto:${this.size}:`;
    this.MAX = Math.max(1, opts.concurrency ?? 4);
  }

  /** Main entry: get a photo URL lazily. Never throws; returns null on failure. */
  async getUrl(user: Pick<Person, 'id' | 'userPrincipalName'>): Promise<string | null> {
    const keyRaw = user.userPrincipalName || user.id;
    const key = encodeURIComponent(keyRaw);

    // hit memory cache first
    if (this.okCache.has(key)) return this.okCache.get(key)!;

    // in-flight de-dup
    if (this.inflight.has(key)) return this.inflight.get(key)!;

    // negative cache (localStorage TTL)
    if (this.isNoPhotoFresh(key)) {
      this.okCache.set(key, null);
      return null;
    }

    const job = this.limit(async () => {
      try {
        const sized = await this.getPhotoDataUrl(`/users/${key}/photos/${this.size}x${this.size}/$value`);
        const url = sized ?? await this.getPhotoDataUrl(`/users/${key}/photo/$value`);
        if (url) {
          this.okCache.set(key, url);
          return url;
        } else {
          this.markNoPhoto(key);
          this.okCache.set(key, null);
          return null;
        }
      } catch {
        // treat as no-photo; do not rethrow
        this.okCache.set(key, null);
        return null;
      } finally {
        this.inflight.delete(key);
      }
    });

    this.inflight.set(key, job);
    return job;
  }

  /** (Optional) prefetch a few visible users without blocking UI */
  prefetch(users: Array<Pick<Person, 'id' | 'userPrincipalName'>>): void {
    for (const u of users) this.getUrl(u).catch(() => void 0);
  }

  // ---- internals ----
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
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onloadend = () => resolve(r.result as string);
        r.onerror = reject;
        r.readAsDataURL(blob);
      });
      return dataUrl;
    } catch {
      return undefined;
    }
  }

  private isNoPhotoFresh(key: string): boolean {
    try {
      const raw = localStorage.getItem(this.lsPrefix + key);
      if (!raw) return false;
      const until = parseInt(raw, 10);
      return Number.isFinite(until) && Date.now() < until;
    } catch {
      return false;
    }
  }

  private markNoPhoto(key: string): void {
    try {
      const until = Date.now() + this.ttlMs;
      localStorage.setItem(this.lsPrefix + key, String(until));
    } catch { /* ignore */ }
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
