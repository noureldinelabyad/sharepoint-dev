// Photo fetching + initials avatar generation with positive/negative caching

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Person } from './models';
import { makeInitialsAvatar } from './utils';

type Options = {
  /** Re-check users with no photo after this many ms (default: 24h). */
  noPhotoTtlMs?: number;
  /** Preferred size for thumbnails (also used in the localStorage key). */
  preferSize?: number;
};

/**
 * Fetches user photos from Graph and enriches Person objects with a data URL.
 * - Tries /photos/{size}x{size}/$value first (smaller).
 * - Falls back to /photo/$value.
 * - Positive cache for found photos.
 * - Negative cache (with TTL) for users without a photo to avoid repeated 404s.
 */
export class PhotosService {
  private okCache = new Map<string, string>();          // upn/id -> data URL
  private ttlMs: number;
  private size: number;
  private lsPrefix: string;

  constructor(private client: MSGraphClientV3, opts: Options = {}) {
    this.ttlMs = opts.noPhotoTtlMs ?? 24 * 60 * 60 * 1000; // 24h
    this.size  = opts.preferSize ?? 96;
    this.lsPrefix = `ss:noPhoto:${this.size}:`;            // localStorage key prefix
  }

  /** Enrich people with photoUrl (Graph photo or initials). */
  async enrich(users: Person[]): Promise<void> {
    await Promise.all(users.map(u => this.attach(u)));
  }

  private async attach(u: Person): Promise<void> {
    const keyRaw = u.userPrincipalName || u.id;
    const key = encodeURIComponent(keyRaw);

    // Positive cache (photo already fetched)
    const cached = this.okCache.get(key);
    if (cached) { u.photoUrl = cached; return; }

    // Negative cache (recently confirmed "no photo")
    if (this.isNoPhotoFresh(key)) {
      u.photoUrl = makeInitialsAvatar(u.displayName, this.size);
      return;
    }

    // Try sized thumbnail first, then full
    const url =
      (await this.getPhotoDataUrl(`/users/${key}/photos/${this.size}x${this.size}/$value`)) ||
      (await this.getPhotoDataUrl(`/users/${key}/photo/$value`));

    if (url) {
      this.okCache.set(key, url);
      u.photoUrl = url;
    } else {
      // Remember "no photo" with TTL and use initials
      this.markNoPhoto(key);
      u.photoUrl = makeInitialsAvatar(u.displayName, this.size);
    }
  }

  /** Graph GET -> data URL (or undefined if 404/denied). */
  private async getPhotoDataUrl(path: string): Promise<string | undefined> {
    try {
      const blob: Blob = await this.client.api(path)
        .responseType(ResponseType.BLOB)
        .get();

      return await new Promise<string>((resolve, reject) => {
        const r = new FileReader();
        r.onloadend = () => resolve(r.result as string);
        r.onerror   = reject;
        r.readAsDataURL(blob);
      });
    } catch {
      // 404/403/etc → treat as "no photo"
      return undefined;
    }
  }

  /** Negative-cache helpers (localStorage) */
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
    } catch {
      /* ignore storage quota errors */
    }
  }
}