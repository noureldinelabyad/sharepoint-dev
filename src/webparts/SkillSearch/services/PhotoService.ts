// Photo fetching + initials avatar generation

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Person } from './models';
import { makeInitialsAvatar } from './utils';

/**
 * Small adapter responsible for fetching user photos and
 * enriching Person objects with a data URL.
 */
export class PhotosService {
  constructor(private client: MSGraphClientV3) {}

  /** Enrich people with photoUrl (Graph photo or initials avatar). */
  async enrich(users: Person[], size = 96): Promise<void> {
    await Promise.all(users.map(async (u) => {
      const key = encodeURIComponent(u.userPrincipalName || u.id);
      u.photoUrl =
        (await this.getPhotoDataUrl(`/users/${key}/photo/$value`)) ||
        (await this.getPhotoDataUrl(`/users/${key}/photos/${size}x${size}/$value`)) ||
        makeInitialsAvatar(u.displayName, size);
    }));
  }

  /** Get a photo as data URL, or undefined if missing/denied. */
  private async getPhotoDataUrl(path: string): Promise<string | undefined> {
    try {
      const blob: Blob = await this.client.api(path).responseType(ResponseType.BLOB).get();
      const dataUrl = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve(reader.result as string);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
      return dataUrl;
    } catch {
      return undefined;
    }
  }
}
