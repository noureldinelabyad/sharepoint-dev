import * as React from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { GraphFacade, Me, Person,  } from "../../services";

/**
 * Strategy:
 * - Load first page (fast initial render).
 * - Immediately prefetch the rest in background into `allPeople` (search index).
 * - Keep infinite scroll for casual browsing; search always searches `allPeople`.
 */

function mergeUniqueById(prev: Person[], incoming: Person[]): Person[] {
  if (!incoming.length) return prev;
  const map = new Map<string, Person>();
  for (const p of prev) map.set(p.id, p);
  for (const p of incoming) if (!map.has(p.id)) map.set(p.id, p);
  return Array.from(map.values());
}

const CHUNK = 12;
const PREFETCH_DELAY_MS = 120;

export function usePeople(msGraphClientFactory: any) {
  const [me, setMe] = React.useState<Me | null>(null);
  const [people, setPeople] = React.useState<Person[]>([]);           // visible list (paged)
  const [allPeople, setAllPeople] = React.useState<Person[]>([]);     // full search index
  const [next, setNext] = React.useState<string | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);
  const [bulkLoading, setBulkLoading] = React.useState(false);
  const [fullyLoaded, setFullyLoaded] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const getClient = React.useCallback(async (): Promise<GraphFacade> => {
    const client: MSGraphClientV3 = await msGraphClientFactory.getClient("3");
    return new GraphFacade(client);
  }, [msGraphClientFactory]);

  /** First paint: me + first page, then kick off background prefetch. */
  const loadFirst = React.useCallback(async () => {
    setLoading(true);
    try {
      const svc = await getClient();

      // me
      const meData = await svc.getMe();
      setMe(meData);

      // fast initial fetch: request CHUNK items so first paint is quick
      try {
        const { items, nextLink } = await svc.getPeoplePage(CHUNK);
        setPeople(items);       // visible: first CHUNK immediately
        setAllPeople(items);    // search index starts with same
        setNext(nextLink);
        setFullyLoaded(!nextLink);
      } catch {
        // fallback (smaller result set) — still show only CHUNK
        const { items } = await svc.getPeopleFallback(50);
        setPeople(items.slice(0, CHUNK));
        setAllPeople(items);
        setNext(undefined);
        setFullyLoaded(true);
      }

      // Stop global loading so UI becomes interactive; prefetch continues in background
    } catch (e: any) {
      setError(e?.message ?? 'Unknown error');
    } finally {
      setLoading(false);
    }
  }, [getClient]);

  /** Infinite scroll for casual browsing (also merges into allPeople). */
  const loadMore = React.useCallback(async () => {
    if (!next || loadingMore) return;
    setLoadingMore(true);
    try {
      const svc = await getClient();
      // request a larger page from the next cursor, append in CHUNK-sized slices
      const { items, nextLink } = await svc.getPeoplePage(100, next);
      setAllPeople(prev => mergeUniqueById(prev, items));
      // append only CHUNK now so visible list grows progressively
      setPeople(prev => mergeUniqueById(prev, items.slice(0, CHUNK)));
      setNext(nextLink);
      if (!nextLink) setFullyLoaded(true);
    } catch (e) {
      console.warn('loadMore failed', e);
    } finally {
      setLoadingMore(false);
    }
  }, [getClient, next, loadingMore]);

  /** Background prefetch: load ALL remaining pages into allPeople for instant search. */
  const prefetchAll = React.useCallback(async () => {
    if (!next || bulkLoading) return;
    setBulkLoading(true);
    try {
      const svc = await getClient();
      let cursor: string | undefined = next;
      while (cursor) {
        const { items, nextLink } = await svc.getPeoplePage(500, cursor);
        // merge to search index immediately
        setAllPeople(prev => mergeUniqueById(prev, items));

        // progressively reveal items in CHUNK batches with tiny pauses
        for (let i = 0; i < items.length; i += CHUNK) {
          const batch = items.slice(i, i + CHUNK);
          setPeople(prev => mergeUniqueById(prev, batch));
          // allow paint and keep UI responsive
          // eslint-disable-next-line no-await-in-loop
          await new Promise(res => setTimeout(res, PREFETCH_DELAY_MS));
        }

        cursor = nextLink;
        setNext(cursor);
      }
      setFullyLoaded(true);
    } catch (e) {
      console.warn('Background prefetch failed', e);
    } finally {
      setBulkLoading(false);
    }
  }, [getClient, next, bulkLoading]);

  /** Auto-start background prefetch shortly after the first page is ready. */
  React.useEffect(() => {
    if (!loading && next && !bulkLoading && !fullyLoaded) {
      const t = setTimeout(() => { prefetchAll(); }, 250);
      return () => clearTimeout(t);
    }
    return;
  }, [loading, next, bulkLoading, fullyLoaded, prefetchAll]);

  return {
    me,
    people,          // visible, paged list
    allPeople,       // full search index (grows in background)
    next,
    loading,
    loadingMore,
    bulkLoading,
    fullyLoaded,
    error,
    loadFirst,
    loadMore,
    prefetchAll
  };
}
