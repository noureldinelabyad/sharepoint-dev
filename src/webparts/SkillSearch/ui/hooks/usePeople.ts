import * as React from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { GraphFacade, Me, Person, PeopleResult } from "../../services";

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

      // Me
      const meData = await svc.getMe();
      setMe(meData);

      // First page (fast)
      try {
        const { items, nextLink } = await svc.getPeoplePage(100);
        setPeople(items);
        setAllPeople(items);
        setNext(nextLink);

       // If there's no next page, we are already fully loaded.
        setFullyLoaded(!nextLink);
      } catch {
        // Fallback when /users is not consented
        const { items } = await svc.getPeopleFallback(50); 
        setPeople(items);
        setAllPeople(items);
        setNext(undefined);
        
        //  /me/people has no paging – we’re done.
        setFullyLoaded(true);
      }
    } catch (e: any) {
      setError(e?.message ?? "Unknown error");
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
      const { items, nextLink }: PeopleResult = await svc.getPeoplePage(100, next);
      setPeople(prev => [...prev, ...items]);
      setAllPeople(prev => mergeUniqueById(prev, items));
      setNext(nextLink);
      if (!nextLink) setFullyLoaded(true);
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
        setAllPeople(prev => mergeUniqueById(prev, items));
        cursor = nextLink;
        setNext(cursor);
      }
      setFullyLoaded(true);
    } catch (e) {
      // Not fatal—users still have first page + fallback paths
      console.warn("Background prefetch failed:", e);
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
