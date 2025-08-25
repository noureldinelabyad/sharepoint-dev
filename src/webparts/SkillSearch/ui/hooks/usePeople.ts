import * as React from "react";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import { GraphFacade, Me, Person, PeopleResult } from "../../services";

export function usePeople(msGraphClientFactory: any) {
  const [me, setMe] = React.useState<Me | null>(null);
  const [people, setPeople] = React.useState<Person[]>([]);
  const [next, setNext] = React.useState<string | undefined>();
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const getClient = React.useCallback(async (): Promise<GraphFacade> => {
    const client: MSGraphClientV3 = await msGraphClientFactory.getClient("3");
    return new GraphFacade(client);
  }, [msGraphClientFactory]);

  const loadFirst = React.useCallback(async () => {
    try {
      const svc = await getClient();
      const meData = await svc.getMe();
      setMe(meData);

      try {
        const { items, nextLink } = await svc.getPeoplePage(24);
        setPeople(items); setNext(nextLink);
      } catch {
        const { items } = await svc.getPeopleFallback(24);
        setPeople(items); setNext(undefined);
      }
    } catch (e: any) {
      setError(e?.message ?? "Unknown error");
    } finally {
      setLoading(false);
    }
  }, [getClient]);

  const loadMore = React.useCallback(async () => {
    if (!next || loadingMore) return;
    setLoadingMore(true);
    try {
      const svc = await getClient();
      const { items, nextLink }: PeopleResult = await svc.getPeoplePage(24, next);
      setPeople(prev => [...prev, ...items]); setNext(nextLink);
    } finally {
      setLoadingMore(false);
    }
  }, [getClient, next, loadingMore]);

  return { me, people, next, loading, loadingMore, error, loadFirst, loadMore };
}
