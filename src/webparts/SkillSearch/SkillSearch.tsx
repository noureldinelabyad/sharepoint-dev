import * as React from "react";
import styles from "./SkillSearch.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Skill } from "./services/models";
import { usePeople } from "./ui";
import { HeroMeCard, PersonCard, SearchBar, SkillsModal } from "./ui";
import { tokenize, matches } from "./utils/search";
import {
  FilterState,
  emptyFilterState,
  applyFilters,
  collectDepartments
} from "./utils/filters";
import { FilterMenu } from "./ui/components/FilterMenu";

export interface SkillSearchProps { context: WebPartContext; }

export default function SkillSearch({ context }: SkillSearchProps) {
  const {
    me, people, allPeople, next, loading, loadingMore,
    bulkLoading, fullyLoaded, error, loadFirst, loadMore
  } = usePeople(context.msGraphClientFactory);

  const [query, setQuery] = React.useState("");
  const [filters, setFilters] = React.useState<FilterState>(emptyFilterState());
  const [skillsModal, setSkillsModal] = React.useState<{ name: string; skills: Skill[] } | null>(null);

  React.useEffect(() => { loadFirst(); }, [loadFirst]);

  const tokens = React.useMemo(() => tokenize(query), [query]);

  // choose source (paged vs full index) then text-search, then filters
  const base = query ? allPeople : people;
  const matchedText = React.useMemo(
    () => base.filter(p => matches(p, tokens)),
    [base, tokens]
  );
  const filtered = React.useMemo(
    () => applyFilters(matchedText, filters),
    [matchedText, filters]
  );

  const deptsForUI = React.useMemo(
    () => collectDepartments(allPeople),
    [allPeople]
  );

  const onScroll = React.useCallback((e: React.UIEvent<HTMLDivElement>) => {
    if (query) return;
    if (!next || loadingMore) return;
    const el = e.currentTarget;
    if (el.scrollTop + el.clientHeight >= el.scrollHeight - 200) loadMore();
  }, [query, next, loadingMore, loadMore]);

  if (loading) return <div>Loading profile…</div>;
  if (error)   return <div style={{ color: "#a80000" }}>Error: {error}</div>;
  if (!me)     return <div>No profile data.</div>;

  const summary =
    query
      ? `${filtered.length} Ergebnis(se) für „${query}“${!fullyLoaded ? "" : ""}`
      : `${people.length}${!fullyLoaded ? " +" : ""} Personen geladen${bulkLoading && !fullyLoaded ? " …" : ""}`;

  const absWebUrl = context.pageContext.web.absoluteUrl;
  const serverRelWebUrl = context.pageContext.web.serverRelativeUrl;

  return (
    <>
      <HeroMeCard
        me={me}
        onOpenSkills={(name, skills) => setSkillsModal({ name, skills })}
        spHttpClient={context.spHttpClient}
        absWebUrl={absWebUrl}
        serverRelWebUrl={serverRelWebUrl} 
        msGraphClientFactory={context.msGraphClientFactory}      
        />

      <SearchBar
        query={query}
        onChange={setQuery}
        summary={summary}
        rightSlot={
          <FilterMenu
            availableDepts={deptsForUI}
            state={filters}
            onChange={setFilters}
          />
        }
      />

      <div className={styles.peopleScroll} onScroll={onScroll}>
        <ul className={styles["templateCards"]}>
          {filtered.length === 0 && query && (
            <li className={styles.noResults}>Keine Treffer. Versuche andere Begriffe.</li>
          )}

          {filtered.map(p => (
            <PersonCard
              key={p.id}
              person={p}
              tokens={tokens}
              onOpenSkills={(name, skills) => setSkillsModal({ name, skills })}
              outlookUrl={(person) =>
                `https://outlook.office.com/calendar/action/compose?rru=addevent&path=/calendar/action/compose&to=${
                  encodeURIComponent(person.mail || person.userPrincipalName)
                }`
              }
              teamsUrl={(person) =>
                `https://teams.microsoft.com/l/chat/0/0?users=${
                  encodeURIComponent(person.mail || person.userPrincipalName)
                }`
              }
              profilesUrl={`${context.pageContext.web.serverRelativeUrl}/_layouts/15/me.aspx`}
                spHttpClient={context.spHttpClient}
                absWebUrl={context.pageContext.web.absoluteUrl}
                serverRelWebUrl={context.pageContext.web.serverRelativeUrl}
                msGraphClientFactory={context.msGraphClientFactory}
            />
          ))}

          {loadingMore && !query && <li>Loading more…</li>}
        </ul>
      </div>

      {skillsModal && (
        <SkillsModal
          name={skillsModal.name}
          skills={skillsModal.skills}
          onClose={() => setSkillsModal(null)}
        />
      )}
    </>
  );
}
