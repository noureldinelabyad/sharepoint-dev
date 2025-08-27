import * as React from "react";
import styles from "./SkillSearch.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Me, Person, Skill } from "./services/models";
import { usePeople } from "./ui";
import { HeroMeCard, PersonCard, SearchBar, SkillsModal } from "./ui";
import { tokenize } from "./utils/search";

export interface SkillSearchProps { context: WebPartContext; }

const CONSULTANT_PROFILES_URL =
  "https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json";

const emailFor = (p: Person | Me) =>
  encodeURIComponent(p.mail || p.userPrincipalName);
const outlookNewMeeting = (p: Person | Me) =>
  `https://outlook.office.com/calendar/deeplink/compose?to=${emailFor(p)}&subject=${encodeURIComponent("Termin mit " + p.displayName)}`;
const teamsChat = (p: Person | Me) =>
  `https://teams.microsoft.com/l/chat/0/0?users=${emailFor(p)}`;

export default function SkillSearch({ context }: SkillSearchProps) : JSX.Element {
  const {
    me,
    people,            // paged list for browsing
    allPeople,         // full index (prefetched)
    next,
    loading,
    loadingMore,
    bulkLoading,
    fullyLoaded,
    error,
    loadFirst,
    loadMore
  } = usePeople(context.msGraphClientFactory);

  const [query, setQuery] = React.useState("");
  const [skillsModal, setSkillsModal] = React.useState<{ name: string; skills: Skill[] } | null>(null);

  React.useEffect(() => { loadFirst(); }, [loadFirst]);

  const tokens = React.useMemo(() => tokenize(query), [query]);

  // Use full index for search; fall back to paged list when query is empty
  const dataSet = query ? allPeople : people;

  const filtered = React.useMemo(
    () => dataSet.filter(p => {
      if (!tokens.length) return true;
      const name  = (p.displayName || "").toLowerCase();
      const mail  = (p.mail || p.userPrincipalName || "").toLowerCase();
      const job   = (p.jobTitle || "").toLowerCase();
      const dept  = (p.department || "").toLowerCase();
      const skill = (p.skills || []).map(s => s.displayName.toLowerCase()).join(" ");
      return tokens.every(x => name.includes(x) || mail.includes(x) || job.includes(x) || dept.includes(x) || skill.includes(x));
    }),
    [dataSet, tokens]
  );

  const onScroll = React.useCallback((e: React.UIEvent<HTMLDivElement>) : void => {
    if (query) return;            // pause infinite scroll while searching (we show from allPeople)
    if (!next || loadingMore) return;
    const el = e.currentTarget;
    if (el.scrollTop + el.clientHeight >= el.scrollHeight - 200) {
      void loadMore();
    }
  }, [query, next, loadingMore, loadMore]);

  if (loading) return <div>Loading profile…</div>;
  if (error) return <div style={{ color: "#a80000" }}>Error: {error}</div>;
  if (!me) return <div>No profile data.</div>;

  const mayHaveMore = !!next || (bulkLoading && !fullyLoaded);

  const summary =
    query
     ? `${filtered.length} Ergebnis(se) für „${query}“${mayHaveMore ? " – weitere Personen werden geladen…" : ""}`
     : `${people.length}${mayHaveMore ? " +" : ""} Personen geladen${bulkLoading ? " …" : ""}`;

  return (
    <>
      <HeroMeCard me={me} onOpenSkills={(name, skills) => setSkillsModal({ name, skills })} />

      <SearchBar
        query={query}
        onChange={setQuery}
        summary={summary}
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
              outlookUrl={outlookNewMeeting}
              teamsUrl={teamsChat}
              profilesUrl={CONSULTANT_PROFILES_URL}
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
