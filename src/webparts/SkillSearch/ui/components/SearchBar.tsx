import * as React from "react";
import styles from "../../SkillSearch.module.scss";

type Props = {
  query: string;
  onChange: (v: string) => void;
  summary: string;
  rightSlot?: React.ReactNode;
};
export const SearchBar: React.FC<Props> = ({ query, onChange, summary, rightSlot }) => (
  <div className={styles.searchBar} aria-label="Personensuche">
    <div className={styles.searchRow}>
      <span className={styles.searchIcon} aria-hidden>🔎</span>
      <input
        className={styles.searchInput}
        type="text"
        value={query}
        onChange={e => onChange(e.target.value)}
        placeholder="Suche nach Name, Skill, Jobtitel, Team/Abteilung, E-Mail …"
        aria-label="Sofortsuche"
      />
      {query && (
        <button className={styles.searchClear} onClick={() => onChange("")} aria-label="Suche löschen">✕</button>
      )}
      {rightSlot /* <— renders our FilterMenu button */ }
    </div>
    {summary && <div className={styles.resultsInfo}>{summary}</div>}
  </div>
);
