import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import filterIcon from '../../assets/filter-solid-full.svg';
import {
  FilterState,
  LEVEL_OPTIONS,
  SkillLevel,
  normDeptKey
} from "../../utils/filters";

type Props = {
  availableDepts: string[];
  state: FilterState;
  onChange: (next: FilterState) => void;
};

export const FilterMenu: React.FC<Props> = ({ availableDepts, state, onChange }) => {
  const [open, setOpen] = React.useState(false);
  const [tab, setTab] = React.useState<"dept" | "level">("dept");
  const ref = React.useRef<HTMLDivElement>(null);

  // close when user clicks outside
  React.useEffect(() => {
    const onDoc = (e: MouseEvent) => { if (open && ref.current && !ref.current.contains(e.target as Node)) setOpen(false); };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  const toggleDept = (name: string) => {
    const key = normDeptKey(name);
    const next = new Set(state.depts);
    next.has(key) ? next.delete(key) : next.add(key);
    onChange({ ...state, depts: next });
  };

  const toggleLevel = (label: SkillLevel) => {
    const next = new Set(state.levels);
    next.has(label) ? next.delete(label) : next.add(label);
    onChange({ ...state, levels: next });
  };

  const clearAll = () => onChange({ depts: new Set(), levels: new Set() });

  const activeCount = state.depts.size + state.levels.size;

  return (
    <div ref={ref} className={styles.filterWrapper}>
      <button
        type="button"
        className={styles.filterButton}
        aria-expanded={open}
        onClick={() => setOpen(o => !o)}
        title="Filter"
      >
        {activeCount ? ` (${activeCount})` : ""}
        <img src={filterIcon} alt="" className={styles.logo} />
      </button>

      {open && (
        <div className={styles.filterDropdown} role="dialog" aria-label="Filter">
          <div className={styles.filterTabs}>
            <button
              className={`${styles.filterTab} ${tab === "dept" ? styles.active : ""}`}
              onClick={() => setTab("dept")}
              type="button"
            >Abteilung</button>
            <button
              className={`${styles.filterTab} ${tab === "level" ? styles.active : ""}`}
              onClick={() => setTab("level")}
              type="button"
            >Skill-Level</button>
            <div className={styles.spacer} />
            <button className={styles.clearLink} onClick={clearAll} type="button">Zurücksetzen</button>
          </div>

          {tab === "dept" && (
            <div className={styles.checkList} role="group" aria-label="Abteilungen">
              {availableDepts.map(d => {
                const key = normDeptKey(d);
                const checked = state.depts.has(key);
                return (
                  <label key={key} className={styles.checkRow}>
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => toggleDept(d)}
                    />
                    <span>{d}</span>
                  </label>
                );
              })}
              {!availableDepts.length && <div className={styles.dim}>Keine Abteilungen gefunden</div>}
            </div>
          )}

          {tab === "level" && (
            <div className={styles.checkList} role="group" aria-label="Skill-Level">
              {LEVEL_OPTIONS.map(lvl => (
                <label key={lvl} className={styles.checkRow}>
                  <input
                    type="checkbox"
                    checked={state.levels.has(lvl)}
                    onChange={() => toggleLevel(lvl)}
                  />
                  <span>{lvl}</span>
                </label>
              ))}
            </div>
          )}
        </div>
      )}
    </div>
  );
};
