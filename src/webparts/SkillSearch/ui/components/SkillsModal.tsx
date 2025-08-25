import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Skill } from "../../services/models";
import { effectiveProficiency, sortSkillsByLevel } from "../../utils/skills";

type Props = {
  name: string;
  skills: Skill[];
  onClose: () => void;
};
export const SkillsModal: React.FC<Props> = ({ name, skills, onClose }) => {
  React.useEffect(() => {
    const onKey = (e: KeyboardEvent) => e.key === "Escape" && onClose();
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onClose]);

  const sorted = sortSkillsByLevel(skills);

  return (
    <div className={styles.modalBackdrop} onClick={onClose}>
      <div className={styles.modalCard} onClick={e => e.stopPropagation()}>
        <div className={styles.modalHeader}>
          Skills — {name}
          <button className={styles.modalClose} onClick={onClose} aria-label="Close">×</button>
        </div>
        <div className={styles.modalBody}>
          <div className={`${styles["card--skills"]} ${styles["card--skills--full"]}`}>
            {sorted.map((s, i) => (
              <span key={i} className={styles.skill}>
                {s.displayName}{effectiveProficiency(s) ? ` • ${effectiveProficiency(s)}` : ""}
              </span>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};
