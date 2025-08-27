import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Me, Skill } from "../../services/models";
import { effectiveProficiency, sortSkillsByLevel } from "../../utils/skills";

type Props = {
  me: Me;
  onOpenSkills: (name: string, skills: Skill[]) => void;
};
export const HeroMeCard: React.FC<Props> = ({ me, onOpenSkills }) => {
  const skills = sortSkillsByLevel(me.skills || []);
  const visible = skills.slice(0, 5);
  return (
    <ul className={styles["templateCards"]} style={{ background: "#fff", gridTemplateColumns: "1fr" }}>
      <li className={styles.card}>
        <div className={styles["cardImage"]}>
          <img src={me.photoUrl ?? "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png"} alt={me.displayName} />
        </div>
        <div className={styles["cardName"]}>{me.displayName}</div>
        <div className={styles["cardMeta"]}>
          {me.jobTitle ?? ""}{me.jobTitle && me.department ? " • " : ""}{me.department ?? ""}
        </div>
        {me.aboutMe && <div style={{ marginBottom: 8, color: "#333" }}>{me.aboutMe}</div>}
        {me.responsibilities?.length ? (
          <div style={{ marginBottom: 8 }}><strong>Ask me about:</strong> {me.responsibilities.slice(0, 6).join(", ")}</div>
        ) : null}
        {/* email */}
        <div className={styles["cardEmail"]}>
          <a href={`mailto:${me.mail || me.userPrincipalName}`}>{me.mail || me.userPrincipalName}</a>
        </div>
        <div className={styles["cardSkills"]}>
          {visible.map((s, i) => (
            <span key={i} className={styles.skill}>
              {s.displayName}{effectiveProficiency(s) ? ` • ${effectiveProficiency(s)}` : ""}
            </span>
          ))}
        </div>
        {skills.length > visible.length && (
          <button className={styles.showAllBtn} onClick={() => onOpenSkills(me.displayName, skills)}>
            Alle ({skills.length}) Skills anzeigen
          </button>
        )}
      </li>
    </ul>
  );
};
