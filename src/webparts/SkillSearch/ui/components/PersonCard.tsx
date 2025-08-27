import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Person, Skill } from "../../services/models";
import { highlightToNodes, prioritiseSkills } from "../../utils/search";
import { sortSkillsByLevel } from "../../utils/skills";

type Props = {
  person: Person;
  tokens: string[];
  onOpenSkills: (name: string, skills: Skill[]) => void;
  outlookUrl: (p: Person) => string;
  teamsUrl: (p: Person) => string;
  profilesUrl: string;
};

const INLINE_LIMIT = 12; // cheap cap to keep DOM light

export const PersonCard: React.FC<Props> = ({
  person, tokens, onOpenSkills, outlookUrl, teamsUrl, profilesUrl
}) => {
  // 1) rank by level  2) bring query hits to the front (stable)
  const ranked = React.useMemo(() => sortSkillsByLevel(person.skills || []), [person.skills]);
  const skills = React.useMemo(() => prioritiseSkills(ranked, tokens), [ranked, tokens]);

  const visible = skills.slice(0, INLINE_LIMIT);

  // detect visual clamping overflow (so we can show "Alle Skills anzeigen")
  const skillsRef = React.useRef<HTMLDivElement>(null);
  const [isClamped, setIsClamped] = React.useState(false);

  React.useLayoutEffect(() => {
    const el = skillsRef.current;
    if (!el) return;
    const check = () => {
      const vertical = el.scrollHeight > el.clientHeight + 1;
      const horizontal = el.scrollWidth > el.clientWidth + 1;
      setIsClamped(vertical || horizontal);
    };
    check();
    const ro = new ResizeObserver(check);
    ro.observe(el);
    const t = requestAnimationFrame(check);
    return () => { ro.disconnect(); cancelAnimationFrame(t); };
  }, [person.id, visible.length]);

  const showAllButton = isClamped || skills.length > INLINE_LIMIT;

  return (
    <li className={styles.card}>
      <div className={styles.cardImage}>
        <img
          src={person.photoUrl ?? "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png"}
          alt={person.displayName}
        />
      </div>

      <div className={styles.cardName}>
        {highlightToNodes(person.displayName, tokens)}
      </div>

      <div className={styles.cardMeta}>
        {highlightToNodes(person.jobTitle ?? "", tokens)}
        {person.jobTitle && person.department ? " • " : ""}
        {highlightToNodes(person.department ?? "", tokens)}
      </div>

      <div className={styles.cardEmail}>
        <a href={`mailto:${person.mail || person.userPrincipalName}`}>
          {highlightToNodes(person.mail || person.userPrincipalName, tokens)}
        </a>
      </div>

      <div className={styles.cardLinks}>
        <a className={styles.linkBtn} href={outlookUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Outlook_(2018%E2%80%93present).svg.png?csf=1&web=1&e=AVZl0q" alt="" className={styles.logo} />
          Termin
        </a>
        <a className={styles.linkBtn} href={teamsUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Teams_(2018%E2%80%93present).svg.png?csf=1&web=1&e=bABdsE" alt="" className={styles.logo} />
          Chat
        </a>
        <a className={styles.linkBtn} href={profilesUrl} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW" alt="" className={styles.logo} />
          Profil anzeigen
        </a>
      </div>

      {/* skills list (chips) */}
      <div ref={skillsRef} className={styles.cardSkills}>
        {visible.map((s, i) => (
          <span key={i} className={styles.skill} title={s.displayName}>
            {highlightToNodes(s.displayName, tokens)}
            {s.proficiency ? <>{" · "}{highlightToNodes(s.proficiency, tokens)}</> : null}
          </span>
        ))}
      </div>

      {showAllButton && (
        <button
          className={styles.showAllBtn}
          onClick={() => onOpenSkills(person.displayName, skills /* pass prioritised list to modal */)}
        >
          Alle ({skills.length}) Skills anzeigen
        </button>
      )}
    </li>
  );
};
