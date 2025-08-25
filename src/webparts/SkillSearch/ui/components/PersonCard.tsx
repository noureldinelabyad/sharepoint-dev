import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Person, Skill } from "../../services/models";
import { highlightToNodes } from "../../utils/search";
import { effectiveProficiency, sortSkillsByLevel } from "../../utils/skills";

type Props = {
  person: Person;
  tokens: string[];
  onOpenSkills: (name: string, skills: Skill[]) => void;
  outlookUrl: (p: Person) => string;
  teamsUrl: (p: Person) => string;
  profilesUrl: string;
};
export const PersonCard: React.FC<Props> = ({ person, tokens, onOpenSkills, outlookUrl, teamsUrl, profilesUrl }) => {
  const skills = sortSkillsByLevel(person.skills || []);
  const visible = skills.slice(0, 12);

  return (
    <li className={styles.card}>
      <div className={styles["card--image"]}>
        <img
          src={person.photoUrl ?? "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png"}
          alt={person.displayName}
        />
      </div>

      <div className={styles["card--name"]}>
        {highlightToNodes(person.displayName, tokens)}
      </div>

      <div className={styles["card--meta"]}>
        {person.jobTitle ?? ""}{person.jobTitle && person.department ? " • " : ""}{person.department ?? ""}
      </div>

      <div className={styles["card--email"]}>
        <a href={`mailto:${person.mail || person.userPrincipalName}`}>
          {highlightToNodes(person.mail || person.userPrincipalName, tokens)}
        </a>
      </div>

      <div className={styles["card--links"]}>
        <a className={styles.linkBtn} href={outlookUrl(person)} target="_blank" rel="noopener noreferrer" title={`Termin mit ${person.displayName}`}>
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Outlook_(2018%E2%80%93present).svg.png?csf=1&web=1&e=AVZl0q" alt="" className={styles.logo} />
          Termin
        </a>
        <a className={styles.linkBtn} href={teamsUrl(person)} target="_blank" rel="noopener noreferrer" title={`Teams-Chat mit ${person.displayName}`}>
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Teams_(2018%E2%80%93present).svg.png?csf=1&web=1&e=bABdsE" alt="" className={styles.logo} />
          Chat
        </a>
        <a className={styles.linkBtn} href={profilesUrl} target="_blank" rel="noopener noreferrer" title="Berater-Profil">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW" alt="" className={styles.logo} />
          Profil anzeigen
        </a>
      </div>

      <div className={styles["card--skills"]}>
        {visible.map((s, i) => (
          <span key={i} className={styles.skill}>
            {s.displayName}{effectiveProficiency(s) ? ` • ${effectiveProficiency(s)}` : ""}
          </span>
        ))}
      </div>

      {skills.length > visible.length && (
        <button className={styles.showAllBtn} onClick={() => onOpenSkills(person.displayName, skills)}>
          Alle ({skills.length}) Skills anzeigen
        </button>
      )}
    </li>
  );
};
