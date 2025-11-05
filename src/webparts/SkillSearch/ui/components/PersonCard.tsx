import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Person, Skill } from "../../services/models";
import { highlightToNodes, prioritiseSkills } from "../../utils/search";
import { sortSkillsByLevel } from "../../utils/skills";
import { GernrateCv } from "./ProfileActions";
import { SPHttpClient } from "@microsoft/sp-http";
import { buildFolderViewUrlAsync } from "../../services/profileRepo";

type Props = {
  person: Person;
  tokens: string[];
  onOpenSkills: (name: string, skills: Skill[]) => void;
  outlookUrl: (p: Person) => string;
  teamsUrl: (p: Person) => string;
  profilesUrl: string;
  spHttpClient: SPHttpClient;
  absWebUrl: string;
  serverRelWebUrl: string;
};

const INLINE_LIMIT = 3;

export const PersonCard: React.FC<Props> = ({
  person, tokens, onOpenSkills, outlookUrl, teamsUrl, profilesUrl,
  spHttpClient, absWebUrl, serverRelWebUrl
}) => {
  const ranked = React.useMemo(() => sortSkillsByLevel(person.skills || []), [person.skills]);
  const skills = React.useMemo(() => prioritiseSkills(ranked, tokens), [ranked, tokens]);

  const visible = skills.slice(0, INLINE_LIMIT);

  const skillsRef = React.useRef<HTMLDivElement>(null);
  const [isClamped, setIsClamped] = React.useState(false);

  // Resolve Profilordner URL for the *target person*
  const [profileFolderUrl, setProfileFolderUrl] = React.useState<string | null>(null);
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const url = await buildFolderViewUrlAsync(spHttpClient, absWebUrl, serverRelWebUrl, person.displayName);
        if (!cancelled) setProfileFolderUrl(url);
      } catch {
        if (!cancelled) setProfileFolderUrl(null);
      }
    })();
    return () => { cancelled = true; };
  }, [spHttpClient, absWebUrl, serverRelWebUrl, person.displayName]);

  // Read privileged flag set by HeroMeCard (fallback to storage)
  const isPrivileged = React.useMemo(() => {
    try {
      if ((window as any).__skillsearch_isPrivileged === true) return true;
      return localStorage.getItem("skillsearch.isPrivileged") === "1";
    } catch { return false; }
  }, []);

  const showFolderBtn = isPrivileged; // privileged users see Profilordner on all person cards
  
  //const showFolderBtn = true;

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
        {/* Actions: open folder + generate CV (download) */}
        {showFolderBtn && (
          <GernrateCv
            spHttpClient={spHttpClient}
            absWebUrl={absWebUrl}
            serverRelWebUrl={serverRelWebUrl}
            displayName={person.displayName}
          />
        )}

        <a className={styles.linkBtn} href={outlookUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Outlook_(2018%E2%80%93present).svg.png?csf=1&web=1&e=AVZl0q" alt="" className={styles.logo} />
          Termin
        </a>
        <a className={styles.linkBtn} href={teamsUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Teams_(2018%E2%80%93present).svg.png?csf=1&web=1&e=bABdsE" alt="" className={styles.logo} />
          Chat
        </a>

        {showFolderBtn && (
          <a
            className={styles.linkBtn}
            role="button"
            href={profileFolderUrl || '#'}
            target="_blank"
            rel="noopener noreferrer"
            aria-label={`Profilordner von ${/* person or me */ person?.displayName} öffnen`}
            onClick={(e) => {
              if (!profileFolderUrl) {
                e.preventDefault();
                alert('Profilordner nicht gefunden oder Zugriff fehlt.');
              }
            }}
          >
            <img
              src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW"
              alt=""
              className={styles.logo}
            />
            Profilordner
          </a>
        )}

      </div>

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
          onClick={() => onOpenSkills(person.displayName, skills)}
        >
          Alle ({skills.length}) Skills anzeigen
        </button>
      )}

    </li>
  );
};
