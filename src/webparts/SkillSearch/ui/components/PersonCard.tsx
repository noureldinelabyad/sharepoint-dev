// src/webparts/SkillSearch/ui/components/PersonCard.tsx
import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Person, Skill } from "../../services/models";
import { highlightToNodes, prioritiseSkills } from "../../utils/search";
import { sortSkillsByLevel } from "../../utils/skills";
import { GenerateCv } from "./ProfileActions";
import { SPHttpClient } from "@microsoft/sp-http";
import { buildFolderViewUrlAsync } from "../../services/profileRepo";
import { makeInitialsAvatar } from "../../services/utils";
import { getPhotosService } from "../../services/PhotoService";
import { useVisibility } from "../hooks/useVisibility";

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
  msGraphClientFactory: any;     // <-- add this so we can fetch photos
};

const INLINE_LIMIT = 3;

export const PersonCard: React.FC<Props> = ({
  person, tokens, onOpenSkills, outlookUrl, teamsUrl, profilesUrl,
  spHttpClient, absWebUrl, serverRelWebUrl, msGraphClientFactory
}) => {
  const ranked = React.useMemo(() => sortSkillsByLevel(person.skills || []), [person.skills]);
  const skills = React.useMemo(() => prioritiseSkills(ranked, tokens), [ranked, tokens]);
  const visible = skills.slice(0, INLINE_LIMIT);

  // create initials once
  const initialsUrl = React.useMemo(() => makeInitialsAvatar(person.displayName, 72), [person.displayName]);


  // ---- FAST FIRST PAINT ----
  const { ref, visible: onScreen } = useVisibility<HTMLLIElement>('300px');
  const [photoUrl, setPhotoUrl] = React.useState<string | null | undefined>(undefined); // undefined = not asked yet

  React.useEffect(() => {
    let cancelled = false;
    if (!onScreen || photoUrl !== undefined) return;

    (async () => {
      const svc = await getPhotosService(msGraphClientFactory, { preferSize: 72, concurrency: 4 });
      const url = await svc.getUrl({ id: person.id, userPrincipalName: person.userPrincipalName });
      if (!cancelled) setPhotoUrl(url ?? null);
    })();

    return () => { cancelled = true; };
  }, [onScreen, photoUrl, msGraphClientFactory, person.id, person.userPrincipalName]);

  // ---- PROFILORDNER: resolve on click, not on mount ----
  const [folderHref, setFolderHref] = React.useState<string | null>(null);
  const [resolvingFolder, setResolvingFolder] = React.useState(false);

  async function handleOpenFolder(e: React.MouseEvent) {
    if (folderHref) return; // already resolved; normal nav
    e.preventDefault();
    if (resolvingFolder) return;

    setResolvingFolder(true);
    try {
      const url = await buildFolderViewUrlAsync(spHttpClient, absWebUrl, serverRelWebUrl, person.displayName);
      if (url) {
        setFolderHref(url);
        // navigate now
        window.open(url, "_blank", "noopener,noreferrer");
      } else {
        alert('Profilordner nicht gefunden oder Zugriff fehlt.');
      }
    } finally {
      setResolvingFolder(false);
    }
  }

  // Read privileged flag set by HeroMeCard (fallback to storage)
  const isPrivileged = React.useMemo(() => {
    try {
      if ((window as any).__skillsearch_isPrivileged === true) return true;
      return localStorage.getItem("skillsearch.isPrivileged") === "1";
    } catch { return false; }
  }, []);
  const showFolderBtn = isPrivileged;

  //const defaultAvatar = "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png";

  return (
    <li ref={ref} className={styles.card}>
      <div className={styles.cardImage}>
        <img
          src={photoUrl ?? initialsUrl} 
          alt={person.displayName}
          // Optional: hide broken image icon if photoUrl === null
          onError={(e) => { if (photoUrl) setPhotoUrl(null); }}
         //onError={() => setPhotoUrl(null)}      // if blob breaks, fall back to initials
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
        
        {showFolderBtn && (
          <a
            className={styles.linkBtn}
            role="button"
            href={folderHref || '#'}
            target="_blank"
            rel="noopener noreferrer"
            aria-label={`Profilordner von ${person.displayName} öffnen`}
            onClick={handleOpenFolder}
          >
            <img
              src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW"
              alt=""
              className={styles.logo}
            />
            {resolvingFolder ? 'Suchen…' : 'Profilordner'}
          </a>
        )}

        <a className={styles.linkBtn} href={outlookUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Outlook_(2018%E2%80%93present).svg.png?csf=1&web=1&e=AVZl0q" alt="" className={styles.logo} />
          Termin
        </a>
        <a className={styles.linkBtn} href={teamsUrl(person)} target="_blank" rel="noopener noreferrer">
          <img src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_Teams_(2018%E2%80%93present).svg.png?csf=1&web=1&e=bABdsE" alt="" className={styles.logo} />
          Chat
        </a>

        {/* Generate CV button (unchanged) */}
        {showFolderBtn && (
          <GenerateCv
            spHttpClient={spHttpClient}
            absWebUrl={absWebUrl}
            serverRelWebUrl={serverRelWebUrl}
            displayName={person.displayName}
          />
        )}
      </div>

      <div className={styles.cardSkills}>
        {visible.map((s, i) => (
          <span key={i} className={styles.skill} title={s.displayName}>
            {highlightToNodes(s.displayName, tokens)}
            {s.proficiency ? <>{" · "}{highlightToNodes(s.proficiency, tokens)}</> : null}
          </span>
        ))}
      </div>

      {(skills.length > INLINE_LIMIT) && (
        <button className={styles.showAllBtn} onClick={() => onOpenSkills(person.displayName, skills)}>
          Alle ({skills.length}) Skills anzeigen
        </button>
      )}
    </li>
  );
};
