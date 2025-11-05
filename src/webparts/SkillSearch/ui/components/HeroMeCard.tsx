import * as React from "react";
import styles from "../../SkillSearch.module.scss";
import { Me, Skill } from "../../services/models";
import { effectiveProficiency, sortSkillsByLevel } from "../../utils/skills";
import { GernrateCv } from "./ProfileActions";
import { SPHttpClient } from "@microsoft/sp-http";
import { buildFolderViewUrlAsync } from "../../services/profileRepo";

type Props = {
  me: Me;
  onOpenSkills: (name: string, skills: Skill[]) => void;
  spHttpClient: SPHttpClient;
  absWebUrl: string;
  serverRelWebUrl: string;
};

export const HeroMeCard: React.FC<Props> = ({ me, onOpenSkills, spHttpClient, absWebUrl, serverRelWebUrl }) => {
  const skills = sortSkillsByLevel(me.skills || []);
  const visible = skills.slice(0, 5);

  const profilesUrl =
    "https://thinformatics.sharepoint.com/sites/Beraterprofile/Freigegebene%20Dokumente/Forms/AllItems.aspx?as=json";

  const displayName = me.displayName;

  // Resolve Profilordner URL for *me*
  const [profileFolderUrl, setProfileFolderUrl] = React.useState<string | null>(null);
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const url = await buildFolderViewUrlAsync(spHttpClient, absWebUrl, serverRelWebUrl, me.displayName);
        if (!cancelled) setProfileFolderUrl(url);
      } catch {
        if (!cancelled) setProfileFolderUrl(null);
      }
    })();
    return () => { cancelled = true; };
  }, [spHttpClient, absWebUrl, serverRelWebUrl, me.displayName]);

  // 1) Is the current viewer privileged?
  const isPrivileged = React.useMemo(() => {
    const jt = (me.jobTitle || "").toLowerCase();
    return ["head of", "hr", "sales", "ceo", "trainee"].some(k => jt.includes(k));
  }, [me.jobTitle]);

  // 2) Publish flag so PersonCard can read it
  React.useEffect(() => {
    try {
      localStorage.setItem("skillsearch.isPrivileged", isPrivileged ? "1" : "0");
      (window as any).__skillsearch_isPrivileged = isPrivileged;
    } catch {}
  }, [isPrivileged]);


  return (
    <ul className={styles["templateCards"]} style={{ background: "#fff", gridTemplateColumns: "1fr" }}>
      <li className={styles.card}>
        <div className={styles["cardImage"]}>
          <img
            src={me.photoUrl ?? "https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/images/persona/size72.png"}
            alt={me.displayName}
          />
        </div>

        <div className={styles["cardName"]}>{me.displayName}</div>

        <div className={styles["cardMeta"]}>
          {me.jobTitle ?? ""}{me.jobTitle && me.department ? " • " : ""}{me.department ?? ""}
        </div>

        <div style={{ width: '100%', margin: '2px', gap: '8px', display: 'flex', justifyContent: 'center' }}>
          {/* Actions: open folder + generate CV (download) */}
          <GernrateCv
            spHttpClient={spHttpClient}
            absWebUrl={absWebUrl}
            serverRelWebUrl={serverRelWebUrl}
            displayName={displayName}
          />

          {/* Always visible for the current user */}
          <a
            className={styles.linkBtn}
            href={profileFolderUrl || '#'}
            target="_blank"
            rel="noopener noreferrer"
            aria-label={`Profilordner von ${me.displayName} öffnen`}
            onClick={(e) => {
              if (!profileFolderUrl) { e.preventDefault(); alert('Profilordner nicht gefunden oder Zugriff fehlt.'); }
            }}
          >
            <img
              src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW"
              alt=""
              className={styles.logo}
            />
            Mein Profilordner
          </a>

          {isPrivileged && (
            <a className={styles.linkBtn} href={profilesUrl} target="_blank" rel="noopener noreferrer">
              <img
                src="https://thinformatics.sharepoint.com/:i:/r/sites/thinformationHub/SiteAssets/SitePages/Skill-Search/32px-Microsoft_Office_SharePoint_(2019%E2%80%93present).svg.png?csf=1&web=1&e=etkaPW"
                alt=""
                className={styles.logo}
              />
              Alle Profile anzeigen
            </a>
          )}

        </div>

        {me.aboutMe && <div style={{ marginBottom: 8, color: "#333" }}>{me.aboutMe}</div>}

        {me.responsibilities?.length ? (
          <div style={{ marginBottom: 8 }}>
            <strong>Ask me about:</strong> {me.responsibilities.slice(0, 6).join(", ")}
          </div>
        ) : null}

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
